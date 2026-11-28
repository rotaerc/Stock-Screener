# -*- coding: utf-8 -*-
"""
미국 주식 거래대금 · 거래량 급증(기관 매집 시그널) · 섹터 자금 흐름 자동 수집기
===========================================================================
- Nasdaq 공식 전체 종목 스크리너 API (7,000+ 미국 상장주: 시총·종가·등락률·거래량·섹터·산업)
- Yahoo Finance v8 일봉 차트 API (12개 섹터/테마 ETF + 3대 지수 ETF + 거래대금/모멘텀 핵심 대형주 ~130종목)
  -> 20일 평균거래량 대비 배수(RVOL), 1일/5일(주간)/1개월 등락률, 기관 매집/이탈 시그널 자동 산출
- 별도 API 키나 고정 IP가 필요 없어 로컬 PC 및 GitHub Actions에서 매일 미국장 마감 후 자동 갱신 가능

사용법:
  python us_market.py               # 수집 -> us_market.json 저장 -> index.html 갱신 -> GitHub 자동 업로드
  python us_market.py --no-push     # 웹페이지 생성까지만 수행 (GitHub 업로드 생략)
"""

import json
import os
import re
import sys
import time
import urllib.parse
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from email.utils import parsedate_to_datetime

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
US_JSON_PATH = os.path.join(HERE, "us_market.json")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

SECTOR_KO_MAP = {
    "Technology": "기술/IT",
    "Health Care": "헬스케어",
    "Finance": "금융",
    "Consumer Discretionary": "임의소비재",
    "Industrials": "산업재",
    "Energy": "에너지",
    "Consumer Staples": "필수소비재",
    "Telecommunications": "통신/미디어",
    "Utilities": "유틸리티",
    "Basic Materials": "소재",
    "Real Estate": "부동산",
    "Miscellaneous": "기타",
}

BENCHMARK_ETFS = [
    ("SPY", "S&P 500", "대형주 지수"),
    ("QQQ", "나스닥 100", "기술·성장 지수"),
    ("IWM", "러셀 2000", "중소형주 지수"),
    ("SMH", "반도체 (SMH)", "미국 반도체 대표"),
]

SECTOR_ETFS = [
    ("XLK", "기술/IT", "Technology"),
    ("SMH", "반도체", "Technology"),
    ("XLC", "통신/미디어", "Telecommunications"),
    ("XLY", "임의소비재", "Consumer Discretionary"),
    ("XLF", "금융", "Finance"),
    ("XLV", "헬스케어", "Health Care"),
    ("XLI", "산업재", "Industrials"),
    ("XLE", "에너지", "Energy"),
    ("XLP", "필수소비재", "Consumer Staples"),
    ("XLU", "유틸리티", "Utilities"),
    ("XLB", "소재", "Basic Materials"),
    ("XLRE", "부동산", "Real Estate"),
]

# 미국 모닝 브리핑: 5대 주요 지수 (티커, 한글명, 보조설명, TradingView 심볼)
MAJOR_INDICES = [
    ("^GSPC", "S&P 500", "미국 대형주 500", "SP:SPX"),
    ("^IXIC", "나스닥 종합", "기술·성장주 지수", "NASDAQ:IXIC"),
    ("^DJI", "다우존스", "우량주 30종목", "DJ:DJI"),
    ("^RUT", "러셀 2000", "중소형주 지수", "TVC:RUT"),
    ("^SOX", "필라델피아 반도체", "한국장 핵심 선행", "NASDAQ:SOX"),
]

# 미국 모닝 브리핑: 핵심 매크로·시장 지표 8종 (티커, 한글명, 분류, 단위, TradingView 심볼, 변동표시타입)
MACRO_INDICATORS = [
    ("^TNX", "미 10년물 금리", "채권", "%", "TVC:US10Y", "bp"),
    ("^VIX", "VIX 공포지수", "변동성", "", "TVC:VIX", "pt"),
    ("DX-Y.NYB", "달러 인덱스", "환율", "", "TVC:DXY", "pct"),
    ("KRW=X", "원/달러 환율", "환율", "원", "FX_IDC:USDKRW", "krw"),
    ("EWY", "MSCI 한국 ETF", "한국선행", "$", "AMEX:EWY", "pct"),
    ("CL=F", "WTI 국제유가", "원자재", "$", "TVC:USOIL", "pct"),
    ("GC=F", "국제 금 (Gold)", "원자재", "$", "TVC:GOLD", "pct"),
    ("BTC-USD", "비트코인 (BTC)", "크립토", "$", "BINANCE:BTCUSDT", "pct"),
]

# 빅테크 M7 + 반도체 핵심주 11종 (항상 수집 보장)
BIG_TECH_CORE = [
    ("NVDA", "엔비디아", "반도체"),
    ("AAPL", "애플", "빅테크"),
    ("MSFT", "마이크로소프트", "빅테크"),
    ("GOOGL", "알파벳", "빅테크"),
    ("AMZN", "아마존", "빅테크"),
    ("META", "메타", "빅테크"),
    ("TSLA", "테슬라", "빅테크"),
    ("AVGO", "브로드컴", "반도체"),
    ("TSM", "TSMC", "반도체"),
    ("MU", "마이크론", "반도체"),
    ("AMD", "AMD", "반도체"),
]

FEAR_GREED_KO = {
    "extreme fear": "극단적 공포",
    "fear": "공포",
    "neutral": "중립",
    "greed": "탐욕",
    "extreme greed": "극단적 탐욕",
}


def _num(x):
    if x is None:
        return None
    s = str(x).replace("$", "").replace("%", "").replace(",", "").strip()
    if s in ("", "-", "N/A", "null", "None", "nan", "NaN"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _is_clean_us_equity(sym, name):
    if not sym or len(sym) > 5:
        return False
    if any(ch in sym for ch in ("^", "/", ".", "-", "+")):
        return False
    if not re.fullmatch(r"[A-Z]{1,5}", sym):
        return False
    up = str(name or "").upper()
    bad_words = ("WARRANT", "RIGHTS", "UNIT", "PREFERRED", "DEPOSITARY", "NOTES", "BOND", "TRUST", "FUND", "ACQUISITION CORP")
    if any(w in up for w in bad_words):
        return False
    return True


def fetch_nasdaq_universe():
    """Nasdaq 공식 스크리너 API에서 미국 전체 상장 주식 일괄 수집"""
    hdr = {
        "User-Agent": UA,
        "Accept": "application/json, text/plain, */*",
        "Origin": "https://www.nasdaq.com",
        "Referer": "https://www.nasdaq.com/",
    }
    r = requests.get(
        "https://api.nasdaq.com/api/screener/stocks?tableonly=true&download=true",
        headers=hdr,
        timeout=25,
    )
    r.raise_for_status()
    rows = (r.json().get("data") or {}).get("rows") or []
    parsed = []
    for row in rows:
        sym = str(row.get("symbol") or "").strip().upper()
        name = str(row.get("name") or "").strip()
        if not _is_clean_us_equity(sym, name):
            continue
        price = _num(row.get("lastsale"))
        pct = _num(row.get("pctchange"))
        vol = _num(row.get("volume"))
        mcap = _num(row.get("marketCap"))
        if not price or price <= 0 or not vol or vol <= 0 or not mcap or mcap <= 0:
            continue
        sector = str(row.get("sector") or "").strip() or "Miscellaneous"
        industry = str(row.get("industry") or "").strip()
        # 회사명 간결화
        short_name = re.sub(
            r"\s+(Common Stock|Ordinary Shares|Class [A-C].*|American Depositary.*|Inc\.?.*|Corp\.?.*|Corporation.*|Holdings?.*|Ltd\.?.*|plc.*)$",
            "",
            name,
            flags=re.IGNORECASE,
        ).strip() or name
        if len(short_name) > 26:
            short_name = short_name[:25].rstrip() + "…"
        dollar_vol = price * vol
        parsed.append({
            "symbol": sym,
            "name": short_name,
            "fullName": name,
            "sector": sector,
            "sectorKo": SECTOR_KO_MAP.get(sector, sector),
            "industry": industry,
            "price": round(price, 2),
            "chg1d": round(pct, 2) if pct is not None else None,
            "volume": int(round(vol)),
            "dollarVol": dollar_vol,
            "dollarVolB": round(dollar_vol / 1e9, 3),
            "mcap": mcap,
            "mcapB": round(mcap / 1e9, 2),
            "turnoverPct": round(dollar_vol / mcap * 100.0, 2) if mcap > 0 else None,
        })
    return parsed


def fetch_yahoo_1mo_stats(symbol):
    """Yahoo Finance v8 일봉(3개월)으로 RVOL(20일 평균 대비), 1일/5일/1개월 등락률, 최신 거래일 계산"""
    try:
        r = requests.get(
            f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}",
            params={"range": "3mo", "interval": "1d"},
            headers={"User-Agent": UA},
            timeout=12,
        )
        if r.status_code != 200:
            return symbol, None
        res_list = (r.json().get("chart") or {}).get("result") or []
        if not res_list:
            return symbol, None
        res = res_list[0]
        meta = res.get("meta") or {}
        ts_list = res.get("timestamp") or []
        q = ((res.get("indicators") or {}).get("quote") or [{}])[0]
        closes = q.get("close") or []
        vols = q.get("volume") or []

        meta_price = _num(meta.get("regularMarketPrice"))
        meta_vol = _num(meta.get("regularMarketVolume"))
        meta_ts = meta.get("regularMarketTime")
        meta_dt = (
            datetime.fromtimestamp(meta_ts, tz=timezone(timedelta(hours=-5))).strftime("%Y-%m-%d")
            if meta_ts else ""
        )

        bars = []
        n_ts = len(ts_list)
        for idx, (ts, c, v) in enumerate(zip(ts_list, closes, vols)):
            dt_us = datetime.fromtimestamp(ts, tz=timezone(timedelta(hours=-5))).strftime("%Y-%m-%d")
            # 조기폐장(블랙프라이데이 등) 또는 장마감 직후 마지막 일봉 close가 None으로 내려오는 경우 meta.regularMarketPrice 보완
            if (c is None or c <= 0) and idx == n_ts - 1 and meta_price and meta_price > 0:
                c = meta_price
                if (v is None or v <= 0) and meta_vol and meta_vol > 0:
                    v = meta_vol
            if c is None or c <= 0:
                continue
            bars.append((dt_us, float(c), float(v or 0)))

        # meta.regularMarketTime이 일봉 배열 마지막 날짜보다 최신이면 최신 종가 바로 추가
        if meta_dt and meta_price and meta_price > 0:
            if not bars or meta_dt > bars[-1][0]:
                bars.append((meta_dt, float(meta_price), float(meta_vol or 0)))
            elif meta_dt == bars[-1][0] and abs(bars[-1][1] - float(meta_price)) > 1e-6:
                bars[-1] = (meta_dt, float(meta_price), float(meta_vol or bars[-1][2]))

        if len(bars) < 2:
            return symbol, None

        last_dt, last_c, last_v = bars[-1]
        prev_c = bars[-2][1]
        chg1d = round((last_c - prev_c) / prev_c * 100.0, 2) if prev_c > 0 else None

        chg5d = None
        if len(bars) >= 6 and bars[-6][1] > 0:
            chg5d = round((last_c - bars[-6][1]) / bars[-6][1] * 100.0, 2)

        chg1m = None
        if len(bars) >= 22 and bars[-22][1] > 0:
            chg1m = round((last_c - bars[-22][1]) / bars[-22][1] * 100.0, 2)
        elif len(bars) >= 15 and bars[0][1] > 0:
            chg1m = round((last_c - bars[0][1]) / bars[0][1] * 100.0, 2)

        # 직전 20영업일 평균 거래량 (당일 제외)
        prev_vols = [b[2] for b in bars[:-1][-20:] if b[2] > 0]
        avg_vol20 = (sum(prev_vols) / len(prev_vols)) if prev_vols else None
        rvol = round(last_v / avg_vol20, 2) if (avg_vol20 and avg_vol20 > 0 and last_v > 0) else None
        dollar_vol_b = round((last_c * last_v) / 1e9, 3)

        return symbol, {
            "tradeDate": last_dt,
            "price": round(last_c, 3 if last_c < 20 else 2),
            "prevClose": round(prev_c, 3 if prev_c < 20 else 2),
            "absChg": round(last_c - prev_c, 3 if last_c < 20 else 2),
            "volume": int(round(last_v)),
            "avgVol20": int(round(avg_vol20)) if avg_vol20 else None,
            "rvol": rvol,
            "dollarVolB": dollar_vol_b,
            "chg1d": chg1d,
            "chg5d": chg5d,
            "chg1m": chg1m,
        }
    except Exception:
        return symbol, None


def fetch_cnn_fear_greed():
    """CNN 공포·탐욕 지수 (Fear & Greed Index) 실시간 조회"""
    try:
        r = requests.get(
            "https://production.dataviz.cnn.io/index/fearandgreed/graphdata",
            headers={
                "User-Agent": UA,
                "Accept": "application/json, text/plain, */*",
                "Referer": "https://edition.cnn.com/",
                "Origin": "https://edition.cnn.com",
            },
            timeout=10,
        )
        if r.status_code != 200:
            return None
        fg = (r.json() or {}).get("fear_and_greed") or {}
        score = _num(fg.get("score"))
        if score is None:
            return None
        rating_raw = str(fg.get("rating") or "").strip().lower()
        return {
            "score": round(score, 1),
            "rating": rating_raw,
            "ratingKo": FEAR_GREED_KO.get(rating_raw, rating_raw or "중립"),
            "prevClose": round(_num(fg.get("previous_close")), 1) if _num(fg.get("previous_close")) is not None else None,
            "prev1Week": round(_num(fg.get("previous_1_week")), 1) if _num(fg.get("previous_1_week")) is not None else None,
            "prev1Month": round(_num(fg.get("previous_1_month")), 1) if _num(fg.get("previous_1_month")) is not None else None,
            "prev1Year": round(_num(fg.get("previous_1_year")), 1) if _num(fg.get("previous_1_year")) is not None else None,
            "timestamp": str(fg.get("timestamp") or "")[:19],
        }
    except Exception:
        return None


def fetch_us_morning_news(limit=10):
    """Google News KR RSS에서 밤사이 뉴욕증시 마감 주요 한국어 헤드라인 수집"""
    kst = timezone(timedelta(hours=9))
    queries = [
        "뉴욕증시 마감 when:2d",
        "미국증시 마감 when:2d",
    ]
    items = []
    seen_keys = set()
    for q in queries:
        if len(items) >= limit:
            break
        try:
            url = (
                "https://news.google.com/rss/search?"
                f"q={urllib.parse.quote(q)}&hl=ko&gl=KR&ceid=KR:ko"
            )
            r = requests.get(url, headers={"User-Agent": UA}, timeout=12)
            if r.status_code != 200:
                continue
            root = ET.fromstring(r.text)
            for it in root.findall("./channel/item"):
                raw_title = (it.findtext("title") or "").strip()
                link = (it.findtext("link") or "").strip()
                pub_raw = (it.findtext("pubDate") or "").strip()
                source_el = it.find("source")
                press = (source_el.text or "").strip() if source_el is not None else ""
                title = raw_title
                if " - " in raw_title:
                    t_part, p_part = raw_title.rsplit(" - ", 1)
                    if t_part.strip():
                        title = t_part.strip()
                    if not press and p_part.strip():
                        press = p_part.strip()
                if not title or not link:
                    continue
                # 유사 제목 중복 제거 (공백·기호 제거 후 앞 18자 비교)
                norm = re.sub(r"[^0-9A-Za-z가-힣]", "", title)[:18]
                if not norm or norm in seen_keys:
                    continue
                seen_keys.add(norm)

                pub_ts = 0
                pub_kst = ""
                if pub_raw:
                    try:
                        dt = parsedate_to_datetime(pub_raw).astimezone(kst)
                        pub_ts = int(dt.timestamp())
                        pub_kst = dt.strftime("%m/%d %H:%M")
                    except Exception:
                        pass

                items.append({
                    "title": title,
                    "press": press or "뉴스",
                    "link": link,
                    "pubKst": pub_kst,
                    "pubTs": pub_ts,
                })
        except Exception:
            continue

    items.sort(key=lambda x: x["pubTs"], reverse=True)
    return [
        {"title": x["title"], "press": x["press"], "link": x["link"], "pubKst": x["pubKst"]}
        for x in items[:limit]
    ]


def build_morning_briefing(major_indices, macro_indicators, fear_greed, big_tech, sector_etfs):
    """수집된 미국 마감 지수·매크로·심리·섹터 데이터를 기반으로 한국어 아침 시황 브리핑 자동 생성"""
    idx_map = {x["symbol"]: x for x in major_indices}
    mac_map = {x["symbol"]: x for x in macro_indicators}

    def _pct(v):
        if v is None:
            return "-"
        return f"{v:+.2f}%"

    spx = idx_map.get("^GSPC") or {}
    ixic = idx_map.get("^IXIC") or {}
    dji = idx_map.get("^DJI") or {}
    sox = idx_map.get("^SOX") or {}
    tnx = mac_map.get("^TNX") or {}
    vix = mac_map.get("^VIX") or {}
    dxy = mac_map.get("DX-Y.NYB") or {}
    krw = mac_map.get("KRW=X") or {}
    ewy = mac_map.get("EWY") or {}
    wti = mac_map.get("CL=F") or {}

    spx_c = spx.get("chg1d")
    ixic_c = ixic.get("chg1d")
    sox_c = sox.get("chg1d")
    ewy_c = ewy.get("chg1d")
    krw_c = krw.get("chg1d")

    avg_main = None
    vals = [v for v in (spx_c, ixic_c) if v is not None]
    if vals:
        avg_main = sum(vals) / len(vals)

    if avg_main is None:
        regime, regime_ko = "NEU", "미국장 마감 요약"
    elif avg_main >= 1.0:
        regime, regime_ko = "BULL", "강세 마감"
    elif avg_main >= 0.25:
        regime, regime_ko = "MILD_UP", "소폭 상승 마감"
    elif avg_main <= -1.0:
        regime, regime_ko = "BEAR", "약세 마감"
    elif avg_main <= -0.25:
        regime, regime_ko = "MILD_DN", "소폭 하락 마감"
    else:
        regime, regime_ko = "MIXED", "혼조 마감"

    # 한국 증시 영향 스코어 (EWY + SOX + 원달러)
    kr_score = 0
    if ewy_c is not None:
        kr_score += 2 if ewy_c >= 0.8 else (1 if ewy_c >= 0.2 else (-2 if ewy_c <= -0.8 else (-1 if ewy_c <= -0.2 else 0)))
    if sox_c is not None:
        kr_score += 2 if sox_c >= 1.5 else (1 if sox_c >= 0.4 else (-2 if sox_c <= -1.5 else (-1 if sox_c <= -0.4 else 0)))
    if krw_c is not None:
        kr_score += 1 if krw_c <= -0.3 else (-1 if krw_c >= 0.4 else 0)

    if kr_score >= 2:
        kr_impact, kr_impact_ko = "POS", "한국장 우호적 · 반도체/외인 수급 기대"
    elif kr_score <= -2:
        kr_impact, kr_impact_ko = "NEG", "한국장 경계 · 변동성·약세 출발 유의"
    else:
        kr_impact, kr_impact_ko = "NEU", "한국장 중립 · 업종/종목 차별화 예상"

    bullets = []

    # 1) 미국 주요 지수 요약
    b1 = (
        f"뉴욕증시는 S&P 500 {_pct(spx_c)}, 나스닥 {_pct(ixic_c)}, 다우 {_pct(dji.get('chg1d'))}로 "
        f"{regime_ko}했습니다."
    )
    if sox_c is not None:
        sox_tone = "강하게 반등하며 기술주 투자심리를 견인했습니다." if sox_c >= 1.5 else (
            "상승 흐름을 보였습니다." if sox_c > 0.2 else (
                "약세를 보이며 반도체 섹터에 부담을 주었습니다." if sox_c <= -1.5 else "보합권 흐름을 나타냈습니다."
            )
        )
        b1 += f" 필라델피아 반도체(SOX)는 {_pct(sox_c)}로 {sox_tone}"
    bullets.append({"tag": "지수 마감", "text": b1})

    # 2) 한국 증시 개장 영향 체크
    kr_parts = []
    if ewy.get("price") is not None:
        kr_parts.append(f"MSCI 한국 ETF(EWY) {_pct(ewy_c)}")
    if krw.get("price") is not None:
        kr_parts.append(f"원/달러 환율 {krw['price']:,.1f}원({_pct(krw_c)})")
    if sox_c is not None:
        kr_parts.append(f"SOX {_pct(sox_c)}")
    if kr_parts:
        bullets.append({
            "tag": "한국장 영향",
            "text": f"{' · '.join(kr_parts)} 흐름으로, 금일 국내 증시는 '{kr_impact_ko}' 여건입니다.",
        })

    # 3) 금리 · VIX · 달러 · 유가 · 심리
    mac_parts = []
    if tnx.get("price") is not None:
        bp = round((tnx.get("absChg") or 0) * 100)
        mac_parts.append(f"미 10년물 금리 {tnx['price']:.2f}%({bp:+d}bp)")
    if vix.get("price") is not None:
        mac_parts.append(f"VIX {vix['price']:.2f}({_pct(vix.get('chg1d'))})")
    if dxy.get("price") is not None:
        mac_parts.append(f"달러인덱스 {dxy['price']:.2f}({_pct(dxy.get('chg1d'))})")
    if wti.get("price") is not None:
        mac_parts.append(f"WTI 유가 ${wti['price']:.2f}({_pct(wti.get('chg1d'))})")
    b3 = ", ".join(mac_parts) + "." if mac_parts else ""
    if fear_greed and fear_greed.get("score") is not None:
        b3 += f" CNN 공포·탐욕 지수는 {fear_greed['score']:.0f}점({fear_greed['ratingKo']}) 구간입니다."
    if b3.strip():
        bullets.append({"tag": "거시·심리", "text": b3.strip()})

    # 4) 주도 섹터 & 빅테크/반도체 핵심 종목
    valid_etfs = [e for e in sector_etfs if e.get("chg1d") is not None]
    valid_etfs.sort(key=lambda x: x["chg1d"], reverse=True)
    top_sec = ", ".join(f"{e['label']}({_pct(e['chg1d'])})" for e in valid_etfs[:3])
    bot_sec = ", ".join(f"{e['label']}({_pct(e['chg1d'])})" for e in valid_etfs[-2:]) if len(valid_etfs) >= 4 else ""

    valid_bt = [b for b in big_tech if b.get("chg1d") is not None]
    valid_bt.sort(key=lambda x: x["chg1d"], reverse=True)
    top_bt = ", ".join(f"{b['symbol']}({_pct(b['chg1d'])})" for b in valid_bt[:3])
    bot_bt = ", ".join(f"{b['symbol']}({_pct(b['chg1d'])})" for b in valid_bt[-2:] if b["chg1d"] < 0)

    b4_parts = []
    if top_sec:
        b4_parts.append(f"강세 섹터는 {top_sec}")
    if bot_sec:
        b4_parts.append(f"상대 약세는 {bot_sec}")
    if top_bt:
        b4_parts.append(f"빅테크·반도체 강세 주도주는 {top_bt}")
    if bot_bt:
        b4_parts.append(f"약세 종목은 {bot_bt}")
    if b4_parts:
        bullets.append({"tag": "섹터·주도주", "text": " / ".join(b4_parts) + "입니다."})

    return {
        "regime": regime,
        "regimeKo": regime_ko,
        "krImpact": kr_impact,
        "krImpactKo": kr_impact_ko,
        "bullets": bullets,
    }


def classify_us_signal(chg1d, rvol):
    if chg1d is None or rvol is None:
        return ""
    if (chg1d >= 1.5 and rvol >= 1.5) or (chg1d >= 3.0 and rvol >= 1.25):
        return "ACCUM"  # 강한 매집 시그널
    if (chg1d <= -1.5 and rvol >= 1.5) or (chg1d <= -3.0 and rvol >= 1.25):
        return "DIST"   # 대량 매도 시그널
    if rvol >= 1.8:
        return "SURGE"  # 거래량 급증
    return ""


def collect_us_market(here=HERE):
    print("=" * 80)
    print(" ▶ 미국 시장 거래대금 · 거래량 급증(RVOL) · 섹터 자금 흐름 수집 시작")
    print("=" * 80)

    print("[1/3] Nasdaq 공식 스크리너 전체 미국 주식 스캔 중...")
    universe = fetch_nasdaq_universe()
    print(f"  - 유효 미국 보통주 {len(universe):,}종목 수집 완료")

    # 2) Yahoo v8 정밀 보강 대상 선별 (시총 $5B↑ 거래대금 TOP 100 + 급등/급락 모멘텀 TOP 35)
    large_caps = [s for s in universe if s["mcapB"] >= 5.0]
    large_caps.sort(key=lambda x: x["dollarVol"], reverse=True)
    cand_map = {}
    for s in large_caps[:100]:
        cand_map[s["symbol"]] = s

    momo_pool = [s for s in universe if s["mcapB"] >= 3.0 and s["dollarVolB"] >= 0.15 and s["chg1d"] is not None]
    momo_pool.sort(key=lambda x: abs(x["chg1d"]), reverse=True)
    for s in momo_pool[:35]:
        cand_map.setdefault(s["symbol"], s)

    # 빅테크 M7 + 반도체 핵심 11종은 항상 cand_map 및 Yahoo 조회 대상에 포함
    uni_by_sym = {s["symbol"]: s for s in universe}
    for sym, _ko, _grp in BIG_TECH_CORE:
        if sym in uni_by_sym:
            cand_map.setdefault(sym, uni_by_sym[sym])

    etf_syms = sorted({x[0] for x in BENCHMARK_ETFS} | {x[0] for x in SECTOR_ETFS})
    morning_syms = sorted(
        {x[0] for x in MAJOR_INDICES}
        | {x[0] for x in MACRO_INDICATORS}
        | {x[0] for x in BIG_TECH_CORE}
    )
    all_fetch_syms = sorted(set(cand_map.keys()) | set(etf_syms) | set(morning_syms))

    print(f"[2/3] Yahoo Finance 일봉 + CNN 공포탐욕 + 모닝뉴스 병렬 조회 중 (총 {len(all_fetch_syms)}티커)...")
    yahoo_stats = {}
    fear_greed = None
    morning_news = []
    with ThreadPoolExecutor(max_workers=16) as ex:
        fut_fg = ex.submit(fetch_cnn_fear_greed)
        fut_news = ex.submit(fetch_us_morning_news, 10)
        futs = [ex.submit(fetch_yahoo_1mo_stats, sym) for sym in all_fetch_syms]
        for fut in as_completed(futs):
            sym, st = fut.result()
            if st:
                yahoo_stats[sym] = st
        try:
            fear_greed = fut_fg.result()
        except Exception:
            fear_greed = None
        try:
            morning_news = fut_news.result() or []
        except Exception:
            morning_news = []
    print(
        f"  - Yahoo 일봉 조회: {len(yahoo_stats)}/{len(all_fetch_syms)}개 · "
        f"CNN 공포탐욕: {fear_greed['score'] if fear_greed else 'N/A'} · "
        f"모닝뉴스: {len(morning_news)}건"
    )

    # Yahoo 최신 거래일 종가/거래대금/등락률을 universe에 우선 반영 (조기폐장/장마감 직후 Nasdaq 지연 보정)
    for s in universe:
        st = yahoo_stats.get(s["symbol"])
        if st:
            if st.get("price"):
                s["price"] = st["price"]
            if st.get("chg1d") is not None:
                s["chg1d"] = st["chg1d"]
            if st.get("dollarVolB") is not None:
                s["dollarVolB"] = st["dollarVolB"]
                s["dollarVol"] = st["dollarVolB"] * 1e9

    # 섹터별 시장 전체 자금 쏠림 집계 (시총 $2B 이상 중대형주 기준)
    sec_agg = {}
    for s in universe:
        if s["mcapB"] < 2.0:
            continue
        sk = s["sector"]
        if sk == "Miscellaneous":
            continue
        ent = sec_agg.setdefault(sk, {
            "sector": sk,
            "sectorKo": s["sectorKo"],
            "totalVolB": 0.0,
            "upVolB": 0.0,
            "downVolB": 0.0,
            "upCount": 0,
            "downCount": 0,
            "count": 0,
            "chgSum": 0.0,
            "leaders": [],
        })
        dv = s["dollarVolB"]
        ent["totalVolB"] += dv
        ent["count"] += 1
        c1 = s.get("chg1d")
        if c1 is not None:
            ent["chgSum"] += c1
            if c1 > 0:
                ent["upVolB"] += dv
                ent["upCount"] += 1
            elif c1 < 0:
                ent["downVolB"] += dv
                ent["downCount"] += 1
        ent["leaders"].append((dv, s["symbol"], s["name"], c1))

    sector_flows = []
    for sk, ent in sec_agg.items():
        ent["leaders"].sort(key=lambda x: x[0], reverse=True)
        net_b = round(ent["upVolB"] - ent["downVolB"], 2)
        sector_flows.append({
            "sector": sk,
            "sectorKo": ent["sectorKo"],
            "totalVolB": round(ent["totalVolB"], 2),
            "upVolB": round(ent["upVolB"], 2),
            "downVolB": round(ent["downVolB"], 2),
            "netFlowB": net_b,
            "upRatio": round(ent["upCount"] / ent["count"] * 100.0, 1) if ent["count"] else 0.0,
            "avgChg1d": round(ent["chgSum"] / ent["count"], 2) if ent["count"] else 0.0,
            "count": ent["count"],
            "leaders": [
                {"symbol": sym, "name": nm, "volB": round(dv, 2), "chg1d": chg}
                for dv, sym, nm, chg in ent["leaders"][:4]
            ],
        })
    sector_flows.sort(key=lambda x: x["netFlowB"], reverse=True)

    # 최신 미국 거래일 판별
    trade_dates = [yahoo_stats[s]["tradeDate"] for s in ("SPY", "QQQ", "NVDA", "AAPL") if s in yahoo_stats]
    trade_date = max(trade_dates) if trade_dates else datetime.now().strftime("%Y-%m-%d")

    indices_out = []
    for sym, label, desc in BENCHMARK_ETFS:
        st = yahoo_stats.get(sym) or {}
        indices_out.append({
            "symbol": sym,
            "label": label,
            "desc": desc,
            "price": st.get("price"),
            "chg1d": st.get("chg1d"),
            "chg5d": st.get("chg5d"),
            "chg1m": st.get("chg1m"),
            "rvol": st.get("rvol"),
            "dollarVolB": st.get("dollarVolB"),
        })

    sector_etfs_out = []
    for sym, ko_lbl, sec_key in SECTOR_ETFS:
        st = yahoo_stats.get(sym) or {}
        sector_etfs_out.append({
            "symbol": sym,
            "label": ko_lbl,
            "sector": sec_key,
            "price": st.get("price"),
            "chg1d": st.get("chg1d"),
            "chg5d": st.get("chg5d"),
            "chg1m": st.get("chg1m"),
            "rvol": st.get("rvol"),
            "dollarVolB": st.get("dollarVolB"),
        })

    stocks_out = []
    for sym, base in cand_map.items():
        st = yahoo_stats.get(sym) or {}
        price = st.get("price") or base["price"]
        vol = st.get("volume") or base["volume"]
        dv_b = st.get("dollarVolB") if st.get("dollarVolB") is not None else base["dollarVolB"]
        chg1d = st.get("chg1d") if st.get("chg1d") is not None else base["chg1d"]
        rvol = st.get("rvol")
        turnover = round((dv_b / base["mcapB"]) * 100.0, 2) if (base["mcapB"] and base["mcapB"] > 0) else base["turnoverPct"]
        sig = classify_us_signal(chg1d, rvol)
        stocks_out.append({
            "symbol": sym,
            "name": base["name"],
            "sector": base["sector"],
            "sectorKo": base["sectorKo"],
            "industry": base["industry"],
            "price": price,
            "chg1d": chg1d,
            "chg5d": st.get("chg5d"),
            "chg1m": st.get("chg1m"),
            "volume": vol,
            "dollarVolB": dv_b,
            "mcapB": base["mcapB"],
            "turnoverPct": turnover,
            "rvol": rvol,
            "signal": sig,
        })
    stocks_out.sort(key=lambda x: x["dollarVolB"], reverse=True)

    accum_cnt = sum(1 for x in stocks_out if x["signal"] == "ACCUM")
    dist_cnt = sum(1 for x in stocks_out if x["signal"] == "DIST")
    surge_cnt = sum(1 for x in stocks_out if (x["rvol"] or 0) >= 1.5)

    # 미국 모닝 브리핑 데이터 조립
    major_indices_out = []
    for sym, label, desc, tv_sym in MAJOR_INDICES:
        st = yahoo_stats.get(sym) or {}
        major_indices_out.append({
            "symbol": sym,
            "label": label,
            "desc": desc,
            "tvSymbol": tv_sym,
            "price": st.get("price"),
            "prevClose": st.get("prevClose"),
            "absChg": st.get("absChg"),
            "chg1d": st.get("chg1d"),
            "chg5d": st.get("chg5d"),
            "chg1m": st.get("chg1m"),
            "tradeDate": st.get("tradeDate"),
        })

    macro_indicators_out = []
    for sym, label, cat, unit, tv_sym, chg_mode in MACRO_INDICATORS:
        st = yahoo_stats.get(sym) or {}
        macro_indicators_out.append({
            "symbol": sym,
            "label": label,
            "category": cat,
            "unit": unit,
            "tvSymbol": tv_sym,
            "chgMode": chg_mode,
            "price": st.get("price"),
            "prevClose": st.get("prevClose"),
            "absChg": st.get("absChg"),
            "chg1d": st.get("chg1d"),
            "chg5d": st.get("chg5d"),
            "tradeDate": st.get("tradeDate"),
        })

    big_tech_out = []
    for sym, ko_name, grp in BIG_TECH_CORE:
        st = yahoo_stats.get(sym) or {}
        base = uni_by_sym.get(sym) or {}
        big_tech_out.append({
            "symbol": sym,
            "nameKo": ko_name,
            "group": grp,
            "price": st.get("price") or base.get("price"),
            "chg1d": st.get("chg1d") if st.get("chg1d") is not None else base.get("chg1d"),
            "chg5d": st.get("chg5d"),
            "dollarVolB": st.get("dollarVolB") if st.get("dollarVolB") is not None else base.get("dollarVolB"),
            "rvol": st.get("rvol"),
        })

    briefing_meta = build_morning_briefing(
        major_indices_out,
        macro_indicators_out,
        fear_greed,
        big_tech_out,
        sector_etfs_out,
    )

    payload = {
        "updatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "tradeDate": trade_date,
        "universeCount": len(universe),
        "indices": indices_out,
        "sectorEtfs": sector_etfs_out,
        "sectorFlows": sector_flows,
        "stocks": stocks_out,
        "summary": {
            "accumCount": accum_cnt,
            "distCount": dist_cnt,
            "surgeCount": surge_cnt,
        },
        "morning": {
            "regime": briefing_meta["regime"],
            "regimeKo": briefing_meta["regimeKo"],
            "krImpact": briefing_meta["krImpact"],
            "krImpactKo": briefing_meta["krImpactKo"],
            "bullets": briefing_meta["bullets"],
            "majorIndices": major_indices_out,
            "macroIndicators": macro_indicators_out,
            "fearGreed": fear_greed,
            "bigTech": big_tech_out,
            "news": morning_news,
        },
    }

    with open(os.path.join(here, "us_market.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    print(
        f"[3/3] 저장 완료: us_market.json (미국 거래일: {trade_date} · 핵심 {len(stocks_out)}종목 · "
        f"매집시그널 {accum_cnt} / 대량매도 {dist_cnt} / 거래량급증 {surge_cnt})"
    )
    print("  - 거래대금 TOP 5:", [(x["symbol"], f"${x['dollarVolB']:.2f}B", f"{x['chg1d']:+.2f}%", f"RVOL {x['rvol']}x") for x in stocks_out[:5]])
    return payload


def main():
    skip_push = any(a.lower() == "--no-push" for a in sys.argv[1:])
    collect_us_market(HERE)

    sys.path.insert(0, HERE)
    import kospi200_yoy as screener

    screener.build_web_dashboard(HERE)
    if not skip_push:
        screener.auto_push_to_github(
            HERE,
            commit_msg=f"Auto-update US market flow ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
        )


if __name__ == "__main__":
    main()
