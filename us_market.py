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
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta

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
            "price": round(last_c, 2),
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

    # 2) Yahoo v8 정밀 보강 대상 선별:
    #    - 시총 $5B 이상 거래대금 TOP 100
    #    - 시총 $3B 이상 & 거래대금 $150M 이상 일간 급등/급락 모멘텀 TOP 35
    large_caps = [s for s in universe if s["mcapB"] >= 5.0]
    large_caps.sort(key=lambda x: x["dollarVol"], reverse=True)
    cand_map = {}
    for s in large_caps[:100]:
        cand_map[s["symbol"]] = s

    momo_pool = [s for s in universe if s["mcapB"] >= 3.0 and s["dollarVolB"] >= 0.15 and s["chg1d"] is not None]
    momo_pool.sort(key=lambda x: abs(x["chg1d"]), reverse=True)
    for s in momo_pool[:35]:
        cand_map.setdefault(s["symbol"], s)

    etf_syms = sorted({x[0] for x in BENCHMARK_ETFS} | {x[0] for x in SECTOR_ETFS})
    all_fetch_syms = sorted(set(cand_map.keys()) | set(etf_syms))

    print(f"[2/3] Yahoo Finance 일봉 조회 중 (ETF {len(etf_syms)}개 + 핵심 종목 {len(cand_map)}개)...")
    yahoo_stats = {}
    with ThreadPoolExecutor(max_workers=14) as ex:
        futs = [ex.submit(fetch_yahoo_1mo_stats, sym) for sym in all_fetch_syms]
        for fut in as_completed(futs):
            sym, st = fut.result()
            if st:
                yahoo_stats[sym] = st
    print(f"  - 일봉 지표(RVOL·5일·1개월) 조회 성공: {len(yahoo_stats)} / {len(all_fetch_syms)}개")

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
