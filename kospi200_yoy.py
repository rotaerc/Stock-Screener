# -*- coding: utf-8 -*-
"""
코스피/코스닥 시총 상위 200개 영업이익 성장·수익성 스크리너 (최종 .xlsx 엑셀 전용 버전)
===================================================================================================
[주요 기능]
1. 네이버 금융 실시간 시가총액 순위 API/웹 연동 (우선주·ETF·스팩 자동 제외)
   -> 1단계 수집 시 네이버 실시간 '현재가' 및 '오늘등락율(%)' 동시 추출
2. 실적 발표 시즌(예: SK하이닉스 등 2026/09 발표 완료 종목)에서 예상실적(OP_2026/09(E), YoY_2026/09(E,%))이
   비어있을 때 3분기 확정 실적 데이터를 예상실적에 동일하게 복사하여 통과 판정 누락 원천 방지
3. 통과 조건 중 E1, E2의 영업이익 성장률(YoY)이 둘 다 마이너스(< 0)인 기업은 통과 대상에서 제외
4. 26년 대비 27년 연간 이익 성장 + 연간 이익률 상승 기업 중:
   - E3 예측치가 있는 기업은 'E0 < E1 < E2 < E3' 연속 성장 기업 추출 ("E0~E3")
   - E3 예측치가 아직 없는(빈칸인) 기업은 'E0 < E1 < E2' 연속 성장 기업도 함께 추출 ("E0~E2(E3미집계)")
5. 한 번 실행으로 KOSPI·KOSDAQ 시총 상위 200개를 연속 수집하여 각각의 통합 엑셀(.xlsx) 파일 자동 생성
"""

import html as html_lib
import io
import json
import os
import re
import shutil
import sys
import subprocess
import time
from datetime import datetime, timedelta

import pandas as pd
import requests

# 엑셀(.xlsx) 저장을 위한 openpyxl 라이브러리 자동 확인 및 설치
try:
    import openpyxl
except ImportError:
    print("[알림] 엑셀(.xlsx) 파일 생성을 위해 openpyxl 라이브러리를 자동 설치합니다...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    import openpyxl

# 기본 실행 대상 시장 리스트 (한 번 실행 시 KOSPI, KOSDAQ 순서로 모두 생성)
MARKETS = ["KOSPI", "KOSDAQ"]
TOP_N = 200

BASE = "https://wcomp.fnguide.com"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
DELAY = 1.0

SNP_CONSOL = "M"
NEG_BASE_YOY = True
HIGH_GROWTH = 50.0
E3_JUMP = 20.0       
E3_MIN_YOY = 20.0    

ETF_PREFIXES = (
    "KODEX ", "TIGER ", "KBSTAR ", "RISE ", "ACE ", "SOL ", "HANARO ",
    "KOSEF ", "ARIRANG ", "PLUS ", "TIMEFOLIO ", "WON ", "1Q ", "히어로즈 ",
    "마이티 ", "파워 ", "에셋플러스 ", "마이다스 ", "트러스톤 ", "엔젤 ",
    "FOCUS ", "TREX ", "MASTER ", "VITA ", "UNICORN ", "KoAct ", "BNK ", "메리츠 "
)

# 재무 Snapshot 데이터가 없는 인프라펀드/ETF성 종목 제외 목록
EXCLUDED_CODES = {
    "088980",  # 맥쿼리인프라
    "415640",  # KB발해인프라
}
EXCLUDED_NAMES = {
    "맥쿼리인프라",
    "KB발해인프라",
}


def make_session():
    s = requests.Session()
    s.headers.update({
        "User-Agent": UA,
        "Accept": "application/json, text/javascript, */*; q=0.01",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "X-Requested-With": "XMLHttpRequest",
    })
    try:
        s.get(BASE + "/", headers={"Accept": "text/html"}, timeout=20)
    except Exception:
        pass
    return s


def is_common_stock(code, name, end_type="stock"):
    """우선주, ETF, ETN, 스팩 및 지정 제외 종목(인프라펀드 등)을 제외하고 순수 보통주만 판별"""
    if not (len(code) == 6 and code.isdigit() and code.endswith("0")):
        return False
    if not name:
        return False
    if code in EXCLUDED_CODES or name in EXCLUDED_NAMES:
        return False
    if end_type and str(end_type).lower() not in ("stock", ""):
        return False
    if name.endswith(("우", "우B", "우C", "우(전환)")):
        return False
    if name.startswith(ETF_PREFIXES) or "ETN" in name or "스팩" in name or "SPAC" in name:
        return False
    return True


def _num(x):
    if pd.isna(x): return None
    s = str(x).replace(",", "").strip()
    if s in ("", "-", "null", "None", "nan", "NaN"): return None
    try: return float(s)
    except ValueError: return None


def get_market_cap_top(market="KOSPI", top_n=200):
    """네이버 금융에서 실시간 시가총액 순위대로 종목(보통주) + 현재가/오늘등락율 수집"""
    print(f"[1단계] {market} 시가총액 상위 {top_n}개 종목(우선주·ETF 제외) 및 실시간 등락율 수집 중...")
    result = []
    naver_quote_map = {}  # {code: (cur_price, change_rate)}
    seen = set()
    headers = {"User-Agent": UA, "Accept": "application/json"}

    # 1차: 네이버 모바일 주식 시가총액 API (현재가 및 등락율 동시 제공)
    for page in range(1, 8):
        try:
            url = f"https://m.stock.naver.com/api/stocks/marketValue/{market}?page={page}&pageSize=100"
            r = requests.get(url, headers=headers, timeout=15)
            if r.status_code != 200:
                break
            stocks = r.json().get("stocks", [])
            if not stocks:
                break
            for item in stocks:
                code = str(item.get("itemCode", "")).strip()
                name = str(item.get("stockName", "")).strip()
                end_type = str(item.get("stockEndType", "stock")).strip()
                if code not in seen and is_common_stock(code, name, end_type):
                    seen.add(code)
                    result.append((code, name))
                    
                    # 네이버 실시간 현재가 및 등락률 추출
                    n_price = _num(item.get("closePrice"))
                    n_rate = _num(item.get("fluctuationsRatio"))
                    cmp_info = item.get("compareToPreviousPrice") or {}
                    if n_rate is not None and n_rate > 0 and str(cmp_info.get("code")) in ("4", "5"):
                        n_rate = -n_rate
                    naver_quote_map[code] = (n_price, n_rate)
                    
                    if len(result) >= top_n:
                        return result, naver_quote_map
        except Exception:
            break

    # 2차(예비): 네이버 PC 금융 시가총액 페이지 파싱
    sosok = 0 if market == "KOSPI" else 1
    for page in range(1, 10):
        try:
            url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok}&page={page}"
            r = requests.get(url, headers={"User-Agent": UA}, timeout=15)
            try:
                html = r.content.decode("utf-8")
            except UnicodeDecodeError:
                html = r.content.decode("cp949", errors="replace")
            
            trs = re.findall(r'<tr[^>]*>(.*?)</tr>', html, re.DOTALL)
            found_any = False
            for tr in trs:
                m_item = re.search(r'<a\s+href="/item/main\.naver\?code=(\d{6})"\s+class="tltle">([^<]+)</a>', tr)
                if not m_item:
                    continue
                found_any = True
                code, name = m_item.group(1).strip(), m_item.group(2).strip()
                if code not in seen and is_common_stock(code, name):
                    seen.add(code)
                    result.append((code, name))
                    
                    nums = re.findall(r'<td\s+class="number">\s*([\d,]+)\s*</td>', tr)
                    n_price = _num(nums[0]) if nums else None
                    m_rt = re.search(r'([+-]?\d+\.\d+)\s*%', tr)
                    n_rate = float(m_rt.group(1)) if m_rt else None
                    naver_quote_map[code] = (n_price, n_rate)
                    
                    if len(result) >= top_n:
                        return result, naver_quote_map
            if not found_any:
                break
        except Exception:
            break

    return result, naver_quote_map


def _fetch(code, sess, freq):
    r = sess.get(BASE + "/CompanyInfo/getCnsPerforTrend",
                 params={"cmp_cd": code, "consol_typ": "C",
                         "freq_typ": freq, "data_typ": "1"},
                 headers={"Referer": f"{BASE}/CompanyInfo/Consensus"
                                     f"?c_id=AA&menu_type=01&cmp_cd={code}"},
                 timeout=25)
    r.raise_for_status()
    ds = r.json().get("dataset", {})
    header, rows = ds.get("header", []), ds.get("data", [])
    if not header:
        raise RuntimeError(f"header 없음(freq={freq})")
    cd2ym = {h["CD"]: h["YYMM"] for h in header}
    periods = sorted(
        [(h["YYMM"], str(h.get("EP_CHK", "")).strip().upper() == "E") for h in header],
        key=lambda t: t[0])

    def row_of(ac, want):
        for x in rows:
            if x.get("AC_CODE") != ac:
                continue
            nm = str(x.get("NAME", "")).strip()
            if want == "value" and nm in ("매출액", "영업이익"):
                return x
            if want == "yoy" and "전년동기대비" in nm:
                return x
        return None

    def series(ac, want):
        row = row_of(ac, want)
        return {ym: _num(row.get(cd)) for cd, ym in cd2ym.items()} if row else {}

    def raw_series(ac, want):
        row = row_of(ac, want)
        return ({ym: (str(row.get(cd)).strip() if row.get(cd) is not None else "")
                 for cd, ym in cd2ym.items()} if row else {})

    return {
        "periods": periods,
        "sales": series("121000", "value"),
        "op":    series("121450", "value"),
        "op_yoy": series("121450", "yoy"),
        "op_yoy_raw": raw_series("121450", "yoy"),
    }


def op_margin(op_val, sales_val):
    if op_val is None or sales_val is None or sales_val <= 0:
        return None
    return round(op_val / sales_val * 100, 2)


def calc_yoy(cur, prev):
    if cur is None or prev is None or prev == 0:
        return None, ""
    num = round((cur - prev) / abs(prev) * 100, 2)
    if prev > 0 and cur > 0:
        return num, f"{num:.2f}"
    label = "적자전환" if prev > 0 else ("흑자전환" if cur > 0 else "적자지속")
    return (num if (NEG_BASE_YOY and cur > 0) else None), label


def prev_year(ym):
    y, mth = ym.split("/")
    return f"{int(y) - 1}/{mth}"


def _parse_sectors_from_snapshot_html(html_text):
    """
    FnGuide Snapshot 상단 corp_group1 영역에서 거래소 업종(섹터1)과 WI26 업종(섹터2)을 추출
    예: '코스피 전기·전자' -> '전기전자', 'WI26 반도체' -> '반도체'
    """
    sector1, sector2 = "", ""
    if not html_text:
        return sector1, sector2

    m_grp = re.search(r'<p\s+class="stxt_group">(.*?)</p>', html_text, re.DOTALL)
    target_html = m_grp.group(1) if m_grp else html_text[:15000]
    spans = re.findall(r'<span\s+class="stxt[^"]*">(.*?)</span>', target_html, re.DOTALL)

    for raw in spans:
        txt = html_lib.unescape(re.sub(r'<[^>]+>', '', raw)).strip()
        if not txt:
            continue
        if not sector1 and (txt.startswith("코스피") or txt.startswith("코스닥")):
            s1 = re.sub(r'^(?:코스피|코스닥)\s*', '', txt).strip()
            s1 = re.sub(r'[·ㆍ\xb7]', '', s1)
            sector1 = s1
        elif not sector2 and txt.upper().startswith("WI"):
            s2 = re.sub(r'^WI\d*\s*', '', txt, flags=re.IGNORECASE).strip()
            s2 = re.sub(r'[·ㆍ\xb7]', '', s2)
            sector2 = s2

    return sector1, sector2


def get_price_info_from_snapshot(code, sess, naver_fallback=None):
    """
    1순위: 네이버 실시간 현재가/등락률 + FnGuide 52주 최고가 및 섹터(업종/WI26) 결합
    2순위: FnGuide Snapshot 페이지의 '종가/전일대비/수익률' 텍스트 정밀 파싱
    3순위: 개별 네이버 모바일 basic API 호출
    """
    n_price, n_rate = naver_fallback if naver_fallback else (None, None)
    cur_price, change_rate, high_52w = n_price, n_rate, None
    sector1, sector2 = "", ""

    try:
        page = f"{BASE}/CompanyInfo/Snapshot?c_id=AA&menu_type=01&cmp_cd={code}"
        r = sess.get(page, headers={"Accept": "text/html", "X-Requested-With": None}, timeout=25)
        if r.status_code == 200:
            html = r.text
            sector1, sector2 = _parse_sectors_from_snapshot_html(html)
            idx_siset = html.find('시세현황')
            if idx_siset != -1:
                sub_html = html[idx_siset:idx_siset+3500]
                clean_text = re.sub(r'<[^>]+>', ' ', sub_html)
                clean_text = re.sub(r'\s+', ' ', clean_text)
                
                m_quote = re.search(
                    r'종가[^0-9]*?([\d,]+)\s*/\s*([+-]?[\d,]+)\s*/\s*([+-]?\d+(?:\.\d+)?)',
                    clean_text
                )
                if m_quote:
                    if cur_price is None:
                        cur_price = _num(m_quote.group(1))
                    if change_rate is None:
                        change_rate = _num(m_quote.group(3))
                else:
                    if cur_price is None:
                        m_close = re.search(r'종가[^0-9]*?([\d,]+)', clean_text)
                        if m_close:
                            cur_price = _num(m_close.group(1))
                    if change_rate is None:
                        m_rate = re.search(r'(?:종가|전일대비|수익률|등락률)[^%]*?([+-]?\d+(?:\.\d+)?)\s*%?', clean_text)
                        if m_rate:
                            change_rate = _num(m_rate.group(1))
                
                # 52주 최고가 파싱
                m_high = re.search(r'52주[^0-9]*?최고가[^0-9]*?([\d,]+)', clean_text)
                if m_high:
                    high_52w = _num(m_high.group(1))
    except Exception:
        pass

    if cur_price is None or change_rate is None:
        try:
            url = f"https://m.stock.naver.com/api/stock/{code}/basic"
            r_nav = requests.get(url, headers={"User-Agent": UA, "Accept": "application/json"}, timeout=10)
            if r_nav.status_code == 200:
                d = r_nav.json()
                if cur_price is None:
                    cur_price = _num(d.get("closePrice"))
                if change_rate is None:
                    rt = _num(d.get("fluctuationsRatio"))
                    cmp_info = d.get("compareToPreviousPrice") or {}
                    if rt is not None and rt > 0 and str(cmp_info.get("code")) in ("4", "5"):
                        rt = -rt
                    change_rate = rt
        except Exception:
            pass

    return cur_price, change_rate, high_52w, sector1, sector2


def _fetch_snp_q(code, sess):
    page = f"{BASE}/CompanyInfo/Snapshot?c_id=AA&menu_type=01&cmp_cd={code}"

    def call():
        return sess.get(BASE + "/CompanyInfo/getSnpFinancial",
                        params={"cmp_cd": code, "consol_typ": SNP_CONSOL, "freq_typ": "Q"},
                        headers={"Referer": page}, timeout=25)

    ds = {}
    try:
        r = call()
        r.raise_for_status()
        ds = r.json().get("dataset") or {}
    except Exception:
        pass
    if not ds.get("header") or not ds.get("data"):
        sess.get(page, headers={"Accept": "text/html", "X-Requested-With": None}, timeout=25)
        time.sleep(0.5)
        r = call()
        r.raise_for_status()
        ds = r.json().get("dataset") or {}
    header, rows = ds.get("header") or [], ds.get("data") or []
    if not header or not rows:
        raise RuntimeError("Snapshot 분기 header/data 없음")

    yms = [h["YYMM"] for h in header]
    if len(set(yms)) != len(yms):
        raise RuntimeError(f"분기 기간 중복 {yms}")
    periods = sorted(
        [(h["YYMM"], str(h.get("EP_CHK") or "").strip().upper() == "E") for h in header],
        key=lambda t: t[0])

    def row(label):
        for x in rows:
            if re.sub(r"\s+", "", str(x.get("NAME", ""))) == label:
                return {h["YYMM"]: _num(x.get(h["CD"])) for h in header}
        return None

    sales, op = row("매출액"), row("영업이익(발표기준)")
    if sales is None or op is None:
        raise RuntimeError("매출액 / 영업이익(발표기준) 행 없음")
    if not any(v is not None for v in op.values()):
        raise RuntimeError("영업이익(발표기준) 값 없음")

    op_yoy, op_yoy_raw = {}, {}
    for ym in yms:
        op_yoy[ym], op_yoy_raw[ym] = calc_yoy(op.get(ym), op.get(prev_year(ym)))
    return {"periods": periods, "sales": sales, "op": op,
            "op_yoy": op_yoy, "op_yoy_raw": op_yoy_raw}


def assign_roles(periods, op=None):
    """
    실제 영업이익(op) 데이터가 존재하는 확정 분기만 actual로 인정하되,
    실적 시즌에 3분기(예: 2026/09)가 방금 발표되어 8개 분기 중 '실적 6개 + 추정 2개'가 된 경우,
    발표된 최신 분기(2026/09)를 예상실적(E1) 자리에 동일하게 배정하여 누락을 방지합니다.
    """
    if op is not None:
        actual_idx = [i for i, (ym, e) in enumerate(periods) if not e and op.get(ym) is not None]
        if not actual_idx:
            actual_idx = [i for i, (_, e) in enumerate(periods) if not e]
    else:
        actual_idx = [i for i, (_, e) in enumerate(periods) if not e]

    if not actual_idx:
        raise RuntimeError("실적 분기가 없음")
    b = actual_idx[-1]
    after = [periods[i][0] for i in range(b + 1, len(periods))]

    # ★ 실적 시즌 핵심 처리: 최신 분기(예: 2026/09)가 실적으로 전환되어 추정치(after)가 2개만 남은 경우
    if len(after) == 2 and len(actual_idx) >= 2:
        b_prev = actual_idx[-2]
        announced_ym = periods[b][0]
        return {
            "actuals": [periods[i][0] for i in actual_idx],
            "base": periods[b_prev][0],       # E0 (예: 2026/06)
            "E1": announced_ym,               # E1 (예: 2026/09 - 발표된 실적을 E1 예상실적 자리에 투입)
            "E2": after[0],                   # E2 (예: 2026/12(E))
            "E3": after[1],                   # E3 (예: 2027/03(E))
            "announced_e1": announced_ym,
            "shifted": {                      # 이동 후 윈도우도 보관하여 교차 검증
                "base": announced_ym,
                "E1": after[0],
                "E2": after[1],
            }
        }

    if len(after) < 2:
        raise RuntimeError("추정 분기가 2개 미만")
    return {
        "actuals": [periods[i][0] for i in actual_idx],
        "base": periods[b][0],
        "E1": after[0],
        "E2": after[1],
        "E3": after[2] if len(after) > 2 else None,
        "announced_e1": None,
        "shifted": None,
    }


def assign_years(periods):
    est = [ym for ym, e in periods if e]
    if len(est) < 2:
        raise RuntimeError("추정 연도가 2개 미만")
    return est[0], est[1]


def _eval_window(base_ym, e1_ym, e2_ym, yoy_map, op_map, sales_map, cond2, is_announced_e1=False):
    """주어진 (base, E1, E2) 분기 조합에 대해 성장조건(①), 분기마진(③), 통과 여부를 판정"""
    yb = yoy_map.get(base_ym)
    y1 = yoy_map.get(e1_ym)
    y2 = yoy_map.get(e2_ym)

    both_neg = (y1 is not None and y2 is not None and y1 < 0 and y2 < 0)
    acc1 = (yb is not None and y1 is not None and y1 > yb) and not both_neg
    acc2 = (yb is not None and y2 is not None and y2 > yb) and not both_neg
    high = (y1 is not None and y2 is not None and y1 >= HIGH_GROWTH and y2 >= HIGH_GROWTH)

    growth_label = "가속" if (acc1 or acc2) else ("고성장" if high else "")
    cond1 = (acc1 or acc2 or high) and not both_neg

    m_base = op_margin(op_map.get(base_ym), sales_map.get(base_ym))
    m_e1 = op_margin(op_map.get(e1_ym), sales_map.get(e1_ym))
    m_e2 = op_margin(op_map.get(e2_ym), sales_map.get(e2_ym))

    cond3 = (m_base is not None and m_e1 is not None and m_e2 is not None
             and m_e1 > m_base and m_e2 > m_e1)

    # ★ 실적 발표 종목 어닝 서프라이즈 보호 로직
    if not cond3 and is_announced_e1:
        op_b, op_1, op_2 = op_map.get(base_ym), op_map.get(e1_ym), op_map.get(e2_ym)
        if (m_base is not None and m_e1 is not None and m_e2 is not None
                and m_e1 > m_base and m_e2 > m_base
                and op_b is not None and op_1 is not None and op_2 is not None
                and op_b < op_1 < op_2):
            cond3 = True

    return {
        "passed": bool(cond1 and cond2 and cond3),
        "cond1": cond1,
        "cond3": cond3,
        "growth_label": growth_label,
        "yb": yb, "y1": y1, "y2": y2,
        "m_base": m_base, "m_e1": m_e1, "m_e2": m_e2,
    }


def build_record(code, name, q, a, price_info):
    op, yoy, qs = dict(q["op"]), dict(q["op_yoy"]), dict(q["sales"])
    R = assign_roles(q["periods"], op)
    Y1, Y2 = assign_years(a["periods"])
    if len(price_info) >= 5:
        cur_price, change_rate, high_52w, sector1, sector2 = price_info[:5]
    else:
        cur_price, change_rate, high_52w = price_info[:3]
        sector1, sector2 = "", ""
    announced_e1 = R.get("announced_e1")
    
    rec = {
        "종목코드": code, 
        "종목명": name,
        "섹터1": sector1,
        "섹터2": sector2,
        "현재가": cur_price,
        "오늘등락율(%)": change_rate,
        "52주최고가": high_52w
    }
    
    if cur_price is not None and high_52w is not None and high_52w > 0:
        drop_pct = round((high_52w - cur_price) / high_52w * 100, 2)
        rec["52주최고가대비하락률(%)"] = drop_pct
        rec["52주최고가대비30%이상하락"] = drop_pct >= 30.0
    else:
        rec["52주최고가대비하락률(%)"] = None
        rec["52주최고가대비30%이상하락"] = False

    for ym, is_e in q["periods"]: 
        rec[f"OP_{ym}" + ("(E)" if is_e else "")] = op.get(ym)

    # ★ 실적 발표로 인해 3분기(예: 2026/09)가 확정 실적으로 들어온 경우 예상실적(E)에도 같은 데이터 입력
    if announced_e1 and op.get(announced_e1) is not None:
        rec[f"OP_{announced_e1}"] = op.get(announced_e1)
        rec[f"OP_{announced_e1}(E)"] = op.get(announced_e1)

    for role in ["base", "E1", "E2"] + (["E3"] if R["E3"] else []):
        ym = R[role]
        pym = prev_year(ym)
        val = op.get(pym)
        
        if val is None and op.get(ym) is not None and yoy.get(ym) is not None:
            f = 1.0 + yoy.get(ym) / 100.0
            if f > 0: 
                val = round(op.get(ym) / f, 2)
        if val is not None:
            op[pym] = val
            
        if f"OP_{pym}" in rec: 
            if rec.get(f"OP_{pym}") is None:
                rec[f"OP_{pym}"] = val
        elif f"OP_{pym}(E)" in rec: 
            if rec.get(f"OP_{pym}(E)") is None:
                rec[f"OP_{pym}(E)"] = val
        else: 
            rec[f"OP_{pym}(전년)"] = val

    def first_valid(*vals):
        for v in vals:
            if v is not None:
                return v
        return None

    for role in ("base", "E1", "E2", "E3"):
        if not R.get(role): continue
        ym = R[role]
        y_val = yoy.get(ym)
        
        if y_val is None:
            pym = prev_year(ym)
            op_curr = first_valid(op.get(ym), rec.get(f"OP_{ym}"), rec.get(f"OP_{ym}(E)"))
            op_prev = first_valid(op.get(pym), rec.get(f"OP_{pym}"), rec.get(f"OP_{pym}(E)"), rec.get(f"OP_{pym}(전년)"))
            
            if op_curr is not None and op_prev is not None and op_prev != 0:
                y_val, _ = calc_yoy(op_curr, op_prev)
        yoy[ym] = y_val
                
        tag = "기준" if role == "base" else "E"
        rec[f"YoY_{ym}({tag},%)"] = y_val

        # ★ 발표된 분기(예: 2026/09)는 YoY(E,%)와 YoY(기준,%) 양쪽에 모두 같은 데이터를 넣어 빈칸 방지
        if ym == announced_e1 and y_val is not None:
            rec[f"YoY_{ym}(E,%)"] = y_val
            rec[f"YoY_{ym}(기준,%)"] = y_val

    # --- 연간 영업이익 및 연간 영업이익률 계산 ---
    op_y1 = a["op"].get(Y1)
    op_y2 = a["op"].get(Y2)
    rec[f"연간OP_{Y1}(E)"] = op_y1
    rec[f"연간OP_{Y2}(E)"] = op_y2
    y_op_growth, _ = calc_yoy(op_y2, op_y1)
    rec["연간OP_성장률(%)"] = y_op_growth
    cond_annual_op_up = (op_y1 is not None and op_y2 is not None and op_y2 > op_y1 and op_y2 > 0)
    rec["연간OP_차년>당해"] = bool(cond_annual_op_up)

    m1 = op_margin(op_y1, a["sales"].get(Y1))
    m2 = op_margin(op_y2, a["sales"].get(Y2))
    rec[f"연간영업이익률_{Y1}(E,%)"] = m1
    rec[f"연간영업이익률_{Y2}(E,%)"] = m2
    rec["연간마진_개선폭(%p)"] = round(m2 - m1, 2) if (m1 is not None and m2 is not None) else None
    cond2 = m1 is not None and m2 is not None and m2 > m1
    rec["연간마진_차년>당해"] = cond2

    qm = {}
    for role in ("base", "E1", "E2", "E3"):
        ym = R.get(role)
        if not ym:
            continue
        qm[role] = op_margin(op.get(ym), qs.get(ym))
        tag = "기준" if role == "base" else "E"
        rec[f"분기영업이익률_{ym}({tag},%)"] = qm[role]
        if ym == announced_e1 and qm[role] is not None:
            rec[f"분기영업이익률_{ym}(E,%)"] = qm[role]
            rec[f"분기영업이익률_{ym}(기준,%)"] = qm[role]

    # ★ 발표 분기를 E1(예상실적)에 넣은 평가(eval_main)와 이동 후 평가(eval_shift) 교차 검증
    eval_main = _eval_window(
        R["base"], R["E1"], R["E2"], yoy, op, qs, cond2,
        is_announced_e1=bool(announced_e1)
    )
    eval_shift = None
    if R.get("shifted"):
        sh = R["shifted"]
        eval_shift = _eval_window(
            sh["base"], sh["E1"], sh["E2"], yoy, op, qs, cond2,
            is_announced_e1=True
        )

    yb = eval_main["yb"]
    y1 = eval_main["y1"]
    y2 = eval_main["y2"]
    y3 = yoy.get(R["E3"]) if R.get("E3") else None

    rec["가속폭_E1-기준(%p)"] = round(y1 - yb, 2) if (yb is not None and y1 is not None) else None
    rec["가속폭_E2-기준(%p)"] = round(y2 - yb, 2) if (yb is not None and y2 is not None) else None
    rec["가속폭_E3-E2(%p)"] = round(y3 - y2, 2) if (y2 is not None and y3 is not None) else None

    rec["성장조건"] = eval_main["growth_label"] or (eval_shift["growth_label"] if eval_shift else "")
    rec["분기마진_기준<E1<E2"] = bool(eval_main["cond3"] or (eval_shift and eval_shift["cond3"]))
    rec["통과"] = bool(eval_main["passed"] or (eval_shift and eval_shift["passed"]))

    rec["E3반등"], rec["E3반등사유"] = False, ""
    if R["E3"]:
        e3 = R["E3"]
        raw3 = q.get("op_yoy_raw", {}).get(e3, "")
        rec["YoY_E3표기"] = raw3
        m_e2 = qm.get("E2")
        m_e3 = qm.get("E3")
        
        turn = "흑자전환" in raw3
        jump = (y2 is not None and y3 is not None
                and (y3 - y2) >= E3_JUMP and y3 >= E3_MIN_YOY)
        margin_up = m_e2 is not None and m_e3 is not None and m_e3 > m_e2
        
        if (turn or jump) and margin_up:
            rec["E3반등"] = True
            rec["E3반등사유"] = "흑자전환" if turn else "급반등"

    # 메인 조건 통과 + E3에서도 추가 반등 성공 여부
    rec["통과+E3반등"] = bool(rec["통과"] and rec["E3반등"])

    # -------------------------------------------------------------------------
    # ★ [연속이익성장 조건]
    # 1) 26년보다 내년(27년) 연간 영업이익 성장 (cond_annual_op_up)
    # 2) 연간 영업이익률 상승 (cond2: 연간마진_차년>당해)
    # 3) 분기 영업이익 연속 성장:
    #    - E3 예측치가 있는 경우: E0 < E1 < E2 < E3 ("E0~E3")
    #    - E3 예측치가 아직 없는(빈칸) 경우: E0 < E1 < E2 ("E0~E2(E3미집계)")
    # -------------------------------------------------------------------------
    op_e0 = first_valid(op.get(R["base"]), rec.get(f"OP_{R['base']}"), rec.get(f"OP_{R['base']}(E)"))
    op_e1 = first_valid(op.get(R["E1"]), rec.get(f"OP_{R['E1']}"), rec.get(f"OP_{R['E1']}(E)"))
    op_e2 = first_valid(op.get(R["E2"]), rec.get(f"OP_{R['E2']}"), rec.get(f"OP_{R['E2']}(E)"))
    op_e3 = first_valid(op.get(R["E3"]), rec.get(f"OP_{R['E3']}"), rec.get(f"OP_{R['E3']}(E)")) if R.get("E3") else None

    cond_e0_to_e2 = (
        op_e0 is not None and op_e1 is not None and op_e2 is not None
        and (op_e0 < op_e1 < op_e2)
        and op_e2 > 0
    )

    if op_e3 is not None:
        cond_q_op_seq = bool(cond_e0_to_e2 and (op_e2 < op_e3) and op_e3 > 0)
        seq_type = "E0~E3" if cond_q_op_seq else ""
        target_end_op = op_e3
    else:
        cond_q_op_seq = bool(cond_e0_to_e2)
        seq_type = "E0~E2(E3미집계)" if cond_q_op_seq else ""
        target_end_op = op_e2

    rec["분기OP_연속성장"] = bool(cond_q_op_seq)
    rec["연속성장구간"] = seq_type

    if op_e0 is not None and target_end_op is not None and op_e0 != 0:
        rec["OP성장률_최종대비E0(%)"] = round((target_end_op - op_e0) / abs(op_e0) * 100, 2)
    else:
        rec["OP성장률_최종대비E0(%)"] = None

    rec["연속이익성장(연간+분기)"] = bool(cond_annual_op_up and cond2 and cond_q_op_seq)
            
    return rec, R, (Y1, Y2)


def fill_missing_base_from_estimate(df):
    """
    실적 시즌 과도기에 동일 분기(예: 2026/09)에 대해
    1) 확정 실적(OP_2026/09)만 있고 예상실적(OP_2026/09(E))이 비어있는 경우 -> 예상실적에 같은 데이터 복사
    2) 예상실적(OP_2026/09(E))만 있고 기준(OP_2026/09 / YoY(기준,%))이 비어있는 경우 -> 상호 보완
    """
    for col in list(df.columns):
        if col.startswith("OP_") and not col.endswith("(E)") and not col.endswith("(전년)"):
            ym = col.replace("OP_", "")
            est_col = f"OP_{ym}(E)"
            if est_col in df.columns:
                df[est_col] = df[est_col].fillna(df[col])

            yoy_base = f"YoY_{ym}(기준,%)"
            yoy_est = f"YoY_{ym}(E,%)"
            if yoy_base in df.columns and yoy_est in df.columns:
                df[yoy_est] = df[yoy_est].fillna(df[yoy_base])
                df[yoy_base] = df[yoy_base].fillna(df[yoy_est])

            opm_base = f"분기영업이익률_{ym}(기준,%)"
            opm_est = f"분기영업이익률_{ym}(E,%)"
            if opm_base in df.columns and opm_est in df.columns:
                df[opm_est] = df[opm_est].fillna(df[opm_base])
                df[opm_base] = df[opm_base].fillna(df[opm_est])

    yoy_base_cols = [c for c in df.columns if c.startswith("YoY_") and c.endswith("(기준,%)")]
    for yoy_base_col in yoy_base_cols:
        ym = yoy_base_col.replace("YoY_", "").replace("(기준,%)", "")
        pym = prev_year(ym)
        
        op_act_col = f"OP_{ym}"
        op_est_col = f"OP_{ym}(E)"
        yoy_est_col = f"YoY_{ym}(E,%)"
        
        if yoy_est_col in df.columns:
            df[yoy_base_col] = df[yoy_base_col].fillna(df[yoy_est_col])
            df[yoy_est_col] = df[yoy_est_col].fillna(df[yoy_base_col])
            
        missing_mask = df[yoy_base_col].isna()
        if missing_mask.any():
            for idx in df[missing_mask].index:
                cur_val = None
                if op_act_col in df.columns and pd.notna(df.at[idx, op_act_col]):
                    cur_val = df.at[idx, op_act_col]
                elif op_est_col in df.columns and pd.notna(df.at[idx, op_est_col]):
                    cur_val = df.at[idx, op_est_col]
                    
                prev_val = None
                for p_col in (f"OP_{pym}", f"OP_{pym}(E)", f"OP_{pym}(전년)"):
                    if p_col in df.columns and pd.notna(df.at[idx, p_col]):
                        prev_val = df.at[idx, p_col]
                        break
                        
                if cur_val is not None and prev_val is not None and prev_val != 0:
                    y_num, _ = calc_yoy(float(cur_val), float(prev_val))
                    if y_num is not None:
                        df.at[idx, yoy_base_col] = y_num
                        if yoy_est_col in df.columns and pd.isna(df.at[idx, yoy_est_col]):
                            df.at[idx, yoy_est_col] = y_num
                        
    return df


def _current_quarter_ym(dt=None):
    d = dt or datetime.now()
    q_month = ((d.month - 1) // 3 + 1) * 3
    return f"{d.year}/{q_month:02d}"


def _find_current_q_op_col(cols, dt=None):
    cur_q = _current_quarter_ym(dt)
    exact = f"OP_{cur_q}(E)"
    if exact in cols:
        return cur_q, exact
    op_e_cols = sorted([c for c in cols if str(c).startswith("OP_") and str(c).endswith("(E)")])
    for c in op_e_cols:
        ym = str(c).replace("OP_", "").replace("(E)", "")
        if ym >= cur_q:
            return ym, c
    if op_e_cols:
        c0 = op_e_cols[0]
        return str(c0).replace("OP_", "").replace("(E)", ""), c0
    return cur_q, None


def _recompute_op_history_entry(entry, quarter_label):
    hist = entry.get("history") or []
    up_cnt, down_cnt = 0, 0
    for i in range(1, len(hist)):
        prev_v = hist[i - 1].get("v")
        cur_v = hist[i].get("v")
        if prev_v is None or cur_v is None:
            continue
        if round(cur_v, 2) > round(prev_v, 2):
            up_cnt += 1
        elif round(cur_v, 2) < round(prev_v, 2):
            down_cnt += 1

    first_v = hist[0].get("v") if hist else None
    last_v = hist[-1].get("v") if hist else None
    total_pct = None
    if first_v is not None and last_v is not None and first_v != 0:
        total_pct = round((last_v - first_v) / abs(first_v) * 100.0, 2)

    net_up = bool(first_v is not None and last_v is not None and round(last_v, 2) > round(first_v, 2))
    level = 0
    if net_up and total_pct is not None:
        if total_pct >= 20.0:
            level = 4
        elif total_pct >= 15.0:
            level = 3
        elif total_pct >= 10.0:
            level = 2
        elif total_pct >= 5.0:
            level = 1

    lvl_labels = {1: "1단계(+5%↑)", 2: "2단계(+10%↑)", 3: "3단계(+15%↑)", 4: "4단계(+20%↑)"}
    name = entry.get("name") or ""
    lines = [f"[{name} · {quarter_label}(E) 분기 예상 영업이익 추적]"]
    if len(hist) <= 1:
        if first_v is not None:
            lines.append(f"• {hist[0]['d']}: {first_v:,.2f}억 (변동 없음)")
        else:
            lines.append("• 추적 데이터 없음")
    else:
        pct_txt = f"{total_pct:+.2f}%" if total_pct is not None else "-"
        lvl_txt = f" · {lvl_labels[level]}" if level in lvl_labels else ""
        lines.append(f"• 누적 변동률: {pct_txt}{lvl_txt} (상향 {up_cnt}회 ▲ / 하향 {down_cnt}회 ▼)")
        lines.append("────────────────────────")
        for idx_h, pt in enumerate(hist[-10:]):
            dt_s, v_s = pt.get("d", ""), pt.get("v")
            if v_s is None:
                continue
            if idx_h == 0 and len(hist) <= 10:
                lines.append(f"• {dt_s}: {v_s:,.2f}억 (기준)")
            else:
                orig_i = len(hist) - min(10, len(hist)) + idx_h
                pv = hist[orig_i - 1].get("v") if orig_i > 0 else None
                if pv is not None:
                    df = round(v_s - pv, 2)
                    arrow = "▲" if df > 0 else ("▼" if df < 0 else "=")
                    p_str = f" / {df / abs(pv) * 100:+.2f}%" if pv != 0 else ""
                    lines.append(f"• {dt_s}: {v_s:,.2f}억 ({arrow} {df:+,.2f}{p_str})")
                else:
                    lines.append(f"• {dt_s}: {v_s:,.2f}억")

    entry["upCount"] = up_cnt
    entry["downCount"] = down_cnt
    entry["firstVal"] = first_v
    entry["lastVal"] = last_v
    entry["totalPct"] = total_pct
    entry["netUp"] = net_up
    entry["level"] = level
    entry["memo"] = "\n".join(lines)
    return entry


def _load_op_estimate_history(here):
    p = os.path.join(here, "op_estimate_history.json")
    if os.path.isfile(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict) and isinstance(data.get("stocks"), dict):
                return data
        except Exception:
            pass
    cur_q = _current_quarter_ym()
    return {"quarter": cur_q, "col": f"OP_{cur_q}(E)", "updatedAt": "", "stocks": {}}


def _save_op_estimate_history(here, data):
    p = os.path.join(here, "op_estimate_history.json")
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def _ingest_snapshot_to_op_history(hist_data, rows_iter, snap_label, quarter_ym, op_col):
    """(code, name, val) 목록을 현재 분기 예상 영업이익 추적 이력에 반영합니다."""
    if hist_data.get("quarter") != quarter_ym:
        hist_data["quarter"] = quarter_ym
        hist_data["col"] = op_col
        hist_data["stocks"] = {}

    stocks = hist_data["stocks"]
    for cd, nm, val in rows_iter:
        if not cd or val is None:
            continue
        v_r = round(float(val), 2)
        ent = stocks.get(cd)
        if not ent:
            ent = {"name": nm, "history": [{"d": snap_label, "v": v_r}]}
            stocks[cd] = ent
        else:
            if nm:
                ent["name"] = nm
            h = ent.get("history") or []
            if not h:
                h.append({"d": snap_label, "v": v_r})
            else:
                last_pt = h[-1]
                if round(float(last_pt.get("v", 0)), 2) != v_r:
                    if last_pt.get("d") == snap_label:
                        last_pt["v"] = v_r
                    else:
                        h.append({"d": snap_label, "v": v_r})
            ent["history"] = h
        _recompute_op_history_entry(ent, quarter_ym)
    hist_data["updatedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    return hist_data


def update_op_estimate_history_from_df(here, df_all, snap_dt=None):
    """스크리너 실행 직후 DataFrame에서 현재 분기 예상 영업이익 변화를 추적·저장합니다."""
    if df_all is None or df_all.empty or "종목코드" not in df_all.columns:
        return _load_op_estimate_history(here)
    cols = df_all.columns.tolist()
    q_ym, op_col = _find_current_q_op_col(cols, snap_dt)
    if not op_col:
        return _load_op_estimate_history(here)

    hist_data = _load_op_estimate_history(here)
    snap_label = (snap_dt or datetime.now()).strftime("%Y-%m-%d %H:%M")
    rows_iter = []
    for _, row in df_all.iterrows():
        raw_cd = str(row.get("종목코드", "") or "").strip()
        if not raw_cd:
            continue
        cd = raw_cd.split(".")[0].zfill(6)
        nm = str(row.get("종목명", "") or "").strip()
        v = _num(row.get(op_col))
        if v is not None:
            rows_iter.append((cd, nm, v))

    _ingest_snapshot_to_op_history(hist_data, rows_iter, snap_label, q_ym, op_col)
    _save_op_estimate_history(here, hist_data)
    return hist_data


def sync_op_history_from_excels(here):
    """
    폴더 및 엑셀백업 내 모든 과거 통합시트 엑셀(.xlsx)을 시간순으로 스캔하여
    현재 분기 예상 영업이익 추적 이력(op_estimate_history.json)을 보강/최신화합니다.
    """
    cur_q = _current_quarter_ym()
    op_col = f"OP_{cur_q}(E)"
    hist_data = _load_op_estimate_history(here)
    if hist_data.get("quarter") != cur_q:
        hist_data = {"quarter": cur_q, "col": op_col, "updatedAt": "", "stocks": {}}

    excel_files = []
    seen_names = set()
    for d in [here, os.path.join(here, "엑셀백업")]:
        if not os.path.isdir(d):
            continue
        for fn in os.listdir(d):
            if fn.startswith("~$") or not fn.lower().endswith(".xlsx"):
                continue
            if "_통합시트_" in fn and fn not in seen_names:
                m_ts = re.search(r"(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})", fn)
                if m_ts:
                    ts_key = f"{m_ts.group(1)}-{m_ts.group(2)}-{m_ts.group(3)} {m_ts.group(4)}:{m_ts.group(5)}"
                    excel_files.append((ts_key, os.path.join(d, fn)))
                    seen_names.add(fn)

    if not excel_files:
        return hist_data

    # 이미 추적 기록이 충분하고 더 새 파일이 없는 경우 중복 스캔 방지 (항상 날짜 오름차순으로 병합)
    excel_files.sort(key=lambda x: x[0])
    # 과거 전체 파일에서 재구축할 때 기존 history의 더 오래된 기록은 보존
    for ts_label, fpath in excel_files:
        try:
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            ws = wb[wb.sheetnames[0]]
            it_rows = ws.iter_rows(values_only=True)
            header = [str(c or "").strip() for c in next(it_rows, [])]
            if "종목코드" not in header or op_col not in header:
                wb.close()
                continue
            c_idx = header.index("종목코드")
            n_idx = header.index("종목명") if "종목명" in header else -1
            o_idx = header.index(op_col)
            rows_iter = []
            for row in it_rows:
                if len(row) <= max(c_idx, o_idx):
                    continue
                raw_cd = str(row[c_idx] or "").strip()
                if not raw_cd:
                    continue
                cd = raw_cd.split(".")[0].zfill(6)
                nm = str(row[n_idx] or "").strip() if (n_idx >= 0 and len(row) > n_idx) else ""
                v = _num(row[o_idx])
                if v is not None:
                    rows_iter.append((cd, nm, v))
            wb.close()
            _ingest_snapshot_to_op_history(hist_data, rows_iter, ts_label, cur_q, op_col)
        except Exception:
            continue

    _save_op_estimate_history(here, hist_data)
    return hist_data


def _load_latest_prev_estimates(here, prefix):
    """
    폴더(here) 내 해당 시장(prefix)의 가장 최근 생성된 엑셀 파일을 찾아
    종목별 예상이익(OP_*(E), 연간OP_*(E)) 데이터를 {(종목코드, 컬럼명): 수치} 형태로 반환
    """
    candidates = []
    latest_fallback = None
    search_dirs = [here, os.path.join(here, "엑셀백업")]
    try:
        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for fn in os.listdir(d):
                if fn.startswith("~$") or not fn.lower().endswith(".xlsx"):
                    continue
                full_p = os.path.join(d, fn)
                if not os.path.isfile(full_p):
                    continue
                if fn.startswith(f"{prefix}_통합시트_") or fn.startswith(f"{prefix}_전체_"):
                    candidates.append(full_p)
                elif fn == f"{prefix}_latest.xlsx" and latest_fallback is None:
                    latest_fallback = full_p
    except Exception:
        return None, {}

    if not candidates and latest_fallback:
        candidates.append(latest_fallback)

    if not candidates:
        return None, {}

    # 파일명 타임스탬프 및 수정시간 기준 가장 최신 파일 선택
    candidates.sort(key=lambda p: (os.path.basename(p), os.path.getmtime(p)), reverse=True)
    latest_path = candidates[0]

    prev_map = {}
    try:
        prev_df = pd.read_excel(latest_path, sheet_name=0, dtype={"종목코드": str})
        if "종목코드" not in prev_df.columns:
            return os.path.basename(latest_path), {}
        est_cols = [
            c for c in prev_df.columns
            if (c.startswith("OP_") or c.startswith("연간OP_")) and c.endswith("(E)")
        ]
        for _, row in prev_df.iterrows():
            raw_cd = str(row.get("종목코드", "")).strip()
            if not raw_cd or raw_cd.lower() == "nan":
                continue
            cd = raw_cd.split(".")[0].zfill(6)
            for col in est_cols:
                v = _num(row.get(col))
                if v is not None:
                    prev_map[(cd, col)] = v
    except Exception as e:
        print(f"  [알림] 이전 엑셀 파일({os.path.basename(latest_path)}) 비교 데이터 로드 실패: {e}")
        return None, {}

    return os.path.basename(latest_path), prev_map


def run_market(market, sess, here):
    print(f"\n{'='*80}")
    print(f" ▶ [{market}] 시가총액 상위 {TOP_N}개 종목 스크리닝 시작")
    print(f"{'='*80}")

    universe, naver_quote_map = get_market_cap_top(market, TOP_N)
    n = len(universe)
    if not n:
        raise RuntimeError(f"{market} 종목 리스트를 가져오지 못했습니다.")
    print(f"  총 {n}개 종목 대상 스크리닝 시작 (시총 1~5위: {', '.join(nm for _, nm in universe[:5])})\n")

    print(f"[2·3단계] [{market}] 분기·연간 데이터 및 주가(현재가/오늘등락율/52주최고가) 수집 + 필터...")
    rows, errors = [], []
    for i, (code, name) in enumerate(universe, 1):
        fallback_quote = naver_quote_map.get(code, (None, None))
        try:
            q = _fetch_snp_q(code, sess)
            time.sleep(0.3)
            a = _fetch(code, sess, "Y")
            time.sleep(0.3)
            price_info = get_price_info_from_snapshot(code, sess, naver_fallback=fallback_quote)
            
            rec, R, YY = build_record(code, name, q, a, price_info)
        except Exception as e:
            rec = {
                "종목코드": code,
                "종목명": name,
                "섹터1": "",
                "섹터2": "",
                "현재가": fallback_quote[0],
                "오늘등락율(%)": fallback_quote[1],
                "통과": False,
                "E3반등": False,
                "통과+E3반등": False,
                "연속이익성장(연간+분기)": False,
                "연속성장구간": ""
            }
            errors.append((code, name, str(e)[:80]))
        rows.append(rec)
        
        tags = []
        if rec.get("통과+E3반등"):
            tags.append("★★통과+E3반등")
        elif rec.get("통과"):
            tags.append("★통과")
        elif rec.get("E3반등"):
            tags.append("E3반등")
        if rec.get("연속이익성장(연간+분기)"):
            tags.append(f"◆연속성장({rec.get('연속성장구간')})")
        status_str = " / ".join(tags) if tags else "일반"
        
        rt_val = rec.get("오늘등락율(%)")
        rt_str = f"{rt_val:+.2f}%" if isinstance(rt_val, (int, float)) else "None"
        sec_str = f"{rec.get('섹터1', '')}/{rec.get('섹터2', '')}".strip("/")
        print(f"  [{market}] ({i}/{n}) {name} ({sec_str}) [{status_str}] 현재가: {rec.get('현재가')}원 (오늘등락율: {rt_str}), 52주최고가: {rec.get('52주최고가')}원 (하락률: {rec.get('52주최고가대비하락률(%)')}%)")
        time.sleep(DELAY)

    full = pd.DataFrame(rows)
    
    # OP_YYYY/09 확정 실적 <-> OP_YYYY/09(E) 예상실적 양방향 보완
    full = fill_missing_base_from_estimate(full)

    stamp = datetime.now().strftime("%Y%m%d_%H%M")

    def custom_sort(cols):
        fixed_front = [
            '종목코드', '종목명', '섹터1', '섹터2', '현재가', '오늘등락율(%)', '52주최고가',
            '52주최고가대비하락률(%)', '52주최고가대비30%이상하락',
            '통과', 'E3반등', '통과+E3반등', '연속이익성장(연간+분기)', '연속성장구간', '성장조건', 'E3반등사유'
        ]
        # 그룹 A: 분기별 영업이익 상세(접힘) + 요약(표시)
        op_detail = sorted([c for c in cols if c.startswith("OP_")])
        op_summary = [c for c in ['분기OP_연속성장', 'OP성장률_최종대비E0(%)'] if c in cols]

        # 그룹 B: 분기별 YoY 상세(접힘) + 가속폭 요약(표시)
        yoy_detail = sorted([c for c in cols if c.startswith("YoY_")])
        yoy_summary = [c for c in ['가속폭_E1-기준(%p)', '가속폭_E2-기준(%p)', '가속폭_E3-E2(%p)'] if c in cols]

        # 그룹 C: 분기별 영업이익률 상세(접힘) + 분기마진 요약(표시)
        qm_detail = sorted([c for c in cols if c.startswith("분기영업이익률_")])
        qm_summary = [c for c in ['분기마진_기준<E1<E2'] if c in cols]

        # 그룹 D: 연간 영업이익·이익률 상세(접힘) + 연간 요약(표시)
        ann_summary_names = ('연간OP_성장률(%)', '연간OP_차년>당해', '연간마진_개선폭(%p)', '연간마진_차년>당해')
        ann_detail = sorted([
            c for c in cols
            if (c.startswith("연간OP_") and c not in ann_summary_names) or c.startswith("연간영업이익률_")
        ])
        ann_summary = [c for c in ann_summary_names if c in cols]

        ordered = (
            [c for c in fixed_front if c in cols]
            + op_detail + op_summary
            + yoy_detail + yoy_summary
            + qm_detail + qm_summary
            + ann_detail + ann_summary
        )
        remaining = sorted([c for c in cols if c not in ordered])
        return ordered + remaining

    def apply_excel_ux(ws, cols, prev_map=None, prev_filename=None, op_hist_stocks=None):
        """
        각 시트에:
        1) 자동 필터
        2) 틀 고정(1행 + 종목코드~섹터2)
        3) 4대 상세 열 그룹화([+] 접기)
        4) 1행 주요 제목 셀에 스크리닝 조건/지표 설명 메모(Comment) 삽입
        5) 직전 엑셀 파일 대비 예상이익(OP_*(E), 연간OP_*(E)) 상향 시 빨간색, 하향 시 파란색 셀 배경 채우기 + 변동 내역 메모 삽입
        6) 분기별·연간 영업이익(OP), YoY, 영업이익률 등 연도별 구간을 진한 세로선 테두리로 구분
        """
        from openpyxl.comments import Comment
        from openpyxl.styles import Border, PatternFill, Side
        from openpyxl.utils import get_column_letter

        # 1) 1행 전체 제목줄 자동 필터
        if ws.max_column >= 1 and ws.max_row >= 1:
            ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 1)}"

        # 2) 틀 고정: 상단 1행(제목) + 좌측 A~D열(종목코드, 종목명, 섹터1, 섹터2) 고정
        ws.freeze_panes = "E2"

        # 3) 오른쪽 상세 데이터 4개 그룹 접기/펼치기([+] / [-]) 설정
        ws.sheet_properties.outlinePr.summaryRight = True
        ann_summary_names = ('연간OP_성장률(%)', '연간OP_차년>당해', '연간마진_개선폭(%p)', '연간마진_차년>당해')
        group_predicates = [
            lambda c: c.startswith("OP_"),                                                          # 그룹 A: 분기 OP
            lambda c: c.startswith("YoY_"),                                                         # 그룹 B: 분기 YoY
            lambda c: c.startswith("분기영업이익률_"),                                                # 그룹 C: 분기 이익률
            lambda c: (c.startswith("연간OP_") and c not in ann_summary_names) or c.startswith("연간영업이익률_"),  # 그룹 D: 연간 상세
        ]
        for pred in group_predicates:
            idxs = [i + 1 for i, c in enumerate(cols) if pred(c)]
            if idxs:
                start_col = get_column_letter(min(idxs))
                end_col = get_column_letter(max(idxs))
                ws.column_dimensions.group(start_col, end_col, outline_level=1, hidden=True)

        # 4) 1행 주요 제목 셀에 조건/지표 설명 메모(Comment) 삽입
        header_comments = {
            "섹터1": "에프앤가이드 거래소 업종 분류\n(예: 전기전자, 제약, 운송장비부품 등)",
            "섹터2": "에프앤가이드 WI26(WICS) 세부 섹터 분류\n(예: 반도체, 건강관리, 자동차 등)",
            "52주최고가대비하락률(%)": "(52주최고가 - 현재가) / 52주최고가 × 100 (%)",
            "52주최고가대비30%이상하락": "52주 최고가 대비 현재가가 30% 이상 하락(조정)한 종목 여부",
            "통과": (
                "[메인 스크리닝 통과 조건 (①+②+③ 동시 충족)]\n"
                "① 분기 영업이익 YoY 가속(E1>E0 또는 E2>E0) 또는 고성장(E1·E2 모두 ≥50%)\n"
                "   * 단, E1·E2 YoY가 둘 다 마이너스(<0)면 제외\n"
                "② 연간 영업이익률 개선 (차년도 > 당해연도)\n"
                "③ 분기 영업이익률 연속 개선 (E0 < E1 < E2)"
            ),
            "E3반등": (
                "[E3(3번째 추정 분기) 반등 조건]\n"
                "• E3 영업이익률 > E2 영업이익률 (분기 마진 상승) 이면서\n"
                "• E3 YoY가 '흑자전환'이거나, E2 대비 +20%p 이상 급반등(이면서 E3 YoY ≥ 20%)"
            ),
            "통과+E3반등": (
                "[통과 + E3 추가반등]\n"
                "메인 '통과' 조건(①+②+③)을 만족하면서 동시에 'E3반등' 조건까지 충족한 핵심 종목"
            ),
            "연속이익성장(연간+분기)": (
                "[연속이익성장 (연간+분기) 통과 조건]\n"
                "① 연간 영업이익 성장 (차년도 > 당해연도 & 차년도 > 0)\n"
                "② 연간 영업이익률 상승 (차년도 > 당해연도)\n"
                "③ 분기 영업이익 연속 증가:\n"
                "   - E3 예측치 존재 시: E0 < E1 < E2 < E3 (4분기 연속)\n"
                "   - E3 미집계(빈칸) 시: E0 < E1 < E2 (3분기 연속)"
            ),
            "연속성장구간": (
                "• E0~E3: 4개 분기 연속 영업이익 증가\n"
                "• E0~E2(E3미집계): E3 컨센서스 미집계 상태에서 E0~E2 3개 분기 연속 영업이익 증가"
            ),
            "성장조건": (
                "• 가속: E1 또는 E2의 영업이익 YoY가 기준 분기(E0) YoY보다 상승\n"
                "• 고성장: E1과 E2의 영업이익 YoY가 모두 50% 이상"
            ),
            "E3반등사유": (
                "• 흑자전환: E3 분기 전년동기대비 흑자전환 + 분기마진 상승\n"
                "• 급반등: E3 YoY가 E2 YoY 대비 +20%p 이상 상승(최소 20% 이상) + 분기마진 상승"
            ),
            "분기OP_연속성장": "분기 영업이익(억원)이 매 분기 연속 증가(E0 < E1 < E2 [< E3])했는지 여부\n(상단 [+] 버튼으로 분기별 영업이익 상세 열을 펼칠 수 있습니다)",
            "OP성장률_최종대비E0(%)": "기준 분기(E0) 영업이익 대비 최종 추정 분기(E3 또는 E2) 영업이익의 성장률(%)",
            "가속폭_E1-기준(%p)": "E1 분기 영업이익 YoY(%) - 기준 분기(E0) 영업이익 YoY(%) (%p 차이)\n(상단 [+] 버튼으로 분기별 YoY 상세 열을 펼칠 수 있습니다)",
            "가속폭_E2-기준(%p)": "E2 분기 영업이익 YoY(%) - 기준 분기(E0) 영업이익 YoY(%) (%p 차이)",
            "가속폭_E3-E2(%p)": "E3 분기 영업이익 YoY(%) - E2 분기 영업이익 YoY(%) (%p 차이)",
            "분기마진_기준<E1<E2": "분기 영업이익률이 기준(E0) < E1 < E2 순서로 연속 상승했는지 여부\n(상단 [+] 버튼으로 분기별 영업이익률 상세 열을 펼칠 수 있습니다)",
            "연간OP_성장률(%)": "당해연도(E) 대비 차년도(E) 연간 영업이익 성장률(%)\n(상단 [+] 버튼으로 연간 상세 열을 펼칠 수 있습니다)",
            "연간OP_차년>당해": "차년도(E) 연간 영업이익이 당해연도(E)보다 크고 흑자(>0)인지 여부",
            "연간마진_개선폭(%p)": "차년도(E) 연간 영업이익률(%) - 당해연도(E) 연간 영업이익률(%) (%p 차이)",
            "연간마진_차년>당해": "차년도(E) 연간 영업이익률이 당해연도(E)보다 상승했는지 여부",
        }
        for c_idx, col_name in enumerate(cols, 1):
            desc = header_comments.get(col_name)
            if not desc:
                if (col_name.startswith("OP_") or col_name.startswith("연간OP_")) and col_name.endswith("(E)"):
                    desc = (
                        f"[{col_name}]\n컨센서스 예상 영업이익(억원)\n"
                        "• 직전 대비 상향(▲): <5% 연빨강 / 5~10% 진빨강 / 10%↑ 매우진빨강\n"
                        "• 직전 대비 하향(▼): <5% 연파랑 / 5~10% 진파랑 / 10%↓ 매우진파랑"
                    )
            if desc:
                line_cnt = desc.count("\n") + 1
                cm = Comment(desc, "Screener")
                cm.width = 380
                cm.height = max(120, (line_cnt + 2) * 26)
                ws.cell(row=1, column=c_idx).comment = cm

        # 5) 직전 엑셀 파일과 예상이익(OP_*(E), 연간OP_*(E)) 비교하여 3단계 강도별 셀 색상 채우기 + 변동 내역 메모 삽입
        up_cnt, down_cnt = 0, 0
        if prev_map and "종목코드" in cols and ws.max_row >= 2:
            from openpyxl.styles import Font
            fill_up_1 = PatternFill(start_color="FFE0E4", end_color="FFE0E4", fill_type="solid")    # 상향 <5%: 연빨강
            fill_up_2 = PatternFill(start_color="FF9AA8", end_color="FF9AA8", fill_type="solid")    # 상향 5~10%: 중간 진빨강
            fill_up_3 = PatternFill(start_color="E11D48", end_color="E11D48", fill_type="solid")    # 상향 10%↑: 아주 진빨강
            fill_down_1 = PatternFill(start_color="D6E9FA", end_color="D6E9FA", fill_type="solid")  # 하향 <5%: 연파랑
            fill_down_2 = PatternFill(start_color="7EB8F0", end_color="7EB8F0", fill_type="solid")  # 하향 5~10%: 중간 진파랑
            fill_down_3 = PatternFill(start_color="1D4ED8", end_color="1D4ED8", fill_type="solid")  # 하향 10%↓: 아주 진파랑
            font_up_2 = Font(color="7F1D1D", bold=True)
            font_down_2 = Font(color="172554", bold=True)
            font_strong = Font(color="FFFFFF", bold=True)

            code_col_idx = cols.index("종목코드") + 1
            est_col_pairs = [
                (i + 1, c) for i, c in enumerate(cols)
                if (c.startswith("OP_") or c.startswith("연간OP_")) and c.endswith("(E)")
            ]
            ref_label = f"\n  (비교: {prev_filename})" if prev_filename else ""
            for r_idx in range(2, ws.max_row + 1):
                raw_cd = str(ws.cell(row=r_idx, column=code_col_idx).value or "").strip()
                if not raw_cd:
                    continue
                cd = raw_cd.split(".")[0].zfill(6)
                for c_idx, col_name in est_col_pairs:
                    old_val = prev_map.get((cd, col_name))
                    if old_val is None:
                        continue
                    cell_obj = ws.cell(row=r_idx, column=c_idx)
                    new_val = _num(cell_obj.value)
                    if new_val is None:
                        continue
                    old_r, new_r = round(old_val, 2), round(new_val, 2)
                    if new_r != old_r:
                        diff = new_r - old_r
                        pct_val = (diff / abs(old_r) * 100.0) if old_r != 0 else None
                        abs_pct = abs(pct_val) if pct_val is not None else 0.0
                        pct_str = f" ({pct_val:+.2f}%)" if pct_val is not None else ""
                        if new_r > old_r:
                            if abs_pct >= 10.0:
                                cell_obj.fill = fill_up_3
                                cell_obj.font = font_strong
                            elif abs_pct >= 5.0:
                                cell_obj.fill = fill_up_2
                                cell_obj.font = font_up_2
                            else:
                                cell_obj.fill = fill_up_1
                            up_cnt += 1
                            head = f"[예상이익 상향 ▲]{ref_label}"
                        else:
                            if abs_pct >= 10.0:
                                cell_obj.fill = fill_down_3
                                cell_obj.font = font_strong
                            elif abs_pct >= 5.0:
                                cell_obj.fill = fill_down_2
                                cell_obj.font = font_down_2
                            else:
                                cell_obj.fill = fill_down_1
                            down_cnt += 1
                            head = f"[예상이익 하향 ▼]{ref_label}"
                        memo_txt = (
                            f"{head}\n"
                            f"• 직전: {old_r:,.2f}\n"
                            f"• 현재: {new_r:,.2f}\n"
                            f"• 증감: {diff:+,.2f}{pct_str}\n"
                        )
                        cm = Comment(memo_txt, "Screener")
                        cm.width = 360
                        cm.height = 160
                        cell_obj.comment = cm

        # 5-2) 현재 분기 예상 영업이익 누적 상향 이력 -> 종목명 셀 단계별 진한 색상 + 변동 이력 메모
        if op_hist_stocks and "종목코드" in cols and "종목명" in cols and ws.max_row >= 2:
            from openpyxl.styles import Font as _Font
            name_fills = {
                1: PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid"),
                2: PatternFill(start_color="FECDD3", end_color="FECDD3", fill_type="solid"),
                3: PatternFill(start_color="FB7185", end_color="FB7185", fill_type="solid"),
                4: PatternFill(start_color="DC2626", end_color="DC2626", fill_type="solid"),
            }
            name_fonts = {
                1: _Font(color="9F1239", bold=True),
                2: _Font(color="881337", bold=True),
                3: _Font(color="4C0519", bold=True),
                4: _Font(color="FFFFFF", bold=True),
            }
            cd_idx_e = cols.index("종목코드") + 1
            nm_idx_e = cols.index("종목명") + 1
            for r_idx in range(2, ws.max_row + 1):
                raw_cd = str(ws.cell(row=r_idx, column=cd_idx_e).value or "").strip()
                if not raw_cd:
                    continue
                cd = raw_cd.split(".")[0].zfill(6)
                ent = op_hist_stocks.get(cd)
                if not ent:
                    continue
                nm_cell = ws.cell(row=r_idx, column=nm_idx_e)
                lvl = int(ent.get("level") or 0)
                if lvl >= 1:
                    nm_cell.fill = name_fills[min(4, lvl)]
                    nm_cell.font = name_fonts[min(4, lvl)]
                memo_h = ent.get("memo") or ""
                if memo_h and len(ent.get("history") or []) >= 2:
                    cm_h = Comment(memo_h, "OP Tracker")
                    cm_h.width = 400
                    cm_h.height = max(140, (memo_h.count("\n") + 2) * 22)
                    nm_cell.comment = cm_h

        # 6) 분기·연간 이익(OP), YoY, 영업이익률 등 연도(YYYY)가 바뀌는 경계마다 진한 세로선(medium) 테두리 적용
        def _metric_year_key(col_name):
            m = re.match(r"^(OP_|YoY_|분기영업이익률_|연간OP_|연간영업이익률_)(\d{4})/", str(col_name))
            return (m.group(1), m.group(2)) if m else None

        col_keys = [_metric_year_key(c) for c in cols]
        thick_side = Side(style="medium", color="000000")
        border_targets = []  # [(c_idx, need_left, need_right)]
        for idx, key in enumerate(col_keys):
            if key is None:
                continue
            prev_key = col_keys[idx - 1] if idx > 0 else None
            next_key = col_keys[idx + 1] if idx + 1 < len(col_keys) else None
            need_left = (prev_key != key)
            need_right = (next_key != key)
            if need_left or need_right:
                border_targets.append((idx + 1, need_left, need_right))

        if border_targets and ws.max_row >= 1:
            for r_idx in range(1, ws.max_row + 1):
                for c_idx, need_left, need_right in border_targets:
                    cell_obj = ws.cell(row=r_idx, column=c_idx)
                    old_b = cell_obj.border
                    cell_obj.border = Border(
                        left=thick_side if need_left else (old_b.left if old_b else None),
                        right=thick_side if need_right else (old_b.right if old_b else None),
                        top=old_b.top if old_b else None,
                        bottom=old_b.bottom if old_b else None,
                    )

        return up_cnt, down_cnt
    
    full = full[custom_sort(full.columns.tolist())]

    # 1. 전체 종목: 시가총액 순서(1위~200위) 그대로 유지
    df_all = full.copy().reset_index(drop=True)
    
    # 2. 통과 종목: 가속폭 내림차순 정렬
    df_pass = full[full["통과"] == True].copy()
    acc_col = "가속폭_E1-기준(%p)"
    if not df_pass.empty and acc_col in df_pass.columns:
        df_pass = df_pass.sort_values(acc_col, ascending=False, na_position="last")
    df_pass = df_pass.reset_index(drop=True)
    
    # 3. E3반등 후보 (미통과 종목 중 E3 반등): 흑자전환 우선, 가속폭_E3-E2 내림차순 정렬
    if "E3반등" in full.columns:
        df_reb = full[(full["E3반등"] == True) & (full["통과"] != True)].copy()
        if not df_reb.empty and "E3반등사유" in df_reb.columns and "가속폭_E3-E2(%p)" in df_reb.columns:
            df_reb["_turn"] = (df_reb["E3반등사유"] == "흑자전환").astype(int)
            df_reb = df_reb.sort_values(["_turn", "가속폭_E3-E2(%p)"],
                                        ascending=[False, False], na_position="last")
            df_reb = df_reb.drop(columns="_turn").reset_index(drop=True)
    else:
        df_reb = full.iloc[0:0]

    # 4. 통과 또는 반등 종목 중 30% 이상 하락한 종목 필터
    passed_or_reb = full[(full["통과"] == True) | (full["E3반등"] == True)].copy()
    target_down30 = passed_or_reb[passed_or_reb["52주최고가대비30%이상하락"] == True].copy()

    if not target_down30.empty and "52주최고가대비하락률(%)" in target_down30.columns:
        target_down30 = target_down30.sort_values("52주최고가대비하락률(%)", ascending=False).reset_index(drop=True)

    # 5. 통과한 종목 중에서 E3에서 추가 반등까지 한 기업 리스트 (통과 == True & E3반등 == True)
    df_pass_and_reb = full[(full["통과"] == True) & (full["E3반등"] == True)].copy()
    if not df_pass_and_reb.empty and "가속폭_E3-E2(%p)" in df_pass_and_reb.columns:
        df_pass_and_reb = df_pass_and_reb.sort_values("가속폭_E3-E2(%p)", ascending=False, na_position="last")
    df_pass_and_reb = df_pass_and_reb.reset_index(drop=True)

    # 6. 연간 이익 성장 + 연간 마진 상승 + 분기 이익(E0~E3 또는 E3미집계시 E0~E2) 연속 성장 기업
    df_seq_growth = full[full["연속이익성장(연간+분기)"] == True].copy()
    if not df_seq_growth.empty:
        df_seq_growth["_is_e3"] = (df_seq_growth["연속성장구간"] == "E0~E3").astype(int)
        sort_cols = ["_is_e3"] + (["OP성장률_최종대비E0(%)"] if "OP성장률_최종대비E0(%)" in df_seq_growth.columns else [])
        df_seq_growth = df_seq_growth.sort_values(sort_cols, ascending=[False] * len(sort_cols), na_position="last")
        df_seq_growth = df_seq_growth.drop(columns="_is_e3").reset_index(drop=True)

    # =========================================================================
    # ★ 8개 시트가 하나로 묶인 통합 엑셀(.xlsx) 파일 1개만 저장
    #   (7_예상이익상향, 8_예상이익하향 시트 포함)
    # =========================================================================
    prefix = f"{market.lower()}_top{TOP_N}"
    prev_filename, prev_est_map = _load_latest_prev_estimates(here, prefix)
    p_xlsx_multi = os.path.join(here, f"{prefix}_통합시트_{stamp}.xlsx")

    # 7·8. 직전 파일 대비 예상이익(OP_*(E), 연간OP_*(E)) 상향 / 하향 기업 추출
    est_compare_cols = [
        c for c in full.columns
        if (c.startswith("OP_") or c.startswith("연간OP_")) and c.endswith("(E)")
    ]
    up_cnt_by_idx, down_cnt_by_idx = {}, {}
    if prev_est_map and "종목코드" in full.columns and est_compare_cols:
        for row_idx, row in full.iterrows():
            raw_cd = str(row.get("종목코드", "") or "").strip()
            if not raw_cd:
                continue
            cd = raw_cd.split(".")[0].zfill(6)
            u_c, d_c = 0, 0
            for col in est_compare_cols:
                old_v = prev_est_map.get((cd, col))
                new_v = _num(row.get(col))
                if old_v is None or new_v is None:
                    continue
                old_r, new_r = round(old_v, 2), round(new_v, 2)
                if new_r > old_r:
                    u_c += 1
                elif new_r < old_r:
                    d_c += 1
            if u_c > 0:
                up_cnt_by_idx[row_idx] = u_c
            if d_c > 0:
                down_cnt_by_idx[row_idx] = d_c

    if up_cnt_by_idx:
        up_order = sorted(up_cnt_by_idx.keys(), key=lambda i: (-up_cnt_by_idx[i], i))
        df_est_up = full.loc[up_order].copy().reset_index(drop=True)
    else:
        df_est_up = full.iloc[0:0].copy()

    if down_cnt_by_idx:
        down_order = sorted(down_cnt_by_idx.keys(), key=lambda i: (-down_cnt_by_idx[i], i))
        df_est_down = full.loc[down_order].copy().reset_index(drop=True)
    else:
        df_est_down = full.iloc[0:0].copy()

    # 현재 분기 예상 영업이익 누적 추적 이력 업데이트
    op_hist_data = update_op_estimate_history_from_df(here, df_all)
    op_hist_stocks = op_hist_data.get("stocks", {})

    all_up_cnt, all_down_cnt = 0, 0
    with pd.ExcelWriter(p_xlsx_multi, engine="openpyxl") as writer:
        sheets_to_write = [
            ("1_전체", df_all),
            ("2_통과", df_pass),
            ("3_E3반등", df_reb),
            ("4_30퍼하락", target_down30),
            ("5_통과후E3반등", df_pass_and_reb),
            ("6_연속이익성장(E3및E2)", df_seq_growth),
            ("7_예상이익상향", df_est_up),
            ("8_예상이익하향", df_est_down),
        ]
        for s_name, s_df in sheets_to_write:
            s_df.to_excel(writer, sheet_name=s_name, index=False)
            u_cnt, d_cnt = apply_excel_ux(
                writer.sheets[s_name],
                s_df.columns.tolist(),
                prev_map=prev_est_map,
                prev_filename=prev_filename,
                op_hist_stocks=op_hist_stocks
            )
            if s_name == "1_전체":
                all_up_cnt, all_down_cnt = u_cnt, d_cnt

    print(f"\n[{market} 통합 엑셀(.xlsx) 저장 완료: {os.path.basename(p_xlsx_multi)}]")
    if prev_filename:
        print(f"  - 직전 파일({prev_filename}) 대비 예상이익 변동(색채우기+메모): 상향(빨간색) {all_up_cnt}셀 / 하향(파란색) {all_down_cnt}셀")
    else:
        print("  - 비교할 이전 엑셀 파일이 없어 예상이익 변동 색채우기는 생략되었습니다.")
    print(f"  - 시트 1 [1_전체] 전체 종목 (시총순): {len(df_all)}개")
    print(f"  - 시트 2 [2_통과] 통과 종목: {len(df_pass)}개")
    print(f"  - 시트 3 [3_E3반등] E3반등 후보 (미통과 중): {len(df_reb)}개")
    print(f"  - 시트 4 [4_30퍼하락] 통과/반등 중 30% 하락: {len(target_down30)}개")
    print(f"  - 시트 5 [5_통과후E3반등] 통과 + E3 추가반등 기업: {len(df_pass_and_reb)}개")
    if not df_pass_and_reb.empty:
        print(f"     ★ 해당 기업: {', '.join(df_pass_and_reb['종목명'].tolist())}")
    print(f"  - 시트 6 [6_연속이익성장(E3및E2)] 연간이익·마진성장 + 분기연속성장: {len(df_seq_growth)}개")
    if not df_seq_growth.empty:
        e3_list = df_seq_growth[df_seq_growth["연속성장구간"] == "E0~E3"]["종목명"].tolist()
        e2_list = df_seq_growth[df_seq_growth["연속성장구간"] == "E0~E2(E3미집계)"]["종목명"].tolist()
        if e3_list:
            print(f"     ◆ [E0~E3 4분기 연속성장 ({len(e3_list)}개)]: {', '.join(e3_list)}")
        if e2_list:
            print(f"     ◇ [E0~E2 3분기 연속성장·E3미집계 ({len(e2_list)}개)]: {', '.join(e2_list)}")
    print(f"  - 시트 7 [7_예상이익상향] 예상이익 상향(▲) 기업: {len(df_est_up)}개")
    print(f"  - 시트 8 [8_예상이익하향] 예상이익 하향(▼) 기업: {len(df_est_down)}개")
    
    if errors:
        print(f"\n      [{market} 조회 실패 {len(errors)}종목]:")
        for c, nm, e in errors[:10]:
            print(f"        {c} {nm}: {e}")

    return p_xlsx_multi


def _clean_json_val(v):
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        if pd.isna(v):
            return None
        return v
    s = str(v).strip()
    if s in ("", "None", "nan", "NaN"):
        return None
    if s.upper() == "TRUE":
        return True
    if s.upper() == "FALSE":
        return False
    return s


def _extract_market_excel_for_web(here, market):
    prefix = f"{market.lower()}_top{TOP_N}"
    candidates = []
    latest_copy_path = os.path.join(here, f"{prefix}_latest.xlsx")
    search_dirs = [here, os.path.join(here, "엑셀백업")]
    try:
        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for fn in os.listdir(d):
                if fn.startswith("~$") or not fn.lower().endswith(".xlsx"):
                    continue
                if fn.startswith(f"{prefix}_통합시트_") or fn.startswith(f"{prefix}_전체_"):
                    full_p = os.path.join(d, fn)
                    if os.path.isfile(full_p):
                        candidates.append(full_p)
    except Exception:
        pass

    if candidates:
        candidates.sort(key=lambda p: (os.path.basename(p), os.path.getmtime(p)), reverse=True)
        newest_cand = candidates[0]
        if os.path.isfile(latest_copy_path) and (os.path.getmtime(latest_copy_path) > os.path.getmtime(newest_cand) + 10):
            src_path = latest_copy_path
        else:
            src_path = newest_cand
            try:
                if os.path.abspath(src_path) != os.path.abspath(latest_copy_path):
                    shutil.copy2(src_path, latest_copy_path)
            except Exception:
                pass
    elif os.path.isfile(latest_copy_path):
        src_path = latest_copy_path
    else:
        return None

    base_fn = os.path.basename(src_path)
    m_ts = re.search(r"(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})", base_fn)
    if m_ts:
        updated_at = f"{m_ts.group(1)}-{m_ts.group(2)}-{m_ts.group(3)} {m_ts.group(4)}:{m_ts.group(5)}"
    else:
        updated_at = datetime.fromtimestamp(os.path.getmtime(src_path)).strftime("%Y-%m-%d %H:%M")

    wb = openpyxl.load_workbook(src_path, data_only=True)
    if not wb.sheetnames:
        return None

    ann_summary_names = ('연간OP_성장률(%)', '연간OP_차년>당해', '연간마진_개선폭(%p)', '연간마진_차년>당해')

    def get_col_group(c_name):
        if c_name.startswith("OP_"):
            return "A"
        if c_name.startswith("YoY_"):
            return "B"
        if c_name.startswith("가속폭_"):
            return "E"
        if c_name.startswith("분기영업이익률_"):
            return "C"
        if (c_name.startswith("연간OP_") and c_name not in ann_summary_names) or c_name.startswith("연간영업이익률_"):
            return "D"
        return ""

    def get_year_key(c_name):
        m = re.match(r"^(OP_|YoY_|분기영업이익률_|연간OP_|연간영업이익률_)(\d{4})/", str(c_name))
        return (m.group(1), m.group(2)) if m else None

    ws0 = wb[wb.sheetnames[0]]
    raw_cols = [str(ws0.cell(1, c).value or "").strip() for c in range(1, ws0.max_column + 1)]
    year_keys = [get_year_key(c) for c in raw_cols]

    columns_meta = []
    for idx, c_name in enumerate(raw_cols):
        cell_h = ws0.cell(1, idx + 1)
        h_comment = cell_h.comment.text.strip() if (cell_h.comment and cell_h.comment.text) else ""
        yk = year_keys[idx]
        prev_yk = year_keys[idx - 1] if idx > 0 else None
        next_yk = year_keys[idx + 1] if idx + 1 < len(year_keys) else None
        columns_meta.append({
            "name": c_name,
            "group": get_col_group(c_name),
            "comment": h_comment,
            "borderL": bool(yk is not None and prev_yk != yk),
            "borderR": bool(yk is not None and next_yk != yk),
        })

    code_col_idx = raw_cols.index("종목코드") if "종목코드" in raw_cols else -1
    sheets_data = []
    total_up, total_down = 0, 0

    for s_name in wb.sheetnames:
        ws = wb[s_name]
        rows_list = []
        s_up, s_down = 0, 0
        for r_idx in range(2, ws.max_row + 1):
            first_val = ws.cell(r_idx, 1).value
            second_val = ws.cell(r_idx, 2).value if ws.max_column >= 2 else None
            if first_val is None and second_val is None:
                continue
            vals = []
            tones = {}
            notes = {}
            for c_idx in range(1, len(raw_cols) + 1):
                cell = ws.cell(r_idx, c_idx)
                v = _clean_json_val(cell.value)
                if (c_idx - 1) == code_col_idx and v is not None:
                    v = str(v).split(".")[0].zfill(6)
                vals.append(v)

                note_txt = cell.comment.text.strip() if (cell.comment and cell.comment.text) else ""
                if note_txt:
                    notes[str(c_idx - 1)] = note_txt

                rgb = ""
                try:
                    if cell.fill and cell.fill.fill_type == "solid" and cell.fill.start_color:
                        rgb = str(cell.fill.start_color.rgb or "").upper()
                except Exception:
                    rgb = ""

                direction = ""
                tier_from_rgb = 1
                if rgb.endswith("E11D48"):
                    direction, tier_from_rgb = "up", 3
                elif rgb.endswith("FF9AA8"):
                    direction, tier_from_rgb = "up", 2
                elif rgb.endswith("FFE0E4") or rgb.endswith("FFC7CE"):
                    direction, tier_from_rgb = "up", 1
                elif rgb.endswith("1D4ED8"):
                    direction, tier_from_rgb = "down", 3
                elif rgb.endswith("7EB8F0"):
                    direction, tier_from_rgb = "down", 2
                elif rgb.endswith("D6E9FA") or rgb.endswith("BDD7EE"):
                    direction, tier_from_rgb = "down", 1

                if direction:
                    tier = tier_from_rgb
                    if note_txt:
                        m_pct = re.search(r"\(([+-]?\d+(?:\.\d+)?)%\)", note_txt)
                        if m_pct:
                            abs_p = abs(float(m_pct.group(1)))
                            if abs_p >= 10.0:
                                tier = 3
                            elif abs_p >= 5.0:
                                tier = 2
                            else:
                                tier = 1
                    tones[str(c_idx - 1)] = f"{direction}{tier}"
                    if direction == "up":
                        s_up += 1
                    else:
                        s_down += 1

            rows_list.append({
                "v": vals,
                "t": tones if tones else None,
                "n": notes if notes else None,
            })

        if s_name == "1_전체":
            total_up, total_down = s_up, s_down

        sheets_data.append({
            "name": s_name,
            "count": len(rows_list),
            "rows": rows_list,
        })

    # 기존 6시트 엑셀 파일이어도 웹페이지에서 '7_예상이익상향', '8_예상이익하향' 탭을 자동 생성
    existing_names = {s["name"] for s in sheets_data}
    all_sheet = next((s for s in sheets_data if s["name"] == "1_전체"), sheets_data[0] if sheets_data else None)
    if all_sheet:
        if "7_예상이익상향" not in existing_names:
            up_rows = [
                r for r in all_sheet["rows"]
                if r.get("t") and any(v.startswith("up") for v in r["t"].values())
            ]
            up_rows.sort(key=lambda r: -sum(1 for v in (r.get("t") or {}).values() if v.startswith("up")))
            sheets_data.append({
                "name": "7_예상이익상향",
                "count": len(up_rows),
                "rows": up_rows,
            })
        if "8_예상이익하향" not in existing_names:
            down_rows = [
                r for r in all_sheet["rows"]
                if r.get("t") and any(v.startswith("down") for v in r["t"].values())
            ]
            down_rows.sort(key=lambda r: -sum(1 for v in (r.get("t") or {}).values() if v.startswith("down")))
            sheets_data.append({
                "name": "8_예상이익하향",
                "count": len(down_rows),
                "rows": down_rows,
            })

    return {
        "market": market,
        "filename": base_fn,
        "downloadFile": f"{prefix}_latest.xlsx",
        "updatedAt": updated_at,
        "upCount": total_up,
        "downCount": total_down,
        "columns": columns_meta,
        "sheets": sheets_data,
    }


def _summarize_pdf_text(text, title=""):
    """PDF 추출 텍스트에서 투자포인트·핵심 문장 3~4개를 자동 정리합니다."""
    if not text:
        return [title] if title else ["PDF 원문 텍스트를 추출하지 못했습니다. 원문 PDF 보기를 이용해 주세요."]

    lines = [re.sub(r"\s+", " ", ln).strip() for ln in text.splitlines()]
    lines = [ln for ln in lines if ln]

    bullets = []
    # ■, , •, – 로 시작하는 불릿 문단 우선 수집
    cur_buf = ""
    for ln in lines:
        if any(k in ln for k in ("Compliance notice", "당보고서공표일", "Stock Indicator", "STOCK DATA", "COMPANY DATA", "K-IFRS", "자료 :")):
            if cur_buf:
                bullets.append(cur_buf.strip())
                cur_buf = ""
            continue
        if re.match(r"^[■•●◆▶\-–]\s*", ln):
            if cur_buf:
                bullets.append(cur_buf.strip())
            cur_buf = re.sub(r"^[■•●◆▶\-–]\s*", "", ln)
        elif cur_buf:
            if len(cur_buf) < 220 and not re.match(r"^[\d\.,\-\+\s%]+$", ln):
                cur_buf += " " + ln
            else:
                bullets.append(cur_buf.strip())
                cur_buf = ""
    if cur_buf:
        bullets.append(cur_buf.strip())

    clean_bullets = []
    for b in bullets:
        b = re.sub(r"\s+", " ", b).strip()
        if len(b) >= 25 and not re.match(r"^[\d\s\.,%\-\+]+$", b):
            if len(b) > 220:
                b = b[:217].rstrip() + "..."
            clean_bullets.append(b)
        if len(clean_bullets) >= 4:
            break

    if clean_bullets:
        return clean_bullets

    # 불릿이 없는 일반 문단형 PDF (SK증권 등) 처리
    joined = " ".join(ln for ln in lines[:45] if len(ln) >= 12 and not re.match(r"^[\d\s\.,%\-\+/]+$", ln))
    sents = re.split(r"(?<=[다요함임음]\.)\s+", joined)
    picked = []
    for s in sents:
        s = s.strip()
        if len(s) >= 25 and not any(skip in s for skip in ("Compliance", "http", "sks.co.kr", "@")):
            if len(s) > 220:
                s = s[:217].rstrip() + "..."
            picked.append(s)
        if len(picked) >= 3:
            break
    return picked if picked else ([title] if title else ["핵심 요약 추출 불가 — 원문 PDF를 참고하세요."])


def _parse_krw_amount(num_str, unit_str=""):
    v = _num(num_str)
    if v is None:
        return None
    if "만" in str(unit_str):
        return v * 10000.0
    return v


def _extract_target_change_info(first_text, last_text, target_price):
    """
    리포트 1~2페이지 본문 및 마지막 페이지 '2년간 목표주가 변동추이' 표에서
    (target_change, prev_target_price, target_change_pct)를 정밀 추출합니다.
    """
    curr_tp = _num(target_price)
    if curr_tp is None or curr_tp <= 0:
        return "없음", None, None

    head = (first_text or "")[:3000]
    head_flat = re.sub(r"\s+", " ", head)

    # 1. 신규 커버리지 개시 여부
    if (
        re.search(r"(?:목표주가|목표가|TP)[^\.]{0,35}(?:신규|개시|편입)", head_flat)
        or "커버리지 개시" in head_flat
        or "커버리지를 개시" in head_flat
        or "신규편입" in head_flat[:1200]
    ):
        return "신규", None, None

    prev_tp = None
    chg_pct = None

    # 2A. 본문에서 '기존 A(만)원에서 B(만)원으로 상향/하향' 패턴 추출
    m_from_to = re.search(
        r"(?:목표주가|목표가|TP)[^\.]{0,35}?(?:기존|직전|이전)?\s*([\d,\.]+)\s*(만?원)\s*(?:에서|→|->|~)\s*([\d,\.]+)\s*(만?원)",
        head_flat,
    )
    if m_from_to:
        p_old = _parse_krw_amount(m_from_to.group(1), m_from_to.group(2))
        p_new = _parse_krw_amount(m_from_to.group(3), m_from_to.group(4))
        if p_old and p_new and p_old > 0:
            prev_tp = round(p_old)
            chg_pct = round((p_new - p_old) / p_old * 100.0, 1)

    # 2B. 요약표 '목표주가 100,000 120,000 ▼/▲' (유진/SK 등) 패턴 추출
    if chg_pct is None:
        m_tbl = re.search(r"목표주가\s+([\d,]{3,})\s+([\d,]{3,})\s*([▲▼↑↓])", head_flat)
        if m_tbl:
            v1, v2 = _num(m_tbl.group(1)), _num(m_tbl.group(2))
            if v1 and v2 and v2 > 0 and abs(v1 - curr_tp) <= max(100, curr_tp * 0.02):
                prev_tp = round(v2)
                chg_pct = round((v1 - v2) / v2 * 100.0, 1)

    # 2C. 본문에서 '목표주가 ... 18% 상향/하향' 명시 퍼센트 추출
    if chg_pct is None:
        m_pct = re.search(
            r"(?:목표주가|목표가|TP)[^\.]{0,55}?([+-]?\d+(?:\.\d+)?)\s*%\s*(?<!우)(상향|하향)",
            head_flat,
        )
        if m_pct:
            pct_v = abs(float(m_pct.group(1)))
            if 0.1 <= pct_v <= 500.0:
                dir_word = m_pct.group(2)
                chg_pct = round(pct_v if dir_word == "상향" else -pct_v, 1)
                denom = 1.0 + chg_pct / 100.0
                if denom > 0:
                    prev_tp = round(curr_tp / denom)

    # 3. 마지막 페이지 '투자의견 및 목표주가 변동추이' 법정 공시 표에서 직전 목표가 추출
    if prev_tp is None and last_text:
        hist_rows = []
        for ln in last_text.splitlines():
            m_h = re.search(r"(20\d{2}[-./]\d{2}[-./]\d{2})[^\d]{1,35}?([\d,]{3,})(?:\s|$)", ln)
            if not m_h:
                continue
            dt_s = m_h.group(1).replace(".", "-").replace("/", "-")
            tp_v = _num(m_h.group(2))
            if tp_v is not None and 500 <= tp_v <= 50000000:
                hist_rows.append((dt_s, tp_v))
        if len(hist_rows) >= 2:
            # 날짜 오름차순 정렬 (중복 날짜 제거 후 마지막 2건 비교)
            dedup = []
            for dt_s, tp_v in hist_rows:
                if dedup and dedup[-1][0] == dt_s and dedup[-1][1] == tp_v:
                    continue
                dedup.append((dt_s, tp_v))
            if len(dedup) >= 2:
                dedup.sort(key=lambda x: x[0])
                last_tp = dedup[-1][1]
                prev_cand = dedup[-2][1]
                if abs(last_tp - curr_tp) <= max(100, curr_tp * 0.02) and prev_cand > 0:
                    prev_tp = round(prev_cand)
                    chg_pct = round((curr_tp - prev_cand) / prev_cand * 100.0, 1)

    # 4. 최종 변동 구분 결정
    if chg_pct is not None:
        if chg_pct > 0.1:
            return "상향", prev_tp, chg_pct
        if chg_pct < -0.1:
            return "하향", prev_tp, chg_pct
        return "유지", prev_tp, 0.0

    # 퍼센트·직전가가 안 잡힌 경우에만 목표주가 인접 텍스트로 방향 판별 ('우상향' 제외)
    m_dir = re.search(r"(?:목표주가|목표가|TP)[^\.]{0,40}?(?<!우)(상향|하향|유지|↑|▲|↓|▼)", head_flat)
    if m_dir:
        tok = m_dir.group(1)
        if tok in ("상향", "↑", "▲"):
            return "상향", prev_tp, None
        if tok in ("하향", "↓", "▼"):
            return "하향", prev_tp, None
        return "유지", prev_tp, 0.0

    return "유지", prev_tp, None


def _detect_target_change(text, target_price):
    chg, _, _ = _extract_target_change_info(text, "", target_price)
    return chg


def _extract_pdf_price_info(text):
    """PDF 본문에서 현재주가와 상승여력(%)을 보조 추출합니다."""
    cur_price, upside_pct = None, None
    if not text:
        return cur_price, upside_pct
    head = text[:2500]
    m_p = re.search(r"(?:현재주가|현재가|종가)\s*(?:\([^)]*\))?\s*[:：]?\s*([\d,]+)\s*원", head)
    if m_p:
        cur_price = _num(m_p.group(1))
    m_u = re.search(r"상승여력\s*[:：]?\s*([+-]?\d+(?:\.\d+)?)\s*%", head)
    if m_u:
        upside_pct = _num(m_u.group(1))
    return cur_price, upside_pct


def fetch_hankyung_reports(here, target_date=None, lookback_days=0):
    """
    한경 컨센서스(https://consensus.hankyung.com/analysis/list?skinType=business)에서
    오늘자(또는 최근 영업일/지정 기간) 기업 리포트 목록 및 PDF를 다운로드·정리하여
    hankyung_reports.json에 최대 60일 치 일자별 이력(byDate)으로 누적 저장합니다.
    """
    json_path = os.path.join(here, "hankyung_reports.json")
    by_date = {}
    existing_by_idx = {}
    if os.path.isfile(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                old_data = json.load(f)
            if isinstance(old_data.get("byDate"), dict):
                for dt_k, rep_list in old_data["byDate"].items():
                    if isinstance(rep_list, list):
                        by_date[dt_k] = rep_list
            elif old_data.get("date") and isinstance(old_data.get("reports"), list):
                by_date[old_data["date"]] = old_data["reports"]
            for rep_list in by_date.values():
                for rep in rep_list:
                    if rep.get("report_idx"):
                        existing_by_idx[str(rep["report_idx"])] = rep
        except Exception:
            pass

    try:
        from pypdf import PdfReader
    except ImportError:
        PdfReader = None

    end_dt = datetime.strptime(target_date, "%Y-%m-%d") if target_date else datetime.now()
    today_str = end_dt.strftime("%Y-%m-%d")
    hdr = {"User-Agent": UA}
    base_list_url = "https://consensus.hankyung.com/analysis/list?skinType=business"

    print(f"\n{'='*80}")
    if lookback_days > 0:
        start_str = (end_dt - timedelta(days=lookback_days)).strftime("%Y-%m-%d")
        print(f" ▶ 한경 컨센서스 기업 리포트 수집 중 ({start_str} ~ {today_str})...")
    else:
        start_str = today_str
        print(f" ▶ 한경 컨센서스 기업 리포트 수집 중 ({today_str})...")

    def parse_page(url_to_get):
        r = requests.get(url_to_get, headers=hdr, timeout=20)
        r.raise_for_status()
        p_html = r.content.decode("utf-8", errors="replace")
        trs = re.findall(r"<tr[^>]*>(.*?)</tr>", p_html, re.S)
        items = []
        for tr in trs[1:]:
            tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, re.S)
            if len(tds) < 6:
                continue

            def cln(s):
                return html_lib.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s))).strip()

            dt = cln(tds[0])
            if not re.match(r"^\d{4}-\d{2}-\d{2}$", dt):
                continue
            m_a = re.search(r'<a[^>]*href="([^"]*downpdf\?report_idx=(\d+))"[^>]*>(.*?)</a>', tds[1], re.S)
            if m_a:
                idx = m_a.group(2)
                raw_title = cln(m_a.group(3))
            else:
                m_idx = re.search(r"downpdf\?report_idx=(\d+)", tr)
                idx = m_idx.group(1) if m_idx else ""
                raw_title = cln(tds[1])

            m_st = re.match(r"^(.+?)\((\d{6})\)\s*(.*)$", raw_title)
            if m_st:
                st_name, st_code, title = m_st.group(1).strip(), m_st.group(2).strip(), m_st.group(3).strip()
            else:
                st_name, st_code, title = "", "", raw_title

            items.append({
                "date": dt,
                "stock_name": st_name,
                "stock_code": st_code,
                "title": title,
                "target_price": cln(tds[2]),
                "opinion": cln(tds[3]),
                "analyst": cln(tds[4]),
                "broker": cln(tds[5]),
                "report_idx": idx,
                "pdf_url": f"https://consensus.hankyung.com/analysis/downpdf?report_idx={idx}" if idx else "",
            })
        return items

    try:
        collected = []
        if lookback_days > 0:
            seen_idx = set()
            for pg in range(1, 6):
                url_pg = f"{base_list_url}&sdate={start_str}&edate={today_str}&pagenum=80&now_page={pg}"
                pg_items = parse_page(url_pg)
                if not pg_items:
                    break
                added = 0
                for it in pg_items:
                    key = it.get("report_idx") or f"{it['date']}_{it['stock_code']}_{it['title']}"
                    if key not in seen_idx and start_str <= it["date"] <= today_str:
                        seen_idx.add(key)
                        collected.append(it)
                        added += 1
                if len(pg_items) < 80 or added == 0:
                    break
        else:
            url_today = f"{base_list_url}&sdate={today_str}&edate={today_str}&pagenum=80&now_page=1"
            collected = [it for it in parse_page(url_today) if it["date"] == today_str]
            # 주말·장전 등으로 오늘 리포트가 아직 없으면 가장 최근 올라온 날짜 리포트 사용
            if not collected:
                latest_items = parse_page(f"{base_list_url}&pagenum=40&now_page=1")
                if latest_items:
                    latest_dt = latest_items[0]["date"]
                    collected = [it for it in latest_items if it["date"] == latest_dt]

        if not collected and not by_date:
            print("  [알림] 한경 컨센서스에서 리포트를 찾지 못해 기존 캐시를 유지합니다.")
            return _load_hankyung_reports(here)

        grouped_new = {}
        for it in collected:
            grouped_new.setdefault(it["date"], []).append(it)

        for dt_key, dt_items in sorted(grouped_new.items(), reverse=True):
            dt_final = []
            for it in dt_items:
                idx = str(it.get("report_idx") or "")
                cached = existing_by_idx.get(idx)
                if cached and (cached.get("curated") or cached.get("summary")):
                    dt_final.append(cached)
                    continue

                # 디스크에 PDF 파일을 남기지 않고 메모리(BytesIO)에서 바로 요약 후 자동 폐기
                pdf_text, last_text = "", ""
                if idx and it.get("pdf_url") and PdfReader is not None:
                    try:
                        rp = requests.get(it["pdf_url"], headers={**hdr, "Referer": base_list_url}, timeout=20)
                        if rp.status_code == 200 and rp.content.startswith(b"%PDF"):
                            reader = PdfReader(io.BytesIO(rp.content))
                            pdf_text = "\n".join((p.extract_text() or "") for p in reader.pages[:2])
                            if len(reader.pages) >= 1:
                                last_text = reader.pages[-1].extract_text() or ""
                    except Exception:
                        pdf_text, last_text = "", ""

                cur_p, up_p = _extract_pdf_price_info(pdf_text)
                if cur_p is not None:
                    it["current_price"] = cur_p
                if up_p is not None:
                    it["upside_pct"] = up_p
                chg, prev_tp, chg_pct = _extract_target_change_info(pdf_text, last_text, it.get("target_price"))
                it["target_change"] = chg
                if prev_tp is not None:
                    it["prev_target_price"] = prev_tp
                if chg_pct is not None:
                    it["target_change_pct"] = chg_pct
                it["summary"] = _summarize_pdf_text(pdf_text, it.get("title", ""))
                it["source"] = "한경"
                it["curated"] = False
                dt_final.append(it)
                existing_by_idx[idx] = it
                time.sleep(0.15)
            # 기존에 병합되어 있던 네이버 리포트 유지 후, 아래에서 중복 제거하며 재병합
            prev_naver_for_dt = [r for r in by_date.get(dt_key, []) if str(r.get("report_idx", "")).startswith("naver_")]
            by_date[dt_key] = dt_final + prev_naver_for_dt
            print(f"  - [{dt_key}] 한경 리포트 {len(dt_final)}건 정리 완료")

        # ---------------------------------------------------------------------
        # ★ 네이버 증권(stock.naver.com/research/company) 비겹침 리포트 추가 병합
        # ---------------------------------------------------------------------
        def _norm_broker(b):
            s = re.sub(r"\s+", "", str(b or "")).lower()
            s = re.sub(r"(투자증권|증권|투자|리서치센터|리서치)$", "", s)
            return s

        def _norm_title(t):
            return re.sub(r"[^0-9a-z가-힣]+", "", str(t or "").lower())

        def _summarize_naver_html(raw_html, fallback_title=""):
            if not raw_html:
                return []
            txt = re.sub(r"<br\s*/?>", "\n", str(raw_html), flags=re.I)
            txt = re.sub(r"</p\s*>", "\n", txt, flags=re.I)
            txt = html_lib.unescape(re.sub(r"<[^>]+>", " ", txt))
            lines = []
            for ln in txt.splitlines():
                ln = re.sub(r'^[\s"“”\'■•●◆▶\-–]+', "", ln).strip()
                ln = re.sub(r"\s+", " ", ln)
                if len(ln) >= 20:
                    if len(ln) > 220:
                        ln = ln[:217].rstrip() + "..."
                    lines.append(ln)
                if len(lines) >= 4:
                    break
            return lines if lines else ([fallback_title] if fallback_title else [])

        def _detect_target_change_strict(txt, tp):
            if not tp or str(tp).strip() in ("0", "", "-"):
                return "없음"
            if not txt:
                return ""
            head = txt[:2200]
            if re.search(r"목표주가[^\n]{0,30}(신규|개시|편입)", head) or "커버리지 개시" in head or "커버리지를 개시" in head:
                return "신규"
            if re.search(r"목표주가[^\n]{0,30}(상향|↑|▲)", head) or ("목표주가" in head and "상향" in head[:1400]):
                return "상향"
            if re.search(r"목표주가[^\n]{0,30}(하향|↓|▼)", head) or ("목표주가" in head and "하향" in head[:1400]):
                return "하향"
            if re.search(r"목표주가[^\n]{0,30}유지", head):
                return "유지"
            return ""

        try:
            naver_min_date = start_str if lookback_days > 0 else (min(grouped_new.keys()) if grouped_new else today_str)
            naver_max_date = today_str
            naver_raw_list = []
            for pg in range(1, 7):
                r_nv = requests.get(
                    f"https://m.stock.naver.com/api/research/company?page={pg}&pageSize=50",
                    headers=hdr,
                    timeout=15,
                )
                if r_nv.status_code != 200:
                    break
                arr = r_nv.json()
                if not isinstance(arr, list) or not arr:
                    break
                oldest_on_page = ""
                for item in arr:
                    wdt = str(item.get("writeDate") or "").strip()
                    if not re.match(r"^\d{4}-\d{2}-\d{2}$", wdt):
                        continue
                    oldest_on_page = wdt
                    if naver_min_date <= wdt <= naver_max_date:
                        naver_raw_list.append(item)
                if oldest_on_page and oldest_on_page < naver_min_date:
                    break
                time.sleep(0.1)

            if naver_raw_list:
                nv_by_dt = {}
                for item in naver_raw_list:
                    nv_by_dt.setdefault(item["writeDate"], []).append(item)

                for dt_key, nv_items in sorted(nv_by_dt.items(), reverse=True):
                    # 해당 날짜의 기존 한경 리포트 목록만 기준(네이버 항목은 최신 목록으로 갱신)
                    hk_list = [r for r in by_date.get(dt_key, []) if not str(r.get("report_idx", "")).startswith("naver_")]
                    for r in hk_list:
                        if not r.get("source"):
                            r["source"] = "한경"

                    seen_broker_keys = {
                        (str(r.get("stock_code") or "").zfill(6), _norm_broker(r.get("broker")))
                        for r in hk_list
                    }
                    seen_titles = [
                        (str(r.get("stock_code") or "").zfill(6), _norm_title(r.get("title")))
                        for r in hk_list
                    ]

                    nv_added = []
                    for nv in nv_items:
                        rid = str(nv.get("researchId") or "").strip()
                        st_code = str(nv.get("itemCode") or "").strip().zfill(6)
                        st_name = str(nv.get("itemName") or "").strip()
                        broker = str(nv.get("brokerName") or "").strip()
                        title = str(nv.get("title") or "").strip()
                        if not rid or not st_code:
                            continue

                        nb = _norm_broker(broker)
                        nt = _norm_title(title)
                        # 한경(또는 이미 추가된 네이버 리포트)과 종목+증권사 또는 종목+제목이 겹치면 제외
                        if (st_code, nb) in seen_broker_keys:
                            continue
                        if nt and any(sc == st_code and st and (nt in st or st in nt) for sc, st in seen_titles):
                            continue

                        nv_idx = f"naver_{rid}"
                        cached_nv = existing_by_idx.get(nv_idx)
                        if cached_nv and cached_nv.get("summary"):
                            cached_nv["source"] = "네이버"
                            nv_added.append(cached_nv)
                            seen_broker_keys.add((st_code, nb))
                            seen_titles.append((st_code, nt))
                            continue

                        # 네이버 상세 API 호출 (목표가, 투자의견, 작성일 현재가, 본문요약, 원문 PDF 링크)
                        rc = {}
                        try:
                            r_det = requests.get(
                                f"https://m.stock.naver.com/api/research/company/{rid}",
                                headers=hdr,
                                timeout=15,
                            )
                            if r_det.status_code == 200:
                                rc = (r_det.json() or {}).get("researchContent") or {}
                        except Exception:
                            rc = {}

                        gp_num = _num(rc.get("goalPrice"))
                        tp_str = f"{int(round(gp_num)):,}" if (gp_num is not None and gp_num > 0) else "0"
                        opinion = str(rc.get("opinion") or "").strip() or "투자의견없음"
                        cur_p = _num(rc.get("priceAtWriteDate"))
                        attach_url = str(rc.get("attachUrl") or "").strip()
                        raw_content = str(rc.get("content") or "")

                        summary_lines = _summarize_naver_html(raw_content, title)
                        plain_content = html_lib.unescape(re.sub(r"<[^>]+>", " ", raw_content))
                        chg, prev_tp, chg_pct = _extract_target_change_info(plain_content, "", tp_str)

                        # 본문에 직전 목표가·상향율이 없거나 요약이 부족하면 PDF 1페이지·마지막 페이지를 메모리에서 확인
                        pdf_text, last_text = "", ""
                        need_pdf = (tp_str != "0" and chg_pct is None and chg != "신규") or len(summary_lines) < 2
                        if need_pdf and attach_url.lower().endswith(".pdf") and PdfReader is not None:
                            try:
                                rp = requests.get(attach_url, headers=hdr, timeout=15)
                                if rp.status_code == 200 and rp.content.startswith(b"%PDF"):
                                    reader = PdfReader(io.BytesIO(rp.content))
                                    pdf_text = "\n".join((p.extract_text() or "") for p in reader.pages[:2])
                                    if len(reader.pages) >= 1:
                                        last_text = reader.pages[-1].extract_text() or ""
                            except Exception:
                                pdf_text, last_text = "", ""

                        if pdf_text or last_text:
                            chg2, prev_tp2, chg_pct2 = _extract_target_change_info(
                                plain_content + "\n" + pdf_text, last_text, tp_str
                            )
                            chg = chg2
                            if prev_tp2 is not None:
                                prev_tp = prev_tp2
                            if chg_pct2 is not None:
                                chg_pct = chg_pct2
                        if len(summary_lines) < 2 and pdf_text:
                            summary_lines = _summarize_pdf_text(pdf_text, title)
                        if not summary_lines:
                            summary_lines = [title]

                        rep_obj = {
                            "date": dt_key,
                            "stock_name": st_name,
                            "stock_code": st_code,
                            "title": title,
                            "target_price": tp_str,
                            "target_change": chg,
                            "opinion": opinion,
                            "analyst": "",
                            "broker": broker,
                            "report_idx": nv_idx,
                            "pdf_url": attach_url,
                            "naver_url": f"https://stock.naver.com/research/company/{rid}",
                            "source": "네이버",
                            "curated": False,
                            "summary": summary_lines,
                        }
                        if prev_tp is not None:
                            rep_obj["prev_target_price"] = prev_tp
                        if chg_pct is not None:
                            rep_obj["target_change_pct"] = chg_pct
                        if cur_p is not None and cur_p > 0:
                            rep_obj["current_price"] = cur_p
                        nv_added.append(rep_obj)
                        existing_by_idx[nv_idx] = rep_obj
                        seen_broker_keys.add((st_code, nb))
                        seen_titles.append((st_code, nt))
                        time.sleep(0.08)

                    by_date[dt_key] = hk_list + nv_added
                    print(f"  - [{dt_key}] 네이버 비겹침 리포트 +{len(nv_added)}건 병합 (총 {len(by_date[dt_key])}건)")
        except Exception as e_nv:
            print(f"  [경고] 네이버 리포트 병합 중 오류(한경 데이터는 정상 유지): {e_nv}")

        # 최대 60일 이전까지만 보관 (60일 초과 데이터 자동 정리)
        cutoff_str = (datetime.now() - timedelta(days=60)).strftime("%Y-%m-%d")
        pruned_by_date = {
            d: lst for d, lst in sorted(by_date.items(), reverse=True)
            if d >= cutoff_str and isinstance(lst, list) and len(lst) > 0
        }
        sorted_dates = sorted(pruned_by_date.keys(), reverse=True)
        latest_date = sorted_dates[0] if sorted_dates else today_str
        latest_reports = pruned_by_date.get(latest_date, [])

        payload = {
            "date": latest_date,
            "updatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "retentionDays": 60,
            "sourceUrl": base_list_url,
            "dates": sorted_dates,
            "byDate": pruned_by_date,
            "reports": latest_reports,
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        total_saved = sum(len(v) for v in pruned_by_date.values())
        print(f" ★ 한경·네이버 통합 기업 리포트 저장 완료: 최신일 {latest_date}({len(latest_reports)}건) / 총 {len(sorted_dates)}일 · {total_saved}건 (최대 60일 보관)")
        print(f"{'='*80}")
        return payload
    except Exception as e:
        print(f"  [경고] 한경 리포트 수집 중 오류 발생(기존 데이터 유지): {e}")
        return _load_hankyung_reports(here)


def _load_hankyung_reports(here):
    json_path = os.path.join(here, "hankyung_reports.json")
    if os.path.isfile(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data.get("byDate"), dict):
                dt = data.get("date") or ""
                reps = data.get("reports") or []
                data["byDate"] = {dt: reps} if (dt and reps) else {}
            data["dates"] = sorted(data["byDate"].keys(), reverse=True)
            return data
        except Exception:
            pass
    return {"date": "", "updatedAt": "", "retentionDays": 60, "sourceUrl": "https://consensus.hankyung.com/analysis/list?skinType=business", "dates": [], "byDate": {}, "reports": []}


def build_web_dashboard(here):
    """
    최신 KOSPI / KOSDAQ 통합 엑셀(.xlsx) 파일 및 한경 기업 리포트(hankyung_reports.json)를 읽어
    웹 브라우저 및 GitHub Pages에서 무료로 바로 열람 가능한 index.html 파일을 생성합니다.
    """
    print(f"\n{'='*80}")
    print(" ▶ 웹페이지(index.html) 대시보드 생성 중...")
    datasets = {}
    for mkt in MARKETS:
        data = _extract_market_excel_for_web(here, mkt)
        if data:
            datasets[mkt] = data
            print(f"  - [{mkt}] 데이터 로드 완료: {data['filename']} (시트 {len(data['sheets'])}개, 상향 {data['upCount']}셀 / 하향 {data['downCount']}셀)")

    if not datasets:
        print("  [경고] 변환할 엑셀(.xlsx) 파일을 찾지 못해 index.html 생성을 건너뜁니다.")
        return None

    reports_data = _load_hankyung_reports(here)
    print(f"  - [리포트] 한경 기업 리포트 로드 완료: {len(reports_data.get('reports', []))}건 (기준일: {reports_data.get('date', '-')})")

    op_hist_raw = _load_op_estimate_history(here)
    op_hist_compact = {
        "quarter": op_hist_raw.get("quarter", _current_quarter_ym()),
        "col": op_hist_raw.get("col", ""),
        "stocks": {
            cd: {
                "lv": int(ent.get("level") or 0),
                "u": int(ent.get("upCount") or 0),
                "d": int(ent.get("downCount") or 0),
                "p": ent.get("totalPct"),
                "m": ent.get("memo") or "",
            }
            for cd, ent in (op_hist_raw.get("stocks") or {}).items()
        },
    }
    up_tracked_cnt = sum(1 for v in op_hist_compact["stocks"].values() if v["lv"] >= 1)
    print(f"  - [분기OP추적] {op_hist_compact['quarter']}(E) 추적 {len(op_hist_compact['stocks'])}종목 (누적 상향 강조 {up_tracked_cnt}종목)")

    json_payload = json.dumps(datasets, ensure_ascii=False).replace("</", "<\\/")
    reports_payload = json.dumps(reports_data, ensure_ascii=False).replace("</", "<\\/")
    op_hist_payload = json.dumps(op_hist_compact, ensure_ascii=False).replace("</", "<\\/")

    html_content = f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>KOSPI · KOSDAQ 200 영업이익 성장·수익성 스크리너</title>
  <style>
    :root {{
      --bg: #f4f6fb;
      --card: #ffffff;
      --text: #1e293b;
      --muted: #64748b;
      --border: #cbd5e1;
      --border-light: #e2e8f0;
      --primary: #2563eb;
      --primary-bg: #eff6ff;
      --kospi: #1d4ed8;
      --kosdaq: #059669;
      --up-bg: #ffc7ce;
      --up-text: #9c0006;
      --down-bg: #bdd7ee;
      --down-text: #0c4a6e;
      --pass-bg: #dcfce7;
      --pass-text: #166534;
      --header-bg: #1e293b;
      --header-text: #f8fafc;
    }}
    * {{ box-sizing: border-box; }}
    html, body {{
      height: 100%;
    }}
    body {{
      margin: 0;
      font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Malgun Gothic', '맑은 고딕', sans-serif;
      background: #f8fafc;
      color: var(--text);
      font-size: 13px;
      display: flex;
      flex-direction: column;
      height: 100vh;
      overflow: hidden;
    }}
    /* 상단 슬림 화이트 헤더 (MONEYLAND 스타일) */
    .topbar {{
      background: #ffffff;
      color: #0f172a;
      height: 48px;
      padding: 0 14px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      flex-shrink: 0;
      border-bottom: 1px solid #e2e8f0;
      z-index: 50;
    }}
    .topbar-left {{
      display: flex;
      align-items: center;
      gap: 10px;
      min-width: 0;
    }}
    .menu-toggle-btn {{
      width: 32px;
      height: 32px;
      border: 1px solid #e2e8f0;
      background: #f8fafc;
      border-radius: 7px;
      font-size: 16px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      color: #334155;
      transition: background 0.15s;
    }}
    .menu-toggle-btn:hover {{
      background: #e2e8f0;
    }}
    .brand-logo {{
      display: inline-flex;
      align-items: center;
      gap: 6px;
      font-size: 15.5px;
      font-weight: 900;
      letter-spacing: -0.4px;
      color: #1e3a8a;
      white-space: nowrap;
      cursor: pointer;
    }}
    .brand-dot {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      width: 24px;
      height: 24px;
      border-radius: 7px;
      background: linear-gradient(135deg, #2563eb, #1d4ed8);
      color: #fff;
      font-size: 13px;
    }}
    .top-date-chip {{
      font-size: 11.5px;
      font-weight: 700;
      color: #475569;
      background: #f1f5f9;
      padding: 4px 9px;
      border-radius: 7px;
      white-space: nowrap;
    }}
    .topbar-center {{
      display: flex;
      align-items: center;
      gap: 10px;
      flex: 1;
      max-width: 640px;
      min-width: 0;
    }}
    .global-search {{
      width: 250px;
      max-width: 100%;
      padding: 6px 11px;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
      font-size: 12px;
      background: #f8fafc;
      outline: none;
    }}
    .global-search:focus {{
      background: #fff;
      border-color: #2563eb;
      box-shadow: 0 0 0 2px rgba(37,99,235,0.14);
    }}
    .top-disclaimer {{
      font-size: 11px;
      color: #94a3b8;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .top-actions {{
      display: flex;
      align-items: center;
      gap: 6px;
      flex-shrink: 0;
    }}
    .dl-btn {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      background: #f1f5f9;
      color: #1e293b;
      text-decoration: none;
      padding: 5px 10px;
      border-radius: 7px;
      font-weight: 700;
      font-size: 11.5px;
      border: 1px solid #cbd5e1;
      transition: all 0.15s;
      white-space: nowrap;
    }}
    .dl-btn:hover {{
      background: #e2e8f0;
      border-color: #94a3b8;
    }}
    /* 좌측 사이드바 + 우측 메인 콘텐츠 셸 */
    .app-shell {{
      flex: 1;
      min-height: 0;
      display: flex;
      position: relative;
      overflow: hidden;
    }}
    .sidebar {{
      width: 186px;
      background: #ffffff;
      border-right: 1px solid #e2e8f0;
      padding: 10px 8px 16px;
      overflow-y: auto;
      flex-shrink: 0;
      display: flex;
      flex-direction: column;
      gap: 12px;
      transition: width 0.18s ease, padding 0.18s ease;
    }}
    body.sidebar-collapsed .sidebar {{
      width: 0;
      padding-left: 0;
      padding-right: 0;
      border-right: none;
      overflow: hidden;
    }}
    .nav-section {{
      display: flex;
      flex-direction: column;
      gap: 2px;
    }}
    .nav-section-title {{
      font-size: 10.5px;
      font-weight: 700;
      color: #94a3b8;
      padding: 4px 10px 3px;
      letter-spacing: 0.2px;
      white-space: nowrap;
    }}
    .nav-item {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 6px;
      width: 100%;
      border: none;
      background: transparent;
      color: #334155;
      padding: 7px 10px;
      border-radius: 8px;
      font-size: 12.5px;
      font-weight: 700;
      cursor: pointer;
      text-align: left;
      text-decoration: none;
      white-space: nowrap;
      transition: background 0.12s, color 0.12s;
    }}
    .nav-item:hover {{
      background: #f1f5f9;
      color: #0f172a;
    }}
    .nav-item.active {{
      background: #1d4ed8;
      color: #ffffff;
      box-shadow: 0 2px 6px rgba(29,78,216,0.28);
    }}
    .nav-item-left {{
      display: inline-flex;
      align-items: center;
      gap: 7px;
      overflow: hidden;
      text-overflow: ellipsis;
    }}
    .nav-badge {{
      font-size: 10.5px;
      font-weight: 800;
      padding: 1px 6px;
      border-radius: 999px;
      background: #f1f5f9;
      color: #475569;
    }}
    .nav-item.active .nav-badge {{
      background: rgba(255,255,255,0.24);
      color: #ffffff;
    }}
    .nav-badge.up {{ background: #ffe4e6; color: #be123c; }}
    .nav-badge.down {{ background: #dbeafe; color: #1d4ed8; }}
    .nav-badge.new {{ background: #ede9fe; color: #6d28d9; }}
    .nav-divider {{
      height: 1px;
      background: #f1f5f9;
      margin: 2px 6px;
    }}
    .sidebar-backdrop {{
      display: none;
      position: fixed;
      inset: 48px 0 0 0;
      background: rgba(15,23,42,0.35);
      z-index: 39;
    }}
    .container {{
      flex: 1;
      min-width: 0;
      min-height: 0;
      display: flex;
      flex-direction: column;
      padding: 10px 14px 10px;
      overflow: hidden;
    }}
    .page-header {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      margin-bottom: 8px;
      flex-shrink: 0;
    }}
    .page-title-group {{
      display: flex;
      align-items: baseline;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .page-title {{
      margin: 0;
      font-size: 18px;
      font-weight: 900;
      color: #0f172a;
      letter-spacing: -0.4px;
    }}
    .page-subtitle {{
      font-size: 12px;
      color: #64748b;
      font-weight: 600;
    }}
    .meta-badges {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      align-items: center;
      font-size: 11.5px;
    }}
    .meta-pill {{
      background: var(--card);
      border: 1px solid var(--border);
      padding: 4px 9px;
      border-radius: 999px;
      font-weight: 600;
    }}
    .meta-pill.up {{
      background: var(--up-bg);
      color: var(--up-text);
      border-color: #f8a5b0;
    }}
    .meta-pill.down {{
      background: var(--down-bg);
      color: var(--down-text);
      border-color: #93c5fd;
    }}
    .sheet-tabs {{
      display: flex;
      flex-wrap: wrap;
      gap: 6px;
      margin-bottom: 8px;
      background: transparent;
      padding: 0;
      flex-shrink: 0;
    }}
    .sheet-btn {{
      border: 1px solid #cbd5e1;
      background: #ffffff;
      color: #334155;
      padding: 5px 11px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 5px;
      transition: all 0.15s;
    }}
    .sheet-btn:hover {{
      background: #f1f5f9;
    }}
    .sheet-btn.active {{
      background: #2563eb;
      color: #fff;
      border-color: #2563eb;
      box-shadow: 0 2px 6px rgba(37,99,235,0.22);
    }}
    .sheet-count {{
      background: rgba(100,116,139,0.14);
      padding: 1px 6px;
      border-radius: 999px;
      font-size: 11px;
    }}
    .sheet-btn.active .sheet-count {{
      background: rgba(255,255,255,0.24);
      color: #fff;
    }}
    .sheet-btn.tab-up {{
      background: #fff1f2;
      color: #9f1239;
      border-color: #fda4af;
    }}
    .sheet-btn.tab-up:hover {{
      background: #ffe4e6;
    }}
    .sheet-btn.tab-up.active {{
      background: #be123c;
      color: #fff;
      border-color: #be123c;
    }}
    .sheet-btn.tab-up.active .sheet-count {{
      background: rgba(255,255,255,0.25);
      color: #fff;
    }}
    .sheet-btn.tab-down {{
      background: #eff6ff;
      color: #1e40af;
      border-color: #93c5fd;
    }}
    .sheet-btn.tab-down:hover {{
      background: #dbeafe;
    }}
    .sheet-btn.tab-down.active {{
      background: #1d4ed8;
      color: #fff;
      border-color: #1d4ed8;
    }}
    .sheet-btn.tab-down.active .sheet-count {{
      background: rgba(255,255,255,0.25);
      color: #fff;
    }}
    .meta-pill.clickable {{
      cursor: pointer;
      transition: transform 0.12s;
    }}
    .meta-pill.clickable:hover {{
      transform: translateY(-1px);
      filter: brightness(0.96);
    }}
    .controls-panel {{
      background: var(--card);
      border: 1px solid var(--border-light);
      border-radius: 9px;
      padding: 7px 10px;
      margin-bottom: 8px;
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      flex-shrink: 0;
    }}
    .filter-left, .group-right {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 6px;
    }}
    .search-input, .select-box {{
      padding: 5px 9px;
      border: 1px solid var(--border);
      border-radius: 7px;
      font-size: 12px;
      outline: none;
    }}
    .search-input {{
      width: 195px;
    }}
    .search-input:focus, .select-box:focus {{
      border-color: var(--primary);
      box-shadow: 0 0 0 2px rgba(37,99,235,0.15);
    }}
    .grp-btn {{
      border: 1px solid var(--border);
      background: #f8fafc;
      color: #334155;
      padding: 5px 9px;
      border-radius: 7px;
      font-size: 11.5px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.15s;
    }}
    .grp-btn.open {{
      background: #dbeafe;
      color: #1e40af;
      border-color: #93c5fd;
    }}
    .table-wrap {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 9px;
      overflow: auto;
      flex: 1;
      min-height: 0;
      position: relative;
      box-shadow: 0 2px 6px rgba(0,0,0,0.04);
    }}
    .pagination-bar {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      margin-top: 8px;
      padding: 6px 12px;
      background: var(--card);
      border: 1px solid var(--border-light);
      border-radius: 9px;
      flex-shrink: 0;
      font-size: 12px;
    }}
    .page-controls {{
      display: flex;
      align-items: center;
      gap: 6px;
    }}
    .pg-btn {{
      border: 1px solid var(--border);
      background: #f8fafc;
      color: var(--text);
      padding: 4px 10px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.12s;
    }}
    .pg-btn:hover:not(:disabled) {{
      background: #e2e8f0;
    }}
    .pg-btn:disabled {{
      opacity: 0.45;
      cursor: not-allowed;
    }}
    .pg-info {{
      font-weight: 700;
      color: #1e293b;
      padding: 0 6px;
    }}
    table {{
      border-collapse: separate;
      border-spacing: 0;
      width: max-content;
      min-width: 100%;
      font-size: 12px;
    }}
    thead th {{
      position: sticky;
      top: 0;
      z-index: 20;
      background: var(--header-bg);
      color: var(--header-text);
      padding: 6px 5px;
      font-weight: 700;
      font-size: 11.5px;
      line-height: 1.25;
      text-align: center;
      white-space: pre-line;
      border-bottom: 2px solid #0f172a;
      border-right: 1px solid #334155;
      cursor: pointer;
      user-select: none;
    }}
    thead th.has-tip {{
      text-decoration: underline dotted rgba(147,197,253,0.75);
      text-underline-offset: 2px;
    }}
    thead th:hover {{
      background: #334155;
    }}
    tbody td {{
      padding: 5px 6px;
      border-bottom: 1px solid var(--border-light);
      border-right: 1px solid var(--border-light);
      white-space: nowrap;
      background: #fff;
    }}
    tbody tr:nth-child(even) td {{
      background: #f8fafc;
    }}
    tbody tr:hover td:not(.tone-up1):not(.tone-up2):not(.tone-up3):not(.tone-down1):not(.tone-down2):not(.tone-down3):not(.name-up1):not(.name-up2):not(.name-up3):not(.name-up4) {{
      background: #eef2ff !important;
    }}
    /* 현재 분기 예상 영업이익 누적 상향 횟수별 종목명 단계별 농도 색상 (1회 → 2회 → 3회 → 4회↑) */
    td.name-up1 {{
      background: #fff1f2 !important;
      color: #9f1239;
      box-shadow: inset 3px 0 0 #fda4af;
      cursor: help;
    }}
    td.name-up2 {{
      background: #fecdd3 !important;
      color: #881337;
      box-shadow: inset 3px 0 0 #f43f5e;
      cursor: help;
    }}
    td.name-up3 {{
      background: #fb7185 !important;
      color: #4c0519;
      box-shadow: inset 3px 0 0 #be123c;
      cursor: help;
    }}
    td.name-up4 {{
      background: #dc2626 !important;
      color: #ffffff !important;
      box-shadow: inset 3px 0 0 #7f1d1d;
      cursor: help;
    }}
    td.name-up4 strong {{ color: #ffffff !important; }}
    td.name-up4 .ext-link {{ background: rgba(255,255,255,0.22); color: #ffffff; }}
    .op-cnt-badge {{
      display: inline-block;
      margin-left: 3px;
      padding: 0px 4px;
      border-radius: 4px;
      font-size: 10px;
      font-weight: 800;
      vertical-align: middle;
    }}
    .op-cnt-1 {{ background: #ffe4e6; color: #be123c; }}
    .op-cnt-2 {{ background: #fda4af; color: #881337; }}
    .op-cnt-3 {{ background: #be123c; color: #ffffff; }}
    .op-cnt-4 {{ background: #7f1d1d; color: #fef08a; }}
    /* 틀 고정 (종목코드~섹터2) */
    .freeze-on th.sticky-col-0, .freeze-on td.sticky-col-0 {{
      position: sticky; left: 0px; z-index: 15; min-width: 58px; max-width: 58px;
    }}
    .freeze-on th.sticky-col-1, .freeze-on td.sticky-col-1 {{
      position: sticky; left: 58px; z-index: 15; min-width: 130px; max-width: 130px;
      overflow: hidden; text-overflow: ellipsis;
    }}
    .freeze-on th.sticky-col-2, .freeze-on td.sticky-col-2 {{
      position: sticky; left: 188px; z-index: 15; min-width: 72px; max-width: 72px;
      overflow: hidden; text-overflow: ellipsis;
    }}
    .freeze-on th.sticky-col-3, .freeze-on td.sticky-col-3 {{
      position: sticky; left: 260px; z-index: 15; min-width: 78px; max-width: 78px;
      overflow: hidden; text-overflow: ellipsis;
      border-right: 2px solid #475569 !important;
    }}
    .freeze-on thead th.sticky-col-0,
    .freeze-on thead th.sticky-col-1,
    .freeze-on thead th.sticky-col-2,
    .freeze-on thead th.sticky-col-3 {{
      z-index: 30;
    }}
    /* 연도별 경계 진한 세로선 */
    th.b-left, td.b-left {{
      border-left: 2px solid #1e293b !important;
    }}
    th.b-right, td.b-right {{
      border-right: 2px solid #1e293b !important;
    }}
    /* 예상이익 상향/하향 3단계 농도 셀 (<5% / 5~10% / 10%↑↓) */
    td.tone-up1 {{
      background: #ffe4e6 !important;
      color: #9f1239;
      font-weight: 700;
      cursor: help;
    }}
    td.tone-up2 {{
      background: #fca5a5 !important;
      color: #7f1d1d;
      font-weight: 800;
      cursor: help;
    }}
    td.tone-up3 {{
      background: #dc2626 !important;
      color: #ffffff !important;
      font-weight: 800;
      cursor: help;
    }}
    td.tone-down1 {{
      background: #dbeafe !important;
      color: #1e40af;
      font-weight: 700;
      cursor: help;
    }}
    td.tone-down2 {{
      background: #93c5fd !important;
      color: #172554;
      font-weight: 800;
      cursor: help;
    }}
    td.tone-down3 {{
      background: #1d4ed8 !important;
      color: #ffffff !important;
      font-weight: 800;
      cursor: help;
    }}
    .legend-bar {{
      display: inline-flex;
      flex-wrap: wrap;
      align-items: center;
      gap: 4px;
      font-size: 11px;
      font-weight: 700;
    }}
    .lg-chip {{
      padding: 2px 6px;
      border-radius: 5px;
      border: 1px solid rgba(0,0,0,0.08);
    }}
    .num {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .center {{ text-align: center; }}
    .pos {{ color: #dc2626; font-weight: 700; }}
    .neg {{ color: #2563eb; font-weight: 700; }}
    .badge-true {{
      display: inline-block;
      background: var(--pass-bg);
      color: var(--pass-text);
      font-weight: 800;
      padding: 1px 6px;
      border-radius: 999px;
      font-size: 11px;
    }}
    .badge-false {{
      color: #cbd5e1;
    }}
    .q-icon {{
      display: inline-block;
      margin-left: 4px;
      width: 15px;
      height: 15px;
      line-height: 15px;
      border-radius: 50%;
      background: #3b82f6;
      color: #fff;
      font-size: 10px;
      text-align: center;
      cursor: help;
    }}
    .ext-link {{
      display: inline-block;
      margin-left: 4px;
      padding: 1px 4px;
      font-size: 10px;
      border-radius: 4px;
      background: #e2e8f0;
      color: #334155;
      text-decoration: none;
      font-weight: 700;
    }}
    .ext-link:hover {{
      background: #cbd5e1;
    }}
    .tooltip-box {{
      position: fixed;
      z-index: 9999;
      background: #0f172a;
      color: #f8fafc;
      padding: 10px 13px;
      border-radius: 9px;
      font-size: 12px;
      line-height: 1.55;
      max-width: 360px;
      white-space: pre-wrap;
      box-shadow: 0 8px 24px rgba(0,0,0,0.3);
      border: 1px solid #334155;
      pointer-events: none;
      display: none;
    }}
    .mkt-btn.active-REPORTS {{
      background: #7c3aed;
      color: #fff;
      border-color: #7c3aed;
      box-shadow: 0 2px 8px rgba(124,58,237,0.28);
    }}
    .rep-badge-link {{
      display: inline-block;
      margin-left: 3px;
      padding: 1px 4px;
      font-size: 10px;
      border-radius: 4px;
      background: #ede9fe;
      color: #6d28d9;
      text-decoration: none;
      font-weight: 800;
      cursor: pointer;
      border: 1px solid #c4b5fd;
    }}
    .rep-badge-link:hover {{
      background: #ddd6fe;
    }}
    .reports-view {{
      display: none;
      flex: 1;
      min-height: 0;
      flex-direction: column;
    }}
    .reports-view.active {{
      display: flex;
    }}
    .reports-grid-wrap {{
      flex: 1;
      min-height: 0;
      overflow-y: auto;
      padding: 2px 2px 14px;
    }}
    .reports-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(460px, 1fr));
      gap: 12px;
    }}
    .rep-card {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 11px;
      padding: 13px 15px;
      display: flex;
      flex-direction: column;
      gap: 8px;
      box-shadow: 0 2px 6px rgba(0,0,0,0.04);
      transition: border-color 0.15s, box-shadow 0.15s;
    }}
    .rep-card:hover {{
      border-color: #93c5fd;
      box-shadow: 0 4px 12px rgba(37,99,235,0.09);
    }}
    .rep-card.highlight {{
      border: 2px solid #7c3aed;
      box-shadow: 0 0 0 4px rgba(124,58,237,0.16);
    }}
    .rep-card-head {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
      border-bottom: 1px solid var(--border-light);
      padding-bottom: 8px;
    }}
    .rep-stock-title {{
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 6px;
    }}
    .rep-stock-name {{
      font-size: 15px;
      font-weight: 800;
      color: #0f172a;
    }}
    .rep-code {{
      font-size: 11.5px;
      color: #64748b;
      font-weight: 600;
    }}
    .rep-pill {{
      font-size: 11px;
      font-weight: 700;
      padding: 2px 7px;
      border-radius: 999px;
    }}
    .rep-pill.chg-상향 {{ background: #ffe4e6; color: #be123c; border: 1px solid #fda4af; }}
    .rep-pill.chg-하향 {{ background: #dbeafe; color: #1d4ed8; border: 1px solid #93c5fd; }}
    .rep-pill.chg-신규 {{ background: #ede9fe; color: #6d28d9; border: 1px solid #c4b5fd; }}
    .rep-pill.chg-유지 {{ background: #f1f5f9; color: #475569; border: 1px solid #cbd5e1; }}
    .rep-pill.chg-없음 {{ background: #f8fafc; color: #94a3b8; border: 1px solid #e2e8f0; }}
    .rep-pill.scr-pass {{ background: #dcfce7; color: #166534; }}
    .rep-pill.scr-seq {{ background: #fef3c7; color: #92400e; }}
    .rep-pill.scr-up {{ background: #fee2e2; color: #b91c1c; }}
    .rep-target-box {{
      text-align: right;
      font-size: 12px;
      line-height: 1.35;
    }}
    .rep-target-price {{
      font-size: 14px;
      font-weight: 800;
      color: #1e3a8a;
    }}
    .rep-headline {{
      font-size: 13.5px;
      font-weight: 800;
      color: #1e293b;
      line-height: 1.4;
    }}
    .rep-meta-row {{
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      font-size: 11.5px;
      color: #475569;
      background: #f8fafc;
      padding: 5px 9px;
      border-radius: 7px;
    }}
    .rep-summary-list {{
      margin: 0;
      padding-left: 17px;
      display: flex;
      flex-direction: column;
      gap: 4px;
      font-size: 12.5px;
      line-height: 1.5;
      color: #334155;
    }}
    .rep-card-foot {{
      display: flex;
      flex-wrap: wrap;
      align-items: center;
      justify-content: space-between;
      gap: 6px;
      margin-top: auto;
      padding-top: 6px;
      border-top: 1px dashed var(--border-light);
    }}
    .rep-btn {{
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 9px;
      border-radius: 6px;
      font-size: 11.5px;
      font-weight: 700;
      text-decoration: none;
      cursor: pointer;
      border: 1px solid var(--border);
      background: #f8fafc;
      color: #1e293b;
    }}
    .rep-btn:hover {{ background: #e2e8f0; }}
    .rep-btn.pdf {{
      background: #fef2f2;
      color: #b91c1c;
      border-color: #fecaca;
    }}
    .rep-btn.pdf:hover {{ background: #fee2e2; }}
    .rep-btn.scr {{
      background: #eff6ff;
      color: #1d4ed8;
      border-color: #bfdbfe;
    }}
    .rep-btn.scr:hover {{ background: #dbeafe; }}
    @media (max-width: 768px) {{
      body {{ height: auto; overflow: auto; }}
      .app-shell {{ overflow: visible; }}
      .top-disclaimer, .top-date-chip {{ display: none; }}
      .global-search {{ width: 150px; }}
      .sidebar {{
        position: fixed;
        top: 48px;
        bottom: 0;
        left: 0;
        z-index: 40;
        width: 210px;
        transform: translateX(-100%);
        transition: transform 0.2s ease;
        box-shadow: 4px 0 18px rgba(0,0,0,0.16);
      }}
      body.mobile-nav-open .sidebar {{
        transform: translateX(0);
        width: 210px;
        padding: 10px 8px 16px;
      }}
      body.mobile-nav-open .sidebar-backdrop {{
        display: block;
      }}
      .container {{ flex: none; width: 100%; overflow: visible; }}
      .table-wrap {{ max-height: 68vh; min-height: 320px; }}
      .search-input {{ width: 150px; }}
      .reports-grid {{ grid-template-columns: 1fr; }}
    }}
  </style>
</head>
<body>
  <header class="topbar">
    <div class="topbar-left">
      <button class="menu-toggle-btn" id="menuToggleBtn" onclick="toggleSidebar()" title="좌측 메뉴 접기/펼치기">☰</button>
      <div class="brand-logo" onclick="selectMarket('KOSPI')">
        <span class="brand-dot">📊</span>
        <span>STOCK SCREENER</span>
      </div>
      <span class="top-date-chip" id="topDateChip"></span>
    </div>
    <div class="topbar-center">
      <input type="text" id="globalSearchInput" class="global-search" placeholder="🔍 한국 종목 통합 검색 (Enter)" />
      <span class="top-disclaimer">본 자료는 투자 참고용이며 수치상 오류가 있을 수 있습니다.</span>
    </div>
    <div class="top-actions" id="dlActions"></div>
  </header>

  <div class="app-shell">
    <aside class="sidebar" id="sidebarNav"></aside>
    <div class="sidebar-backdrop" onclick="closeMobileNav()"></div>

    <main class="container">
      <div class="page-header">
        <div class="page-title-group">
          <h1 class="page-title" id="pageTitle">KOSPI 상위 200</h1>
          <span class="page-subtitle" id="pageSubtitle">영업이익 성장·수익성 스크리너</span>
        </div>
        <div class="meta-badges" id="metaBadges"></div>
      </div>

      <div class="sheet-tabs" id="sheetTabs"></div>

    <div class="controls-panel" id="screenerControls">
      <div class="filter-left">
        <input type="text" id="searchInput" class="search-input" placeholder="🔍 종목명 / 코드 / 섹터 검색..." />
        <select id="sectorSelect" class="select-box">
          <option value="">전체 섹터</option>
        </select>
        <select id="toneSelect" class="select-box">
          <option value="">전체 (예상이익 변동 무관)</option>
          <option value="qop1">🔥 현재분기 OP 누적 +5% 이상 (1단계↑)</option>
          <option value="qop2">🔥 현재분기 OP 누적 +10% 이상 (2단계↑)</option>
          <option value="qop3">🔥🔥 현재분기 OP 누적 +15% 이상 (3단계↑)</option>
          <option value="qop4">🔥🔥 현재분기 OP 누적 +20% 이상 (4단계)</option>
          <option value="up">▲ 예상이익 상향 전체 (직전 대비)</option>
          <option value="up5">▲ 예상이익 +5% 이상 상향</option>
          <option value="up10">▲▲ 예상이익 +10% 이상 대폭 상향</option>
          <option value="down">▼ 예상이익 하향 전체 (직전 대비)</option>
          <option value="down5">▼ 예상이익 -5% 이하 하향</option>
          <option value="down10">▼▼ 예상이익 -10% 이하 대폭 하향</option>
        </select>
        <span id="rowCountLabel" style="font-weight:700; color:#475569;"></span>
      </div>
      <div class="group-right">
        <button class="grp-btn" id="btnGrpA" onclick="toggleGroup('A')">➕ 분기OP 상세</button>
        <button class="grp-btn" id="btnGrpB" onclick="toggleGroup('B')">➕ 분기YoY 상세</button>
        <button class="grp-btn" id="btnGrpE" onclick="toggleGroup('E')">➕ 가속폭 상세</button>
        <button class="grp-btn" id="btnGrpC" onclick="toggleGroup('C')">➕ 분기이익률 상세</button>
        <button class="grp-btn" id="btnGrpD" onclick="toggleGroup('D')">➕ 연간 상세</button>
        <button class="grp-btn" id="btnGrpAll" onclick="toggleAllGroups()">📂 상세 전체 펼치기</button>
        <button class="grp-btn open" id="btnFreeze" onclick="toggleFreeze()">📌 좌측 틀고정 ON</button>
      </div>
    </div>

    <div class="table-wrap freeze-on" id="tableWrap">
      <table id="dataTable">
        <thead id="tableHead"></thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>

    <div class="pagination-bar" id="paginationBar">
      <div class="page-controls">
        <span style="font-weight:700; color:#475569;">보기 방식:</span>
        <select id="pageSizeSelect" class="select-box" onchange="onPageSizeChange()">
          <option value="0" selected>전체 스크롤 보기 (기본 · 세로 스크롤)</option>
          <option value="20">20줄씩 페이지 나누기</option>
          <option value="30">30줄씩 페이지 나누기</option>
          <option value="50">50줄씩 페이지 나누기</option>
        </select>
        <span id="pageRangeInfo" style="color:#64748b; font-weight:600;"></span>
      </div>
      <div class="legend-bar">
        <span style="color:#475569;">종목명(분기OP 누적상향):</span>
        <span class="lg-chip" style="background:#fff1f2; color:#9f1239; border-left:3px solid #fda4af;">▲5%↑</span>
        <span class="lg-chip" style="background:#fecdd3; color:#881337; border-left:3px solid #f43f5e;">▲10%↑</span>
        <span class="lg-chip" style="background:#fb7185; color:#4c0519; border-left:3px solid #be123c;">▲15%↑</span>
        <span class="lg-chip" style="background:#dc2626; color:#fff; border-left:3px solid #7f1d1d;">🔥20%↑</span>
        <span style="color:#cbd5e1; margin:0 2px;">|</span>
        <span style="color:#475569;">셀 변동:</span>
        <span class="lg-chip" style="background:#ffe4e6; color:#9f1239;">▲&lt;5%</span>
        <span class="lg-chip" style="background:#fca5a5; color:#7f1d1d;">▲5~10%</span>
        <span class="lg-chip" style="background:#dc2626; color:#fff;">▲10%↑</span>
        <span class="lg-chip" style="background:#dbeafe; color:#1e40af;">▼&lt;5%</span>
        <span class="lg-chip" style="background:#93c5fd; color:#172554;">▼5~10%</span>
        <span class="lg-chip" style="background:#1d4ed8; color:#fff;">▼10%↓</span>
      </div>
      <div class="page-controls" id="pageNavGroup">
        <button class="pg-btn" id="btnFirstPage" onclick="goToPage(1)">⏮ 처음</button>
        <button class="pg-btn" id="btnPrevPage" onclick="goToPage(currentPage - 1)">◀ 이전</button>
        <span class="pg-info" id="pageNumInfo">1 / 1</span>
        <button class="pg-btn" id="btnNextPage" onclick="goToPage(currentPage + 1)">다음 ▶</button>
        <button class="pg-btn" id="btnLastPage" onclick="goToPage(totalPages)">끝 ⏭</button>
      </div>
      <div class="page-controls" id="scrollTopGroup" style="display:none;">
        <button class="pg-btn" onclick="document.getElementById('tableWrap').scrollTop = 0">⬆ 맨 위로</button>
      </div>
    </div>

    <div class="reports-view" id="reportsView">
      <div class="controls-panel" style="margin-bottom:10px;">
        <div class="filter-left">
          <div style="display:inline-flex; align-items:center; gap:4px; background:#eff6ff; padding:3px 6px; border-radius:8px; border:1px solid #bfdbfe;">
            <button class="pg-btn" id="btnPrevRepDate" onclick="stepReportDate(1)" title="이전 날짜 리포트">◀</button>
            <select id="repDateSelect" class="select-box" style="font-weight:800; color:#1e3a8a; border-color:#93c5fd;" onchange="onReportDateChange()"></select>
            <button class="pg-btn" id="btnNextRepDate" onclick="stepReportDate(-1)" title="다음 날짜 리포트">▶</button>
          </div>
          <input type="text" id="repSearchInput" class="search-input" style="width:225px;" placeholder="🔍 종목명 / 증권사 / 내용 (전체기간 검색)" />
          <select id="repBrokerSelect" class="select-box">
            <option value="">전체 증권사</option>
          </select>
          <select id="repChangeSelect" class="select-box">
            <option value="">목표가 변동 전체</option>
            <option value="상향">▲ 목표가 상향</option>
            <option value="신규">★ 신규편입</option>
            <option value="유지">유지</option>
            <option value="하향">▼ 목표가 하향</option>
          </select>
          <select id="repSortSelect" class="select-box" style="font-weight:700; color:#1e293b;" onchange="renderReportsGrid()">
            <option value="priority" selected>정렬: 목표가상향율순 → 예상이익상향 → 괴리율순</option>
            <option value="gap">정렬: 괴리율 높은순</option>
            <option value="estUp">정렬: 예상이익 상향 우선</option>
            <option value="latest">정렬: 원문 등록순</option>
          </select>
          <span id="repCountLabel" style="font-weight:700; color:#475569;"></span>
        </div>
        <div class="group-right">
          <span style="font-size:11px; color:#64748b; font-weight:600;">🗓️ 최대 60일 보관</span>
          <a class="rep-btn" href="https://consensus.hankyung.com/analysis/list?skinType=business" target="_blank">🌐 한경 원문</a>
          <a class="rep-btn" href="https://stock.naver.com/research/company" target="_blank">💚 네이버 리서치</a>
        </div>
      </div>
      <div class="reports-grid-wrap" id="reportsGridWrap">
        <div class="reports-grid" id="reportsGrid"></div>
      </div>
    </div>
    </main>
  </div>

  <div class="tooltip-box" id="tooltipBox"></div>

  <script>
    const DB = {json_payload};
    const REPORTS_DATA = {reports_payload};
    const OP_HIST = {op_hist_payload};
    const OP_HIST_STOCKS = (OP_HIST && OP_HIST.stocks) ? OP_HIST.stocks : {{}};
    const marketKeys = Object.keys(DB);
    let currentViewMode = 'SCREENER'; // 'SCREENER' | 'REPORTS'
    let currentMarket = marketKeys.includes('KOSPI') ? 'KOSPI' : marketKeys[0];
    let currentSheetIdx = 0;
    let openGroups = {{ A: false, B: false, E: false, C: false, D: false }};
    let freezeOn = true;
    let sortColIdx = -1;
    let sortAsc = true;
    let pageSize = 0;
    let currentPage = 1;
    let totalPages = 1;
    let highlightReportCode = '';

    const REPORTS_BY_DATE = (REPORTS_DATA.byDate && Object.keys(REPORTS_DATA.byDate).length > 0)
      ? REPORTS_DATA.byDate
      : (REPORTS_DATA.date ? {{ [REPORTS_DATA.date]: (REPORTS_DATA.reports || []) }} : {{}});
    const REPORT_DATES = Object.keys(REPORTS_BY_DATE).sort().reverse();
    let currentReportDate = REPORT_DATES[0] || REPORTS_DATA.date || '';

    const REPORTS_BY_CODE = {{}};
    REPORT_DATES.forEach(dt => {{
      (REPORTS_BY_DATE[dt] || []).forEach(rep => {{
        if (rep.stock_code) {{
          if (!REPORTS_BY_CODE[rep.stock_code]) REPORTS_BY_CODE[rep.stock_code] = [];
          REPORTS_BY_CODE[rep.stock_code].push(rep);
        }}
      }});
    }});

    const STOCK_LOOKUP = {{}};
    marketKeys.forEach(mkt => {{
      const d = DB[mkt];
      const cols = d.columns.map(c => c.name);
      const priceIdx = cols.indexOf('현재가');
      const rateIdx = cols.indexOf('오늘등락율(%)');
      const passIdx = cols.indexOf('통과');
      const seqIdx = cols.indexOf('연속이익성장(연간+분기)');
      const allSheet = d.sheets[0];
      if (!allSheet) return;
      allSheet.rows.forEach(r => {{
        const cd = String(r.v[0] || '');
        if (!cd) return;
        const hasUp = r.t ? Object.values(r.t).some(v => String(v).startsWith('up')) : false;
        STOCK_LOOKUP[cd] = {{
          market: mkt,
          price: priceIdx >= 0 ? r.v[priceIdx] : null,
          rate: rateIdx >= 0 ? r.v[rateIdx] : null,
          passed: passIdx >= 0 ? !!r.v[passIdx] : false,
          seqGrowth: seqIdx >= 0 ? !!r.v[seqIdx] : false,
          hasEstUp: hasUp
        }};
      }});
    }});

    const SHEET_LABELS = {{
      "1_전체": "1_전체 (시총순)",
      "2_통과": "2_통과 ★",
      "3_E3반등": "3_E3반등",
      "4_30퍼하락": "4_30%이상하락",
      "5_통과후E3반등": "5_통과+E3반등 ★★",
      "6_연속이익성장(E3및E2)": "6_연속이익성장 ◆",
      "7_예상이익상향": "7_예상이익상향 ▲",
      "8_예상이익하향": "8_예상이익하향 ▼"
    }};

    function initApp() {{
      if (window.innerWidth <= 768) {{
        freezeOn = false;
        const wrap = document.getElementById('tableWrap');
        const btn = document.getElementById('btnFreeze');
        wrap.classList.remove('freeze-on');
        btn.classList.remove('open');
        btn.textContent = '📌 좌측 틀고정 OFF';
      }}
      renderDownloadButtons();
      renderTopDateChip();
      renderSidebar();
      selectMarket(currentMarket);

      document.getElementById('searchInput').addEventListener('input', () => {{ currentPage = 1; renderTableBody(); }});
      document.getElementById('sectorSelect').addEventListener('change', () => {{ currentPage = 1; renderTableBody(); }});
      document.getElementById('toneSelect').addEventListener('change', () => {{ currentPage = 1; renderTableBody(); }});

      initReportsFilters();
      document.getElementById('repSearchInput').addEventListener('input', renderReportsGrid);
      document.getElementById('repBrokerSelect').addEventListener('change', renderReportsGrid);
      document.getElementById('repChangeSelect').addEventListener('change', () => {{ renderSidebar(); renderReportsGrid(); }});

      document.getElementById('globalSearchInput').addEventListener('keydown', (e) => {{
        if (e.key === 'Enter') handleGlobalSearch(e.target.value);
      }});
    }}

    function toggleSidebar() {{
      if (window.innerWidth <= 768) {{
        document.body.classList.toggle('mobile-nav-open');
      }} else {{
        document.body.classList.toggle('sidebar-collapsed');
      }}
    }}

    function closeMobileNav() {{
      document.body.classList.remove('mobile-nav-open');
    }}

    function renderTopDateChip() {{
      const d = DB[currentMarket] || DB[marketKeys[0]];
      const dtStr = (d && d.updatedAt) ? d.updatedAt : (REPORTS_DATA.date || '');
      document.getElementById('topDateChip').textContent = `🕒 ${{dtStr}}`;
    }}

    function handleGlobalSearch(rawQuery) {{
      const q = String(rawQuery || '').trim();
      if (!q) return;
      const ql = q.toLowerCase();
      // 현재 시장에 매칭되면 현재 시장, 아니면 다른 시장에서 검색
      const checkMkt = (mkt) => {{
        const d = DB[mkt];
        if (!d || !d.sheets[0]) return false;
        return d.sheets[0].rows.some(r =>
          String(r.v[0] || '').toLowerCase().includes(ql) ||
          String(r.v[1] || '').toLowerCase().includes(ql)
        );
      }};
      let targetMkt = checkMkt(currentMarket) ? currentMarket : marketKeys.find(m => checkMkt(m));
      if (targetMkt) {{
        selectMarket(targetMkt);
        currentSheetIdx = 0;
        renderSheetTabs();
        document.getElementById('searchInput').value = q;
        document.getElementById('sectorSelect').value = '';
        document.getElementById('toneSelect').value = '';
        currentPage = 1;
        renderTable();
      }} else {{
        selectReportsView();
        document.getElementById('repSearchInput').value = q;
        renderReportsGrid();
      }}
    }}

    function getAllStoredReportsList() {{
      const merged = [];
      REPORT_DATES.forEach(dt => {{
        (REPORTS_BY_DATE[dt] || []).forEach(r => merged.push(r));
      }});
      return merged;
    }}

    function getActiveReportsList() {{
      if (currentReportDate === 'ALL') {{
        return getAllStoredReportsList();
      }}
      return REPORTS_BY_DATE[currentReportDate] || REPORTS_DATA.reports || [];
    }}

    function initReportsFilters() {{
      const dateSel = document.getElementById('repDateSelect');
      const totalAll = REPORT_DATES.reduce((acc, dt) => acc + (REPORTS_BY_DATE[dt] || []).length, 0);
      const dateOpts = REPORT_DATES.map((dt, i) => {{
        const cnt = (REPORTS_BY_DATE[dt] || []).length;
        const tag = i === 0 ? ' · 최신' : '';
        return `<option value="${{dt}}">📅 ${{dt}} (${{cnt}}건)${{tag}}</option>`;
      }}).join('');
      const allOpt = REPORT_DATES.length > 1
        ? `<option value="ALL">🗂️ 전체 보관기간 (최근 ${{REPORT_DATES.length}}일 · ${{totalAll}}건)</option>`
        : '';
      dateSel.innerHTML = dateOpts + allOpt;
      if (currentReportDate) dateSel.value = currentReportDate;
      refreshReportBrokers();
      updateReportDateNavButtons();
    }}

    function refreshReportBrokers() {{
      const reps = getActiveReportsList();
      const brokers = Array.from(new Set(reps.map(r => r.broker).filter(Boolean))).sort();
      const sel = document.getElementById('repBrokerSelect');
      const prev = sel.value;
      sel.innerHTML = `<option value="">전체 증권사 (${{brokers.length}})</option>` +
        brokers.map(b => `<option value="${{escapeHtml(b)}}">${{escapeHtml(b)}}</option>`).join('');
      if (brokers.includes(prev)) sel.value = prev;
    }}

    function updateReportDateNavButtons() {{
      const btnPrev = document.getElementById('btnPrevRepDate');
      const btnNext = document.getElementById('btnNextRepDate');
      if (!btnPrev || !btnNext) return;
      if (currentReportDate === 'ALL' || REPORT_DATES.length <= 1) {{
        btnPrev.disabled = true;
        btnNext.disabled = true;
        return;
      }}
      const idx = REPORT_DATES.indexOf(currentReportDate);
      btnPrev.disabled = (idx < 0 || idx >= REPORT_DATES.length - 1);
      btnNext.disabled = (idx <= 0);
    }}

    function stepReportDate(delta) {{
      if (currentReportDate === 'ALL') return;
      const idx = REPORT_DATES.indexOf(currentReportDate);
      const nextIdx = idx + delta;
      if (nextIdx >= 0 && nextIdx < REPORT_DATES.length) {{
        currentReportDate = REPORT_DATES[nextIdx];
        document.getElementById('repDateSelect').value = currentReportDate;
        onReportDateChange();
      }}
    }}

    function onReportDateChange() {{
      currentReportDate = document.getElementById('repDateSelect').value || REPORT_DATES[0] || '';
      highlightReportCode = '';
      refreshReportBrokers();
      updateReportDateNavButtons();
      if (currentViewMode === 'REPORTS') {{
        updateReportsHeaderAndBadges();
        renderSidebar();
        renderReportsGrid();
      }}
    }}

    function onPageSizeChange() {{
      pageSize = parseInt(document.getElementById('pageSizeSelect').value, 10) || 0;
      currentPage = 1;
      renderTableBody();
    }}

    function goToPage(p) {{
      if (p < 1) p = 1;
      if (p > totalPages) p = totalPages;
      currentPage = p;
      renderTableBody();
      document.getElementById('tableWrap').scrollTop = 0;
    }}

    function renderDownloadButtons() {{
      const box = document.getElementById('dlActions');
      box.innerHTML = marketKeys.map(m => {{
        const d = DB[m];
        return `<a class="dl-btn" href="${{d.downloadFile}}" download="${{d.filename}}">📥 ${{m}} 엑셀</a>`;
      }}).join('');
    }}

    function renderSidebar() {{
      const nav = document.getElementById('sidebarNav');
      const reps = getActiveReportsList();
      const repUp = reps.filter(r => r.target_change === '상향').length;
      const repNew = reps.filter(r => r.target_change === '신규').length;
      const repDown = reps.filter(r => r.target_change === '하향').length;
      const curRepChg = document.getElementById('repChangeSelect') ? document.getElementById('repChangeSelect').value : '';

      const curD = DB[currentMarket] || DB[marketKeys[0]];
      const curSheetName = (curD && curD.sheets[currentSheetIdx]) ? curD.sheets[currentSheetIdx].name : '';
      const sheetCountOf = (name) => {{
        const s = curD ? curD.sheets.find(x => x.name === name) : null;
        return s ? s.count : 0;
      }};

      const isScr = (currentViewMode === 'SCREENER');
      const isRep = (currentViewMode === 'REPORTS');

      const mktItems = marketKeys.map(m => {{
        const icon = m === 'KOSPI' ? '📊' : '🚀';
        const active = (isScr && currentMarket === m && curSheetName === '1_전체') ? 'active' : '';
        return `<button class="nav-item ${{active}}" onclick="selectMarketAndSheet('${{m}}', '1_전체')">
          <span class="nav-item-left"><span>${{icon}}</span><span>${{m}} 상위 200</span></span>
          <span class="nav-badge">${{DB[m].sheets[0] ? DB[m].sheets[0].count : 0}}</span>
        </button>`;
      }}).join('');

      const quickSheets = [
        {{ name: '2_통과', icon: '⭐', label: '메인 통과' }},
        {{ name: '6_연속이익성장(E3및E2)', icon: '💎', label: '연속이익성장' }},
        {{ name: '7_예상이익상향', icon: '🔺', label: '예상이익 상향', bCls: 'up' }},
        {{ name: '8_예상이익하향', icon: '🔻', label: '예상이익 하향', bCls: 'down' }},
        {{ name: '4_30퍼하락', icon: '📉', label: '고점대비 30%↓' }},
      ].map(q => {{
        const active = (isScr && curSheetName === q.name) ? 'active' : '';
        return `<button class="nav-item ${{active}}" onclick="selectMarketAndSheet(currentMarket, '${{q.name}}')">
          <span class="nav-item-left"><span>${{q.icon}}</span><span>${{q.label}}</span></span>
          <span class="nav-badge ${{q.bCls || ''}}">${{sheetCountOf(q.name)}}</span>
        </button>`;
      }}).join('');

      const dlLinks = marketKeys.map(m => {{
        const d = DB[m];
        return `<a class="nav-item" href="${{d.downloadFile}}" download="${{d.filename}}">
          <span class="nav-item-left"><span>📥</span><span>${{m}} 엑셀(.xlsx)</span></span>
        </a>`;
      }}).join('');

      nav.innerHTML = `
        <div class="nav-section">
          <div class="nav-section-title">한국 스크리너</div>
          ${{mktItems}}
        </div>
        <div class="nav-divider"></div>
        <div class="nav-section">
          <div class="nav-section-title">${{currentMarket}} 빠른 필터</div>
          ${{quickSheets}}
        </div>
        <div class="nav-divider"></div>
        <div class="nav-section">
          <div class="nav-section-title">기업 리포트</div>
          <button class="nav-item ${{isRep && !curRepChg ? 'active' : ''}}" onclick="selectReportsView('', '')">
            <span class="nav-item-left"><span>📰</span><span>오늘의 리포트</span></span>
            <span class="nav-badge">${{reps.length}}</span>
          </button>
          <button class="nav-item ${{isRep && curRepChg === '상향' ? 'active' : ''}}" onclick="selectReportsView('', '상향')">
            <span class="nav-item-left"><span>📈</span><span>목표가 상향</span></span>
            <span class="nav-badge up">${{repUp}}</span>
          </button>
          <button class="nav-item ${{isRep && curRepChg === '신규' ? 'active' : ''}}" onclick="selectReportsView('', '신규')">
            <span class="nav-item-left"><span>✨</span><span>신규 커버리지</span></span>
            <span class="nav-badge new">${{repNew}}</span>
          </button>
          <button class="nav-item ${{isRep && curRepChg === '하향' ? 'active' : ''}}" onclick="selectReportsView('', '하향')">
            <span class="nav-item-left"><span>📉</span><span>목표가 하향</span></span>
            <span class="nav-badge down">${{repDown}}</span>
          </button>
        </div>
        <div class="nav-divider"></div>
        <div class="nav-section">
          <div class="nav-section-title">다운로드 · 원문</div>
          ${{dlLinks}}
          <a class="nav-item" href="https://consensus.hankyung.com/analysis/list?skinType=business" target="_blank">
            <span class="nav-item-left"><span>🌐</span><span>한경 컨센서스</span></span>
          </a>
          <a class="nav-item" href="https://stock.naver.com/research/company" target="_blank">
            <span class="nav-item-left"><span>💚</span><span>네이버 리서치</span></span>
          </a>
        </div>
      `;
    }}

    function selectMarketAndSheet(mkt, sheetName) {{
      closeMobileNav();
      selectMarket(mkt);
      if (sheetName) selectSheetByName(sheetName);
    }}

    function setScreenerVisibility(showScreener) {{
      document.getElementById('sheetTabs').style.display = showScreener ? 'flex' : 'none';
      document.getElementById('screenerControls').style.display = showScreener ? 'flex' : 'none';
      document.getElementById('tableWrap').style.display = showScreener ? 'block' : 'none';
      document.getElementById('paginationBar').style.display = showScreener ? 'flex' : 'none';
      document.getElementById('reportsView').classList.toggle('active', !showScreener);
    }}

    function updateReportsHeaderAndBadges() {{
      const reps = getActiveReportsList();
      const dateLabel = currentReportDate === 'ALL'
        ? `전체 보관기간 (최근 ${{REPORT_DATES.length}}일)`
        : `${{currentReportDate || REPORTS_DATA.date || '-'}} 기준`;
      document.getElementById('pageTitle').textContent = '기업 리포트 요약';
      document.getElementById('pageSubtitle').textContent = `한경 · 네이버 증권 통합 · ${{dateLabel}} · 목표가상향 → 예상이익상향 → 괴리율순`;

      const upCnt = reps.filter(r => r.target_change === '상향').length;
      const newCnt = reps.filter(r => r.target_change === '신규').length;
      const downCnt = reps.filter(r => r.target_change === '하향').length;
      document.getElementById('metaBadges').innerHTML = `
        <span class="meta-pill">📅 ${{dateLabel}}</span>
        <span class="meta-pill">📄 총 ${{reps.length}}건</span>
        <span class="meta-pill up clickable" onclick="selectReportsView('', '상향')">▲ 목표가 상향 ${{upCnt}}건</span>
        <span class="meta-pill clickable" style="background:#ede9fe; color:#6d28d9; border-color:#c4b5fd;" onclick="selectReportsView('', '신규')">★ 신규 ${{newCnt}}건</span>
        <span class="meta-pill down clickable" onclick="selectReportsView('', '하향')">▼ 목표가 하향 ${{downCnt}}건</span>
      `;
    }}

    function selectReportsView(focusCode = '', filterChange = null) {{
      closeMobileNav();
      currentViewMode = 'REPORTS';
      highlightReportCode = focusCode || '';
      setScreenerVisibility(false);

      if (focusCode) {{
        const matched = REPORTS_BY_CODE[focusCode];
        if (matched && matched.length > 0 && matched[0].date) {{
          currentReportDate = matched[0].date;
          const dateSel = document.getElementById('repDateSelect');
          if (dateSel) dateSel.value = currentReportDate;
          refreshReportBrokers();
          updateReportDateNavButtons();
        }}
        document.getElementById('repSearchInput').value = '';
        document.getElementById('repBrokerSelect').value = '';
        document.getElementById('repChangeSelect').value = '';
      }} else if (filterChange !== null) {{
        document.getElementById('repChangeSelect').value = filterChange;
      }}
      updateReportsHeaderAndBadges();
      renderSidebar();
      renderReportsGrid();
      if (focusCode) {{
        const targetEl = document.getElementById('rep-card-' + focusCode);
        if (targetEl) targetEl.scrollIntoView({{ behavior: 'smooth', block: 'center' }});
      }}
    }}

    function jumpToStockInScreener(code, market) {{
      selectMarket(market);
      currentSheetIdx = 0;
      renderSheetTabs();
      document.getElementById('searchInput').value = code;
      document.getElementById('sectorSelect').value = '';
      document.getElementById('toneSelect').value = '';
      currentPage = 1;
      renderTable();
    }}

    function computeReportRankMeta(r) {{
      const info = STOCK_LOOKUP[r.stock_code] || null;
      const tpNum = parseFloat(String(r.target_price || '').replace(/,/g, ''));
      const hasTp = !isNaN(tpNum) && tpNum > 0;
      const curPrice = (info && info.price) ? info.price : (r.current_price || null);
      let gap = null;
      if (hasTp && curPrice && curPrice > 0) {{
        gap = ((tpNum - curPrice) / curPrice * 100);
      }} else if (r.upside_pct !== null && r.upside_pct !== undefined && !isNaN(Number(r.upside_pct))) {{
        gap = Number(r.upside_pct);
      }}
      const tpChgPct = (r.target_change_pct !== null && r.target_change_pct !== undefined && !isNaN(Number(r.target_change_pct)))
        ? Number(r.target_change_pct)
        : null;
      return {{
        info,
        tpNum,
        hasTp,
        curPrice,
        gap,
        tpChgPct,
        isTargetUp: r.target_change === '상향' ? 1 : 0,
        isEstUp: (info && info.hasEstUp) ? 1 : 0,
      }};
    }}

    function renderReportsGrid() {{
      const q = document.getElementById('repSearchInput').value.trim().toLowerCase();
      const brokerF = document.getElementById('repBrokerSelect').value;
      const chgF = document.getElementById('repChangeSelect').value;
      const sortMode = document.getElementById('repSortSelect') ? document.getElementById('repSortSelect').value : 'priority';
      const isGlobalSearch = q.length > 0 && currentReportDate !== 'ALL';
      const reps = isGlobalSearch ? getAllStoredReportsList() : getActiveReportsList();

      let items = reps.map((r, origIdx) => ({{ r, origIdx, meta: computeReportRankMeta(r) }})).filter(({{ r }}) => {{
        if (brokerF && r.broker !== brokerF) return false;
        if (chgF && r.target_change !== chgF) return false;
        if (q) {{
          const hay = [r.stock_name, r.stock_code, r.broker, r.analyst, r.title, r.date, ...(r.summary || [])].join(' ').toLowerCase();
          if (!hay.includes(q)) return false;
        }}
        return true;
      }});

      const gapVal = (m) => (m.gap !== null && !isNaN(m.gap)) ? m.gap : -999999;
      const raiseVal = (m) => (m.tpChgPct !== null && !isNaN(m.tpChgPct)) ? m.tpChgPct : -999999;
      if (sortMode === 'priority') {{
        // 1순위: 목표가 상향 (그 중에서도 상향율이 높은 기업 우선!)
        // 2순위: 예상이익 상향
        // 3순위: 괴리율 큰 순서 (동률 시 최신 날짜 우선)
        items.sort((a, b) => {{
          if (b.meta.isTargetUp !== a.meta.isTargetUp) return b.meta.isTargetUp - a.meta.isTargetUp;
          if (a.meta.isTargetUp === 1 && b.meta.isTargetUp === 1) {{
            const dr = raiseVal(b.meta) - raiseVal(a.meta);
            if (dr !== 0) return dr;
          }}
          if (b.meta.isEstUp !== a.meta.isEstUp) return b.meta.isEstUp - a.meta.isEstUp;
          const dg = gapVal(b.meta) - gapVal(a.meta);
          if (dg !== 0) return dg;
          const dd = String(b.r.date || '').localeCompare(String(a.r.date || ''));
          if (dd !== 0) return dd;
          return a.origIdx - b.origIdx;
        }});
      }} else if (sortMode === 'gap') {{
        items.sort((a, b) => {{
          const dg = gapVal(b.meta) - gapVal(a.meta);
          return dg !== 0 ? dg : (a.origIdx - b.origIdx);
        }});
      }} else if (sortMode === 'estUp') {{
        items.sort((a, b) => {{
          if (b.meta.isEstUp !== a.meta.isEstUp) return b.meta.isEstUp - a.meta.isEstUp;
          const dg = gapVal(b.meta) - gapVal(a.meta);
          return dg !== 0 ? dg : (a.origIdx - b.origIdx);
        }});
      }}

      const countEl = document.getElementById('repCountLabel');
      if (isGlobalSearch) {{
        countEl.innerHTML = `<span style="color:#1d4ed8;">전체 보관기간(${{REPORT_DATES.length}}일) 검색: ${{items.length}}건</span>`;
      }} else {{
        countEl.textContent = `표시: ${{items.length}} / ${{reps.length}}건`;
      }}
      const grid = document.getElementById('reportsGrid');
      if (!items.length) {{
        grid.innerHTML = `<div style="grid-column:1/-1; text-align:center; padding:40px; color:#64748b; background:#fff; border-radius:10px; border:1px solid #e2e8f0;">조건에 맞는 기업 리포트가 없습니다.</div>`;
        return;
      }}

      grid.innerHTML = items.map(({{ r, meta }}) => {{
        const {{ info, tpNum, hasTp, curPrice, gap }} = meta;
        let upsideHtml = '';
        if (gap !== null && !isNaN(gap)) {{
          const gapCls = gap >= 0 ? 'pos' : 'neg';
          const priceTxt = curPrice ? `현재가 ${{Math.round(curPrice).toLocaleString('ko-KR')}}원 · ` : '';
          upsideHtml = `<div style="font-size:11.5px; color:#475569;">${{priceTxt}}괴리율 <span class="${{gapCls}}" style="font-size:12.5px;">${{gap >= 0 ? '+' : ''}}${{gap.toFixed(1)}}%</span></div>`;
        }} else if (curPrice) {{
          upsideHtml = `<div style="font-size:11px; color:#64748b;">현재가 ${{Math.round(curPrice).toLocaleString('ko-KR')}}원</div>`;
        }}

        const scrBadges = [];
        if (isGlobalSearch || currentReportDate === 'ALL') {{
          scrBadges.push(`<span class="rep-pill" style="background:#eff6ff; color:#1d4ed8; border:1px solid #bfdbfe;">📅 ${{escapeHtml(r.date || '')}}</span>`);
        }}
        if (info) {{
          scrBadges.push(`<span class="rep-pill" style="background:#f1f5f9; color:#334155;">${{info.market}}</span>`);
          if (info.hasEstUp) scrBadges.push(`<span class="rep-pill scr-up">▲예상이익상향</span>`);
          if (info.passed) scrBadges.push(`<span class="rep-pill scr-pass">★통과</span>`);
          if (info.seqGrowth) scrBadges.push(`<span class="rep-pill scr-seq">◆연속성장</span>`);
        }}

        const chg = r.target_change || '유지';
        const tpPct = meta.tpChgPct;
        let chgText = chg === '상향' ? '▲ 목표가상향' : (chg === '하향' ? '▼ 목표가하향' : (chg === '신규' ? '★ 신규' : chg));
        if (chg === '상향' && tpPct !== null) {{
          chgText = `▲ 목표가 +${{Math.abs(tpPct).toFixed(1)}}%`;
        }} else if (chg === '하향' && tpPct !== null) {{
          chgText = `▼ 목표가 -${{Math.abs(tpPct).toFixed(1)}}%`;
        }}
        const prevTpTip = (r.prev_target_price && hasTp)
          ? `data-tip="직전 목표가 ${{Math.round(r.prev_target_price).toLocaleString('ko-KR')}}원 → ${{Math.round(tpNum).toLocaleString('ko-KR')}}원${{tpPct !== null ? ' (' + (tpPct >= 0 ? '+' : '') + tpPct.toFixed(1) + '%)' : ''}}"`
          : '';
        const hlCls = (highlightReportCode && highlightReportCode === r.stock_code) ? 'highlight' : '';
        const summaryLis = (r.summary || []).map(s => `<li>${{escapeHtml(s)}}</li>`).join('');
        const naverUrl = `https://finance.naver.com/item/main.naver?code=${{r.stock_code}}`;
        const fnUrl = `https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A${{r.stock_code}}`;

        return `<div class="rep-card ${{hlCls}}" id="rep-card-${{r.stock_code}}">
          <div class="rep-card-head">
            <div class="rep-stock-title">
              <span class="rep-stock-name">${{escapeHtml(r.stock_name || '기업명')}}</span>
              <span class="rep-code">(${{escapeHtml(r.stock_code || '-')}})</span>
              <span class="rep-pill chg-${{chg}}" ${{prevTpTip}}>${{chgText}}</span>
              ${{scrBadges.join('')}}
            </div>
            <div class="rep-target-box">
              <div class="rep-target-price">${{hasTp ? '목표가 ' + Math.round(tpNum).toLocaleString('ko-KR') + '원' : '목표가 미제시'}} <span style="font-size:11.5px; color:#475569;">(${{escapeHtml(r.opinion || '-')}})</span></div>
              ${{upsideHtml}}
            </div>
          </div>
          <div class="rep-headline">${{escapeHtml(r.title || '')}}</div>
          <div class="rep-meta-row">
            <span>🏢 <strong>${{escapeHtml(r.broker || '-')}}</strong></span>
            ${{r.analyst ? `<span>✍️ ${{escapeHtml(r.analyst)}}</span>` : ''}}
            <span>📅 ${{escapeHtml(r.date || '-')}}</span>
            <span style="margin-left:auto; padding:1px 6px; border-radius:4px; font-weight:800; font-size:10.5px; ${{r.source === '네이버' ? 'background:#dcfce7; color:#15803d;' : 'background:#fef3c7; color:#92400e;'}}">${{r.source === '네이버' ? '네이버' : '한경'}}</span>
          </div>
          <ul class="rep-summary-list">${{summaryLis}}</ul>
          <div class="rep-card-foot">
            <div style="display:flex; gap:6px; flex-wrap:wrap;">
              ${{r.pdf_url ? `<a class="rep-btn pdf" href="${{r.pdf_url}}" target="_blank">📄 원문 PDF 보기</a>` : ''}}
              ${{r.naver_url ? `<a class="rep-btn" href="${{r.naver_url}}" target="_blank">💚 네이버 원문</a>` : ''}}
              <a class="rep-btn" href="${{naverUrl}}" target="_blank">네이버 금융</a>
              <a class="rep-btn" href="${{fnUrl}}" target="_blank">에프앤가이드</a>
            </div>
            ${{info ? `<button class="rep-btn scr" onclick="jumpToStockInScreener('${{r.stock_code}}', '${{info.market}}')">📊 ${{info.market}} 표에서 보기</button>` : ''}}
          </div>
        </div>`;
      }}).join('');
      bindTooltips();
    }}

    function selectMarket(mkt) {{
      currentViewMode = 'SCREENER';
      currentMarket = mkt;
      sortColIdx = -1;
      currentPage = 1;
      setScreenerVisibility(true);
      const d = DB[currentMarket];

      document.getElementById('pageTitle').textContent = `${{currentMarket}} 상위 200`;
      document.getElementById('pageSubtitle').textContent = '영업이익 성장 · 수익성 · 예상이익 변동 스크리너';
      renderTopDateChip();

      document.getElementById('metaBadges').innerHTML = `
        <span class="meta-pill">🕒 업데이트: ${{d.updatedAt}}</span>
        <span class="meta-pill">📄 ${{d.filename}}</span>
        <span class="meta-pill up clickable" onclick="selectSheetByName('7_예상이익상향')" title="예상이익 상향 기업 탭으로 바로 이동">▲ 예상이익 상향 ${{d.upCount}}셀</span>
        <span class="meta-pill down clickable" onclick="selectSheetByName('8_예상이익하향')" title="예상이익 하향 기업 탭으로 바로 이동">▼ 예상이익 하향 ${{d.downCount}}셀</span>
      `;

      if (currentSheetIdx >= d.sheets.length) currentSheetIdx = 0;
      renderSidebar();
      renderSheetTabs();
      populateSectorFilter();
      renderTable();
    }}

    function selectSheetByName(targetName) {{
      const d = DB[currentMarket];
      const idx = d.sheets.findIndex(s => s.name === targetName);
      if (idx >= 0) selectSheet(idx);
    }}

    function renderSheetTabs() {{
      const d = DB[currentMarket];
      const box = document.getElementById('sheetTabs');
      box.innerHTML = d.sheets.map((s, idx) => {{
        const active = idx === currentSheetIdx ? 'active' : '';
        const extraCls = s.name === '7_예상이익상향' ? 'tab-up' : (s.name === '8_예상이익하향' ? 'tab-down' : '');
        const label = SHEET_LABELS[s.name] || s.name;
        return `<button class="sheet-btn ${{extraCls}} ${{active}}" onclick="selectSheet(${{idx}})">
          <span>${{label}}</span>
          <span class="sheet-count">${{s.count}}</span>
        </button>`;
      }}).join('');
    }}

    function selectSheet(idx) {{
      currentSheetIdx = idx;
      sortColIdx = -1;
      currentPage = 1;
      renderSidebar();
      renderSheetTabs();
      populateSectorFilter();
      renderTable();
    }}

    function populateSectorFilter() {{
      const d = DB[currentMarket];
      const sheet = d.sheets[currentSheetIdx];
      const cols = d.columns.map(c => c.name);
      const s1Idx = cols.indexOf('섹터1');
      const s2Idx = cols.indexOf('섹터2');
      const sectors = new Set();

      sheet.rows.forEach(r => {{
        if (s1Idx >= 0 && r.v[s1Idx]) sectors.add('섹터1:' + r.v[s1Idx]);
        if (s2Idx >= 0 && r.v[s2Idx]) sectors.add('섹터2:' + r.v[s2Idx]);
      }});

      const sel = document.getElementById('sectorSelect');
      const prev = sel.value;
      const sorted = Array.from(sectors).sort();
      sel.innerHTML = `<option value="">전체 섹터 (${{sorted.length}})</option>` +
        sorted.map(s => `<option value="${{s}}">${{s.replace(':', ' · ')}}</option>`).join('');
      if (sorted.includes(prev)) sel.value = prev;
    }}

    function toggleGroup(g) {{
      openGroups[g] = !openGroups[g];
      updateGroupButtons();
      renderTable();
    }}

    function toggleAllGroups() {{
      const anyClosed = !openGroups.A || !openGroups.B || !openGroups.E || !openGroups.C || !openGroups.D;
      openGroups.A = anyClosed;
      openGroups.B = anyClosed;
      openGroups.E = anyClosed;
      openGroups.C = anyClosed;
      openGroups.D = anyClosed;
      updateGroupButtons();
      renderTable();
    }}

    function updateGroupButtons() {{
      const labels = {{ A: '분기OP 상세', B: '분기YoY 상세', E: '가속폭 상세', C: '분기이익률 상세', D: '연간 상세' }};
      ['A', 'B', 'E', 'C', 'D'].forEach(g => {{
        const btn = document.getElementById('btnGrp' + g);
        const isOpen = openGroups[g];
        btn.classList.toggle('open', isOpen);
        btn.textContent = (isOpen ? '➖ ' : '➕ ') + labels[g];
      }});
      const allOpen = openGroups.A && openGroups.B && openGroups.E && openGroups.C && openGroups.D;
      const btnAll = document.getElementById('btnGrpAll');
      btnAll.classList.toggle('open', allOpen);
      btnAll.textContent = allOpen ? '📁 상세 전체 접기' : '📂 상세 전체 펼치기';
    }}

    function toggleFreeze() {{
      freezeOn = !freezeOn;
      const wrap = document.getElementById('tableWrap');
      const btn = document.getElementById('btnFreeze');
      wrap.classList.toggle('freeze-on', freezeOn);
      btn.classList.toggle('open', freezeOn);
      btn.textContent = freezeOn ? '📌 좌측 틀고정 ON' : '📌 좌측 틀고정 OFF';
    }}

    function onHeaderClick(colIdx) {{
      if (sortColIdx === colIdx) {{
        sortAsc = !sortAsc;
      }} else {{
        sortColIdx = colIdx;
        sortAsc = false;
      }}
      currentPage = 1;
      renderTable();
    }}

    const SHORT_HEADERS = {{
      "종목코드": "코드",
      "종목명": "종목명",
      "섹터1": "섹터1",
      "섹터2": "섹터2",
      "현재가": "현재가",
      "오늘등락율(%)": "등락\\n(%)",
      "52주최고가": "52주\\n최고가",
      "52주최고가대비하락률(%)": "고점대비\\n하락(%)",
      "52주최고가대비30%이상하락": "30%↓",
      "통과": "통과",
      "E3반등": "E3\\n반등",
      "통과+E3반등": "통과+\\nE3",
      "연속이익성장(연간+분기)": "연속\\n성장",
      "연속성장구간": "성장\\n구간",
      "성장조건": "성장\\n조건",
      "E3반등사유": "E3\\n사유",
      "분기OP_연속성장": "분기OP\\n연속",
      "OP성장률_최종대비E0(%)": "OP성장\\n(%)",
      "가속폭_E1-기준(%p)": "가속폭\\nE1-E0",
      "가속폭_E2-기준(%p)": "가속폭\\nE2-E0",
      "가속폭_E3-E2(%p)": "가속폭\\nE3-E2",
      "분기마진_기준<E1<E2": "분기\\n마진↑",
      "연간OP_성장률(%)": "연간OP\\n성장(%)",
      "연간OP_차년>당해": "연간OP\\n증가",
      "연간마진_개선폭(%p)": "연간마진\\n(%p)",
      "연간마진_차년>당해": "연간\\n마진↑"
    }};

    function getShortHeader(name) {{
      if (SHORT_HEADERS[name]) return SHORT_HEADERS[name];
      let s = name
        .replace(/^OP_20/, 'OP\\n')
        .replace(/^YoY_20/, 'YoY\\n')
        .replace(/^분기영업이익률_20/, '분기마진\\n')
        .replace(/^연간OP_20/, '연간OP\\n')
        .replace(/^연간영업이익률_20/, '연간마진\\n')
        .replace(/\\(기준,%\\)/g, '(기준)')
        .replace(/\\(E,%\\)/g, '(E)');
      return s;
    }}

    function formatCell(colName, val, rowVals, colIdx) {{
      if (val === null || val === undefined || val === '') return '<span class="badge-false">-</span>';
      if (typeof val === 'boolean') {{
        return val ? '<span class="badge-true">✔</span>' : '<span class="badge-false">-</span>';
      }}
      if (colName === '종목명') {{
        const code = rowVals[0] || '';
        const naverUrl = `https://finance.naver.com/item/main.naver?code=${{code}}`;
        const fnUrl = `https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp?pGB=1&gicode=A${{code}}`;
        const reps = REPORTS_BY_CODE[code] || [];
        let repBtn = '';
        if (reps.length > 0) {{
          const r0 = reps[0];
          const repTip = `[오늘의 기업 리포트 · ${{r0.broker}}]\\n${{r0.title}}\\n• 목표가: ${{r0.target_price}}원 (${{r0.target_change || r0.opinion}})\\n(클릭 시 리포트 요약 보기)`;
          repBtn = `<span class="rep-badge-link" onclick="selectReportsView('${{code}}')" data-tip="${{escapeHtml(repTip)}}">📄</span>`;
        }}
        const oh = OP_HIST_STOCKS[code];
        let cntBadge = '';
        if (oh && oh.lv >= 1) {{
          const bIcon = oh.lv >= 4 ? '🔥' : '▲';
          const pctLabel = (oh.p !== null && oh.p !== undefined) ? `+${{Math.round(oh.p)}}%` : `${{oh.lv}}`;
          cntBadge = `<span class="op-cnt-badge op-cnt-${{oh.lv}}">${{bIcon}}${{pctLabel}}</span>`;
        }}
        const titleAttr = (oh && oh.m) ? '' : `title="${{escapeHtml(String(val))}}"`;
        return `<strong ${{titleAttr}}>${{escapeHtml(String(val))}}</strong>${{cntBadge}}
          <a class="ext-link" href="${{naverUrl}}" target="_blank" title="네이버 금융 열기">N</a>
          <a class="ext-link" href="${{fnUrl}}" target="_blank" title="에프앤가이드 열기">Fn</a>${{repBtn}}`;
      }}
      if (colName === '섹터1' || colName === '섹터2') {{
        const s = escapeHtml(String(val));
        return `<span title="${{s}}">${{s}}</span>`;
      }}
      if (colName === '연속성장구간') {{
        const s = String(val);
        if (s.includes('E3미집계')) return `<span title="${{escapeHtml(s)}}">E0~E2</span>`;
        return escapeHtml(s);
      }}
      if (typeof val === 'number') {{
        if (colName === '현재가' || colName === '52주최고가') {{
          return Math.round(val).toLocaleString('ko-KR');
        }}
        if (colName === '오늘등락율(%)') {{
          const cls = val > 0 ? 'pos' : (val < 0 ? 'neg' : '');
          const sign = val > 0 ? '+' : '';
          return `<span class="${{cls}}">${{sign}}${{val.toFixed(2)}}%</span>`;
        }}
        return val.toLocaleString('ko-KR', {{ minimumFractionDigits: 0, maximumFractionDigits: 2 }});
      }}
      return escapeHtml(String(val));
    }}

    function escapeHtml(s) {{
      return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }}

    function getAutoRevealColSet(sheet) {{
      const s = new Set();
      if (sheet && (sheet.name === '7_예상이익상향' || sheet.name === '8_예상이익하향')) {{
        const wantPrefix = sheet.name === '7_예상이익상향' ? 'up' : 'down';
        sheet.rows.forEach(r => {{
          if (!r.t) return;
          Object.entries(r.t).forEach(([colIdxStr, toneVal]) => {{
            if (String(toneVal).startsWith(wantPrefix)) s.add(parseInt(colIdxStr, 10));
          }});
        }});
      }}
      return s;
    }}

    function renderTable() {{
      const d = DB[currentMarket];
      const cols = d.columns;
      const sheet = d.sheets[currentSheetIdx];
      const autoRevealCols = getAutoRevealColSet(sheet);
      const thead = document.getElementById('tableHead');

      let trHtml = '<tr>';
      cols.forEach((c, idx) => {{
        if (c.group && !openGroups[c.group] && !autoRevealCols.has(idx)) return;
        const classes = [];
        if (idx <= 3) classes.push(`sticky-col-${{idx}}`);
        if (c.borderL) classes.push('b-left');
        if (c.borderR) classes.push('b-right');
        const sortIcon = (sortColIdx === idx) ? (sortAsc ? ' ▲' : ' ▼') : '';
        const shortLabel = getShortHeader(c.name);
        const fullTip = c.comment
          ? `[${{c.name}}]\\n${{c.comment}}`
          : (shortLabel.replace(/\\n/g, '') !== c.name ? `[${{c.name}}]` : '');
        if (fullTip) classes.push('has-tip');
        const commentAttr = fullTip ? `data-tip="${{escapeHtml(fullTip)}}"` : '';
        trHtml += `<th class="${{classes.join(' ')}}" onclick="onHeaderClick(${{idx}})" ${{commentAttr}}>${{escapeHtml(shortLabel)}}${{sortIcon}}</th>`;
      }});
      trHtml += '</tr>';
      thead.innerHTML = trHtml;

      renderTableBody();
      bindTooltips();
    }}

    function renderTableBody() {{
      const d = DB[currentMarket];
      const cols = d.columns;
      const sheet = d.sheets[currentSheetIdx];
      const tbody = document.getElementById('tableBody');

      const q = document.getElementById('searchInput').value.trim().toLowerCase();
      const secFilter = document.getElementById('sectorSelect').value;
      const toneFilter = document.getElementById('toneSelect').value;

      const s1Idx = cols.findIndex(c => c.name === '섹터1');
      const s2Idx = cols.findIndex(c => c.name === '섹터2');

      let filtered = sheet.rows.filter(r => {{
        if (q) {{
          const code = String(r.v[0] || '').toLowerCase();
          const name = String(r.v[1] || '').toLowerCase();
          const s1 = s1Idx >= 0 ? String(r.v[s1Idx] || '').toLowerCase() : '';
          const s2 = s2Idx >= 0 ? String(r.v[s2Idx] || '').toLowerCase() : '';
          if (!code.includes(q) && !name.includes(q) && !s1.includes(q) && !s2.includes(q)) return false;
        }}
        if (secFilter) {{
          const [secType, secVal] = secFilter.split(':');
          if (secType === '섹터1' && String(r.v[s1Idx] || '') !== secVal) return false;
          if (secType === '섹터2' && String(r.v[s2Idx] || '') !== secVal) return false;
        }}
        if (toneFilter) {{
          if (toneFilter.startsWith('qop')) {{
            const cd = String(r.v[0] || '');
            const oh = OP_HIST_STOCKS[cd];
            const minLv = parseInt(toneFilter.replace('qop', ''), 10) || 1;
            if (!oh || oh.lv < minLv) return false;
          }} else {{
            if (!r.t) return false;
            const tones = Object.values(r.t);
            if (toneFilter === 'up' && !tones.some(t => t.startsWith('up'))) return false;
            if (toneFilter === 'up5' && !tones.some(t => t === 'up2' || t === 'up3')) return false;
            if (toneFilter === 'up10' && !tones.some(t => t === 'up3')) return false;
            if (toneFilter === 'down' && !tones.some(t => t.startsWith('down'))) return false;
            if (toneFilter === 'down5' && !tones.some(t => t === 'down2' || t === 'down3')) return false;
            if (toneFilter === 'down10' && !tones.some(t => t === 'down3')) return false;
          }}
        }}
        return true;
      }});

      if (sortColIdx >= 0) {{
        filtered = filtered.slice().sort((a, b) => {{
          const va = a.v[sortColIdx];
          const vb = b.v[sortColIdx];
          if (va === null || va === undefined) return 1;
          if (vb === null || vb === undefined) return -1;
          if (typeof va === 'number' && typeof vb === 'number') {{
            return sortAsc ? (va - vb) : (vb - va);
          }}
          return sortAsc
            ? String(va).localeCompare(String(vb), 'ko')
            : String(vb).localeCompare(String(va), 'ko');
        }});
      }}

      document.getElementById('rowCountLabel').textContent = `표시 종목: ${{filtered.length}} / ${{sheet.count}}개`;

      let pageRows = filtered;
      const pageNavGroup = document.getElementById('pageNavGroup');
      const scrollTopGroup = document.getElementById('scrollTopGroup');
      if (pageSize > 0) {{
        totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
        if (currentPage > totalPages) currentPage = totalPages;
        const startIdx = (currentPage - 1) * pageSize;
        const endIdx = Math.min(filtered.length, startIdx + pageSize);
        pageRows = filtered.slice(startIdx, endIdx);
        document.getElementById('pageRangeInfo').textContent =
          filtered.length > 0 ? `총 ${{filtered.length}}개 중 ${{startIdx + 1}} ~ ${{endIdx}}번 표시` : '표시할 종목 없음';
        pageNavGroup.style.display = 'flex';
        scrollTopGroup.style.display = 'none';
      }} else {{
        totalPages = 1;
        currentPage = 1;
        document.getElementById('pageRangeInfo').textContent =
          `총 ${{filtered.length}}개 종목 · 세로 스크롤바/마우스 휠로 전체 확인`;
        pageNavGroup.style.display = 'none';
        scrollTopGroup.style.display = 'flex';
      }}

      document.getElementById('pageNumInfo').textContent = `${{currentPage}} / ${{totalPages}} 페이지`;
      document.getElementById('btnFirstPage').disabled = (currentPage <= 1);
      document.getElementById('btnPrevPage').disabled = (currentPage <= 1);
      document.getElementById('btnNextPage').disabled = (currentPage >= totalPages);
      document.getElementById('btnLastPage').disabled = (currentPage >= totalPages);

      const autoRevealCols = getAutoRevealColSet(sheet);
      const rowsHtml = pageRows.map(r => {{
        let cells = '<tr>';
        cols.forEach((c, idx) => {{
          if (c.group && !openGroups[c.group] && !autoRevealCols.has(idx)) return;
          const val = r.v[idx];
          const classes = [];
          if (idx <= 3) classes.push(`sticky-col-${{idx}}`);
          if (c.borderL) classes.push('b-left');
          if (c.borderR) classes.push('b-right');

          if (typeof val === 'number') classes.push('num');
          else if (typeof val === 'boolean' || idx === 0) classes.push('center');

          const tone = r.t ? r.t[String(idx)] : null;
          if (tone) classes.push(`tone-${{tone}}`);

          let note = r.n ? r.n[String(idx)] : null;
          if (c.name === '종목명') {{
            const cd = String(r.v[0] || '');
            const oh = OP_HIST_STOCKS[cd];
            if (oh) {{
              if (oh.lv >= 1) classes.push(`name-up${{oh.lv}}`);
              if (oh.m) note = oh.m;
            }}
          }}
          const tipAttr = note ? `data-tip="${{escapeHtml(note)}}"` : '';
          const arrow = tone ? (tone.startsWith('up') ? ' ▲' : ' ▼') : '';

          cells += `<td class="${{classes.join(' ')}}" ${{tipAttr}}>${{formatCell(c.name, val, r.v, idx)}}${{arrow}}</td>`;
        }});
        cells += '</tr>';
        return cells;
      }}).join('');

      tbody.innerHTML = rowsHtml || `<tr><td colspan="30" style="text-align:center; padding:30px; color:#64748b;">조건에 맞는 종목이 없습니다.</td></tr>`;
      bindTooltips();
    }}

    function bindTooltips() {{
      const tipBox = document.getElementById('tooltipBox');
      document.querySelectorAll('[data-tip]').forEach(el => {{
        el.onmouseenter = (e) => {{
          const txt = el.getAttribute('data-tip');
          if (!txt) return;
          tipBox.textContent = txt;
          tipBox.style.display = 'block';
          positionTip(e);
        }};
        el.onmousemove = (e) => positionTip(e);
        el.onmouseleave = () => {{ tipBox.style.display = 'none'; }};
      }});
    }}

    function positionTip(e) {{
      const tipBox = document.getElementById('tooltipBox');
      let x = e.clientX + 14;
      let y = e.clientY + 14;
      if (x + 360 > window.innerWidth) x = Math.max(10, e.clientX - 350);
      if (y + 180 > window.innerHeight) y = Math.max(10, e.clientY - 170);
      tipBox.style.left = x + 'px';
      tipBox.style.top = y + 'px';
    }}

    window.addEventListener('DOMContentLoaded', initApp);
  </script>
</body>
</html>
"""

    out_html = os.path.join(here, "index.html")
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(html_content)

    print(f" ★ 웹페이지 생성 완료: {out_html}")
    print(f"{'='*80}")
    return out_html


def auto_push_to_github(here, commit_msg=None):
    """
    로컬 PC에서 실행 완료 후 변경된 웹페이지(index.html), 소스, 최신 엑셀 파일을
    GitHub 저장소(origin/main)에 자동으로 커밋 & 푸시합니다.
    (GitHub Actions 환경에서는 워크플로우가 직접 푸시하므로 자동 건너뜁니다.)
    """
    if os.environ.get("GITHUB_ACTIONS") == "true":
        return True

    git_dir = os.path.join(here, ".git")
    if not os.path.isdir(git_dir):
        print("[GitHub 자동 업로드] .git 폴더가 없어 업로드를 건너뜁니다.")
        return False

    print(f"\n{'='*80}")
    print(" ▶ GitHub 자동 업로드(git push) 시작...")

    def run_git(args):
        return subprocess.run(
            ["git", "-C", here] + args,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

    try:
        upload_targets = [
            "kospi200_yoy.py",
            "index.html",
            "hankyung_reports.json",
            "op_estimate_history.json",
            "requirements.txt",
            "kospi_top200_latest.xlsx",
            "kosdaq_top200_latest.xlsx",
            ".gitignore",
            os.path.join(".github", "workflows", "main.yml"),
        ]
        existing_targets = [f for f in upload_targets if os.path.exists(os.path.join(here, f))]
        if existing_targets:
            r_add = run_git(["add"] + existing_targets)
            if r_add.returncode != 0:
                print(f"  [경고] git add 실패: {r_add.stderr.strip()}")
                return False

        r_diff = run_git(["diff", "--staged", "--quiet"])
        if r_diff.returncode != 0:
            msg = commit_msg or f"Update stock screener ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
            r_commit = run_git(["commit", "-m", msg])
            if r_commit.returncode != 0:
                print(f"  [경고] git commit 실패: {r_commit.stderr.strip()}")
                return False
            print(f"  - 변경사항 커밋 완료: {msg}")
        else:
            print("  - 새로 커밋할 변경사항이 없습니다. 원격 동기화만 확인합니다.")

        # GitHub Actions가 먼저 업데이트한 커밋이 있어도 충돌 없이 최신 로컬 결과 우선 반영
        r_fetch = run_git(["fetch", "origin", "main"])
        if r_fetch.returncode == 0:
            r_rev = run_git(["rev-list", "--count", "HEAD..origin/main"])
            behind = int((r_rev.stdout or "0").strip() or "0")
            if behind > 0:
                print(f"  - 원격(GitHub)의 최신 변경사항 {behind}건과 자동 병합 중...")
                r_rb = run_git(["rebase", "-X", "theirs", "origin/main"])
                if r_rb.returncode != 0:
                    run_git(["rebase", "--abort"])
                    print(f"  [경고] 자동 병합 실패: {r_rb.stderr.strip()}")
                    return False

        r_push = subprocess.run(["git", "-C", here, "push", "origin", "main"])
        if r_push.returncode == 0:
            print(" ★ GitHub 자동 업로드 완료! (약 1분 뒤 웹페이지에 자동 반영됩니다)")
            print("   👉 https://rotaerc.github.io/Stock-Screener/")
            print(f"{'='*80}")
            return True
        else:
            print("  [경고] GitHub 업로드(push) 실패: GitHub 로그인이 필요하거나 네트워크 오류가 발생했습니다.")
            print(f"{'='*80}")
            return False
    except Exception as e:
        print(f"  [경고] GitHub 자동 업로드 중 오류: {e}")
        print(f"{'='*80}")
        return False


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    args = [a for a in sys.argv[1:] if a.lower() != "--no-push"]
    skip_push = any(a.lower() == "--no-push" for a in sys.argv[1:])

    # GitHub 업로드만 즉시 실행 (예: python kospi200_yoy.py --push)
    if args and args[0].lower() in ("--push", "push"):
        auto_push_to_github(here)
        return

    # 오늘자(또는 최근 N일) 한경 기업 리포트 수집·정리 후 웹페이지 최신화 + GitHub 업로드
    # 예: python kospi200_yoy.py --reports  또는  python kospi200_yoy.py --reports 7
    if args and args[0].lower() in ("--reports", "reports", "--report", "--reports-backfill"):
        lb_days = int(args[1]) if (len(args) > 1 and args[1].isdigit()) else 0
        fetch_hankyung_reports(here, lookback_days=lb_days)
        build_web_dashboard(here)
        if not skip_push:
            auto_push_to_github(here)
        return

    # 기존 최신 엑셀(.xlsx) 파일만으로 웹페이지(index.html) 즉시 생성 + GitHub 자동 업로드
    if args and args[0].lower() in ("--html", "html", "--web", "web"):
        build_web_dashboard(here)
        if not skip_push:
            auto_push_to_github(here)
        return

    sess = make_session()

    # 단일 종목 빠른 테스트 지원 (예: python kospi200_yoy.py 000660)
    if args and args[0].isdigit():
        code = args[0].zfill(6)
        q = _fetch_snp_q(code, sess)
        a = _fetch(code, sess, "Y")
        price_info = get_price_info_from_snapshot(code, sess)
        rec, R, (Y1, Y2) = build_record(code, code, q, a, price_info)
        print(f"[{code}] 섹터1: {rec.get('섹터1')} | 섹터2: {rec.get('섹터2')}")
        print(f"[{code}] 인식된 분기: {[p[0] + ('E' if p[1] else '') for p in q['periods']]}")
        print(f"  적용 역할: E0(기준)={R['base']} | E1={R['E1']} | E2={R['E2']} | E3={R['E3']}")
        if R.get("announced_e1"):
            ym_a = R["announced_e1"]
            print(f"  ★실적시즌 동기화: OP_{ym_a}={rec.get(f'OP_{ym_a}')} -> OP_{ym_a}(E)={rec.get(f'OP_{ym_a}(E)')}")
        print(f"  연간OP({Y1}->{Y2}): {rec.get(f'연간OP_{Y1}(E)')} -> {rec.get(f'연간OP_{Y2}(E)')} (성장: {rec.get('연간OP_차년>당해')})")
        print(f"  연간마진개선: {rec.get('연간마진_차년>당해')} | 분기OP연속성장: {rec.get('분기OP_연속성장')} ({rec.get('연속성장구간')})")
        print(f"  최종 판정: 메인통과={'★통과' if rec.get('통과') else '미통과'} | 연속이익성장(연간+분기)={'◆통과' if rec.get('연속이익성장(연간+분기)') else '미통과'}")
        return

    # 명령줄 인자로 특정 시장만 지정 가능 (예: python kospi200_yoy.py KOSPI), 미지정 시 KOSPI + KOSDAQ 모두 실행
    if args and args[0].upper() in ("KOSPI", "KOSDAQ"):
        target_markets = [args[0].upper()]
    else:
        target_markets = MARKETS

    saved_files = []
    for market in target_markets:
        saved_path = run_market(market, sess, here)
        saved_files.append((market, saved_path))

    print(f"\n{'='*80}")
    print(" ★ 모든 시장 통합 엑셀(.xlsx) 생성 완료 요약")
    for mkt, path in saved_files:
        print(f"  - [{mkt}] {os.path.basename(path)}")
    print(f"{'='*80}")

    # 크롤링 및 엑셀 생성 완료 후 오늘자 한경 기업 리포트 수집 + 웹페이지(index.html) 최신화 + GitHub 자동 업로드
    fetch_hankyung_reports(here)
    build_web_dashboard(here)
    if not skip_push:
        auto_push_to_github(here)


if __name__ == "__main__":
    main()