# -*- coding: utf-8 -*-
"""
키움증권 REST API 투자자별 수급(외국인 · 연기금 · 사모펀드) 자동 수집기
- 매일(월~금) 16:20 장마감 확정 수급(일간) + 주간 누적(매주 금요일 1주 확정 집계) 수집
- 수집 결과를 investor_flow.json 에 저장하고 index.html 웹 대시보드 생성 + GitHub Pages 자동 업로드

사용법:
  python kiwoom_flow.py --myip          # 키움 API에 등록할 현재 PC의 공인 IP 확인
  python kiwoom_flow.py --test          # .env 키 및 API 연결·샘플 응답 테스트
  python kiwoom_flow.py                 # 오늘 일간 + 이번주 주간 수급 수집 + 웹페이지 갱신 + GitHub 업로드
  python kiwoom_flow.py --date 20261126 # 특정 날짜 수급 수집
  python kiwoom_flow.py --no-push       # 웹페이지 생성까지만 수행 (GitHub 업로드 생략)
"""

import json
import os
import re
import sys
import time
from datetime import datetime, timedelta

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
TOKEN_CACHE_PATH = os.path.join(HERE, "kiwoom_token_cache.json")
FLOW_JSON_PATH = os.path.join(HERE, "investor_flow.json")

# 키움 REST API ka10058 (투자자별일별매매종목요청) 투자자 코드
INVESTORS = [
    ("foreign", "외국인", "9000"),
    ("pension", "연기금", "6000"),
    ("private", "사모펀드", "3100"),
]

MARKETS = [
    ("KOSPI", "001"),
    ("KOSDAQ", "101"),
]

RETENTION_DAYS = 60
RETENTION_WEEKS = 14
TOP_N_PER_SIDE = 20

# 국내 상장 ETF / ETN 제외 필터 (개별 종목 수급만 순수 집계)
ETF_BRAND_PREFIXES = (
    "KODEX ", "TIGER ", "RISE ", "KBSTAR ", "ACE ", "KINDEX ", "SOL ", "HANARO ",
    "PLUS ", "ARIRANG ", "KIWOOM ", "KOSEF ", "히어로즈 ", "TIMEFOLIO ", "TIME ",
    "WON ", "KOACT ", "FOCUS ", "TREX ", "HK ", "UNICORN ", "BNK ", "1Q ",
    "TRUSTON ", "마이티 ", "에셋플러스 ", "파워 ", "마이다스 ", "WOORI ", "SMART ",
    "VITA ", "MASTER ", "KTOP ",
)
ETF_KEYWORDS = (
    "ETF", "ETN", "액티브", "(합성", "인버스", "레버리지", "200TR",
    "선물(", "채권(", "국고채", "통안채", "회사채", "CD금리", "KOFR", "머니마켓", "단기채",
)


def is_etf_or_etn(name):
    nm = str(name or "").strip()
    if not nm:
        return False
    up = nm.upper()
    if any(up.startswith(p.upper()) for p in ETF_BRAND_PREFIXES):
        return True
    if any(k.upper() in up for k in ETF_KEYWORDS):
        return True
    if re.search(r"\bTR$", up):
        return True
    return False


def load_env(env_path=ENV_PATH):
    """외부 패키지 없이 .env 파일을 안전하게 읽습니다."""
    cfg = {}
    if os.path.isfile(env_path):
        with open(env_path, "r", encoding="utf-8-sig") as f:
            for raw_line in f:
                line = raw_line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k:
                    cfg[k] = v
    # 환경변수가 직접 설정된 경우 우선 적용
    for key in ("KIWOOM_APPKEY", "KIWOOM_SECRETKEY", "KIWOOM_ENV"):
        if os.environ.get(key):
            cfg[key] = os.environ[key].strip()
    return cfg


def get_base_url(env_mode="prod"):
    if str(env_mode or "").lower() in ("mock", "vts", "paper", "모의"):
        return "https://mockapi.kiwoom.com"
    return "https://api.kiwoom.com"


def get_public_ip():
    for url in ("https://api.ipify.org?format=json", "https://ifconfig.me/ip"):
        try:
            r = requests.get(url, timeout=8)
            if r.status_code == 200:
                txt = r.text.strip()
                if txt.startswith("{"):
                    return r.json().get("ip", "")
                return txt
        except Exception:
            continue
    return ""


def _parse_signed_num(raw):
    if raw is None:
        return None
    s = str(raw).strip().replace(",", "")
    if not s:
        return None
    # 키움 API가 '--1234' 또는 '+-1234' 형태로 부호를 중복 반환하는 경우 정규화
    neg = s.count("-") > 0
    s_clean = re.sub(r"[^0-9.]", "", s)
    if not s_clean:
        return None
    try:
        val = float(s_clean)
        return -val if neg else val
    except Exception:
        return None


def _parse_abs_num(raw):
    v = _parse_signed_num(raw)
    return abs(v) if v is not None else None


def get_kiwoom_token(appkey, secretkey, env_mode="prod", force_refresh=False):
    """키움 REST API OAuth2 접근토큰 발급 (au10001) + 로컬 캐시"""
    base_url = get_base_url(env_mode)
    now = datetime.now()

    if not force_refresh and os.path.isfile(TOKEN_CACHE_PATH):
        try:
            with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
                cached = json.load(f)
            if (
                cached.get("env") == env_mode
                and cached.get("appkey_tail") == appkey[-6:]
                and cached.get("token")
                and cached.get("expires_at")
            ):
                exp = datetime.strptime(cached["expires_at"], "%Y-%m-%d %H:%M:%S")
                if exp > now + timedelta(minutes=10):
                    return cached["token"]
        except Exception:
            pass

    url = f"{base_url}/oauth2/token"
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "api-id": "au10001",
    }
    payload = {
        "grant_type": "client_credentials",
        "appkey": appkey,
        "secretkey": secretkey,
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=15)
    try:
        data = resp.json()
    except Exception:
        raise RuntimeError(f"토큰 응답 파싱 실패 (HTTP {resp.status_code}): {resp.text[:300]}")

    token = data.get("token") or data.get("access_token")
    if not token:
        msg = data.get("return_msg") or data.get("msg") or str(data)
        raise RuntimeError(
            f"키움 토큰 발급 실패 (HTTP {resp.status_code}, code={data.get('return_code')}): {msg}\n"
            "  ※ openapi.kiwoom.com 에서 현재 PC 공인 IP가 등록되었는지, AppKey/SecretKey가 정확한지 확인해 주세요."
        )

    # 만료 시각 파싱 (예: '20261127083000' 또는 기본 12시간)
    exp_dt = now + timedelta(hours=12)
    raw_exp = str(data.get("expires_dt") or "").strip()
    if re.fullmatch(r"\d{14}", raw_exp):
        try:
            exp_dt = datetime.strptime(raw_exp, "%Y%m%d%H%M%S")
        except Exception:
            pass

    try:
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(
                {
                    "env": env_mode,
                    "appkey_tail": appkey[-6:],
                    "token": token,
                    "expires_at": exp_dt.strftime("%Y-%m-%d %H:%M:%S"),
                },
                f,
                ensure_ascii=False,
                indent=2,
            )
    except Exception:
        pass

    return token


def fetch_ka10058(token, start_dt, end_dt, invsr_tp, trde_tp, mrkt_tp, env_mode="prod", stex_tp="3"):
    """
    키움 REST API ka10058 (투자자별일별매매종목요청) 호출
    - start_dt, end_dt: 'YYYYMMDD'
    - trde_tp: '2' (순매수), '1' (순매도)
    - mrkt_tp: '001' (코스피), '101' (코스닥)
    - invsr_tp: '9000' (외국인), '6000' (연기금), '3100' (사모펀드)
    - stex_tp: '3' (통합), '1' (KRX)
    """
    base_url = get_base_url(env_mode)
    url = f"{base_url}/api/dostk/stkinfo"
    headers = {
        "Content-Type": "application/json;charset=UTF-8",
        "authorization": f"Bearer {token}",
        "api-id": "ka10058",
        "cont-yn": "N",
        "next-key": "",
    }
    body = {
        "strt_dt": start_dt,
        "end_dt": end_dt,
        "trde_tp": str(trde_tp),
        "mrkt_tp": str(mrkt_tp),
        "invsr_tp": str(invsr_tp),
        "stex_tp": str(stex_tp),
    }

    resp = requests.post(url, headers=headers, json=body, timeout=15)
    try:
        data = resp.json()
    except Exception:
        raise RuntimeError(f"ka10058 응답 파싱 실패 (HTTP {resp.status_code}): {resp.text[:300]}")

    ret_code = data.get("return_code")
    if ret_code not in (None, 0, "0"):
        raise RuntimeError(f"ka10058 오류 (code={ret_code}): {data.get('return_msg') or data}")

    # 응답 리스트 키 탐색 (기본: invsr_daly_trde_stk)
    items = data.get("invsr_daly_trde_stk")
    if not isinstance(items, list):
        items = next((v for v in data.values() if isinstance(v, list)), [])

    mkt_label = "KOSPI" if str(mrkt_tp) == "001" else "KOSDAQ"
    parsed = []
    for row in items:
        if not isinstance(row, dict):
            continue
        raw_cd = str(row.get("stk_cd") or row.get("isu_cd") or "").strip()
        # 키움 종목코드 접두사 'A' 또는 '_AL' 등 제거
        m_cd = re.search(r"(\d{6})", raw_cd)
        if not m_cd:
            continue
        code = m_cd.group(1)
        name = str(row.get("stk_nm") or "").strip()
        if not name or is_etf_or_etn(name):
            continue

        # 순매매금액 (ka10058 기본 단위: 백만원 -> 억원 변환: ÷ 100)
        raw_amt = (
            row.get("netslmt_amt")
            or row.get("netprps_amt")
            or row.get("net_amt")
            or row.get("trde_amt")
        )
        raw_qty = (
            row.get("netslmt_qty")
            or row.get("netprps_qty")
            or row.get("net_qty")
        )
        amt_mil = _parse_signed_num(raw_amt)
        qty_val = _parse_signed_num(raw_qty)
        avg_prc = _parse_abs_num(row.get("prsm_avg_pric"))
        cur_prc = _parse_abs_num(row.get("cur_prc"))
        flu_rt = _parse_signed_num(row.get("flu_rt"))

        # trde_tp='2'(순매수)는 양수, '1'(순매도)는 음수로 부호 일관화
        if amt_mil is not None:
            amt_eok = round(abs(amt_mil) / 100.0, 1)
            amt_eok = amt_eok if str(trde_tp) == "2" else -amt_eok
        else:
            amt_eok = 0.0

        if qty_val is not None:
            qty_int = int(round(abs(qty_val)))
            qty_int = qty_int if str(trde_tp) == "2" else -qty_int
        else:
            qty_int = 0

        if amt_eok == 0.0 and qty_int == 0:
            continue

        parsed.append(
            {
                "code": code,
                "name": name,
                "market": mkt_label,
                "netAmt": amt_eok,       # 단위: 억원 (+순매수 / -순매도)
                "netQty": qty_int,       # 단위: 주
                "avgPrice": int(round(avg_prc)) if avg_prc is not None else None,
                "price": int(round(cur_prc)) if cur_prc is not None else None,
                "chgPct": round(flu_rt, 2) if flu_rt is not None else None,
            }
        )

    # 금액 절대값 기준 상위 정렬
    parsed.sort(key=lambda x: (abs(x["netAmt"]), abs(x["netQty"])), reverse=True)
    return parsed


def collect_period_snapshot(token, start_dt, end_dt, env_mode="prod", top_n=TOP_N_PER_SIDE):
    """지정 기간(start_dt~end_dt, YYYYMMDD)에 대해 외국인·연기금·사모펀드의 KOSPI/KOSDAQ 순매수·순매도 수집"""
    s_fmt = f"{start_dt[:4]}-{start_dt[4:6]}-{start_dt[6:8]}"
    e_fmt = f"{end_dt[:4]}-{end_dt[4:6]}-{end_dt[6:8]}"
    investors_out = {}

    for inv_key, inv_label, inv_code in INVESTORS:
        buy_merged = []
        sell_merged = []
        for mkt_name, mkt_code in MARKETS:
            try:
                rows_buy = fetch_ka10058(token, start_dt, end_dt, inv_code, "2", mkt_code, env_mode=env_mode)
                buy_merged.extend(rows_buy[:top_n])
            except Exception as e:
                print(f"  [경고] {inv_label} {mkt_name} 순매수 조회 실패: {e}")
            time.sleep(0.28)

            try:
                rows_sell = fetch_ka10058(token, start_dt, end_dt, inv_code, "1", mkt_code, env_mode=env_mode)
                sell_merged.extend(rows_sell[:top_n])
            except Exception as e:
                print(f"  [경고] {inv_label} {mkt_name} 순매도 조회 실패: {e}")
            time.sleep(0.28)

        # 종목코드 중복 제거 후 시장별 최대 top_n개 유지 + 금액순 정렬
        def _dedup_sort(lst, is_buy):
            by_code = {}
            for it in lst:
                cd = it["code"]
                if cd not in by_code or abs(it["netAmt"]) > abs(by_code[cd]["netAmt"]):
                    by_code[cd] = it
            kp = [x for x in by_code.values() if x.get("market") == "KOSPI"]
            kq = [x for x in by_code.values() if x.get("market") != "KOSPI"]
            kp.sort(key=lambda x: x["netAmt"], reverse=is_buy)
            kq.sort(key=lambda x: x["netAmt"], reverse=is_buy)
            res = kp[:top_n] + kq[:top_n]
            res.sort(key=lambda x: x["netAmt"], reverse=is_buy)
            return res

        investors_out[inv_key] = {
            "label": inv_label,
            "buy": _dedup_sort(buy_merged, True),
            "sell": _dedup_sort(sell_merged, False),
        }
        print(
            f"  - [{inv_label}] ({s_fmt}~{e_fmt}) 순매수 {len(investors_out[inv_key]['buy'])}건 / "
            f"순매도 {len(investors_out[inv_key]['sell'])}건"
        )

    return {
        "startDate": s_fmt,
        "endDate": e_fmt,
        "updatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "investors": investors_out,
    }


def load_investor_flow(here=HERE):
    p = os.path.join(here, "investor_flow.json")
    if os.path.isfile(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                data.setdefault("byDate", {})
                data.setdefault("byWeek", {})
                data["dates"] = sorted(data["byDate"].keys(), reverse=True)
                data["weeks"] = sorted(data["byWeek"].keys(), reverse=True)
                return data
        except Exception:
            pass
    return {
        "updatedAt": "",
        "latestDate": "",
        "latestWeek": "",
        "dates": [],
        "weeks": [],
        "byDate": {},
        "byWeek": {},
    }


def _resolve_trading_date(raw_yyyymmdd=None):
    if raw_yyyymmdd and re.fullmatch(r"\d{8}", str(raw_yyyymmdd)):
        dt = datetime.strptime(str(raw_yyyymmdd), "%Y%m%d")
    else:
        dt = datetime.now()
        # 주말 실행 시 직전 금요일로 보정
        if dt.weekday() == 5:
            dt -= timedelta(days=1)
        elif dt.weekday() == 6:
            dt -= timedelta(days=2)
    return dt


def update_investor_flow(here=HERE, target_yyyymmdd=None, include_week=True):
    cfg = load_env(os.path.join(here, ".env"))
    appkey = cfg.get("KIWOOM_APPKEY", "")
    secretkey = cfg.get("KIWOOM_SECRETKEY", "")
    env_mode = cfg.get("KIWOOM_ENV", "prod")

    if not appkey or not secretkey or "여기에_" in appkey or "여기에_" in secretkey:
        raise RuntimeError(
            "키움 REST API 키가 설정되지 않았습니다.\n"
            f"  1) '{os.path.join(here, '.env.example')}' 파일을 복사해 '.env' 로 저장하세요.\n"
            "  2) KIWOOM_APPKEY / KIWOOM_SECRETKEY 값을 입력한 뒤 다시 실행해 주세요."
        )

    dt = _resolve_trading_date(target_yyyymmdd)
    ymd = dt.strftime("%Y%m%d")
    ymd_dash = dt.strftime("%Y-%m-%d")
    is_friday = (dt.weekday() == 4)

    print(f"\n{'='*80}")
    print(f" ▶ 키움 REST API 투자자 수급 수집 시작 ({ymd_dash}, env={env_mode})")
    print(f"{'='*80}")

    token = get_kiwoom_token(appkey, secretkey, env_mode=env_mode)
    store = load_investor_flow(here)

    # 1) 일간 수급 수집 (당일 장마감 확정치)
    print(f"[1/2] 일간 수급 집계 ({ymd_dash})...")
    daily_snap = collect_period_snapshot(token, ymd, ymd, env_mode=env_mode)
    has_daily = any(
        len(v.get("buy") or []) + len(v.get("sell") or []) > 0
        for v in daily_snap["investors"].values()
    )
    if has_daily:
        store["byDate"][ymd_dash] = daily_snap
    else:
        print(f"  [알림] {ymd_dash} 일간 수급 데이터가 아직 없습니다(장마감 16:20 이후 확정). 직전 영업일 일간 수급을 확인합니다...")
        prev_dt = dt - timedelta(days=1)
        while prev_dt.weekday() >= 5:
            prev_dt -= timedelta(days=1)
        prev_ymd = prev_dt.strftime("%Y%m%d")
        prev_dash = prev_dt.strftime("%Y-%m-%d")
        if prev_dash not in store["byDate"]:
            prev_snap = collect_period_snapshot(token, prev_ymd, prev_ymd, env_mode=env_mode)
            has_prev = any(
                len(v.get("buy") or []) + len(v.get("sell") or []) > 0
                for v in prev_snap["investors"].values()
            )
            if has_prev:
                store["byDate"][prev_dash] = prev_snap

    # 기존 저장된 과거 기록에도 ETF/ETN 항목이 남아 있으면 정리
    for sec_key in ("byDate", "byWeek"):
        for _k, _snap in (store.get(sec_key) or {}).items():
            for _iv in (_snap.get("investors") or {}).values():
                _iv["buy"] = [r for r in (_iv.get("buy") or []) if not is_etf_or_etn(r.get("name"))]
                _iv["sell"] = [r for r in (_iv.get("sell") or []) if not is_etf_or_etn(r.get("name"))]

    # 2) 주간 수급 수집 (매주 월~오늘 누적, 금요일에는 1주 전체 확정 집계!)
    if include_week:
        mon_dt = dt - timedelta(days=dt.weekday())
        mon_ymd = mon_dt.strftime("%Y%m%d")
        mon_dash = mon_dt.strftime("%Y-%m-%d")
        week_tag = "금요일 주간 확정" if is_friday else "이번주 누적(진행중)"
        print(f"[2/2] 주간 수급 집계 ({mon_dash} ~ {ymd_dash} · {week_tag})...")
        weekly_snap = collect_period_snapshot(token, mon_ymd, ymd, env_mode=env_mode)
        weekly_snap["isFinalWeek"] = is_friday
        weekly_snap["weekLabel"] = f"{mon_dash} ~ {ymd_dash}" + (" (주간 확정)" if is_friday else " (진행중)")
        has_weekly = any(
            len(v.get("buy") or []) + len(v.get("sell") or []) > 0
            for v in weekly_snap["investors"].values()
        )
        if has_weekly:
            # 같은 주(월요일 시작일 기준) 키로 덮어써서 금요일에 주간 확정치로 완성
            store["byWeek"][mon_dash] = weekly_snap

    _save_investor_flow(here, store, update_mcaps=True)

    print(
        f" ★ 키움 수급 저장 완료: 일간 {len(store['dates'])}일 / 주간 {len(store['weeks'])}주 "
        f"(최신일: {store['latestDate'] or '-'})"
    )
    print(f"{'='*80}")
    return store


def fetch_naver_market_caps(target_codes=None):
    """
    네이버 증권 시가총액 API에서 KOSPI / KOSDAQ 종목별 시가총액(억원)을 조회합니다.
    target_codes가 주어지면 필요한 종목만 추려 반환합니다.
    """
    need = set(target_codes) if target_codes else None
    mcaps = {}
    hdr = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"}
    for mkt in ("KOSPI", "KOSDAQ"):
        page = 1
        max_pages = 30
        while page <= max_pages:
            try:
                url = f"https://m.stock.naver.com/api/stocks/marketValue/{mkt}?page={page}&pageSize=100"
                r = requests.get(url, headers=hdr, timeout=12)
                if r.status_code != 200:
                    break
                payload = r.json() or {}
                stocks = payload.get("stocks") or []
                if not stocks:
                    break
                for st in stocks:
                    cd = str(st.get("itemCode") or "").strip().zfill(6)
                    if not cd:
                        continue
                    if need is not None and cd not in need:
                        continue
                    mv_eok = None
                    raw_mv_won = st.get("marketValueRaw")
                    if raw_mv_won is not None:
                        won_num = _parse_abs_num(raw_mv_won)
                        if won_num is not None and won_num > 0:
                            mv_eok = won_num / 1e8  # 원 -> 억원
                    if mv_eok is None:
                        mv_eok = _parse_abs_num(st.get("marketValue"))  # 억원
                    if mv_eok is not None and mv_eok > 0:
                        mcaps[cd] = round(mv_eok, 1)
                total_cnt = int(payload.get("totalCount") or 0)
                if page * 100 >= total_cnt or (need is not None and len(mcaps) >= len(need)):
                    break
                page += 1
                time.sleep(0.05)
            except Exception:
                break
    return mcaps


def refresh_market_caps_in_store(store):
    """store 내 등장하는 모든 종목의 시가총액(억원)을 최신화하여 store['marketCaps']에 저장합니다."""
    codes = set()
    for sec_key in ("byDate", "byWeek"):
        for snap in (store.get(sec_key) or {}).values():
            for iv in (snap.get("investors") or {}).values():
                for r in (iv.get("buy") or []) + (iv.get("sell") or []):
                    cd = str(r.get("code") or "").strip()
                    if cd:
                        codes.add(cd)
    if not codes:
        store.setdefault("marketCaps", {})
        return store
    existing = dict(store.get("marketCaps") or {})
    fetched = fetch_naver_market_caps(codes)
    if fetched:
        existing.update(fetched)
    store["marketCaps"] = {cd: existing[cd] for cd in sorted(codes) if cd in existing}
    print(f"  - [시가총액] 수급 종목 {len(codes)}개 중 {len(store['marketCaps'])}개 시가총액(억원) 매칭 완료")
    return store


def _save_investor_flow(here, store, update_mcaps=False):
    cutoff_d = (datetime.now() - timedelta(days=RETENTION_DAYS)).strftime("%Y-%m-%d")
    store["byDate"] = {
        k: v for k, v in sorted(store.get("byDate", {}).items(), reverse=True)
        if k >= cutoff_d
    }
    week_keys = sorted(store.get("byWeek", {}).keys(), reverse=True)[:RETENTION_WEEKS]
    store["byWeek"] = {k: store["byWeek"][k] for k in week_keys}

    # 과거 기록 ETF/ETN 정리
    for sec_key in ("byDate", "byWeek"):
        for _k, _snap in (store.get(sec_key) or {}).items():
            for _iv in (_snap.get("investors") or {}).values():
                _iv["buy"] = [r for r in (_iv.get("buy") or []) if not is_etf_or_etn(r.get("name"))]
                _iv["sell"] = [r for r in (_iv.get("sell") or []) if not is_etf_or_etn(r.get("name"))]

    if update_mcaps or not store.get("marketCaps"):
        refresh_market_caps_in_store(store)

    store["dates"] = sorted(store["byDate"].keys(), reverse=True)
    store["weeks"] = sorted(store["byWeek"].keys(), reverse=True)
    store["latestDate"] = store["dates"][0] if store["dates"] else ""
    store["latestWeek"] = store["weeks"][0] if store["weeks"] else ""
    store["updatedAt"] = datetime.now().strftime("%Y-%m-%d %H:%M")

    with open(os.path.join(here, "investor_flow.json"), "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, separators=(",", ":"))
    return store


def backfill_investor_flow(here=HERE, daily_days=31, num_weeks=13):
    """최근 1개월 일간 수급 + 최근 3개월(13주) 주간 수급을 키움 API로 일괄 소급 수집합니다."""
    cfg = load_env(os.path.join(here, ".env"))
    appkey = cfg.get("KIWOOM_APPKEY", "")
    secretkey = cfg.get("KIWOOM_SECRETKEY", "")
    env_mode = cfg.get("KIWOOM_ENV", "prod")
    if not appkey or not secretkey or "여기에_" in appkey:
        raise RuntimeError(".env 에 키움 REST API 키를 먼저 설정해 주세요.")

    token = get_kiwoom_token(appkey, secretkey, env_mode=env_mode)
    store = load_investor_flow(here)
    today = _resolve_trading_date()

    print(f"\n{'='*80}")
    print(f" ▶ 키움 수급 과거 소급 수집 시작 (최근 {daily_days}일 일간 + 최근 {num_weeks}주 주간)")
    print(f"{'='*80}")

    # 1) 최근 1개월 평일 일간 수급 수집
    day_targets = []
    for delta in range(daily_days + 1):
        d = today - timedelta(days=delta)
        if d.weekday() < 5:
            day_targets.append(d)
    print(f"[1/2] 일간 수급 점검 대상: {len(day_targets)}영업일 후보")
    added_days = 0
    for idx, d in enumerate(day_targets, 1):
        d_dash = d.strftime("%Y-%m-%d")
        d_ymd = d.strftime("%Y%m%d")
        if d_dash in store["byDate"] and any(
            len(v.get("buy") or []) > 0 for v in (store["byDate"][d_dash].get("investors") or {}).values()
        ):
            continue
        print(f"  ({idx}/{len(day_targets)}) 일간 {d_dash} 조회 중...")
        snap = collect_period_snapshot(token, d_ymd, d_ymd, env_mode=env_mode)
        has_any = any(
            len(v.get("buy") or []) + len(v.get("sell") or []) > 0
            for v in snap["investors"].values()
        )
        if has_any:
            store["byDate"][d_dash] = snap
            added_days += 1
            if added_days % 4 == 0:
                _save_investor_flow(here, store)
        else:
            print(f"    - {d_dash}: 데이터 없음(휴장일 또는 장마감 전) 건너뜀")

    _save_investor_flow(here, store)

    # 2) 최근 3개월(13주) 주간 수급 수집 (각 주의 월~금)
    cur_mon = today - timedelta(days=today.weekday())
    print(f"\n[2/2] 주간 수급 점검 대상: 최근 {num_weeks}주")
    added_weeks = 0
    for w_idx in range(num_weeks):
        mon = cur_mon - timedelta(weeks=w_idx)
        fri = mon + timedelta(days=4)
        end_d = min(fri, today)
        mon_dash = mon.strftime("%Y-%m-%d")
        end_dash = end_d.strftime("%Y-%m-%d")
        is_final = (w_idx > 0) or (today.weekday() == 4)

        existing_w = store["byWeek"].get(mon_dash)
        if existing_w and existing_w.get("isFinalWeek") and w_idx > 0:
            continue

        print(f"  ({w_idx + 1}/{num_weeks}) 주간 {mon_dash} ~ {end_dash} 조회 중...")
        wsnap = collect_period_snapshot(token, mon.strftime("%Y%m%d"), end_d.strftime("%Y%m%d"), env_mode=env_mode)
        has_any = any(
            len(v.get("buy") or []) + len(v.get("sell") or []) > 0
            for v in wsnap["investors"].values()
        )
        if has_any:
            wsnap["isFinalWeek"] = is_final
            wsnap["weekLabel"] = f"{mon_dash} ~ {end_dash}" + (" (주간 확정)" if is_final else " (진행중)")
            store["byWeek"][mon_dash] = wsnap
            added_weeks += 1
            if added_weeks % 3 == 0:
                _save_investor_flow(here, store)

    _save_investor_flow(here, store, update_mcaps=True)
    print(
        f"\n ★ 과거 소급 수집 완료: 일간 총 {len(store['dates'])}일(+{added_days}) / "
        f"주간 총 {len(store['weeks'])}주(+{added_weeks})"
    )
    print(f"{'='*80}")
    return store


def run_test():
    ip = get_public_ip()
    print(f"[현재 PC 공인 IP] {ip or '확인 실패'}")
    cfg = load_env()
    appkey = cfg.get("KIWOOM_APPKEY", "")
    secretkey = cfg.get("KIWOOM_SECRETKEY", "")
    env_mode = cfg.get("KIWOOM_ENV", "prod")
    if not appkey or not secretkey or "여기에_" in appkey:
        print("[안내] .env 파일에 KIWOOM_APPKEY / KIWOOM_SECRETKEY 를 먼저 입력해 주세요.")
        return 1

    print(f"[1/2] 키움 토큰 발급 테스트 (env={env_mode})...")
    token = get_kiwoom_token(appkey, secretkey, env_mode=env_mode, force_refresh=True)
    print(f"  ✓ 토큰 발급 성공! (앞 12자리: {token[:12]}...)")

    dt = _resolve_trading_date()
    ymd = dt.strftime("%Y%m%d")
    print(f"[2/2] ka10058 외국인 코스피 순매수 샘플 조회 ({ymd})...")
    rows = fetch_ka10058(token, ymd, ymd, "9000", "2", "001", env_mode=env_mode)
    print(f"  ✓ 수신 건수: {len(rows)}건")
    for idx, r in enumerate(rows[:5], 1):
        print(
            f"   {idx}. {r['name']}({r['code']}) | 순매수 {r['netAmt']:+.1f}억 ({r['netQty']:+,}주) | "
            f"현재가 {r['price']} ({r['chgPct']}%) | 추정평균가 {r['avgPrice']}"
        )
    return 0


def main():
    args = sys.argv[1:]
    if any(a.lower() in ("--myip", "myip") for a in args):
        ip = get_public_ip()
        print(f"현재 PC 공인 IP: {ip or '확인 실패'}")
        print("→ 키움 REST API 포털(https://openapi.kiwoom.com) 마이페이지에 위 IP를 등록하세요.")
        return

    if any(a.lower() in ("--test", "test") for a in args):
        sys.exit(run_test())

    skip_push = any(a.lower() == "--no-push" for a in args)
    if any(a.lower() in ("--mcaps", "mcaps", "--refresh-mcaps") for a in args):
        store = load_investor_flow(HERE)
        _save_investor_flow(HERE, store, update_mcaps=True)
    elif any(a.lower() in ("--backfill", "backfill") for a in args):
        nums = [int(a) for a in args if a.isdigit()]
        d_days = nums[0] if len(nums) >= 1 else 31
        n_weeks = nums[1] if len(nums) >= 2 else 13
        backfill_investor_flow(HERE, daily_days=d_days, num_weeks=n_weeks)
    else:
        target_date = None
        for i, a in enumerate(args):
            if a.lower() == "--date" and i + 1 < len(args):
                target_date = args[i + 1].replace("-", "")
        update_investor_flow(HERE, target_yyyymmdd=target_date, include_week=True)

    # 웹 대시보드(index.html) 최신화 + GitHub 자동 업로드
    sys.path.insert(0, HERE)
    import kospi200_yoy as screener

    screener.build_web_dashboard(HERE)
    if not skip_push:
        screener.auto_push_to_github(
            HERE,
            commit_msg=f"Auto-update Kiwoom investor flow ({datetime.now().strftime('%Y-%m-%d %H:%M')})",
        )


if __name__ == "__main__":
    main()
