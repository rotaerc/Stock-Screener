# -*- coding: utf-8 -*-
"""
통계청(KOSIS) 반도체 출하·재고·생산 지수 & 재고순환지표 + 삼성전자/SK하이닉스 주가 수집기
=====================================================================================
- 공식 통계청 KOSIS Open API(https://kosis.kr/openapi/)를 통해 '반도체' 월별 출하지수·재고지수·생산지수 수집
- 출하증가율(YoY), 재고증가율(YoY), 재고순환지표(출하YoY - 재고YoY), 재고율(재고/출하×100), 4사분면 국면 자동 산출
- 네이버 금융 월봉 시세에서 삼성전자(005930)·SK하이닉스(000660) 월말 종가를 자동 결합하여 semicon_cycle.json 저장
- 수집 후 index.html 웹 대시보드 자동 갱신 및 GitHub Pages 자동 배포

사용법:
  python kosis_semicon.py --test              # .env 의 KOSIS_API_KEY 연결 및 통계표/반도체 코드 자동 탐색 테스트
  python kosis_semicon.py                     # 통계청 KOSIS 수집 + 삼전/하이닉스 주가 결합 + 웹 갱신 + GitHub 업로드
  python kosis_semicon.py --import 파일.xlsx   # KOSIS 웹에서 직접 다운로드한 엑셀/CSV로 반영 (API 키 없이 가능)
  python kosis_semicon.py --prices-only       # 삼성전자·SK하이닉스 월말 주가만 최신화
  python kosis_semicon.py --no-push           # 웹페이지 생성까지만 수행 (GitHub 업로드 생략)
"""

import ast
import json
import os
import re
import sys
import time
from datetime import datetime

import pandas as pd
import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
SEMICON_JSON_PATH = os.path.join(HERE, "semicon_cycle.json")

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

KOSIS_ORG_ID = "101"  # 통계청
# 광업제조업동향조사 산업별 생산·출하·재고지수 주요 후보 통계표 ID (2020=100 기준 우선)
CANDIDATE_TBL_IDS = [
    "DT_1F02001",  # 시도/산업별 광공업생산지수 (생산·출하·재고 포함)
    "DT_1F02011",  # 산업별 광공업생산지수
    "DT_1F02012",  # 산업별 출하·재고지수
    "DT_1F02003",
    "DT_1F02004",
    "DT_1F01511",
    "DT_1F01501",
]


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
    for key in ("KOSIS_API_KEY", "KOSIS_TBL_ID", "KOSIS_ORG_ID"):
        if os.environ.get(key):
            cfg[key] = os.environ[key].strip()
    return cfg


def _num(x):
    if x is None:
        return None
    s = str(x).replace(",", "").strip()
    if s in ("", "-", "...", "null", "None", "nan", "NaN", "p"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _kosis_get_json(url, params, timeout=25):
    r = requests.get(url, params=params, headers={"User-Agent": UA}, timeout=timeout)
    r.raise_for_status()
    txt = r.text.strip()
    try:
        return r.json()
    except Exception:
        # 일부 KOSIS 오류 응답은 비표준 JSON ({err:"10",errMsg:"..."}) 형태
        fixed = re.sub(r'([{\s,])(\w+)\s*:', r'\1"\2":', txt)
        return json.loads(fixed)


def search_kosis_tables(api_key, keyword="광공업생산지수"):
    """KOSIS 공식 통계표 검색 API로 광업제조업동향조사 출하·재고 포함 통계표 목록 탐색"""
    url = "https://kosis.kr/openapi/statisticsSearch.do"
    params = {
        "method": "getList",
        "apiKey": api_key,
        "searchNm": keyword,
        "orgId": KOSIS_ORG_ID,
        "format": "json",
        "jsonVD": "Y",
        "startCount": "1",
        "resultCount": "30",
    }
    try:
        res = _kosis_get_json(url, params)
        if isinstance(res, list):
            return res
    except Exception:
        pass
    return []


def get_kosis_meta(api_key, org_id, tbl_id, meta_type="ITM"):
    """KOSIS 통계표 메타정보 조회 (type: TBL / ITM / PRD)"""
    url = "https://kosis.kr/openapi/statisticsData.do"
    params = {
        "method": "getMeta",
        "type": meta_type,
        "apiKey": api_key,
        "orgId": org_id,
        "tblId": tbl_id,
        "format": "json",
    }
    res = _kosis_get_json(url, params)
    if isinstance(res, dict) and res.get("err"):
        return None, res.get("errMsg") or str(res)
    if isinstance(res, list):
        return res, ""
    return None, "예상치 못한 응답 형식"


def _analyze_table_meta(meta_rows):
    """
    getMeta(type=ITM) 응답에서
    - 지역(전국) 분류코드,
    - 반도체 산업 분류코드(예: 261 / C261 / 반도체),
    - 생산지수 / 출하지수 / 재고지수 항목(ITM) 코드(원지수 우선, 계절조정 보조)
    를 자동으로 매핑합니다.
    """
    if not meta_rows:
        return None

    by_obj = {}
    items = []
    for row in meta_rows:
        obj_id = str(row.get("OBJ_ID") or "").strip().upper()
        itm_id = str(row.get("ITM_ID") or "").strip()
        itm_nm = str(row.get("ITM_NM") or "").strip()
        obj_nm = str(row.get("OBJ_NM") or "").strip()
        if obj_id == "ITEM" or obj_id == "ITM" or "항목" in obj_nm:
            items.append((itm_id, itm_nm))
        elif obj_id:
            by_obj.setdefault(obj_id, {"name": obj_nm, "vals": []})
            by_obj[obj_id]["vals"].append((itm_id, itm_nm))

    # 1) 항목(출하지수, 재고지수, 생산지수) 매칭
    def pick_item(keywords, prefer_orig=True):
        cands = []
        for cid, nm in items:
            nm_c = re.sub(r"\s+", "", nm)
            if not all(k in nm_c for k in keywords):
                continue
            if "증감" in nm_c or "비중" in nm_c or "가중치" in nm_c:
                continue
            score = 0
            if "원지수" in nm_c:
                score += 10 if prefer_orig else 1
            elif "계절조정" in nm_c:
                score += 1 if prefer_orig else 10
            else:
                score += 6
            # 총출하 우선 (내수/수출 세부 제외)
            if "내수" in nm_c or "수출" in nm_c:
                score -= 5
            if "율" in nm_c:
                score -= 8
            cands.append((score, cid, nm))
        if not cands:
            return None
        cands.sort(key=lambda x: (-x[0], len(x[2])))
        return cands[0][1], cands[0][2]

    ship_itm = pick_item(["출하"], prefer_orig=True)
    inv_itm = pick_item(["재고"], prefer_orig=True)
    prod_itm = pick_item(["생산"], prefer_orig=True)
    ship_sa = pick_item(["출하", "계절조정"], prefer_orig=False)
    inv_sa = pick_item(["재고", "계절조정"], prefer_orig=False)

    if not ship_itm or not inv_itm:
        return None

    # 2) 분류(OBJ_ID) 매핑: 전국(시도별 테이블인 경우) + 반도체 산업코드
    obj_selections = {}
    semi_found = None
    sorted_objs = sorted(by_obj.keys())

    for oid in sorted_objs:
        onm = by_obj[oid]["name"]
        vals = by_obj[oid]["vals"]
        # 반도체 산업 분류 탐색
        semi_cands = []
        for cid, nm in vals:
            nm_c = re.sub(r"\s+", "", nm)
            if "반도체" in nm_c and "제조용" not in nm_c and "기계" not in nm_c and "장비" not in nm_c:
                sc = 0
                if nm_c in ("반도체", "반도체제조업") or nm_c.startswith("반도체"):
                    sc += 20
                if "261" in cid or "261" in nm_c:
                    sc += 15
                if "C261" in cid.upper():
                    sc += 10
                # 전자부품·컴퓨터 등 상위 대분류보다는 소분류 '반도체' 우선
                if "전자부품" in nm_c or "통신" in nm_c:
                    sc -= 8
                semi_cands.append((sc, cid, nm))
        if semi_cands:
            semi_cands.sort(key=lambda x: (-x[0], len(x[2])))
            semi_found = (oid, semi_cands[0][1], semi_cands[0][2])
            obj_selections[oid] = semi_cands[0][1]
            continue

        # 시도/지역 분류인 경우 '전국' 선택
        if any(k in onm for k in ("시도", "지역", "행정구역")):
            nat = next(((cid, nm) for cid, nm in vals if "전국" in nm), None)
            if nat:
                obj_selections[oid] = nat[0]
            elif vals:
                obj_selections[oid] = vals[0][0]
            continue

        # 기타 분류가 있으면 첫 번째(전체/계) 선택
        tot = next(((cid, nm) for cid, nm in vals if nm.strip() in ("전체", "계", "총계", "전국")), None)
        obj_selections[oid] = tot[0] if tot else (vals[0][0] if vals else "ALL")

    if not semi_found:
        return None

    return {
        "ship": ship_itm,
        "inv": inv_itm,
        "prod": prod_itm,
        "ship_sa": ship_sa,
        "inv_sa": inv_sa,
        "semi": semi_found,
        "objs": obj_selections,
        "obj_order": sorted_objs,
    }


def discover_kosis_semicon_source(api_key, preferred_tbl_id=""):
    """공식 KOSIS Open API에서 반도체 출하·재고 지수 통계표 및 분류코드를 자동 탐색합니다."""
    cand_list = []
    if preferred_tbl_id:
        cand_list.append((KOSIS_ORG_ID, preferred_tbl_id, "지정 통계표"))
    for tid in CANDIDATE_TBL_IDS:
        if tid != preferred_tbl_id:
            cand_list.append((KOSIS_ORG_ID, tid, ""))

    # 검색 API로도 최신 2020=100 광공업 생산·출하·재고 통계표 보강
    for kw in ("시도 산업별 광공업생산지수", "산업별 광공업생산 출하 재고"):
        for hit in search_kosis_tables(api_key, kw):
            tid = str(hit.get("TBL_ID") or "").strip()
            oid = str(hit.get("ORG_ID") or KOSIS_ORG_ID).strip()
            tnm = str(hit.get("TBL_NM") or "").strip()
            if tid and not any(c[1] == tid for c in cand_list):
                if "2020=100" in tnm or "광공업" in tnm or "출하" in tnm or "재고" in tnm:
                    cand_list.append((oid, tid, tnm))

    last_err = ""
    for oid, tid, tnm in cand_list:
        meta, err = get_kosis_meta(api_key, oid, tid, meta_type="ITM")
        if err:
            last_err = err
            # 인증키 오류면 즉시 중단
            if "인증" in str(err) or "KEY" in str(err).upper():
                raise RuntimeError(f"KOSIS API 인증 오류: {err}")
            continue
        parsed = _analyze_table_meta(meta)
        if parsed:
            tbl_meta, _ = get_kosis_meta(api_key, oid, tid, meta_type="TBL")
            real_nm = tnm
            if tbl_meta and isinstance(tbl_meta, list) and tbl_meta[0].get("TBL_NM"):
                real_nm = tbl_meta[0]["TBL_NM"]
            parsed["orgId"] = oid
            parsed["tblId"] = tid
            parsed["tblNm"] = real_nm or tid
            return parsed

    raise RuntimeError(
        f"KOSIS에서 반도체 출하·재고 지수 통계표를 찾지 못했습니다. (마지막 오류: {last_err or '매칭 항목 없음'})"
    )


def fetch_kosis_semicon_monthly(api_key, src_info, start_ym="201501"):
    """탐색된 KOSIS 통계표에서 월별 반도체 생산·출하·재고 지수를 수집합니다."""
    url = "https://kosis.kr/openapi/Param/statisticsParameterData.do"
    end_ym = datetime.now().strftime("%Y%m")

    itm_ids = [src_info["ship"][0], src_info["inv"][0]]
    if src_info.get("prod"):
        itm_ids.append(src_info["prod"][0])
    if src_info.get("ship_sa") and src_info["ship_sa"][0] not in itm_ids:
        itm_ids.append(src_info["ship_sa"][0])
    if src_info.get("inv_sa") and src_info["inv_sa"][0] not in itm_ids:
        itm_ids.append(src_info["inv_sa"][0])

    # KOSIS 1회 최대 셀 제한 예방을 위해 4년 단위로 분할 조회
    start_year = int(start_ym[:4])
    end_year = int(end_ym[:4])
    all_rows = []

    for y0 in range(start_year, end_year + 1, 4):
        y1 = min(y0 + 3, end_year)
        p_start = f"{y0}01" if y0 > start_year else start_ym
        p_end = f"{y1}12" if y1 < end_year else end_ym
        params = {
            "method": "getList",
            "apiKey": api_key,
            "orgId": src_info["orgId"],
            "tblId": src_info["tblId"],
            "itmId": "+".join(itm_ids),
            "format": "json",
            "jsonVD": "Y",
            "prdSe": "M",
            "startPrdDe": p_start,
            "endPrdDe": p_end,
        }
        # objL1 ~ objL8 매핑
        for idx, oid in enumerate(src_info["obj_order"], 1):
            params[f"objL{idx}"] = src_info["objs"][oid]

        res = _kosis_get_json(url, params)
        if isinstance(res, dict) and res.get("err"):
            # 데이터 없음(30)은 스킵, 그 외 오류는 출력
            if str(res.get("err")) != "30":
                print(f"  [주의] KOSIS 구간({p_start}~{p_end}) 응답: {res}")
            continue
        if isinstance(res, list):
            all_rows.extend(res)
        time.sleep(0.2)

    ship_code = src_info["ship"][0]
    inv_code = src_info["inv"][0]
    prod_code = src_info["prod"][0] if src_info.get("prod") else None
    ship_sa_code = src_info["ship_sa"][0] if src_info.get("ship_sa") else None
    inv_sa_code = src_info["inv_sa"][0] if src_info.get("inv_sa") else None

    monthly = {}
    for r in all_rows:
        prd = str(r.get("PRD_DE") or "").strip()
        if not re.match(r"^\d{6}$", prd):
            continue
        ym = f"{prd[:4]}-{prd[4:6]}"
        val = _num(r.get("DT"))
        if val is None:
            continue
        iid = str(r.get("ITM_ID") or "").strip()
        rec = monthly.setdefault(ym, {})
        if iid == ship_code:
            rec["shipIdx"] = round(val, 2)
        elif iid == inv_code:
            rec["invIdx"] = round(val, 2)
        elif prod_code and iid == prod_code:
            rec["prodIdx"] = round(val, 2)
        elif ship_sa_code and iid == ship_sa_code:
            rec["shipSaIdx"] = round(val, 2)
        elif inv_sa_code and iid == inv_sa_code:
            rec["invSaIdx"] = round(val, 2)

    return monthly


def fetch_naver_monthly_prices(start_date="20150101"):
    """네이버 금융 월봉 시세 API에서 삼성전자(005930)·SK하이닉스(000660) 월말 종가 수집"""
    end_date = datetime.now().strftime("%Y%m%d")
    out = {}
    for key, code, label in (("samsung", "005930", "삼성전자"), ("hynix", "000660", "SK하이닉스")):
        try:
            r = requests.get(
                "https://api.finance.naver.com/siseJson.naver",
                params={
                    "symbol": code,
                    "requestType": "1",
                    "startTime": start_date,
                    "endTime": end_date,
                    "timeframe": "month",
                },
                headers={"User-Agent": UA},
                timeout=15,
            )
            r.raise_for_status()
            rows = ast.literal_eval(r.text.strip())
            cnt = 0
            for row in rows[1:]:
                if not row or len(row) < 5:
                    continue
                dt_str = str(row[0]).strip()
                if len(dt_str) < 6:
                    continue
                ym = f"{dt_str[:4]}-{dt_str[4:6]}"
                close_v = _num(row[4])
                if close_v is not None and close_v > 0:
                    out.setdefault(ym, {})[key] = int(round(close_v))
                    cnt += 1
            print(f"  - [주가] {label}({code}) 월봉 {cnt}개월 수집 완료")
        except Exception as e:
            print(f"  - [경고] {label}({code}) 월봉 수집 실패: {e}")
    return out


def classify_inventory_phase(ship_yoy, inv_yoy):
    """표준 재고순환 4국면 자동 판정"""
    if ship_yoy is None or inv_yoy is None:
        return ""
    if ship_yoy >= 0 and inv_yoy < 0:
        return "회복기"
    if ship_yoy >= 0 and inv_yoy >= 0:
        return "확장기" if ship_yoy >= inv_yoy else "둔화·후퇴기"
    if ship_yoy < 0 and inv_yoy >= 0:
        return "둔화·후퇴기"
    return "수축·조정기"


def _next_ym(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    return f"{y + 1:04d}-01" if m == 12 else f"{y:04d}-{m + 1:02d}"


def _prev_ym(ym):
    y, m = int(ym[:4]), int(ym[5:7])
    return f"{y - 1:04d}-12" if m == 1 else f"{y:04d}-{m - 1:02d}"


def build_semicon_payload(monthly_kosis, monthly_prices, meta_info):
    """월별 지수 + 주가를 결합해 YoY, 재고순환지표, 재고율, 4국면을 계산하고 다음달까지 10년 월평균 예측치를 산출합니다."""
    actual_kosis_yms = sorted(
        ym for ym, v in monthly_kosis.items()
        if v.get("shipIdx") is not None and v.get("invIdx") is not None
    )
    has_kosis = bool(actual_kosis_yms)
    latest_actual_ym = actual_kosis_yms[-1] if actual_kosis_yms else ""

    # 현재월 + 1개월(다음달)까지 미발표 구간에 대해 최근 10년 동월 전월비(MoM%) 평균으로 예측치 자동 연장
    now_ym = datetime.now().strftime("%Y-%m")
    forecast_end_ym = _next_ym(now_ym)

    combined_kosis = {ym: dict(v) for ym, v in monthly_kosis.items()}
    forecast_yms = []

    if latest_actual_ym and latest_actual_ym < forecast_end_ym:
        # 월별(1~12월) 최근 10개년 확정치 전월비(MoM%) 평균 산출
        mom_avg = {}
        for field in ("shipIdx", "invIdx", "prodIdx"):
            mom_avg[field] = {}
            for mo in range(1, 13):
                rates = []
                for ym in actual_kosis_yms:
                    if int(ym[5:7]) != mo:
                        continue
                    pym = _prev_ym(ym)
                    cur_v = (monthly_kosis.get(ym) or {}).get(field)
                    prv_v = (monthly_kosis.get(pym) or {}).get(field)
                    if cur_v is not None and prv_v is not None and prv_v > 0:
                        rates.append((cur_v / prv_v) - 1.0)
                recent_10 = rates[-10:]
                mom_avg[field][mo] = (sum(recent_10) / len(recent_10)) if recent_10 else 0.0

        cur_ym = _next_ym(latest_actual_ym)
        while cur_ym <= forecast_end_ym:
            pym = _prev_ym(cur_ym)
            mo = int(cur_ym[5:7])
            prev_rec = combined_kosis.get(pym) or {}
            fc_rec = {"isForecast": True}
            for field in ("shipIdx", "invIdx", "prodIdx"):
                pv = prev_rec.get(field)
                if pv is not None and pv > 0:
                    fc_rec[field] = round(pv * (1.0 + mom_avg[field].get(mo, 0.0)), 2)
            combined_kosis[cur_ym] = fc_rec
            forecast_yms.append(cur_ym)
            cur_ym = _next_ym(cur_ym)

    all_yms = sorted(set(combined_kosis.keys()) | set(monthly_prices.keys()))

    series = []
    for ym in all_yms:
        k = combined_kosis.get(ym) or {}
        p = monthly_prices.get(ym) or {}
        y, m = int(ym[:4]), int(ym[5:7])
        prev_ym = f"{y - 1:04d}-{m:02d}"
        kp = combined_kosis.get(prev_ym) or {}

        ship = k.get("shipIdx")
        inv = k.get("invIdx")
        prod = k.get("prodIdx")

        ship_prev = kp.get("shipIdx")
        inv_prev = kp.get("invIdx")
        prod_prev = kp.get("prodIdx")

        ship_yoy = round((ship - ship_prev) / ship_prev * 100.0, 2) if (ship is not None and ship_prev and ship_prev > 0) else None
        inv_yoy = round((inv - inv_prev) / inv_prev * 100.0, 2) if (inv is not None and inv_prev and inv_prev > 0) else None
        prod_yoy = round((prod - prod_prev) / prod_prev * 100.0, 2) if (prod is not None and prod_prev and prod_prev > 0) else None

        cycle = round(ship_yoy - inv_yoy, 2) if (ship_yoy is not None and inv_yoy is not None) else None
        inv_ratio = round(inv / ship * 100.0, 1) if (inv is not None and ship and ship > 0) else None
        phase = classify_inventory_phase(ship_yoy, inv_yoy)
        is_fc = bool(k.get("isForecast"))

        row_obj = {
            "ym": ym,
            "shipIdx": ship,
            "invIdx": inv,
            "prodIdx": prod,
            "shipYoY": ship_yoy,
            "invYoY": inv_yoy,
            "prodYoY": prod_yoy,
            "cycle": cycle,
            "invRatio": inv_ratio,
            "phase": phase,
            "samsung": p.get("samsung"),
            "hynix": p.get("hynix"),
        }
        if is_fc:
            row_obj["isForecast"] = True
        series.append(row_obj)

    valid_actual_months = [s["ym"] for s in series if s["cycle"] is not None and not s.get("isForecast")]
    latest_kosis_ym = valid_actual_months[-1] if valid_actual_months else ""

    return {
        "updatedAt": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "hasKosisData": bool(has_kosis and valid_actual_months),
        "latestMonth": latest_kosis_ym,
        "forecastMonths": forecast_yms,
        "forecastMethod": "최근 10년 동월 전월비(MoM%) 평균 반영",
        "source": {
            "org": "통계청 (KOSIS 광업제조업동향조사)",
            "tblId": meta_info.get("tblId", ""),
            "tblNm": meta_info.get("tblNm", ""),
            "industry": meta_info.get("industry", "반도체"),
            "shipLabel": meta_info.get("shipLabel", "출하지수"),
            "invLabel": meta_info.get("invLabel", "재고지수"),
            "prodLabel": meta_info.get("prodLabel", "생산지수"),
        },
        "series": series,
    }


def import_kosis_file(filepath):
    """
    KOSIS 웹에서 직접 다운로드한 엑셀(.xlsx/.xls) 또는 CSV(.csv) 파일을 파싱해
    월별 반도체 생산·출하·재고 지수 dict를 추출합니다.
    """
    if not os.path.isfile(filepath):
        raise FileNotFoundError(f"파일을 찾을 수 없습니다: {filepath}")

    ext = os.path.splitext(filepath)[1].lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
            try:
                df = pd.read_csv(filepath, encoding=enc)
                break
            except Exception:
                df = None
        if df is None:
            raise RuntimeError("CSV 파일 인코딩을 읽지 못했습니다.")
    else:
        df = pd.read_excel(filepath)

    # 날짜 컬럼 탐색 (예: '2024.01 월', '2024. 01', '202401', '2024-01')
    ym_cols = []
    for col in df.columns:
        m = re.search(r"(20\d{2})\s*[.\-/]?\s*(0[1-9]|1[0-2])", str(col))
        if m:
            ym_cols.append((col, f"{m.group(1)}-{m.group(2)}"))

    if not ym_cols:
        raise RuntimeError(
            "엑셀/CSV에서 월별 시점 컬럼(예: 2024.01, 2024-01)을 찾지 못했습니다. "
            "KOSIS 다운로드 시 시점(월)을 열(가로)로 배치해 저장해 주세요."
        )

    text_cols = [c for c in df.columns if c not in [x[0] for x in ym_cols]]
    monthly = {}

    for _, row in df.iterrows():
        row_txt = " ".join(str(row[c]) for c in text_cols if pd.notna(row[c]))
        row_c = re.sub(r"\s+", "", row_txt)
        # 반도체 행만 선별 (파일 전체가 반도체 단일 파일이면 출하/재고 키워드만으로 허용)
        if "제조용" in row_c or "장비" in row_c:
            continue
        if "증감" in row_c or "가중치" in row_c or "내수" in row_c or "수출" in row_c:
            continue
        if "계절조정" in row_c:
            continue

        field = None
        if "출하" in row_c:
            field = "shipIdx"
        elif "재고" in row_c and "재고율" not in row_c:
            field = "invIdx"
        elif "생산" in row_c:
            field = "prodIdx"
        if not field:
            continue

        # 시도 컬럼이 있으면 전국 행만
        if any(city in row_c for city in ("서울", "부산", "대구", "인천", "광주", "대전", "울산", "세종", "경기", "강원", "충북", "충남", "전북", "전남", "경북", "경남", "제주")):
            continue
        # 산업 컬럼에 다른 산업이 섞여 있을 때 반도체만 필터
        if any("반도체" in str(v) for v in df[text_cols].values.ravel()):
            if "반도체" not in row_c:
                continue

        for col, ym in ym_cols:
            val = _num(row[col])
            if val is not None:
                monthly.setdefault(ym, {})[field] = round(val, 2)

    if not monthly:
        raise RuntimeError("파일에서 반도체 출하지수/재고지수 행을 찾지 못했습니다.")

    return monthly


def save_and_deploy(payload, skip_push=False):
    with open(SEMICON_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))
    print(
        f"  ★ 저장 완료: {os.path.basename(SEMICON_JSON_PATH)} "
        f"(총 {len(payload['series'])}개월 / 통계청 최신월: {payload['latestMonth'] or '미연동'})"
    )

    try:
        import kospi200_yoy
        kospi200_yoy.build_web_dashboard(HERE)
        if not skip_push:
            kospi200_yoy.auto_push_to_github(HERE, commit_msg=f"Update semiconductor inventory cycle ({datetime.now().strftime('%Y-%m-%d %H:%M')})")
            try:
                import telegram_notify
                telegram_notify.notify_update("semicon", HERE)
            except Exception:
                pass
    except Exception as e:
        print(f"  [경고] 웹 대시보드 재생성/푸시 중 오류: {e}")


def main():
    raw_args = sys.argv[1:]
    skip_push = any(a.lower() == "--no-push" for a in raw_args)
    is_test = any(a.lower() == "--test" for a in raw_args)
    prices_only = any(a.lower() == "--prices-only" for a in raw_args)

    import_path = None
    for i, a in enumerate(raw_args):
        if a.lower() == "--import" and i + 1 < len(raw_args):
            import_path = raw_args[i + 1]

    print("=" * 80)
    print(" ▶ 통계청(KOSIS) 반도체 재고순환지표 + 삼전/하이닉스 주가 수집기")
    print("=" * 80)

    if import_path:
        print(f"[1/2] KOSIS 파일 수동 임포트 중: {import_path}")
        monthly_kosis = import_kosis_file(import_path)
        print(f"  - 파싱 완료: {len(monthly_kosis)}개월 ({min(monthly_kosis)} ~ {max(monthly_kosis)})")
        print("[2/2] 삼성전자·SK하이닉스 월봉 주가 수집 중...")
        monthly_prices = fetch_naver_monthly_prices()
        payload = build_semicon_payload(
            monthly_kosis,
            monthly_prices,
            {"tblId": "FILE_IMPORT", "tblNm": os.path.basename(import_path), "industry": "반도체"},
        )
        save_and_deploy(payload, skip_push=skip_push)
        return

    if prices_only:
        existing = {}
        meta_info = {}
        if os.path.isfile(SEMICON_JSON_PATH):
            try:
                old = json.load(open(SEMICON_JSON_PATH, encoding="utf-8"))
                meta_info = old.get("source") or {}
                for s in old.get("series") or []:
                    if s.get("isForecast"):
                        continue
                    if s.get("shipIdx") is not None or s.get("invIdx") is not None:
                        existing[s["ym"]] = {
                            k: s[k] for k in ("shipIdx", "invIdx", "prodIdx") if s.get(k) is not None
                        }
            except Exception:
                pass
        monthly_prices = fetch_naver_monthly_prices()
        payload = build_semicon_payload(existing, monthly_prices, meta_info)
        save_and_deploy(payload, skip_push=skip_push)
        return

    cfg = load_env()
    api_key = cfg.get("KOSIS_API_KEY", "").strip()
    if not api_key or "여기에" in api_key:
        print("[안내] .env 파일에 KOSIS_API_KEY 가 설정되지 않았습니다.")
        print("  1) https://kosis.kr/openapi/ 무료 회원가입 후 Open API 인증키 즉시 발급")
        print("  2) 이 폴더의 .env 파일에 아래 한 줄 추가:")
        print("     KOSIS_API_KEY=발급받은_인증키")
        print("  3) python kosis_semicon.py 다시 실행")
        print("  ※ 또는 KOSIS 웹에서 엑셀 다운로드 후: python kosis_semicon.py --import 파일.xlsx")
        if not is_test:
            # 주가 데이터라도 최신화해서 semicon_cycle.json 생성 (기존 통계청 데이터가 있다면 보존)
            existing = {}
            meta_info = {}
            if os.path.isfile(SEMICON_JSON_PATH):
                try:
                    old = json.load(open(SEMICON_JSON_PATH, encoding="utf-8"))
                    meta_info = old.get("source") or {}
                    for s in old.get("series") or []:
                        if s.get("shipIdx") is not None or s.get("invIdx") is not None:
                            existing[s["ym"]] = {
                                k: s[k] for k in ("shipIdx", "invIdx", "prodIdx") if s.get(k) is not None
                            }
                except Exception:
                    pass
            monthly_prices = fetch_naver_monthly_prices()
            payload = build_semicon_payload(existing, monthly_prices, meta_info)
            save_and_deploy(payload, skip_push=skip_push)
        return

    print("[1/3] KOSIS 반도체 출하·재고 통계표 및 분류코드 자동 탐색 중...")
    src_info = discover_kosis_semicon_source(api_key, preferred_tbl_id=cfg.get("KOSIS_TBL_ID", ""))
    print(f"  - 통계표: [{src_info['tblId']}] {src_info['tblNm']}")
    print(f"  - 산업분류: {src_info['semi'][2]} (코드: {src_info['semi'][1]})")
    print(f"  - 출하지수: {src_info['ship'][1]} ({src_info['ship'][0]})")
    print(f"  - 재고지수: {src_info['inv'][1]} ({src_info['inv'][0]})")
    if src_info.get("prod"):
        print(f"  - 생산지수: {src_info['prod'][1]} ({src_info['prod'][0]})")

    if is_test:
        print("\n★ KOSIS API 연결 및 반도체 통계표 탐색 테스트 성공!")
        return

    print("[2/3] 통계청 KOSIS 월별 반도체 생산·출하·재고 지수 수집 중 (2015-01 ~ 최신)...")
    monthly_kosis = fetch_kosis_semicon_monthly(api_key, src_info, start_ym="201501")
    if not monthly_kosis:
        raise RuntimeError("KOSIS API에서 월별 수치를 가져오지 못했습니다.")
    print(f"  - 수집 완료: {len(monthly_kosis)}개월 ({min(monthly_kosis)} ~ {max(monthly_kosis)})")

    print("[3/3] 네이버 금융 삼성전자·SK하이닉스 월봉 주가 결합 중...")
    monthly_prices = fetch_naver_monthly_prices(start_date="20150101")

    payload = build_semicon_payload(
        monthly_kosis,
        monthly_prices,
        {
            "tblId": src_info["tblId"],
            "tblNm": src_info["tblNm"],
            "industry": src_info["semi"][2],
            "shipLabel": src_info["ship"][1],
            "invLabel": src_info["inv"][1],
            "prodLabel": src_info["prod"][1] if src_info.get("prod") else "",
        },
    )
    valid = [s for s in payload["series"] if s["cycle"] is not None and not s.get("isForecast")]
    if valid:
        last = valid[-1]
        print(
            f"  - 최신 확정월({last['ym']}): 출하YoY {last['shipYoY']:+.2f}% / "
            f"재고YoY {last['invYoY']:+.2f}% / 재고순환지표 {last['cycle']:+.2f}%p [{last['phase']}]"
        )
    fcs = [s for s in payload["series"] if s.get("isForecast") and s.get("cycle") is not None]
    if fcs:
        fc_txt = ", ".join(f"{x['ym']}(E) {x['cycle']:+.1f}%p" for x in fcs)
        print(f"  - 10년 월평균 예측치({len(fcs)}개월): {fc_txt}")
    save_and_deploy(payload, skip_push=skip_push)


if __name__ == "__main__":
    main()
