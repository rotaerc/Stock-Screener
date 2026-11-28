# -*- coding: utf-8 -*-
"""
텔레그램 업데이트 완료 알림 발송 모듈
===========================================================================
- .env 파일 또는 환경변수(GitHub Secrets)에서 TELEGRAM_BOT_TOKEN, TELEGRAM_CHAT_ID 로드
- 토큰이 설정되어 있지 않으면 오류 없이 조용히 건너뜁니다.
- 사용법:
    python telegram_notify.py test       # 연결 테스트 (CHAT_ID 미설정 시 최근 메시지에서 숫자 ID 자동 조회)
    python telegram_notify.py us         # 미국 모닝 브리핑 업데이트 요약 알림
    python telegram_notify.py reports    # 기업 리포트 업데이트 요약 알림
    python telegram_notify.py screener   # 한국 스크리너 업데이트 요약 알림
    python telegram_notify.py flow       # 키움 수급 업데이트 요약 알림
    python telegram_notify.py semicon    # 반도체 재고순환 업데이트 요약 알림
    python telegram_notify.py all        # 전체 업데이트 요약 알림
"""

import json
import os
import sys
from datetime import datetime, timezone, timedelta

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = os.path.join(HERE, ".env")
DASHBOARD_URL = "https://rotaerc.github.io/Stock-Screener/"
KST = timezone(timedelta(hours=9))


def _load_telegram_cfg(env_path=ENV_PATH):
    cfg = {}
    if os.path.isfile(env_path):
        try:
            with open(env_path, "r", encoding="utf-8-sig") as f:
                for raw in f:
                    line = raw.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    k, v = line.split("=", 1)
                    k = k.strip()
                    v = v.strip().strip('"').strip("'")
                    if k in ("TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID") and v:
                        cfg[k] = v
        except Exception:
            pass
    for k in ("TELEGRAM_BOT_TOKEN", "TELEGRAM_CHAT_ID"):
        val = (os.environ.get(k) or "").strip()
        if val:
            cfg[k] = val
    return cfg.get("TELEGRAM_BOT_TOKEN", ""), cfg.get("TELEGRAM_CHAT_ID", "")


def _read_json(path):
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _fmt_pct(v):
    if v is None:
        return "-"
    try:
        n = float(v)
        return f"{n:+.2f}%"
    except Exception:
        return "-"


def send_telegram_message(text, token=None, chat_id=None):
    tok, cid = _load_telegram_cfg()
    tok = (token or tok or "").strip()
    cid = (chat_id or cid or "").strip()
    if not tok or not cid:
        return False
    try:
        r = requests.post(
            f"https://api.telegram.org/bot{tok}/sendMessage",
            json={
                "chat_id": cid,
                "text": text,
                "disable_web_page_preview": True,
            },
            timeout=12,
        )
        if r.status_code == 200:
            print("  ★ 텔레그램 알림 발송 완료!")
            return True
        print(f"  [경고] 텔레그램 발송 실패 (HTTP {r.status_code})")
        return False
    except Exception as e:
        print(f"  [경고] 텔레그램 발송 오류: {e}")
        return False


def discover_chat_id(token):
    """봇에게 보낸 최근 메시지에서 숫자 chat_id 자동 탐색"""
    if not token:
        return None
    try:
        r = requests.get(f"https://api.telegram.org/bot{token}/getUpdates", timeout=10)
        if r.status_code != 200:
            return None
        results = (r.json() or {}).get("result") or []
        for item in reversed(results):
            msg = item.get("message") or item.get("edited_message") or {}
            chat = msg.get("chat") or {}
            cid = chat.get("id")
            if cid is not None:
                return str(cid)
        return None
    except Exception:
        return None


def _build_us_text(here=HERE):
    d = _read_json(os.path.join(here, "us_market.json")) or {}
    m = d.get("morning") or {}
    td = d.get("tradeDate") or "-"
    idx_map = {x["symbol"]: x for x in (m.get("majorIndices") or [])}
    mac_map = {x["symbol"]: x for x in (m.get("macroIndicators") or [])}
    spx = idx_map.get("^GSPC") or {}
    ixic = idx_map.get("^IXIC") or {}
    sox = idx_map.get("^SOX") or {}
    ewy = mac_map.get("EWY") or {}
    krw = mac_map.get("KRW=X") or {}
    tnx = mac_map.get("^TNX") or {}
    fg = m.get("fearGreed") or {}
    ss = sorted(m.get("sectorStrength") or [], key=lambda x: x.get("avgChg1d") or -999, reverse=True)
    top_themes = ", ".join(f"{x['name']}({_fmt_pct(x.get('avgChg1d'))})" for x in ss[:3])

    lines = [
        f"🌅 [미국 모닝 브리핑 업데이트 · {td}]",
        f"• S&P500 {_fmt_pct(spx.get('chg1d'))} | 나스닥 {_fmt_pct(ixic.get('chg1d'))} | SOX {_fmt_pct(sox.get('chg1d'))}",
    ]
    sub = []
    if ewy.get("chg1d") is not None:
        sub.append(f"EWY {_fmt_pct(ewy.get('chg1d'))}")
    if krw.get("price") is not None:
        sub.append(f"환율 {krw['price']:,.1f}원")
    if tnx.get("price") is not None:
        sub.append(f"미10년 {tnx['price']:.2f}%")
    if fg.get("score") is not None:
        sub.append(f"공포탐욕 {round(fg['score'])}({fg.get('ratingKo', '')})")
    if sub:
        lines.append("• " + " · ".join(sub))
    if top_themes:
        lines.append(f"• 강세 테마: {top_themes}")
    if m.get("krImpactKo"):
        lines.append(f"• 한국장 영향: {m['krImpactKo']}")
    lines.append(f"👉 {DASHBOARD_URL}")
    return "\n".join(lines)


def _build_reports_text(here=HERE):
    d = _read_json(os.path.join(here, "hankyung_reports.json")) or {}
    by_dt = d.get("byDate") or {}
    latest_dt = sorted(by_dt.keys(), reverse=True)[0] if by_dt else (d.get("date") or "-")
    reps = by_dt.get(latest_dt) or d.get("reports") or []
    ups = [r for r in reps if r.get("target_change") == "상향"]
    news = [r for r in reps if r.get("target_change") == "신규"]
    up_names = []
    seen = set()
    for r in ups:
        nm = r.get("stock_name") or ""
        if nm and nm not in seen:
            seen.add(nm)
            up_names.append(nm)
        if len(up_names) >= 6:
            break

    lines = [
        f"📰 [기업 리포트 업데이트 · {latest_dt}]",
        f"• 총 {len(reps)}건 (목표가 상향 {len(ups)}건 · 신규 {len(news)}건)",
    ]
    if up_names:
        lines.append(f"• 목표가 상향: {', '.join(up_names)}")
    lines.append(f"👉 {DASHBOARD_URL}")
    return "\n".join(lines)


def _build_screener_text(here=HERE):
    now_kst = datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
    lines = [
        f"📊 [한국 스크리너 업데이트 완료]",
        f"• 코스피/코스닥 상위 200 영업이익 추정치 및 반도체 사이클 갱신 ({now_kst})",
        f"👉 {DASHBOARD_URL}",
    ]
    return "\n".join(lines)


def _build_flow_text(here=HERE):
    d = _read_json(os.path.join(here, "investor_flow.json")) or {}
    by_dt = d.get("byDate") or {}
    latest_dt = sorted(by_dt.keys(), reverse=True)[0] if by_dt else "-"
    snap = by_dt.get(latest_dt) or {}
    invs = snap.get("investors") or snap
    lines = [f"🏦 [키움 수급 업데이트 · {latest_dt}]"]
    for key, label in (("foreign", "외국인"), ("pension", "연기금"), ("private", "사모펀드")):
        buys = ((invs.get(key) or {}).get("buy") or [])[:3]
        if buys:
            names = ", ".join(b.get("name", "") for b in buys if b.get("name"))
            lines.append(f"• {label} TOP3: {names}")
    lines.append(f"👉 {DASHBOARD_URL}")
    return "\n".join(lines)


def _build_semicon_text(here=HERE):
    d = _read_json(os.path.join(here, "semicon_cycle.json")) or {}
    ser = d.get("series") or []
    acts = [s for s in ser if s.get("cycle") is not None and not s.get("isForecast")]
    fcs = [s for s in ser if s.get("cycle") is not None and s.get("isForecast")]
    lines = ["🔄 [반도체 재고순환지표 업데이트]"]
    if acts:
        a = acts[-1]
        lines.append(f"• 확정({a['ym']}): 재고순환 {a['cycle']:+.2f}%p [{a.get('phase', '')}]")
    if fcs:
        fc_txt = ", ".join(f"{x['ym']}(E) {x['cycle']:+.1f}%p" for x in fcs)
        lines.append(f"• 10년 평균 예측: {fc_txt}")
    lines.append(f"👉 {DASHBOARD_URL}")
    return "\n".join(lines)


def _build_kr_text(here=HERE):
    d = _read_json(os.path.join(here, "kr_market.json")) or {}
    td = d.get("tradeDate") or "-"
    sess = d.get("sessionLabel") or "국내 시장 · NXT"
    nxt_val = (d.get("nxtTopVal") or [])[:5]
    nxt_up = (d.get("nxtTopUp") or [])[:4]
    ss = sorted(d.get("sectorStrength") or [], key=lambda x: x.get("avgNxtChg") or -999, reverse=True)
    lines = [f"🇰🇷 [국내 시장 · NXT 업데이트 · {td}]", f"• 구분: {sess}"]
    if nxt_val:
        v_txt = ", ".join(f"{x['name']}({_fmt_pct(x.get('nxtChg'))}, {int(round(x.get('nxtValEok') or 0)):,}억)" for x in nxt_val[:4])
        lines.append(f"• NXT 거래대금 TOP: {v_txt}")
    if nxt_up:
        u_txt = ", ".join(f"{x['name']}({_fmt_pct(x.get('nxtChg'))})" for x in nxt_up[:4])
        lines.append(f"• NXT 급등 TOP: {u_txt}")
    if ss:
        t_txt = ", ".join(f"{x['name']}({_fmt_pct(x.get('avgNxtChg'))})" for x in ss[:3])
        lines.append(f"• 강세 섹터: {t_txt}")
    lines.append(f"👉 {DASHBOARD_URL}")
    return "\n".join(lines)


def notify_update(mode="all", here=HERE):
    """외부 스크립트 또는 GitHub Actions에서 호출하는 통합 알림 함수"""
    tok, cid = _load_telegram_cfg()
    if not tok or not cid:
        print("  [안내] TELEGRAM_BOT_TOKEN / TELEGRAM_CHAT_ID 미설정으로 텔레그램 알림 생략")
        return False

    m = str(mode or "all").strip().lower()
    if m == "us":
        text = _build_us_text(here)
    elif m == "kr":
        text = _build_kr_text(here)
    elif m == "reports":
        text = _build_reports_text(here)
    elif m == "screener":
        text = _build_screener_text(here)
    elif m == "flow":
        text = _build_flow_text(here)
    elif m == "semicon":
        text = _build_semicon_text(here)
    else:
        now_kst = datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")
        us_part = _build_us_text(here).splitlines()[1] if os.path.isfile(os.path.join(here, "us_market.json")) else ""
        lines = [
            f"🚀 [전체 대시보드 업데이트 완료 · {now_kst}]",
            "• 미국 모닝 브리핑 · 기업 리포트 · 코스피/코스닥 200 스크리너 갱신 완료",
        ]
        if us_part:
            lines.append(us_part)
        lines.append(f"👉 {DASHBOARD_URL}")
        text = "\n".join(lines)

    return send_telegram_message(text, token=tok, chat_id=cid)


def main():
    arg = (sys.argv[1] if len(sys.argv) > 1 else "test").strip().lower()
    if arg == "test":
        tok, cid = _load_telegram_cfg()
        if not tok:
            print("[오류] .env 파일에 TELEGRAM_BOT_TOKEN 이 설정되지 않았습니다.")
            print("  예) TELEGRAM_BOT_TOKEN=123456789:ABCdef...")
            sys.exit(1)
        if not cid:
            print("TELEGRAM_CHAT_ID 가 비어있어 최근 봇 메시지에서 Chat ID를 자동 탐색합니다...")
            found = discover_chat_id(tok)
            if not found:
                print("[안내] Chat ID를 찾지 못했습니다. 방금 만든 본인 텔레그램 봇 대화방에 '안녕' 메시지를 한 번 보낸 뒤 다시 실행해 주세요.")
                sys.exit(1)
            cid = found
            print(f"★ Chat ID 자동 탐색 성공: {cid}")
            print(f"  -> .env 파일과 GitHub Secrets에 TELEGRAM_CHAT_ID={cid} 를 등록해 주세요!")
        now_kst = datetime.now(KST).strftime("%Y-%m-%d %H:%M:%S KST")
        ok = send_telegram_message(
            f"✅ [Stock Screener 텔레그램 알림 테스트 성공]\n• 시각: {now_kst}\n• 앞으로 업데이트가 완료되면 이 대화방으로 자동 알림이 옵니다.\n👉 {DASHBOARD_URL}",
            token=tok,
            chat_id=cid,
        )
        sys.exit(0 if ok else 1)

    notify_update(arg, HERE)


if __name__ == "__main__":
    main()
