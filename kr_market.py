# -*- coding: utf-8 -*-
"""
국내 시장 브리핑 · NXT(넥스트레이드) 프리/애프터마켓 집중 매수 · 45개 국내 섹터 강도 보드 수집기
=================================================================================================
- 네이버 증권 공식 API (별도 API 키·고정 IP 불필요 -> 로컬 PC 및 GitHub Actions에서 08:30 KST 자동 실행 가능)
- 코스피/코스닥 시총 상위 ~800종목 + 45개 국내 핵심 테마 구성 종목을 50종목씩 배치 폴링하여
  KRX 정규장 및 NXT(넥스트레이드 프리마켓 08:00~08:50 / 애프터마켓 15:30~20:00) 실시간 체결가·등락률·거래대금 수집
- 결과 저장: kr_market.json -> index.html 웹 대시보드 자동 반영 + GitHub 자동 업로드 + 텔레그램 알림

사용법:
  python kr_market.py             # 수집 -> kr_market.json -> index.html -> GitHub 업로드 + 텔레그램 알림
  python kr_market.py --no-push   # 웹페이지 생성까지만 수행 (GitHub 업로드/알림 생략)
"""

import json
import os
import re
import sys
import urllib.parse
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone, timedelta
from email.utils import parsedate_to_datetime

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
KR_JSON_PATH = os.path.join(HERE, "kr_market.json")
KST = timezone(timedelta(hours=9))

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
NAVER_HDR = {"User-Agent": UA, "Referer": "https://stock.naver.com/"}

# 국내 주요 지수 4종
KR_INDICES = [
    ("KOSPI", "코스피", "유가증권시장", "https://stock.naver.com/domestic/index/KOSPI/price"),
    ("KOSDAQ", "코스닥", "기술·성장주 시장", "https://stock.naver.com/domestic/index/KOSDAQ/price"),
    ("KPI200", "코스피 200", "국내 대표 대형주 200", "https://stock.naver.com/domestic/index/KPI200/price"),
]

# 45개 국내 핵심 테마·섹터 강도 보드 (미국 섹터 강도 보드와 동일 UX)
# (id, 테마명, 아이콘, Google뉴스검색어, [(종목코드, 종목명, 1줄 핵심설명), ...])
KR_THEME_BASKETS = [
    ("kr_semi_big", "반도체 대형·HBM", "🧩", "삼성전자 SK하이닉스 HBM 반도체 주가", [
        ("005930", "삼성전자", "글로벌 메모리(DRAM·낸드)·HBM·파운드리 대표 대장주"),
        ("000660", "SK하이닉스", "엔비디아 HBM 핵심 독주 · AI 메모리 글로벌 선도"),
        ("042700", "한미반도체", "HBM 필수 후공정 TC본더 글로벌 1위 장비사"),
        ("402340", "SK스퀘어", "SK하이닉스 최대주주(약 20%) · 반도체 지주 대장"),
        ("000990", "DB하이텍", "국내 대표 8인치 시스템반도체 파운드리 전문"),
        ("058470", "리노공업", "AI·모바일 반도체 테스트 소켓(리노핀) 글로벌 1위"),
    ]),
    ("kr_semi_equip", "반도체 전공정 장비", "🛠️", "주성엔지니어링 원익IPS 유진테크 HPSP 반도체장비", [
        ("403870", "HPSP", "첨단 미세공정 필수 고압수소어닐링(HPA) 장비 독점"),
        ("036930", "주성엔지니어링", "ALD(원자층증착) 반도체·디스플레이 핵심 장비사"),
        ("240810", "원익IPS", "삼성전자 핵심 증착(PECVD·ALD)·식각 전공정 장비"),
        ("084370", "유진테크", "메모리 반도체용 저압화학기상증착(LPCVD) 전문"),
        ("039030", "이오테크닉스", "반도체 레이저 어닐링·마커·그루빙 장비 선도"),
        ("095610", "테스", "반도체 낸드·D램 증착(PECVD)·건식식각 장비"),
    ]),
    ("kr_semi_osat", "반도체 후공정·소부장", "🔬", "하나마이크론 이수페타시스 ISC 반도체 소부장", [
        ("007660", "이수페타시스", "글로벌 빅테크 AI 가속기용 초다층 PCB(MLB) 핵심"),
        ("095340", "ISC", "AI 서버·HBM용 실리콘 러버 테스트 소켓 글로벌 선도"),
        ("067310", "하나마이크론", "국내 대표 반도체 패키징·테스트(OSAT) 외주 전문"),
        ("131970", "두산테스나", "국내 1위 시스템반도체(CIS·AP) 웨이퍼 테스트 전문"),
        ("036830", "솔브레인", "반도체 식각액(불산)·CMP 슬러리 핵심 공정 소재"),
        ("357780", "솔브레인홀딩스", "반도체 소재·2차전지 전해액 지주사"),
    ]),
    ("kr_glass_sub", "유리 기판", "🪟", "SKC 필옵틱스 와이씨켐 유리기판 주가", [
        ("011790", "SKC", "자회사 앱솔릭스 세계 최초 AI용 반도체 유리기판 상용화"),
        ("161580", "필옵틱스", "유리기판 핵심 TGV(유리관통전극) 레이저 장비 선도"),
        ("112290", "와이씨켐", "반도체 유리기판 전용 포토레지스트·현상액 소재 개발"),
        ("204270", "제이앤티씨", "커버글라스 기술 기반 반도체 유리기판(TGV) 사업 확장"),
        ("078150", "HB테크놀러지", "반도체 유리기판 외관 검사·리페어 장비 공급"),
    ]),
    ("kr_optical", "광통신·AI 네트워크", "🔦", "오이솔루션 대한광통신 쏠리드 광통신 주가", [
        ("138080", "오이솔루션", "AI 데이터센터·5G용 초고속 광트랜시버 국내 대표"),
        ("010170", "대한광통신", "광섬유·광케이블 일괄생산 · AI 데이터센터 광케이블 수혜"),
        ("050890", "쏠리드", "북미·유럽 중계기(DAS)·오픈랜 통신장비 수출 강자"),
        ("032500", "케이엠더블유", "5G 기지국 안테나·필터(MMR) 통신장비 전문"),
        ("218410", "RFHIC", "질화갈륨(GaN) 전력증폭기 · 방산 레이더·통신 핵심"),
    ]),
    ("kr_power_grid", "전력기기·변압기", "🏗️", "HD현대일렉트릭 효성중공업 LS일렉트릭 산일전기", [
        ("267260", "HD현대일렉트릭", "북미 초고압 변압기·AI 데이터센터 전력망 대장주"),
        ("298040", "효성중공업", "초고압 변압기·차단기 북미·유럽 수주 급증 수혜"),
        ("010120", "LS ELECTRIC", "배전반·초고압 변압기·미국 빅테크 데이터센터 배전"),
        ("062040", "산일전기", "특수변압기·신재생·북미 전력망용 변압기 고수익 강자"),
        ("001440", "대한전선", "초고압 전력케이블·해저케이블 글로벌 전력망 수주"),
        ("103590", "일진전기", "초고압 변압기·전력케이블 동시 보유 종합 중전기"),
    ]),
    ("kr_nuclear", "원자력·SMR", "☢️", "두산에너빌리티 한전기술 우리기술 원전 SMR", [
        ("034020", "두산에너빌리티", "원전 주기기·미국 뉴스케일/X에너지 SMR 핵심 제작사"),
        ("052690", "한전기술", "국내 원전 종합설계·원자로계통설계 독점 공기업"),
        ("000720", "현대건설", "대형 원전·미국 홀텍 SMR 시공 글로벌 원전 건설 대장"),
        ("051600", "한전KPS", "원자력·화력 발전설비 정비 전문 공기업"),
        ("032820", "우리기술", "원전 핵심 제어계측시스템(MMIS) 국산화 공급"),
        ("105840", "우진", "원자력 발전소용 핵심 계측기(ICI 등) 독점 공급"),
    ]),
    ("kr_defense", "방산·K-방산", "🛡️", "한화에어로스페이스 현대로템 LIG 방산 주가", [
        ("012450", "한화에어로스페이스", "K9 자주포·천무·항공엔진·우주발사체 K-방산 대장주"),
        ("064350", "현대로템", "K2 전차 폴란드·중동 대규모 수출 및 철도차량"),
        ("079550", "LIG디펜스앤에어로스페이스", "천궁-II 지대공 유도무기·정밀타격 방산 핵심"),
        ("047810", "한국항공우주", "KF-21 전투기·FA-50 경공격기·우주위성 국내 유일 완제기"),
        ("272210", "한화시스템", "AESA 레이더·방산전자·저궤도 위성통신(원웹) 핵심"),
        ("103140", "풍산", "대구경 탄약 수출 급증 및 구리(신동) 사업 병행"),
    ]),
    ("kr_ship", "조선·해양·LNG선", "🚢", "HD한국조선해양 한화오션 삼성중공업 조선주", [
        ("009540", "HD한국조선해양", "글로벌 1위 조선 지주사 · LNG선·친환경선 수주 선도"),
        ("329180", "HD현대중공업", "초대형 LNG선·특수선(함정)·선박엔진 글로벌 1위"),
        ("042660", "한화오션", "LNG선·미 해군 함정 MRO·잠수함·해양플랜트 강자"),
        ("010140", "삼성중공업", "FLNG(부유식 액화천연가스 설비) 글로벌 독보적 경쟁력"),
        ("443060", "HD현대마린솔루션", "친환경 선박 개조·AM(애프터마켓) 고마진 조선 서비스 1위"),
        ("082740", "한화엔진", "선박용 저속 디젤·이중연료(DF) 친환경 추진엔진 전문"),
    ]),
    ("kr_battery", "2차전지 셀·소재", "🔋", "LG에너지솔루션 삼성SDI 에코프로비엠 2차전지", [
        ("373220", "LG에너지솔루션", "글로벌 전기차·북미 ESS 배터리 국내 시총 1위 셀 제조"),
        ("006400", "삼성SDI", "각형 프리미엄 배터리·ESS·전고체(ASB) 배터리 선도"),
        ("247540", "에코프로비엠", "하이니켈 양극재 글로벌 선도 · 코스닥 2차전지 대장"),
        ("003670", "포스코퓨처엠", "국내 유일 양극재·음극재(인조흑연) 동시 생산 소재사"),
        ("066970", "엘앤에프", "테슬라향 하이니켈 NCMA 양극재 핵심 공급사"),
        ("086520", "에코프로", "에코프로비엠·머티리얼즈 보유 2차전지 수직계열 지주"),
    ]),
    ("kr_solid_bat", "전고체·차세대 배터리", "⚡", "이수스페셜티케미컬 레이크머티리얼즈 전고체 배터리", [
        ("006400", "삼성SDI", "2027년 양산 목표 황화물계 전고체 배터리 국내 최선두"),
        ("457190", "이수스페셜티케미컬", "황화물계 전고체 핵심 원료 황화리튬(Li2S) 국내 선도"),
        ("281740", "레이크머티리얼즈", "반도체 전구체 및 전고체 배터리용 황화리튬 소재 확장"),
        ("278280", "천보", "차세대 2차전지 특수 전해질(LiFSI 등) 소재 전문"),
        ("348370", "엔켐", "북미 현지 공장 보유 글로벌 2차전지 전해액 핵심"),
    ]),
    ("kr_ess", "ESS·전력저장장치", "🔌", "서진시스템 한중엔시에스 LG에너지솔루션 ESS", [
        ("373220", "LG에너지솔루션", "북미 AI 데이터센터·전력망용 LFP ESS 배터리 대량 수주"),
        ("178320", "서진시스템", "글로벌 1위 플루언스향 ESS 함체·인클로저 핵심 공급"),
        ("107640", "한중엔시에스", "삼성SDI향 수냉식 ESS 냉각모듈(Chiller·HVAC) 공급"),
        ("006400", "삼성SDI", "삼성 배터리 박스(SBB) 북미 유틸리티 ESS 고속 성장"),
        ("010120", "LS ELECTRIC", "ESS 전력변환장치(PCS) 및 배전 통합 솔루션"),
    ]),
    ("kr_robot", "로봇·휴머노이드", "🤖", "레인보우로보틱스 두산로보틱스 로보티즈 로봇주", [
        ("277810", "레인보우로보틱스", "삼성전자 자회사 · 휴머노이드·협동로봇 국내 대장주"),
        ("454910", "두산로보틱스", "글로벌 상위권 협동로봇 라인업 및 북미 자동화 확장"),
        ("108490", "로보티즈", "로봇 핵심 관절 액추에이터(다이나믹셀)·자율주행 로봇"),
        ("058610", "에스피지", "로봇용 정밀감속기(SH/SR) 국산화 선도 핵심 부품사"),
        ("090710", "휴림로봇", "산업용 직교·스카라 로봇 및 스마트팩토리 자동화"),
        ("348340", "뉴로메카", "협동로봇·자율이동로봇(AMR)·F&B 자동화 전문"),
    ]),
    ("kr_auto", "자동차·전장·자율주행", "🚗", "현대차 기아 현대모비스 현대오토에버 주가", [
        ("005380", "현대차", "글로벌 Top3 완성차 · 하이브리드·보스턴다이내믹스 로봇"),
        ("000270", "기아", "글로벌 고수익 완성차 · RV·HEV·EV 라인업 강점"),
        ("012330", "현대모비스", "현대차그룹 핵심 전동화·자율주행 모듈 및 A/S 부품"),
        ("307950", "현대오토에버", "현대차그룹 차량용 SW(모빌진)·자율주행 정밀지도·SDV"),
        ("204320", "HL만도", "자율주행 조향(SbW)·제동(EMB) 핵심 전장 부품사"),
        ("005850", "에스엘", "글로벌 완성차용 LED 헤드램프·전동화 부품 강자"),
    ]),
    ("kr_space_drone", "우주항공·드론", "🚀", "쎄트렉아이 AP위성 한화시스템 우주항공 주가", [
        ("012450", "한화에어로스페이스", "누리호 체계종합기업 · 한국판 스페이스X 주도"),
        ("047810", "한국항공우주", "차세대 중형위성·군 정찰위성·발사체 체계 개발"),
        ("099320", "쎄트렉아이", "초고해상도 지구관측 인공위성 본체·탑재체 수출 강자"),
        ("211270", "AP위성", "위성통신 단말기·다누리 달 탐사 탑재컴퓨터 개발"),
        ("010820", "퍼스텍", "방산 유도무기·드론·무인기 핵심 부품 공급"),
    ]),
    ("kr_bio_big", "대형 바이오·CDMO", "💊", "삼성바이오로직스 셀트리온 알테오젠 바이오", [
        ("207940", "삼성바이오로직스", "글로벌 1위 바이오의약품 위탁생산(CDMO) 압도적 수주"),
        ("068270", "셀트리온", "짐펜트라 미국 직판·자가면역/항암 바이오시밀러 선도"),
        ("196170", "알테오젠", "정맥주사→피하주사(SC) 제형변경 플랫폼(ALT-B4) 글로벌 독주"),
        ("000100", "유한양행", "국산 최초 FDA 승인 항암신약 렉라자(레이저티닙) 로열티"),
        ("326030", "SK바이오팜", "뇌전증 신약 엑스코프리 미국 직판 고성장·방사성의약품"),
    ]),
    ("kr_bio_new", "신약·ADC·바이오텍", "🧬", "리가켐바이오 에이비엘바이오 HLB 올릭스", [
        ("141080", "리가켐바이오", "국내 대표 ADC(항체약물접합체) 플랫폼·글로벌 기술수출"),
        ("298380", "에이비엘바이오", "이중항체·뇌혈관장벽(BBB) 셔틀 플랫폼(그랩바디-B)"),
        ("028300", "HLB", "간암 1차 치료제 리보세라닙 글로벌 허가 추진"),
        ("226950", "올릭스", "RNA 간섭(siRNA) 기반 비만·탈모·황반변성 신약 개발"),
        ("009420", "한올바이오파마", "자가면역질환 FcRn 항체 신약(바토클리맙·IMVT-1402)"),
    ]),
    ("kr_obesity", "비만치료제·미용의료", "💉", "한미약품 펩트론 디앤디파마텍 에이피알", [
        ("128940", "한미약품", "한국인 맞춤 GLP-1 비만신약 에페글레나타이드 임상 선도"),
        ("087010", "펩트론", "장기지속형 주사제(스마트데포) 비만약 플랫폼 기술"),
        ("347850", "디앤디파마텍", "경구용 GLP-1 비만치료제 미국 기술이전(멧세라)"),
        ("278470", "에이피알", "메디큐브 뷰티 디바이스 글로벌 폭발 성장 K-뷰티 대장"),
        ("214150", "클래시스", "슈링크·볼뉴머 피부미용 의료기기 글로벌 고마진 성장"),
        ("145020", "휴젤", "보툴리눔 톡신(보툴렉스) 미국·중국·유럽 글로벌 허가"),
    ]),
    ("kr_ai_sw", "AI 플랫폼·소프트웨어", "💻", "NAVER 카카오 솔트룩스 이스트소프트 AI", [
        ("035420", "NAVER", "검색·커머스·하이퍼클로바X AI·클라우드 국내 플랫폼 대장"),
        ("035720", "카카오", "카카오톡 기반 AI 에이전트(카나나)·페이·모빌리티"),
        ("304100", "솔트룩스", "국내 대표 생성형 AI(루시아)·기업용 LLM 구축"),
        ("047560", "이스트소프트", "AI 휴먼·알툴즈·페르소 AI 글로벌 SaaS 확장"),
        ("402030", "코난테크놀로지", "국방·공공 AI LLM 및 빅데이터 분석 소프트웨어"),
        ("030520", "한컴", "한컴오피스 기반 AI 문서작성(한컴어시스턴트) 사업"),
    ]),
    ("kr_quantum", "양자암호·보안", "⚛️", "엑스게이트 케이씨에스 우리로 양자컴퓨팅", [
        ("356680", "엑스게이트", "SK텔레콤 협력 양자난수생성기(QRNG) 탑재 양자VPN 선도"),
        ("115500", "케이씨에스", "SKT와 양자암호 원칩(Q-HSM) 공동개발 보안칩 전문"),
        ("046970", "우리로", "양자암호통신 핵심 단일광자검출기(SPAD) 원천기술 보유"),
        ("056360", "코위버", "양자암호화 적용 광전송장비(ROADM·POTN) 공급"),
        ("053800", "안랩", "국내 1위 사이버보안 소프트웨어·클라우드 보안 관제"),
    ]),
    ("kr_solar_wind", "태양광·풍력·신재생", "☀️", "한화솔루션 씨에스윈드 OCI홀딩스 신재생", [
        ("009830", "한화솔루션", "미국 조지아 태양광 통합생산단지(솔라허브) IRA 최대 수혜"),
        ("010060", "OCI홀딩스", "말레이시아 비중국 태양광용 폴리실리콘 글로벌 핵심"),
        ("112610", "씨에스윈드", "글로벌 1위 풍력 타워 및 해상풍력 하부구조물(블라트)"),
        ("100090", "SK오션플랜트", "아시아 선도 해상풍력 하부구조물(재킷)·특수선 건조"),
        ("322000", "HD현대에너지솔루션", "고효율 태양광 모듈·인버터·솔루션 공급"),
    ]),
    ("kr_hydrogen", "수소·연료전지", "💧", "두산퓨얼셀 비나텍 범한퓨얼셀 수소", [
        ("336260", "두산퓨얼셀", "국내 1위 발전용 인산형·고체산화물 연료전지(SOFC)"),
        ("126340", "비나텍", "수소연료전지 핵심 부품(MEA·지지체) 및 슈퍼커패시터"),
        ("382900", "범한퓨얼셀", "잠수함용 수소연료전지 독점 및 건물용·수소충전소"),
        ("271940", "일진하이솔루스", "현대차 넥쏘 수소저장탱크(Type4) 독점 공급"),
    ]),
    ("kr_cooling", "데이터센터 냉각·열관리", "❄️", "워트 GST 유니셈 LG전자 액침냉각", [
        ("396470", "워트", "반도체 초정밀 온습도 제어장비(THC) 및 냉각 솔루션"),
        ("083450", "GST", "반도체 칠러(온도조절)·스크러버 및 데이터센터 액침냉각 개발"),
        ("036200", "유니셈", "반도체용 칠러·스크러버 장비 핵심 공급사"),
        ("066570", "LG전자", "북미 AI 데이터센터 초대형 냉동기(칠러)·HVAC 급성장"),
    ]),
    ("kr_bank", "은행·금융지주", "🏦", "KB금융 신한지주 하나금융지주 밸류업 은행", [
        ("105560", "KB금융", "국내 리딩 금융지주 · 주주환원(밸류업) 배당/자사주 선도"),
        ("055550", "신한지주", "균형 잡힌 은행·비은행 포트폴리오 및 분기배당·자사주"),
        ("086790", "하나금융지주", "외환·기업금융 강점 및 높은 배당수익률 대표 금융주"),
        ("316140", "우리금융지주", "기업금융 명가 · 보험사 인수로 비은행 시너지 확대"),
        ("138040", "메리츠금융지주", "압도적 ROE·자사주 매입소각 국내 밸류업 교과서"),
    ]),
    ("kr_sec_ins", "증권·보험", "📈", "키움증권 미래에셋증권 삼성화재 삼성생명", [
        ("039490", "키움증권", "국내 리테일 브로커리지(개인 점유율) 1위 · 해외주식 강자"),
        ("006800", "미래에셋증권", "국내 최대 자기자본 · 글로벌 IB·인도/미국 해외법인 확장"),
        ("071050", "한국금융지주", "한국투자증권 보유 · IB·운용 압도적 수익성"),
        ("000810", "삼성화재", "국내 손해보험 압도적 1위 · K-ICS 최고 자본비율"),
        ("032830", "삼성생명", "국내 생명보험 1위 · 삼성전자 지분 보유 금융 대장"),
    ]),
    ("kr_steel_metal", "철강·비철금속·구리", "⛏️", "POSCO홀딩스 고려아연 풍산 대원전선", [
        ("005490", "POSCO홀딩스", "국내 최대 철강사 및 아르헨티나 리튬·2차전지 소재"),
        ("010130", "고려아연", "글로벌 1위 아연·은·전략광물(안티모니·인듐) 제련사"),
        ("004020", "현대제철", "현대차그룹 일관제철소 · 미국 전기로 제철소 투자"),
        ("006340", "대원전선", "전력·통신·자동차용 전선 제조 · 구리가격/전력망 민감"),
        ("000500", "가온전선", "LS그룹 전력·배전 케이블 핵심 계열사"),
        ("006260", "LS", "LS전선(해저케이블)·LS MnM(구리제련)·LS일렉트릭 지주"),
    ]),
    ("kr_oil_chem", "정유·석유화학", "🛢️", "S-Oil SK이노베이션 LG화학 정유주", [
        ("010950", "S-Oil", "아람코 대주주 · 정제마진·샤힌 프로젝트 석유화학 확장"),
        ("096770", "SK이노베이션", "SK에너지(정유)·SK E&S(LNG·전력)·SK온(배터리) 통합"),
        ("078930", "GS", "GS칼텍스(정유)·GS EPS(발전)·GS리테일 에너지 지주"),
        ("051910", "LG화학", "LG에너지솔루션 모회사 · 양극재·생명과학·석유화학"),
    ]),
    ("kr_const_mach", "건설·건설기계", "🏗️", "현대건설 두산밥캣 HD건설기계 DL이앤씨", [
        ("000720", "현대건설", "원전·해외 플랜트·국내 주택 대표 종합건설사"),
        ("241560", "두산밥캣", "북미 소형 건설장비(스کید로더) 점유율 1위 · 미국 인프라 수혜"),
        ("267270", "HD건설기계", "HD현대건설기계·인프라코어 통합 글로벌 중장비 시너지"),
        ("375500", "DL이앤씨", "재무건전성 최상위 건설사 · 플랜트·SMR 협력"),
        ("006360", "GS건설", "자이 브랜드 주택 및 수처리·모듈러 신사업"),
    ]),
    ("kr_elec_comp", "전자부품·IT하드웨어", "📱", "삼성전기 LG이노텍 삼성전자 전자부품", [
        ("009150", "삼성전기", "AI 서버·전장용 초고용량 MLCC 및 FC-BGA 기판 핵심"),
        ("011070", "LG이노텍", "애플 아이폰 카메라모듈(폴디드줌) 글로벌 1위 공급사"),
        ("034220", "LG디스플레이", "대형 OLED·애플향 중소형 OLED 실적 턴어라운드"),
        ("009155", "삼성전기우", "삼성전기 우선주 (배당 + MLCC·유리기판 모멘텀)"),
    ]),
    ("kr_cosmetic", "화장품·K-뷰티", "💄", "아모레퍼시픽 코스맥스 실리콘투 한국콜마", [
        ("090430", "아모레퍼시픽", "코스알엑스(COSRX) 인수·북미/유럽 중심 K-뷰티 대표"),
        ("192820", "코스맥스", "글로벌 1위 화장품 ODM · 인디브랜드 수출 폭발 수혜"),
        ("257720", "실리콘투", "미국·유럽 글로벌 K-뷰티 역직구 유통 플랫폼 대장"),
        ("161890", "한국콜마", "선케어·기초화장품 글로벌 최상위권 ODM 제조사"),
        ("051900", "LG생활건강", "더후·화장품·생활용품·코카콜라음료 종합 소비재"),
    ]),
    ("kr_food", "K-푸드·음식료", "🍜", "삼양식품 농심 CJ제일제당 오리온", [
        ("003230", "삼양식품", "불닭볶음면 글로벌 신드롬 · 밀양 신공장 증설 K-푸드 대장"),
        ("004370", "농심", "신라면·짜파게티 미국·유럽 글로벌 라면 수출 성장"),
        ("097950", "CJ제일제당", "비비고 만두 글로벌 1위 · 바이오·식품 종합 대기업"),
        ("271560", "오리온", "중국·베트남·러시아 해외 제과 고수익 성장"),
        ("005180", "빙그레", "메로나·바나나맛우유 북미·글로벌 아이스크림 수출"),
    ]),
    ("kr_ent", "엔터·K-팝·미디어", "🎤", "하이브 에스엠 JYP YG 스튜디오드래곤", [
        ("352820", "하이브", "BTS·세븐틴·위버스 플랫폼 글로벌 K-팝 1위 엔터"),
        ("041510", "에스엠", "에스파·라이즈·NCT 등 IP 및 카카오 시너지"),
        ("035900", "JYP Ent.", "스트레이키즈·트와이스 북미·일본 투어 고수익"),
        ("122870", "와이지엔터테인먼트", "블랙핑크·베이비몬스터 글로벌 투어 모멘텀"),
        ("253450", "스튜디오드래곤", "글로벌 OTT(넷플릭스·디즈니+)향 K-드라마 제작 1위"),
    ]),
    ("kr_game", "게임", "🎮", "크래프톤 시프트업 엔씨소프트 넷마블", [
        ("259960", "크래프톤", "배틀그라운드(PUBG) 글로벌 흥행·인조이 신작 게임 대장"),
        ("462870", "시프트업", "승리의 여신: 니케·스텔라 블레이드 글로벌 고마진 흥행"),
        ("251270", "넷마블", "나 혼자만 레벨업 등 글로벌 IP 신작 턴어라운드"),
        ("036570", "엔씨소프트", "리니지 IP 및 글로벌 신작 라인업 구조조정 반등"),
        ("263750", "펄어비스", "검은사막 글로벌 서비스 및 대작 '붉은사막' 모멘텀"),
    ]),
    ("kr_air_travel", "항공·여행·카지노", "✈️", "대한항공 파라다이스 롯데관광개발 하나투어", [
        ("003490", "대한항공", "아시아나 합병 글로벌 메가 캐리어 · 여객·화물 1위"),
        ("034230", "파라다이스", "외국인 전용 카지노·영종도 복합리조트 1위"),
        ("032350", "롯데관광개발", "제주 드림타워 외국인 카지노·호텔 매출 급성장"),
        ("114090", "GKL", "서울·부산 세븐럭 외국인 전용 카지노 공기업"),
        ("039130", "하나투어", "국내 1위 여행사 · 아웃바운드 패키지 고마진 전환"),
    ]),
    ("kr_shipping", "해운·물류", "📦", "HMM 팬오션 현대글로비스 대한해운", [
        ("011200", "HMM", "국내 최대 컨테이너 국적선사 · 해상운임(SCFI) 직결"),
        ("086280", "현대글로비스", "현대차·기아 완성차 해상운송(PCTC)·물류 고배당 성장"),
        ("028670", "팬오션", "국내 최대 벌크선사(BDI 민감) 및 LNG선 장기계약 확대"),
        ("000120", "CJ대한통운", "국내 택배 1위 · 이커머스(알리·신세계) 풀필먼트 수혜"),
    ]),
    ("kr_telecom", "통신·배당 방어주", "📶", "SK텔레콤 KT LG유플러스 통신주", [
        ("017670", "SK텔레콤", "국내 무선 1위 · AI 데이터센터(AIDC)·고배당 대표주"),
        ("030200", "KT", "마이크로소프트 AI 클라우드 협력 · 부동산·고배당 밸류업"),
        ("032640", "LG유플러스", "유무선 통신·AI 콜센터(익시)·안정적 배당 수익률"),
    ]),
    ("kr_retail", "유통·편의점·백화점", "🛒", "이마트 BGF리테일 신세계 현대백화점", [
        ("139480", "이마트", "국내 1위 대형마트·트레이더스·스타벅스코리아 보유"),
        ("282330", "BGF리테일", "CU 편의점 1위 · 해외 진출 및 안정적 현금창출"),
        ("004170", "신세계", "강남·명품 중심 백화점·면세점·신세계인터내셔날"),
        ("069960", "현대백화점", "더현대서울 흥행 · 지누스·면세점 종합 유통"),
        ("007070", "GS리테일", "GS25 편의점·GS더프레시 슈퍼마켓·홈쇼핑"),
    ]),
    ("kr_crypto", "가상자산·비트코인 관련", "₿", "우리기술투자 한화투자증권 위지트 비트코인", [
        ("041190", "우리기술투자", "두나무(업비트) 지분 약 7.2% 보유 국내 대표 가상자산주"),
        ("003530", "한화투자증권", "두나무(업비트) 지분 약 5.9% 보유 증권사"),
        ("036090", "위지트", "빗썸코리아 관계사 티사이언티픽 대주주"),
        ("094480", "갤럭시아머니트리", "전자결제·STO(토큰증권)·블록체인 신사업"),
    ]),
    ("kr_fintech", "핀테크·간편결제", "💳", "카카오페이 NHN KCP 다날 결제", [
        ("377300", "카카오페이", "카카오톡 기반 국내 대표 간편결제·증권·보험 핀테크"),
        ("060250", "NHN KCP", "국내 온라인 전자결제(PG) 점유율 1위 · 해외가맹점 강점"),
        ("035600", "KG이니시스", "국내 대표 온라인 PG 결제사 및 자회사 시너지"),
        ("064260", "다날", "휴대폰 결제 선도 및 스테이블코인·글로벌 결제 확장"),
    ]),
    ("kr_holdco", "대기업 지주사·밸류업", "🏛️", "삼성물산 SK 한화 두산 지주사", [
        ("028260", "삼성물산", "삼성그룹 실질 지주 · 건설·상사·삼성바이오 지분 보유"),
        ("034730", "SK", "SK하이닉스·SK이노베이션·SKT 보유 SK그룹 지주"),
        ("000880", "한화", "한화에어로·한화오션·한화솔루션 핵심 방산/조선 지주"),
        ("000150", "두산", "엔비디아향 동박적층판(CCL) 전자BG 폭풍성장 + 두산에너빌"),
        ("003550", "LG", "LG전자·LG화학·LG CNS 보유 안정적 현금·배당 지주"),
    ]),
    ("kr_util", "전력·가스 유틸리티", "💡", "한국전력 한국가스공사 지역난방공사", [
        ("015760", "한국전력", "국내 독점 전력공급 공기업 · 전기요금·연료비 안정 수혜"),
        ("036460", "한국가스공사", "국내 LNG 도입·배관망 독점 · 미수금 회수 및 배당"),
        ("071320", "지역난방공사", "집단에너지 열·전기 공급 · 요금 정상화 및 고배당"),
    ]),
    ("kr_top10", "코스피 시총 TOP10", "👑", "코스피 시가총액 상위 대형주", [
        ("005930", "삼성전자", "코스피 시총 1위 · 글로벌 반도체·모바일·가전"),
        ("000660", "SK하이닉스", "코스피 시총 2위 · 글로벌 AI HBM 메모리 선도"),
        ("373220", "LG에너지솔루션", "글로벌 2차전지·ESS 배터리 대표주"),
        ("207940", "삼성바이오로직스", "글로벌 1위 바이오 CDMO 대표주"),
        ("005380", "현대차", "글로벌 Top3 완성차·로보틱스 대표주"),
        ("000270", "기아", "글로벌 고수익 완성차·RV 대표주"),
        ("005935", "삼성전자우", "삼성전자 우선주 · 높은 배당수익률"),
        ("068270", "셀트리온", "글로벌 바이오시밀러·신약 대표주"),
        ("105560", "KB금융", "국내 리딩 금융지주 · 밸류업 대표주"),
        ("012450", "한화에어로스페이스", "K-방산·우주항공 시총 대장주"),
    ]),
]


def _num(x):
    if x is None:
        return None
    s = str(x).replace(",", "").replace("%", "").replace("+", "").strip()
    if s in ("", "-", "N/A", "null", "None", "nan", "NaN"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _is_clean_kr_stock(name, code):
    if not code or not re.fullmatch(r"\d{6}", str(code)):
        return False
    nm = str(name or "").strip()
    if not nm:
        return False
    bad_prefixes = (
        "KODEX", "TIGER", "KBSTAR", "RISE", "ACE", "SOL", "HANARO", "ARIRANG",
        "PLUS", "KOSEF", "TIMEFOLIO", "WOORI", "WON", "KIWOOM", "BNK", "1Q",
        "KoAct", "히어로즈", "파워", "마이티", "FOCUS", "KINDEX", "TREX", "SMART",
    )
    up = nm.upper()
    if any(up.startswith(p.upper() + " ") or up.startswith(p.upper()) for p in bad_prefixes):
        return False
    if any(w in up for w in ("ETN", "스팩", "SPAC", "리츠", "REIT", "인버스", "레버리지", "선물")):
        return False
    return True


def fetch_kr_indices():
    """네이버 증권 API에서 코스피·코스닥·코스피200 지수 및 원/달러 환율 수집"""
    out = []
    for code, label, desc, url in KR_INDICES:
        try:
            r = requests.get(f"https://m.stock.naver.com/api/index/{code}/basic", headers=NAVER_HDR, timeout=8)
            if r.status_code != 200:
                continue
            j = r.json()
            price = _num(j.get("closePrice"))
            chg_pct = _num(j.get("fluctuationsRatio"))
            abs_chg = _num(j.get("compareToPreviousClosePrice"))
            cmp_code = str(((j.get("compareToPreviousPrice") or {}).get("code")) or "")
            if cmp_code in ("4", "5") and abs_chg is not None and abs_chg > 0:
                abs_chg = -abs_chg
            if cmp_code in ("4", "5") and chg_pct is not None and chg_pct > 0:
                chg_pct = -chg_pct
            out.append({
                "code": code,
                "label": label,
                "desc": desc,
                "url": url,
                "price": price,
                "absChg": abs_chg,
                "chgPct": chg_pct,
                "tradedAt": str(j.get("localTradedAt") or "")[:16].replace("T", " "),
                "status": str(j.get("marketStatus") or ""),
            })
        except Exception:
            continue

    # 원/달러 환율 (네이버 외환 API)
    try:
        r = requests.get(
            "https://m.stock.naver.com/front-api/marketIndex/productDetail",
            params={"category": "exchange", "reutersCode": "FX_USDKRW"},
            headers=NAVER_HDR,
            timeout=8,
        )
        if r.status_code == 200:
            res = (r.json() or {}).get("result") or {}
            price = _num(res.get("calcPrice") or res.get("closePrice"))
            chg_pct = _num(res.get("fluctuationsRatio"))
            abs_chg = _num(res.get("fluctuations"))
            ft = str(((res.get("fluctuationsType") or {}).get("code")) or "")
            if ft in ("4", "5") and abs_chg is not None and abs_chg > 0:
                abs_chg = -abs_chg
            if ft in ("4", "5") and chg_pct is not None and chg_pct > 0:
                chg_pct = -chg_pct
            if price:
                out.append({
                    "code": "USDKRW",
                    "label": "원/달러 환율",
                    "desc": "하나은행 고시환율",
                    "url": "https://stock.naver.com/marketindex/exchange/FX_USDKRW",
                    "price": price,
                    "absChg": abs_chg,
                    "chgPct": chg_pct,
                    "tradedAt": str(res.get("localTradedAt") or "")[:16].replace("T", " "),
                    "status": "FX",
                })
    except Exception:
        pass
    return out


def fetch_kr_top_universe(pages_per_market=4):
    """코스피·코스닥 시가총액 상위 ~800개 보통주 코드·종목명·시총 목록 수집"""
    uni = {}
    tasks = []
    for mkt in ("KOSPI", "KOSDAQ"):
        for pg in range(1, pages_per_market + 1):
            tasks.append((mkt, pg))

    def _fetch_page(mkt, pg):
        try:
            r = requests.get(
                f"https://m.stock.naver.com/api/stocks/marketValue/{mkt}",
                params={"page": pg, "pageSize": 100},
                headers=NAVER_HDR,
                timeout=10,
            )
            if r.status_code != 200:
                return []
            rows = (r.json() or {}).get("stocks") or []
            parsed = []
            for s in rows:
                cd = str(s.get("itemCode") or "").strip()
                nm = str(s.get("stockName") or "").strip()
                if not _is_clean_kr_stock(nm, cd):
                    continue
                mcap_eok = round((_num(s.get("marketValueRaw")) or 0) / 1e8, 1)
                parsed.append({"code": cd, "name": nm, "market": mkt, "mcapEok": mcap_eok})
            return parsed
        except Exception:
            return []

    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(_fetch_page, m, p) for m, p in tasks]
        for fut in as_completed(futs):
            for it in fut.result():
                uni.setdefault(it["code"], it)
    return uni


def fetch_realtime_batch(codes):
    """네이버 실시간 배치 폴링 API(50종목씩)로 KRX 정규장 + NXT(프리/애프터마켓) 체결가·등락률·거래대금 일괄 수집"""
    chunks = [codes[i:i + 50] for i in range(0, len(codes), 50)]
    results = {}

    def _poll_chunk(chunk):
        try:
            r = requests.get(
                f"https://polling.finance.naver.com/api/realtime/domestic/stock/{','.join(chunk)}",
                headers=NAVER_HDR,
                timeout=12,
            )
            if r.status_code != 200:
                return []
            return (r.json() or {}).get("datas") or []
        except Exception:
            return []

    with ThreadPoolExecutor(max_workers=12) as ex:
        futs = [ex.submit(_poll_chunk, ch) for ch in chunks]
        for fut in as_completed(futs):
            for d in fut.result():
                cd = str(d.get("itemCode") or "").strip()
                if not cd:
                    continue
                krx_price = _num(d.get("closePrice"))
                krx_chg = _num(d.get("fluctuationsRatio"))
                krx_cmp = str(((d.get("compareToPreviousPrice") or {}).get("code")) or "")
                if krx_cmp in ("4", "5") and krx_chg is not None and krx_chg > 0:
                    krx_chg = -krx_chg
                # 정규장 거래대금 (억)
                krx_val_str = str(d.get("accumulatedTradingValue") or "")
                krx_vol = int(_num(d.get("accumulatedTradingVolume")) or 0)
                krx_val_eok = None
                if krx_price and krx_vol:
                    krx_val_eok = round((krx_price * krx_vol) / 1e8, 1)

                # NXT (넥스트레이드 프리/애프터마켓) 정보
                ov = d.get("overMarketPriceInfo") or {}
                has_nxt = bool(ov and ov.get("overPrice"))
                nxt_price = _num(ov.get("overPrice")) if has_nxt else None
                nxt_chg = _num(ov.get("fluctuationsRatio")) if has_nxt else None
                nxt_cmp = str(((ov.get("compareToPreviousPrice") or {}).get("code")) or "")
                if nxt_cmp in ("4", "5") and nxt_chg is not None and nxt_chg > 0:
                    nxt_chg = -nxt_chg
                nxt_val_raw = _num(ov.get("accumulatedTradingValueRaw")) if has_nxt else None
                nxt_vol_raw = _num(ov.get("accumulatedTradingVolumeRaw")) if has_nxt else None
                nxt_val_eok = round(nxt_val_raw / 1e8, 1) if nxt_val_raw is not None else (
                    round((nxt_price * nxt_vol_raw) / 1e8, 1) if (nxt_price and nxt_vol_raw) else None
                )
                nxt_session = str(ov.get("tradingSessionType") or "") if has_nxt else ""
                nxt_status = str(ov.get("overMarketStatus") or "") if has_nxt else ""
                nxt_time = str(ov.get("localTradedAt") or "")[:16].replace("T", " ") if has_nxt else ""

                ex_name = str(((d.get("stockExchangeType") or {}).get("name")) or "")
                results[cd] = {
                    "code": cd,
                    "name": str(d.get("stockName") or "").strip(),
                    "market": "KOSDAQ" if "KOSDAQ" in ex_name.upper() else "KOSPI",
                    "krxPrice": int(round(krx_price)) if krx_price is not None else None,
                    "krxChg": round(krx_chg, 2) if krx_chg is not None else None,
                    "krxVol": krx_vol,
                    "krxValEok": krx_val_eok,
                    "krxValTxt": krx_val_str,
                    "krxStatus": str(d.get("marketStatus") or ""),
                    "krxTime": str(d.get("localTradedAt") or "")[:16].replace("T", " "),
                    "hasNxt": has_nxt,
                    "nxtPrice": int(round(nxt_price)) if nxt_price is not None else None,
                    "nxtChg": round(nxt_chg, 2) if nxt_chg is not None else None,
                    "nxtValEok": nxt_val_eok,
                    "nxtValTxt": str(ov.get("accumulatedTradingValue") or "") if has_nxt else "",
                    "nxtVol": int(round(nxt_vol_raw)) if nxt_vol_raw is not None else 0,
                    "nxtSession": nxt_session,
                    "nxtStatus": nxt_status,
                    "nxtTime": nxt_time,
                }
    return results


def fetch_google_kr_news(query, limit=3, days=3):
    """Google News KR RSS에서 국내 시장/테마 최신 뉴스 수집"""
    try:
        url = (
            "https://news.google.com/rss/search?"
            f"q={urllib.parse.quote(f'{query} when:{days}d')}&hl=ko&gl=KR&ceid=KR:ko"
        )
        r = requests.get(url, headers={"User-Agent": UA}, timeout=9)
        if r.status_code != 200:
            return []
        root = ET.fromstring(r.text)
        items = []
        seen = set()
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
            norm = re.sub(r"[^0-9A-Za-z가-힣]", "", title)[:16]
            if not norm or norm in seen:
                continue
            seen.add(norm)
            pub_ts = 0
            pub_kst = ""
            if pub_raw:
                try:
                    dt = parsedate_to_datetime(pub_raw).astimezone(KST)
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
        items.sort(key=lambda x: x["pubTs"], reverse=True)
        return [
            {"title": x["title"], "press": x["press"], "link": x["link"], "pubKst": x["pubKst"]}
            for x in items[:limit]
        ]
    except Exception:
        return []


def collect_kr_market(here=HERE):
    print("=" * 80)
    print(" ▶ 국내 시장 브리핑 · NXT(넥스트레이드) 집중 매수 · 국내 섹터 강도 수집 시작")
    print("=" * 80)

    print("[1/4] 코스피/코스닥 지수 · 환율 및 시총 상위 ~800종목 목록 스캔 중...")
    indices = fetch_kr_indices()
    uni = fetch_kr_top_universe(pages_per_market=4)

    theme_codes = {
        cd for _tid, _tnm, _ico, _nq, members in KR_THEME_BASKETS
        for cd, _nm, _desc in members
    }
    all_codes = sorted(set(uni.keys()) | theme_codes)
    print(f"  - 지수 {len(indices)}개 · 실시간 폴링 대상 {len(all_codes)}종목 (유니버스 {len(uni)} + 테마 보강)")

    print("[2/4] 네이버 실시간 배치 폴링 (KRX 정규장 + NXT 프리/애프터마켓) 실행 중...")
    rt_map = fetch_realtime_batch(all_codes)
    nxt_stocks = [v for v in rt_map.values() if v.get("hasNxt") and (v.get("nxtValEok") or 0) > 0]
    print(f"  - 실시간 시세 수신: {len(rt_map)}/{len(all_codes)}종목 · NXT 거래 종목: {len(nxt_stocks)}종목")

    # 현재 NXT 세션 판별 (PRE_MARKET / AFTER_MARKET / CLOSE)
    sam = rt_map.get("005930") or (nxt_stocks[0] if nxt_stocks else {})
    nxt_session = sam.get("nxtSession") or ""
    nxt_status = sam.get("nxtStatus") or ""
    nxt_time = sam.get("nxtTime") or ""
    krx_status = sam.get("krxStatus") or ""
    krx_time = sam.get("krxTime") or ""

    if nxt_session == "PRE_MARKET":
        session_label = "🌅 NXT 프리마켓 실시간 (08:00~08:50)" if nxt_status == "OPEN" else "🌅 NXT 프리마켓 마감"
        default_mode = "NXT"
    elif krx_status == "OPEN":
        session_label = "🏛️ KRX 정규장 진행 중 (09:00~15:30)"
        default_mode = "KRX"
    elif nxt_session == "AFTER_MARKET":
        session_label = "🌙 NXT 애프터마켓 진행 중 (15:30~20:00)" if nxt_status == "OPEN" else "🌙 NXT 애프터마켓 마감"
        default_mode = "NXT"
    else:
        session_label = "🏛️ 국내 정규장 마감"
        default_mode = "KRX"

    # 코드 -> 소속 테마명 매핑 (NXT TOP 리스트에 테마 태그 표시용)
    code_to_theme = {}
    for _tid, tname, _ico, _nq, members in KR_THEME_BASKETS:
        if _tid == "kr_top10":
            continue
        for cd, _nm, _desc in members:
            code_to_theme.setdefault(cd, tname)

    # NXT 거래대금 TOP 25 & NXT 상승률 급등 TOP 25 (거래대금 5억 이상 & 상승 종목 우선)
    def _nxt_row(v):
        return {
            "code": v["code"],
            "name": v["name"],
            "market": v["market"],
            "theme": code_to_theme.get(v["code"], ""),
            "nxtPrice": v["nxtPrice"],
            "nxtChg": v["nxtChg"],
            "nxtValEok": v["nxtValEok"],
            "nxtValTxt": v["nxtValTxt"],
            "nxtVol": v["nxtVol"],
            "krxPrice": v["krxPrice"],
            "krxChg": v["krxChg"],
            "krxValEok": v["krxValEok"],
            "mcapEok": (uni.get(v["code"]) or {}).get("mcapEok"),
        }

    nxt_by_val = sorted(nxt_stocks, key=lambda x: x.get("nxtValEok") or 0, reverse=True)
    nxt_top_val = [_nxt_row(x) for x in nxt_by_val[:25]]

    # 프리마켓(08:30)에는 거래대금 1억↑, 애프터/마감에는 10억↑ 기준 적용
    min_val_eok = 1.0 if nxt_session == "PRE_MARKET" else 10.0
    nxt_surge_pool = [
        x for x in nxt_stocks
        if (x.get("nxtValEok") or 0) >= min_val_eok and x.get("nxtChg") is not None
    ]
    if len(nxt_surge_pool) < 15:
        nxt_surge_pool = [x for x in nxt_stocks if x.get("nxtChg") is not None]
    nxt_surge_pool.sort(key=lambda x: x.get("nxtChg") or -999, reverse=True)
    nxt_top_up = [_nxt_row(x) for x in nxt_surge_pool[:25]]

    # KRX 정규장 거래대금 TOP 20도 함께 저장
    krx_by_val = sorted(
        [v for v in rt_map.values() if (v.get("krxValEok") or 0) > 0],
        key=lambda x: x.get("krxValEok") or 0,
        reverse=True,
    )
    krx_top_val = [_nxt_row(x) for x in krx_by_val[:20]]

    print("[3/4] 국내 45개 테마 한국어 뉴스 & 시장 핵심 뉴스 병렬 수집 중...")
    theme_news_map = {}
    market_news = []
    with ThreadPoolExecutor(max_workers=14) as ex:
        fut_mkt = ex.submit(fetch_google_kr_news, "코스피 코스닥 특징주 OR 넥스트레이드 프리마켓", 10, 2)
        futs_tn = {
            ex.submit(fetch_google_kr_news, nq, 3, 5): tid
            for tid, _tnm, _ico, nq, _mems in KR_THEME_BASKETS
        }
        for fut in as_completed(futs_tn):
            tid = futs_tn[fut]
            try:
                theme_news_map[tid] = fut.result() or []
            except Exception:
                theme_news_map[tid] = []
        try:
            market_news = fut_mkt.result() or []
        except Exception:
            market_news = []
    print(f"  - 테마 뉴스 수집: {sum(1 for v in theme_news_map.values() if v)}/{len(KR_THEME_BASKETS)}개 · 시장 뉴스: {len(market_news)}건")

    # 45개 국내 섹터 강도 보드 집계 (NXT 기준 + KRX 정규장 기준 + 테마별 NXT 합산 거래대금)
    sector_strength = []
    for tid, tname, ticon, _nq, members in KR_THEME_BASKETS:
        rows = []
        for cd, nm, desc in members:
            v = rt_map.get(cd) or {}
            krx_c = v.get("krxChg")
            nxt_c = v.get("nxtChg") if v.get("hasNxt") else None
            eff_nxt_c = nxt_c if nxt_c is not None else krx_c
            rows.append({
                "code": cd,
                "name": v.get("name") or nm,
                "market": v.get("market") or "KOSPI",
                "desc": desc,
                "krxPrice": v.get("krxPrice"),
                "krxChg": krx_c,
                "krxValEok": v.get("krxValEok"),
                "hasNxt": bool(v.get("hasNxt")),
                "nxtPrice": v.get("nxtPrice"),
                "nxtChg": nxt_c,
                "effNxtChg": eff_nxt_c,
                "nxtValEok": v.get("nxtValEok") or 0.0,
            })
        krx_vals = [r["krxChg"] for r in rows if r["krxChg"] is not None]
        nxt_vals = [r["effNxtChg"] for r in rows if r["effNxtChg"] is not None]
        nxt_only_vals = [r["nxtChg"] for r in rows if r["nxtChg"] is not None]
        avg_krx = round(sum(krx_vals) / len(krx_vals), 2) if krx_vals else 0.0
        avg_nxt = round(sum(nxt_vals) / len(nxt_vals), 2) if nxt_vals else avg_krx
        tot_nxt_val_eok = round(sum(r["nxtValEok"] or 0.0 for r in rows), 1)
        tot_krx_val_eok = round(sum(r["krxValEok"] or 0.0 for r in rows), 1)

        rows.sort(key=lambda x: (x["effNxtChg"] if default_mode == "NXT" else x["krxChg"]) or -999, reverse=True)
        sector_strength.append({
            "id": tid,
            "name": tname,
            "icon": ticon,
            "avgKrxChg": avg_krx,
            "avgNxtChg": avg_nxt,
            "nxtCovered": len(nxt_only_vals),
            "count": len(rows),
            "upKrxCount": sum(1 for v in krx_vals if v > 0),
            "dnKrxCount": sum(1 for v in krx_vals if v < 0),
            "upNxtCount": sum(1 for v in nxt_vals if v > 0),
            "dnNxtCount": sum(1 for v in nxt_vals if v < 0),
            "nxtValEok": tot_nxt_val_eok,
            "krxValEok": tot_krx_val_eok,
            "news": theme_news_map.get(tid) or [],
            "stocks": rows,
        })

    # NXT 거래대금이 가장 많이 몰린 테마 TOP 8 (코스피 TOP10 제외)
    nxt_theme_flow = sorted(
        [s for s in sector_strength if s["id"] != "kr_top10" and s["nxtValEok"] > 0],
        key=lambda x: x["nxtValEok"],
        reverse=True,
    )[:8]

    # 자동 한국어 4줄 브리핑 생성
    idx_map = {x["code"]: x for x in indices}
    kp = idx_map.get("KOSPI") or {}
    kq = idx_map.get("KOSDAQ") or {}
    fx = idx_map.get("USDKRW") or {}

    def _fp(v):
        return f"{v:+.2f}%" if v is not None else "-"

    bullets = []
    bullets.append({
        "tag": "국내 지수",
        "text": (
            f"코스피 {kp.get('price', 0):,.2f}({_fp(kp.get('chgPct'))}) · "
            f"코스닥 {kq.get('price', 0):,.2f}({_fp(kq.get('chgPct'))})"
            + (f" · 원/달러 환율 {fx['price']:,.1f}원({_fp(fx.get('chgPct'))})" if fx.get("price") else "")
            + f" 기준입니다. ({session_label})"
        ),
    })
    if nxt_top_val:
        top5_txt = ", ".join(
            f"{x['name']}({_fp(x['nxtChg'])}, {int(round(x['nxtValEok'])):,}억)"
            for x in nxt_top_val[:5]
        )
        bullets.append({
            "tag": "NXT 자금집중",
            "text": f"NXT 거래대금 상위 핵심 종목은 {top5_txt} 순으로 매수·거래가 집중되었습니다.",
        })
    if nxt_top_up:
        up5_txt = ", ".join(
            f"{x['name']}({_fp(x['nxtChg'])}" + (f" · {x['theme']}" if x.get("theme") else "") + ")"
            for x in nxt_top_up[:5]
        )
        bullets.append({
            "tag": "NXT 급등주",
            "text": f"NXT 강세 급등 상위 종목은 {up5_txt} 등입니다.",
        })
    sort_fld = "avgNxtChg" if default_mode == "NXT" else "avgKrxChg"
    sorted_ss = sorted(sector_strength, key=lambda x: x[sort_fld], reverse=True)
    top_ss = ", ".join(f"{x['name']}({_fp(x[sort_fld])})" for x in sorted_ss[:4])
    bot_ss = ", ".join(f"{x['name']}({_fp(x[sort_fld])})" for x in sorted_ss[-2:])
    bullets.append({
        "tag": "주도 섹터",
        "text": f"강세 상위 섹터는 {top_ss} / 상대 약세는 {bot_ss} 흐름입니다.",
    })

    total_nxt_val_eok = round(sum(x.get("nxtValEok") or 0 for x in nxt_stocks), 1)
    trade_date = (nxt_time or krx_time or datetime.now(KST).strftime("%Y-%m-%d"))[:10]

    payload = {
        "updatedAt": datetime.now(KST).strftime("%Y-%m-%d %H:%M"),
        "tradeDate": trade_date,
        "sessionLabel": session_label,
        "defaultMode": default_mode,
        "nxtSession": nxt_session,
        "nxtStatus": nxt_status,
        "nxtTime": nxt_time,
        "krxStatus": krx_status,
        "krxTime": krx_time,
        "universeCount": len(rt_map),
        "nxtActiveCount": len(nxt_stocks),
        "totalNxtValEok": total_nxt_val_eok,
        "indices": indices,
        "bullets": bullets,
        "nxtTopVal": nxt_top_val,
        "nxtTopUp": nxt_top_up,
        "krxTopVal": krx_top_val,
        "nxtThemeFlow": [
            {
                "id": x["id"],
                "name": x["name"],
                "icon": x["icon"],
                "nxtValEok": x["nxtValEok"],
                "avgNxtChg": x["avgNxtChg"],
                "leaders": [
                    {"name": s["name"], "code": s["code"], "nxtChg": s["effNxtChg"], "nxtValEok": s["nxtValEok"]}
                    for s in sorted(x["stocks"], key=lambda r: r.get("nxtValEok") or 0, reverse=True)[:3]
                ],
            }
            for x in nxt_theme_flow
        ],
        "sectorStrength": sorted(sector_strength, key=lambda x: abs(x[sort_fld]), reverse=True),
        "news": market_news,
    }

    with open(os.path.join(here, "kr_market.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    print(
        f"[4/4] 저장 완료: kr_market.json (기준일 {trade_date} · {session_label} · "
        f"NXT {len(nxt_stocks)}종목 / 총 {total_nxt_val_eok:,.0f}억 · 국내 {len(sector_strength)}테마)"
    )
    print("  - NXT 거래대금 TOP 5:", [(x["name"], f"{x['nxtChg']:+.2f}%", f"{x['nxtValEok']:,.0f}억") for x in nxt_top_val[:5]])
    return payload


def main():
    skip_push = any(a.lower() == "--no-push" for a in sys.argv[1:])
    collect_kr_market(HERE)

    sys.path.insert(0, HERE)
    import kospi200_yoy as screener

    screener.build_web_dashboard(HERE)
    if not skip_push:
        screener.auto_push_to_github(
            HERE,
            commit_msg=f"Auto-update KR market & NXT ({datetime.now(KST).strftime('%Y-%m-%d %H:%M')})",
        )
        try:
            import telegram_notify
            telegram_notify.notify_update("kr", HERE)
        except Exception:
            pass


if __name__ == "__main__":
    main()
