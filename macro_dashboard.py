#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
macro_dashboard.py — 재무부-연준 개입 체제 변곡점 대시보드
============================================================
FRED(무료, API 키 불필요) + TreasuryDirect 공개 API에서 지표를 수집해
신호등(green/yellow/red) HTML 대시보드를 생성한다.

사용법:
    python macro_dashboard.py                 # ./output/index.html 생성
    python macro_dashboard.py --out docs      # GitHub Pages용 docs/ 출력
    python macro_dashboard.py --json          # 신호를 JSON으로도 저장

의존성: requests (pip install requests)
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import datetime as dt
from dataclasses import dataclass, field, asdict

import requests

# ------------------------------------------------------------------
# 데이터 소스 정의
# ------------------------------------------------------------------

# fred.stlouisfed.org(그래프 CSV)는 GitHub Actions 등 데이터센터 IP를 차단한다
# (연결 즉시 거부, http=000). api.stlouisfed.org는 도달되므로 키가 있으면 그쪽을 쓴다.
# 키가 없어도 돌아가도록, FRED가 재가공하는 원천(재무부·연준·뉴욕연준)을 1차 소스로 쓴다.
# OAS(ICE BofA)만은 원천이 유료라 FRED 외 대안이 없다 → 키 없으면 결측 카드.
FRED_CSV = "https://fred.stlouisfed.org/graph/fredgraph.csv?id={sid}&cosd={start}"
FRED_API = ("https://api.stlouisfed.org/fred/series/observations"
            "?series_id={sid}&api_key={key}&file_type=json&observation_start={start}")

SERIES = {
    # H.4.1 계열 (주간, 수요일 기준)
    "FIMA":      dict(sid="H41RESPPALGTRFNWW", name="FIMA 레포 사용액",        unit="$B",  scale=1e-3),   # millions -> $B
    "FRPOOL":    dict(sid="WLRRAFOIAL",        name="외국공적 역레포 풀",       unit="$B",  scale=1e-3),
    "SWAP":      dict(sid="SWPT",              name="중앙은행 통화스왑",        unit="$B",  scale=1e-3),
    "WALCL":     dict(sid="WALCL",             name="연준 총자산",             unit="$T",  scale=1e-6),
    # 일간 계열
    "DGS10":     dict(sid="DGS10",             name="미 10년물",              unit="%",   scale=1.0),
    "DGS30":     dict(sid="DGS30",             name="미 30년물",              unit="%",   scale=1.0),
    "T5YIFR":    dict(sid="T5YIFR",            name="5y5y 포워드 브레이크이븐", unit="%",   scale=1.0),
    "ACMTP10":   dict(sid="THREEFYTP10",       name="ACM 10년 텀프리미엄",     unit="%",   scale=1.0),  # FRED엔 ACM이 없어 Kim-Wright로 폴백
    "USDJPY":    dict(sid="DEXJPUS",           name="엔/달러",                unit="",    scale=1.0),
    "IG_OAS":    dict(sid="BAMLC0A0CM",        name="IG 전체 OAS",            unit="bp",  scale=100.0),
    "AA_OAS":    dict(sid="BAMLC0A2CAA",       name="AA OAS (하이퍼스케일러 프록시)", unit="bp", scale=100.0),
}

# --- 키 없는 원천 소스 (FRED와 동일 데이터) ---
# 재무부 일별 파 수익률 곡선(명목/실질). 연도별 CSV, 최신일이 위.
TREASURY_CSV = ("https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
                "daily-treasury-rates.csv/{year}/all?type={kind}&field_tdr_date_value={year}&page&_format=csv")
# 연준 H.4.1 전체 패키지(SDMX XML zip, ~30MB). 주간 수요일 레벨 계열의 DDP 식별자.
FED_H41_ZIP = "https://www.federalreserve.gov/datadownload/Output.aspx?rel=H41&filetype=zip"
H41_SERIES = {
    "FIMA":   "RESPPALGTRF_N.WW",   # Assets: Other: Repurchase agreements - Foreign official
    "SWAP":   "RESH4SCS_N.WW",      # Central bank liquidity swaps
    "WALCL":  "RESPPMA_N.WW",       # Total assets (less eliminations)
    "FRPOOL": "RESPPLLRF_N.WW",     # Reverse repos: Foreign official and international accounts
}
# 뉴욕연준 ACM 텀프리미엄(xls, 'ACM Daily' 시트, ACMTP10 열). xlrd 필요.
NYFED_ACM_XLS = "https://www.newyorkfed.org/medialibrary/media/research/data_indicators/ACMTermPremium.xls"

TD_AUCTION = ("https://www.treasurydirect.gov/TA_WS/securities/search"
              "?type={typ}&days=60&format=json")

# FRED DEXJPUS는 4영업일 지연 발표라 개입 존 감시에 못 쓴다. 야후를 1차 소스로,
# 실패 시 FRED로 폴백한다.
YF_CHART = "https://query1.finance.yahoo.com/v8/finance/chart/{sym}?range={rng}&interval=1d"

# 재발행(reopening)은 잔존만기로 표기된다: 10Y 재발행 = "9-Year 10-Month" 등.
# 신규발행만 잡으면 월 3회 중 1회만 걸리므로 잔존만기 표기를 함께 인식한다.
TENOR_PATTERNS = {
    "10Y": re.compile(r"^(10-year$|9-year\s+1[01]-month)"),
    "20Y": re.compile(r"^(20-year$|19-year\s+1[01]-month)"),
    "30Y": re.compile(r"^(30-year$|29-year\s+1[01]-month)"),
}

# 만기별 응찰배수(BTC) 임계치 — 텐서마다 정상 수준이 다르다.
BTC_BANDS = {"10Y": (2.50, 2.35), "20Y": (2.55, 2.40), "30Y": (2.40, 2.25)}

# ------------------------------------------------------------------
# 수집 유틸
# ------------------------------------------------------------------

def http_get(url, timeout=(10, 45), retries=2):
    """(connect, read) 타임아웃 + 재시도. CI 러너에서 간헐적으로 느려지는 소스 대비."""
    last = None
    for attempt in range(retries + 1):
        try:
            r = requests.get(url, timeout=timeout,
                             headers={"User-Agent": "macro-dashboard/1.0"})
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            last = e
            if attempt < retries:
                print(f"    재시도 {attempt + 1}/{retries}: {type(e).__name__}", flush=True)
    raise last


_FRED_STATE = {}

def fetch_fred(sid, lookback_days=420):
    """FRED 관측치 -> [(date, float), ...] 오름차순. 결측('.')은 제외.

    FRED_API_KEY가 있으면 공식 API를, 없으면 그래프 CSV를 쓴다.
    CI(데이터센터 IP)에서는 CSV 쪽이 차단되므로 키가 사실상 필수.
    """
    start = (dt.date.today() - dt.timedelta(days=lookback_days)).isoformat()
    key = os.environ.get("FRED_API_KEY", "").strip()
    if _FRED_STATE.get("dead"):
        raise RuntimeError("FRED 차단(이번 실행에서 이미 실패) — 재시도 생략")
    if key:
        obs = http_get(FRED_API.format(sid=sid, key=key, start=start)).json()
        rows = []
        for o in obs.get("observations", []):
            if o.get("value") in (".", "", "NA", None):
                continue
            try:
                rows.append((dt.date.fromisoformat(o["date"]), float(o["value"])))
            except (ValueError, KeyError):
                continue
        rows.sort()
        return rows
    # 키 없는 CSV 경로: 데이터센터 IP는 차단(무응답)되므로 짧게 한 번만 시도하고,
    # 실패하면 이번 실행의 나머지 FRED 호출은 모두 생략한다.
    try:
        txt = http_get(FRED_CSV.format(sid=sid, start=start), timeout=(10, 20), retries=0).text
    except requests.RequestException:
        _FRED_STATE["dead"] = True
        raise
    rows = []
    rdr = csv.reader(io.StringIO(txt))
    header = next(rdr, None)
    for row in rdr:
        if len(row) < 2 or row[1] in (".", "", "NA"):
            continue
        try:
            rows.append((dt.date.fromisoformat(row[0]), float(row[1])))
        except ValueError:
            continue
    return rows


# ------------------------------------------------------------------
# 키 없는 원천 소스 fetcher
# ------------------------------------------------------------------

def _treasury_curve(kind, years):
    """재무부 파 수익률 CSV -> {date: {col_lower: float}}. 명목(daily_treasury_yield_curve)
    또는 실질(daily_treasury_real_yield_curve). 연초 데이터 부족 대비 전년도까지 합친다."""
    out = {}
    for y in years:
        txt = http_get(TREASURY_CSV.format(year=y, kind=kind)).text
        rdr = csv.reader(io.StringIO(txt))
        header = [h.strip().lower() for h in (next(rdr, None) or [])]
        for row in rdr:
            if not row or len(row) != len(header):
                continue
            try:
                d = dt.datetime.strptime(row[0].strip(), "%m/%d/%Y").date()
            except ValueError:
                continue
            vals = {}
            for h, v in zip(header[1:], row[1:]):
                try:
                    vals[h] = float(v)
                except ValueError:
                    pass
            if vals:
                out[d] = vals
    return out


_TREASURY_CACHE = {}

def treasury_curves():
    if not _TREASURY_CACHE:
        yr = dt.date.today().year
        _TREASURY_CACHE["nom"]  = _treasury_curve("daily_treasury_yield_curve", (yr - 1, yr))
        _TREASURY_CACHE["real"] = _treasury_curve("daily_treasury_real_yield_curve", (yr - 1, yr))
    return _TREASURY_CACHE


def fetch_treasury_yield(col):
    """명목 곡선 한 컬럼('10 yr', '30 yr') -> [(date, pct)] 오름차순."""
    nom = treasury_curves()["nom"]
    rows = [(d, v[col]) for d, v in nom.items() if col in v]
    rows.sort()
    if not rows:
        raise ValueError(f"재무부 곡선에 '{col}' 컬럼 없음")
    return rows


def fetch_t5yifr():
    """5y5y 포워드 브레이크이븐 — FRED T5YIFR과 같은 공식으로 재무부 명목·실질 곡선에서 계산.
    T5YIFR = (((1+BE10)^10 / (1+BE5)^5)^(1/5) - 1), BEn = DGSn - DFIIn."""
    cv = treasury_curves()
    rows = []
    for d, n in cv["nom"].items():
        r = cv["real"].get(d)
        if not r or "5 yr" not in n or "10 yr" not in n or "5 yr" not in r or "10 yr" not in r:
            continue
        be5, be10 = (n["5 yr"] - r["5 yr"]) / 100, (n["10 yr"] - r["10 yr"]) / 100
        fwd = ((1 + be10) ** 10 / (1 + be5) ** 5) ** (1 / 5) - 1
        rows.append((d, fwd * 100))
    rows.sort()
    if not rows:
        raise ValueError("명목·실질 곡선 교집합 없음")
    return rows


_H41_CACHE = {}

def fetch_h41(key, lookback_days=420):
    """연준 H.4.1 데이터패키지(zip/SDMX XML)에서 주간 계열 -> [(date, millions)]."""
    import zipfile
    if "xml" not in _H41_CACHE:
        r = http_get(FED_H41_ZIP, timeout=(10, 120))
        z = zipfile.ZipFile(io.BytesIO(r.content))
        name = next(n for n in z.namelist() if n.lower().endswith("_data.xml"))
        _H41_CACHE["xml"] = z.read(name).decode("utf-8", "replace")
    xml = _H41_CACHE["xml"]
    sid = H41_SERIES[key]
    i = xml.find(f'SERIES_NAME="{sid}"')
    if i < 0:
        raise ValueError(f"H.4.1 패키지에 {sid} 없음")
    j = xml.find("</kf:Series>", i)
    chunk = xml[i:j]
    cutoff = dt.date.today() - dt.timedelta(days=lookback_days)
    rows = []
    for val, per in re.findall(r'OBS_VALUE="([^"]*)"[^>]*TIME_PERIOD="(\d{4}-\d{2}-\d{2})"', chunk):
        try:
            d = dt.date.fromisoformat(per)
            if d >= cutoff:
                rows.append((d, float(val)))
        except ValueError:
            continue
    rows.sort()
    if not rows:
        raise ValueError(f"{sid} 관측치 없음")
    return rows


def fetch_acm_tp10(lookback_days=420):
    """뉴욕연준 ACM 일별 10년 텀프리미엄 -> [(date, pct)]."""
    import xlrd  # pip install xlrd (xls 구형식)
    r = http_get(NYFED_ACM_XLS, timeout=(10, 90))
    wb = xlrd.open_workbook(file_contents=r.content)
    sh = wb.sheet_by_name("ACM Daily")
    header = [str(c).strip() for c in sh.row_values(0)]
    ci = header.index("ACMTP10")
    cutoff = dt.date.today() - dt.timedelta(days=lookback_days)
    rows = []
    for ri in range(1, sh.nrows):
        rv = sh.row_values(ri)
        try:
            d = dt.datetime.strptime(str(rv[0]).strip(), "%d-%b-%Y").date()
            v = float(rv[ci])
        except (ValueError, TypeError):
            continue
        if d >= cutoff:
            rows.append((d, v))
    rows.sort()
    if not rows:
        raise ValueError("ACM 시트 관측치 없음")
    return rows


# 계열별 1차(원천) 소스. 없으면 FRED만 사용.
ALT_SOURCES = {
    "FIMA":    lambda: fetch_h41("FIMA"),
    "SWAP":    lambda: fetch_h41("SWAP"),
    "WALCL":   lambda: fetch_h41("WALCL"),
    "FRPOOL":  lambda: fetch_h41("FRPOOL"),
    "DGS10":   lambda: fetch_treasury_yield("10 yr"),
    "DGS30":   lambda: fetch_treasury_yield("30 yr"),
    "T5YIFR":  fetch_t5yifr,
    "ACMTP10": fetch_acm_tp10,
}
ALT_LABEL = {"FIMA": "H.4.1", "SWAP": "H.4.1", "WALCL": "H.4.1", "FRPOOL": "H.4.1",
             "DGS10": "재무부", "DGS30": "재무부", "T5YIFR": "재무부(계산)", "ACMTP10": "뉴욕연준"}


def fetch_series(key):
    """원천 소스 우선, 실패 시 FRED. 반환 (rows, source_label)."""
    meta = SERIES[key]
    errs = []
    if key in ALT_SOURCES:
        try:
            return ALT_SOURCES[key](), ALT_LABEL[key]
        except Exception as e:
            errs.append(f"{ALT_LABEL[key]}: {type(e).__name__} {e}")
            print(f"    {key} 원천 실패({errs[-1]}) -> FRED 폴백", flush=True)
    try:
        return fetch_fred(meta["sid"]), "FRED"
    except Exception as e:
        errs.append(f"FRED: {type(e).__name__} {e}")
    raise RuntimeError(" / ".join(errs))


def utc_date(ts):
    return dt.datetime.fromtimestamp(ts, dt.timezone.utc).date()


def fetch_yahoo_fx(sym="JPY=X", rng="6mo"):
    """야후 일별 종가 -> [(date, float), ...] 오름차순. 마지막 관측은 실시간 최종가로 갱신."""
    res = http_get(YF_CHART.format(sym=sym, rng=rng)).json()["chart"]["result"][0]
    ts = res.get("timestamp") or []
    close = (res.get("indicators", {}).get("quote") or [{}])[0].get("close") or []
    rows = [(utc_date(t), float(c)) for t, c in zip(ts, close) if c is not None]
    meta = res.get("meta") or {}
    px, mt = meta.get("regularMarketPrice"), meta.get("regularMarketTime")
    if px and mt:  # 장중 최종가로 당일 값을 덮어쓴다
        d = utc_date(mt)
        rows = [r for r in rows if r[0] < d] + [(d, float(px))]
    rows.sort()
    if not rows:
        raise ValueError("야후 응답에 관측치 없음")
    return rows


def classify_tenor(term):
    t = (term or "").strip().lower()
    for tag, pat in TENOR_PATTERNS.items():
        if pat.match(t):
            return tag
    return None


def fetch_auctions():
    """최근 10Y Note / 20Y·30Y Bond 낙찰 결과(재발행 포함)."""
    out = []
    for typ in ("Note", "Bond"):
        try:
            data = http_get(TD_AUCTION.format(typ=typ)).json()
        except Exception:
            continue
        for a in data:
            tenor = classify_tenor(a.get("securityTerm"))
            if tenor is None:
                continue
            if typ == "Note" and tenor != "10Y":
                continue
            try:
                btc = float(a.get("bidToCoverRatio") or 0)
            except ValueError:
                btc = 0.0
            if btc <= 0:
                continue
            dealer_pct = None
            try:
                pd_amt  = float(a.get("primaryDealerAccepted") or 0)
                cmp_amt = float(a.get("competitiveAccepted") or 0)
                if cmp_amt > 0:
                    dealer_pct = 100.0 * pd_amt / cmp_amt
            except (ValueError, TypeError):
                pass
            out.append(dict(
                date=(a.get("auctionDate") or "")[:10],
                term=a.get("securityTerm", ""),
                tenor=tenor,
                reopen=not (a.get("securityTerm") or "").strip().lower().endswith("-year"),
                high_yield=a.get("highYield") or a.get("highDiscountRate") or "",
                btc=btc, dealer_pct=dealer_pct,
            ))
    out.sort(key=lambda x: x["date"], reverse=True)
    return out[:8]

# ------------------------------------------------------------------
# 신호 계산
# ------------------------------------------------------------------

GREEN, YELLOW, RED, NA = "green", "yellow", "red", "na"
RANK = {GREEN: 0, YELLOW: 1, RED: 2, NA: -1}


@dataclass
class Signal:
    key: str
    name: str
    status: str          # green / yellow / red / na
    value: str           # 표시용 최신값
    asof: str            # 기준일
    delta: str           # 추세 문자열
    note: str            # 판독 코멘트
    group: str           # 카드 그룹
    history: list = field(default_factory=list)  # 스파크라인용 최근값


def latest(series):
    return series[-1] if series else None


def value_at_offset(series, days_back):
    """days_back일 이전 시점과 가장 가까운 관측치."""
    if not series:
        return None
    target = series[-1][0] - dt.timedelta(days=days_back)
    best = None
    for d, v in series:
        if d <= target:
            best = (d, v)
        else:
            break
    return best


def fmt_delta(cur, prev, unit, digits=1):
    if prev is None:
        return "—"
    d = cur - prev
    arrow = "▲" if d > 0 else ("▼" if d < 0 else "＝")
    return f"{arrow} {abs(d):.{digits}f}{unit}/4주"


def spark(series, n=26):
    return [v for _, v in series[-n:]]


def build_signals(data, auctions):
    S = []

    def get(key):
        raw = data.get(key) or []
        meta = SERIES[key]
        return [(d, v * meta["scale"]) for d, v in raw]

    # -------- 1. FIMA 레포 --------
    fima = get("FIMA")
    swap = get("SWAP")
    if fima:
        d, v = latest(fima)
        swap_v = latest(swap)[1] if swap else 0.0
        if v <= 0.05:
            st, note = GREEN, "평시(0). 체제는 대기 상태."
        elif v <= 10:
            st, note = YELLOW, "가동 시작 — 엔화 개입 뉴스와 시점 대조 필요(시나리오 A 추정)."
        elif v <= 45:
            st, note = YELLOW, "규모 확대 — 지속성(눌어붙는지)과 스왑라인 동반 여부 확인."
        else:
            st, note = RED, "한도(600억$) 근접 — 증액 안건이 FOMC 이벤트로 격상."
        if v > 0.05 and swap_v > 5:
            st, note = RED, "FIMA+스왑라인 동반 상승 = 스트레스성(시나리오 B) 의심."
        prev = value_at_offset(fima, 28)
        S.append(Signal("fima", "FIMA 레포 사용액", st, f"${v:,.1f}B", d.isoformat(),
                        fmt_delta(v, prev[1] if prev else None, "B"), note,
                        "개입 체제", spark(fima)))
    else:
        S.append(Signal("fima", "FIMA 레포 사용액", NA, "—", "", "—",
                        "수집 실패 — H.4.1 원문 확인", "개입 체제"))

    # -------- 1b. 중앙은행 통화스왑 (FIMA와 짝) --------
    if swap:
        d, v = latest(swap)
        prev = value_at_offset(swap, 28)
        if v < 1:
            st, note = GREEN, "평시 잔액(스탠딩 라인 소액 사용)."
        elif v < 20:
            st, note = YELLOW, "잔액 증가 — 달러 조달 스트레스 초기."
        else:
            st, note = RED, "스왑라인 본격 가동 = 시나리오 B(스트레스성)."
        S.append(Signal("swap", "중앙은행 통화스왑", st, f"${v:,.2f}B", d.isoformat(),
                        fmt_delta(v, prev[1] if prev else None, "B", 2), note,
                        "개입 체제", spark(swap)))

    # -------- 2. 엔/달러 --------
    jpy = get("USDJPY")
    if jpy:
        d, v = latest(jpy)
        if v < 155:
            st, note = GREEN, "개입 존 밖."
        elif v < 158:
            st, note = YELLOW, "155~158 — 구두개입·레이트체크 감시 구간."
        else:
            st, note = RED, "158 상회 — 2차 공동개입 트리거 존."
        src = data.get("_jpy_src", "FRED")
        lag = (dt.date.today() - d).days
        if lag >= 3:
            note += f" ⚠ {lag}일 지연 데이터({src}) — 실시간 확인 필요."
        else:
            note += f" ({src})"
        prev = value_at_offset(jpy, 28)
        S.append(Signal("jpy", "엔/달러", st, f"{v:,.2f}", d.isoformat(),
                        fmt_delta(v, prev[1] if prev else None, ""), note,
                        "개입 체제", spark(jpy)))

    # -------- 3. 30년물 (베센트 반응함수) --------
    y30 = get("DGS30")
    y10 = get("DGS10")
    if y30:
        d, v = latest(y30)
        if v < 5.0:
            st, note = GREEN, "박스 하단부 — 개입 압력 낮음."
        elif v < 5.2:
            st, note = YELLOW, "박스 중상단."
        else:
            st, note = RED, "5.2% 상회 — 베센트 반응함수 발동 존(서프라이즈 조치 확률↑)."
        prev = value_at_offset(y30, 28)
        S.append(Signal("y30", "미 30년물", st, f"{v:.3f}%", d.isoformat(),
                        fmt_delta(v * 100, prev[1] * 100 if prev else None, "bp", 0), note,
                        "장기금리 박스권", spark(y30)))
    if y10:
        d, v = latest(y10)
        prev = value_at_offset(y10, 28)
        S.append(Signal("y10", "미 10년물", GREEN if v < 4.7 else YELLOW,
                        f"{v:.3f}%", d.isoformat(),
                        fmt_delta(v * 100, prev[1] * 100 if prev else None, "bp", 0),
                        "베센트의 명시적 관리 대상.", "장기금리 박스권", spark(y10)))

    # -------- 4. 5y5y 브레이크이븐 --------
    be = get("T5YIFR")
    if be:
        d, v = latest(be)
        if v < 2.4:
            st, note = GREEN, "장기 기대 앵커 유지."
        elif v < 2.6:
            st, note = YELLOW, "앵커 이탈 조짐 — 물가지표일 30년물 반응과 교차확인."
        else:
            st, note = RED, "탈앵커 — 재정우위 프라이싱 경계."
        prev = value_at_offset(be, 28)
        S.append(Signal("be55", "5y5y 브레이크이븐", st, f"{v:.2f}%", d.isoformat(),
                        fmt_delta(v * 100, prev[1] * 100 if prev else None, "bp", 0), note,
                        "인플레·재정우위", spark(be)))

    # -------- 5. ACM 텀프리미엄 --------
    tp = get("ACMTP10")
    if tp:
        d, v = latest(tp)
        prev3m = value_at_offset(tp, 90)
        rise3m = (v - prev3m[1]) * 100 if prev3m else 0
        if v < 0.8 and rise3m < 25:
            st, note = GREEN, "프리미엄 안정."
        elif v < 1.1 and rise3m < 40:
            st, note = YELLOW, f"완만한 상승(3개월 {rise3m:+.0f}bp) — 상승분의 주범인지 분해 확인."
        else:
            st, note = RED, f"레벨/속도 경보(3개월 {rise3m:+.0f}bp) — 금리상승이 기대가 아닌 프리미엄 주도."
        prev = value_at_offset(tp, 28)
        src = (data.get("_sources") or {}).get("ACMTP10", "")
        if src == "FRED":
            note += " ⚠ FRED THREEFYTP10(Kim-Wright) 폴백 — ACM보다 ~20bp 높게 나오는 모형."
        elif src:
            note += f" ({src})"
        S.append(Signal("tp10", "ACM 10년 텀프리미엄", st, f"{v:.2f}%", d.isoformat(),
                        fmt_delta(v * 100, prev[1] * 100 if prev else None, "bp", 0), note,
                        "인플레·재정우위", spark(tp)))

    # -------- 6. AA vs IG OAS (하이퍼스케일러 크레딧 프록시) --------
    aa, ig = get("AA_OAS"), get("IG_OAS")
    if aa and ig:
        d, va = latest(aa)
        vi = latest(ig)[1]
        pa, pi = value_at_offset(aa, 28), value_at_offset(ig, 28)
        if pa and pi:
            rel = (va - pa[1]) - (vi - pi[1])  # AA 확대분 - IG 확대분 (bp)
            if rel < 3:
                st, note = GREEN, "AA가 시장 대비 상대 확대 없음."
            elif rel < 8:
                st, note = YELLOW, f"AA 상대 확대 +{rel:.0f}bp/4주 — 발행 캘린더와 대조(소화불량 vs 매도)."
            else:
                st, note = RED, f"AA 상대 확대 +{rel:.0f}bp/4주 — 신규발행 없는 주간이면 크레딧 균열 초기신호."
            delta_str = f"AA {fmt_delta(va, pa[1], 'bp', 0)}"
        else:
            st, note, delta_str = NA, "추세 계산 불가", "—"
        S.append(Signal("credit", "AA OAS vs IG OAS", st,
                        f"AA {va:.0f}bp · IG {vi:.0f}bp", d.isoformat(),
                        delta_str, note, "AI 크레딧", spark(aa)))
    else:
        S.append(Signal("credit", "AA OAS vs IG OAS", NA, "—", "", "—",
                        "ICE BofA OAS는 FRED 전용 계열 — FRED_API_KEY 시크릿 등록 시 활성.", "AI 크레딧"))

    # -------- 7. 장기물 입찰 --------
    if auctions:
        a = auctions[0]
        btc = a["btc"]
        hi, lo = BTC_BANDS.get(a["tenor"], (2.40, 2.20))
        if btc >= hi:
            st, note = GREEN, "수요 견조."
        elif btc >= lo:
            st, note = YELLOW, "응찰 둔화 — 다음 입찰 테일 주시."
        else:
            st, note = RED, "매수자 파업 재발 신호 — 딜러 인수비중 동반 상승 시 확정적."
        if a.get("dealer_pct") and a["dealer_pct"] >= 20 and st == YELLOW:
            note += " 딜러 인수비중 20%+ 동반."
        dealer = f" · 딜러 {a['dealer_pct']:.0f}%" if a.get("dealer_pct") else ""
        label = f"{a['tenor']}{' 재발행' if a['reopen'] else ''}"
        S.append(Signal("auction", f"최근 장기물 입찰 ({label})", st,
                        f"BTC {btc:.2f} · 낙찰 {a['high_yield']}%{dealer}",
                        a["date"], f"임계 {hi:.2f}/{lo:.2f}", note, "장기금리 박스권"))
    else:
        S.append(Signal("auction", "최근 장기물 입찰", NA, "—", "", "—",
                        "TreasuryDirect 수집 실패", "장기금리 박스권"))

    # -------- 8. 컨텍스트: 연준 총자산 / 역레포 풀 --------
    bs = get("WALCL")
    if bs:
        d, v = latest(bs)
        prev = value_at_offset(bs, 28)
        S.append(Signal("walcl", "연준 총자산", GREEN, f"${v:.2f}T", d.isoformat(),
                        fmt_delta(v * 1000, prev[1] * 1000 if prev else None, "B", 0),
                        "QT 종료/재개·FIMA 가동 시 여기서 확인.", "컨텍스트", spark(bs)))
    fp = get("FRPOOL")
    if fp:
        d, v = latest(fp)
        prev = value_at_offset(fp, 28)
        S.append(Signal("frpool", "외국공적 역레포 풀", GREEN, f"${v:.0f}B", d.isoformat(),
                        fmt_delta(v, prev[1] if prev else None, "B", 0),
                        "급감 = 외국 공적부문이 달러를 빼가는 중(FIMA의 선행 그림자).", "컨텍스트", spark(fp)))

    return S


def overall(signals):
    # 컨텍스트 카드는 상시 green 고정이라 종합판정에서 제외한다.
    live = [s for s in signals if s.status != NA and s.group != "컨텍스트"]
    reds = sum(1 for s in live if s.status == RED)
    yels = sum(1 for s in live if s.status == YELLOW)
    if reds >= 2:
        return RED, "복수 경보 — 변곡 국면 진입 가능성. 개별 신호의 조합(어느 시나리오인지)을 판독할 것."
    if reds == 1:
        return YELLOW, "단일 경보 — 해당 축의 동행지표를 교차확인."
    if yels >= 3:
        return YELLOW, "다수 주의 — 박스권 상단부 압력 축적 중."
    return GREEN, "박스권 레짐 정상 작동 중."

# ------------------------------------------------------------------
# HTML 렌더링
# ------------------------------------------------------------------

STATUS_KO = {GREEN: "정상", YELLOW: "주의", RED: "경보", NA: "결측"}

def svg_spark(vals, w=140, h=34):
    if not vals or len(vals) < 2:
        return ""
    lo, hi = min(vals), max(vals)
    rng = (hi - lo) or 1.0
    pts = []
    for i, v in enumerate(vals):
        x = i * w / (len(vals) - 1)
        y = h - 3 - (v - lo) / rng * (h - 6)
        pts.append(f"{x:.1f},{y:.1f}")
    return (f'<svg class="spark" viewBox="0 0 {w} {h}" preserveAspectRatio="none">'
            f'<polyline points="{" ".join(pts)}" />'
            f'<circle cx="{pts[-1].split(",")[0]}" cy="{pts[-1].split(",")[1]}" r="2.4"/></svg>')


def render_html(signals, auctions, gen_at):
    ov_st, ov_note = overall(signals)
    groups = []
    for g in ("개입 체제", "장기금리 박스권", "인플레·재정우위", "AI 크레딧", "컨텍스트"):
        items = [s for s in signals if s.group == g]
        if items:
            groups.append((g, items))

    def card(s):
        return f"""
      <div class="card st-{s.status}">
        <div class="card-top">
          <span class="lamp"></span>
          <span class="card-name">{s.name}</span>
          <span class="card-status">{STATUS_KO[s.status]}</span>
        </div>
        <div class="card-val">{s.value}</div>
        <div class="card-meta"><span>{s.asof}</span><span>{s.delta}</span></div>
        {svg_spark(s.history)}
        <div class="card-note">{s.note}</div>
      </div>"""

    group_html = ""
    for g, items in groups:
        group_html += f'\n    <h2 class="grp">{g}</h2>\n    <div class="grid">'
        group_html += "".join(card(s) for s in items)
        group_html += "</div>"

    auct_rows = ""
    for a in auctions:
        dealer = f"{a['dealer_pct']:.0f}%" if a.get("dealer_pct") else "—"
        auct_rows += (f"<tr><td>{a['date']}</td><td>{a['term']}</td>"
                      f"<td>{a['high_yield']}</td><td>{a['btc']:.2f}</td><td>{dealer}</td></tr>")
    auct_table = f"""
    <h2 class="grp">최근 장기물 입찰 이력</h2>
    <table class="auct"><thead><tr>
      <th>입찰일</th><th>만기</th><th>낙찰금리</th><th>응찰배수</th><th>딜러비중</th>
    </tr></thead><tbody>{auct_rows or '<tr><td colspan="5">데이터 없음</td></tr>'}</tbody></table>"""

    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>변곡점 대시보드 — 재무부·연준 개입 체제</title>
<style>
:root {{
  --navy-deep:#0f1b2a; --navy:#1C3D5A; --orange:#E5601F;
  --green:#2f9e6e; --yellow:#e0a921; --red:#d64541;
  --paper:#f4f6f9; --card:#ffffff; --ink:#152230; --mut:#6b7a8c;
  --mono:'SF Mono','JetBrains Mono',Consolas,monospace;
}}
* {{ box-sizing:border-box; margin:0; }}
body {{ background:var(--paper); color:var(--ink);
  font-family:'Apple SD Gothic Neo','Malgun Gothic',sans-serif; }}
.hero {{ background:linear-gradient(135deg,var(--navy-deep),var(--navy));
  color:#fff; padding:34px 28px 28px; }}
.hero .eyebrow {{ font-family:var(--mono); font-size:11px; letter-spacing:.18em;
  color:var(--orange); text-transform:uppercase; }}
.hero h1 {{ font-size:24px; margin:8px 0 4px; font-weight:800; }}
.hero .sub {{ color:#b9c6d4; font-size:13px; }}
.regime {{ display:flex; align-items:center; gap:12px; margin-top:18px;
  background:rgba(255,255,255,.06); border-left:4px solid; border-radius:6px;
  padding:12px 16px; font-size:14px; }}
.regime.st-green {{ border-color:var(--green); }}
.regime.st-yellow {{ border-color:var(--yellow); }}
.regime.st-red {{ border-color:var(--red); }}
.regime b {{ font-family:var(--mono); }}
.wrap {{ max-width:1080px; margin:0 auto; padding:8px 28px 48px; }}
.grp {{ font-size:13px; letter-spacing:.12em; color:var(--navy); margin:30px 0 12px;
  text-transform:uppercase; font-family:var(--mono); border-bottom:2px solid var(--navy);
  display:inline-block; padding-bottom:3px; }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(240px,1fr)); gap:14px; }}
.card {{ background:var(--card); border-radius:10px; padding:14px 16px;
  border-top:4px solid var(--mut); box-shadow:0 1px 4px rgba(15,27,42,.08); }}
.card.st-green  {{ border-top-color:var(--green); }}
.card.st-yellow {{ border-top-color:var(--yellow); }}
.card.st-red    {{ border-top-color:var(--red); }}
.card-top {{ display:flex; align-items:center; gap:8px; }}
.lamp {{ width:9px; height:9px; border-radius:50%; background:var(--mut); flex:none; }}
.st-green .lamp  {{ background:var(--green); }}
.st-yellow .lamp {{ background:var(--yellow); }}
.st-red .lamp    {{ background:var(--red); box-shadow:0 0 0 4px rgba(214,69,65,.15); }}
.card-name {{ font-size:13px; font-weight:700; flex:1; }}
.card-status {{ font-size:11px; font-family:var(--mono); color:var(--mut); }}
.st-red .card-status {{ color:var(--red); font-weight:700; }}
.card-val {{ font-family:var(--mono); font-size:21px; font-weight:700; margin:8px 0 2px; }}
.card-meta {{ display:flex; justify-content:space-between; font-size:11px;
  color:var(--mut); font-family:var(--mono); margin-bottom:6px; }}
.spark {{ width:100%; height:34px; }}
.spark polyline {{ fill:none; stroke:var(--navy); stroke-width:1.6; opacity:.75; }}
.spark circle {{ fill:var(--orange); }}
.card-note {{ font-size:12px; color:#3d4c5c; margin-top:6px; line-height:1.5;
  border-top:1px dashed #dde4ec; padding-top:6px; }}
.auct {{ width:100%; border-collapse:collapse; background:var(--card);
  border-radius:10px; overflow:hidden; font-size:13px;
  box-shadow:0 1px 4px rgba(15,27,42,.08); }}
.auct th {{ background:var(--navy); color:#fff; padding:9px 12px; text-align:left;
  font-size:12px; }}
.auct td {{ padding:8px 12px; border-top:1px solid #edf1f6; font-family:var(--mono); }}
.foot {{ margin-top:36px; font-size:11px; color:var(--mut); line-height:1.7; }}
@media (max-width:560px) {{ .hero{{padding:24px 18px;}} .wrap{{padding:8px 18px 40px;}} }}
</style>
</head>
<body>
<div class="hero">
  <div class="wrap" style="padding:0;">
    <div class="eyebrow">Treasury–Fed Regime Watch</div>
    <h1>변곡점 대시보드</h1>
    <div class="sub">생성 {gen_at} KST · 소스: FRED (H.4.1 계열 포함) · TreasuryDirect</div>
    <div class="regime st-{ov_st}"><b>[{STATUS_KO[ov_st]}]</b> {ov_note}</div>
  </div>
</div>
<div class="wrap">
  {group_html}
  {auct_table}
  <div class="foot">
    판독 규칙 요약 — FIMA: 0=평시, 가동 시 스왑라인 동반 여부로 개입성(A)/스트레스성(B) 구분 ·
    30년물 5.2%+입찰부진 = 베센트 반응함수 존 · 5y5y 2.6% 또는 텀프리미엄 급등 = 재정우위 프라이싱 ·
    AA의 IG 대비 상대 확대 = 하이퍼스케일러 크레딧 프록시(신규발행 유무와 교차확인).
    수치 확정치는 원 소스 기준이며 본 대시보드는 개인 리서치 보조용.
  </div>
</div>
</body>
</html>"""

# ------------------------------------------------------------------
# 텔레그램 (선택) — 환경변수가 있을 때만 동작
# ------------------------------------------------------------------

LAMP = {GREEN: "🟢", YELLOW: "🟡", RED: "🔴", NA: "⚪"}


def telegram_text(signals, gen_at, page_url=""):
    ov_st, ov_note = overall(signals)
    lines = [f"{LAMP[ov_st]} <b>변곡점 대시보드 — {STATUS_KO[ov_st]}</b>",
             f"<i>{ov_note}</i>", ""]
    alerts = [s for s in signals if s.status in (RED, YELLOW) and s.group != "컨텍스트"]
    for s in alerts:
        lines.append(f"{LAMP[s.status]} <b>{s.name}</b>  {s.value}")
        lines.append(f"   └ {s.note}")
    if not alerts:
        lines.append("경보·주의 없음. 박스권 정상.")
    lines += ["", f"<code>{gen_at} KST</code>"]
    if page_url:
        lines.append(page_url)
    return "\n".join(lines)


def send_telegram(text):
    token = os.environ.get("TELEGRAM_BOT_TOKEN")
    chat = os.environ.get("TELEGRAM_CHAT_ID")
    if not token or not chat:
        print("[--] 텔레그램 환경변수 없음 — 발송 생략")
        return False
    r = requests.post(f"https://api.telegram.org/bot{token}/sendMessage",
                      data={"chat_id": chat, "text": text,
                            "parse_mode": "HTML", "disable_web_page_preview": "true"},
                      timeout=20)
    ok = r.ok and r.json().get("ok")
    print(f"[{'ok' if ok else '!!'}] 텔레그램 발송 {'성공' if ok else r.text[:200]}")
    return bool(ok)

# ------------------------------------------------------------------
# main
# ------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="output", help="출력 디렉토리 (기본 output/)")
    ap.add_argument("--json", action="store_true", help="signals.json도 저장")
    ap.add_argument("--telegram", action="store_true",
                    help="요약 텔레그램 발송 (TELEGRAM_BOT_TOKEN/CHAT_ID 필요)")
    ap.add_argument("--url", default="", help="텔레그램 본문에 붙일 대시보드 URL")
    args = ap.parse_args()

    data, errors, sources = {}, [], {}
    has_key = bool(os.environ.get("FRED_API_KEY", "").strip())
    for key, meta in SERIES.items():
        if key == "USDJPY" and not has_key:
            data[key] = []   # 야후가 1차. 키 없으면 FRED 지연값 시도조차 생략
            continue
        try:
            data[key], sources[key] = fetch_series(key)
            print(f"[ok] {key:8s} {meta['sid']:<20s} n={len(data[key]):<4d} <- {sources[key]}", flush=True)
        except Exception as e:
            errors.append(f"{key}: {e}")
            data[key] = []
            print(f"[!!] {key:8s} 수집 실패: {e}", file=sys.stderr, flush=True)
    data["_sources"] = sources

    # 엔/달러는 야후 실시간을 우선 사용(FRED는 4영업일 지연)
    data["_jpy_src"] = "FRED"
    try:
        yf = fetch_yahoo_fx("JPY=X")
        data["USDJPY"] = yf
        data["_jpy_src"] = "Yahoo 실시간"
        print(f"[ok] USDJPY  Yahoo JPY=X          n={len(yf)} (FRED 대체)", flush=True)
    except Exception as e:
        errors.append(f"USDJPY(Yahoo): {e}")
        print(f"[!!] USDJPY  야후 실패, FRED 지연값 사용: {e}", file=sys.stderr)

    try:
        auctions = fetch_auctions()
        print(f"[ok] auctions n={len(auctions)}", flush=True)
    except Exception as e:
        auctions = []
        errors.append(f"auctions: {e}")
        print(f"[!!] auctions 수집 실패: {e}", file=sys.stderr)

    signals = build_signals(data, auctions)

    os.makedirs(args.out, exist_ok=True)
    gen_at = (dt.datetime.now(dt.timezone.utc) + dt.timedelta(hours=9)).strftime("%Y-%m-%d %H:%M")
    html = render_html(signals, auctions, gen_at)
    out_html = os.path.join(args.out, "index.html")
    with open(out_html, "w", encoding="utf-8") as f:
        f.write(html)

    if args.json:
        with open(os.path.join(args.out, "signals.json"), "w", encoding="utf-8") as f:
            json.dump([asdict(s) for s in signals], f, ensure_ascii=False, indent=2, default=str)

    # 콘솔 요약
    ov_st, ov_note = overall(signals)
    print("\n" + "=" * 62)
    print(f"종합: [{STATUS_KO[ov_st]}] {ov_note}")
    print("=" * 62)
    for s in signals:
        print(f"  {STATUS_KO[s.status]:2s} | {s.name:<22s} {s.value:<26s} {s.delta}")
    if errors:
        print(f"\n수집 실패 {len(errors)}건: " + "; ".join(errors))
    print(f"\n-> {out_html}")

    if args.telegram:
        send_telegram(telegram_text(signals, gen_at, args.url))

    # 절반 이상 결측이면 대시보드가 사실상 무의미하므로 실패로 끝낸다
    missing = sum(1 for s in signals if s.status == NA)
    if missing * 2 >= len(signals):
        print("", file=sys.stderr)
        print(f"[!!] 신호 {len(signals)}개 중 {missing}개 결측 — 실패 처리", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
