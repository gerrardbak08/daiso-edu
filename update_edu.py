#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
정기안전보건교육 이수현황 페이지 자동 업데이트 스크립트
사용법:  python update_edu.py [미이수현황.xls] [대상연월일 YYYYMMDD]
  - 인자 생략 시: 폴더에서 가장 최근 .xls 자동 선택, 기준일은 파일명 끝 8자리(YYYYMMDD)에서 추출
동작: 미이수자 .xls 를 읽어 26상반기_안전보건교육_이수현황.html 의 모든 수치/차트/모달/다운로드엑셀을 갱신.
주의: 부서별 "대상(전체인원)"은 미이수 명단에 없으므로 기존 HTML의 부서 구조에서 승계함(조직개편 시 ORPHAN 경고).
"""
import sys, os, re, json, glob, base64, io, datetime
import xlrd, openpyxl

HTML = "26상반기_안전보건교육_이수현황.html"

# 조직개편 병합 맵: 옛 부서들 -> 새 부서명 (대상 합산 승계). 필요 시 여기에 추가.
MERGE = {
    "상품본부": {"name": "미용/패션/식품/인테리어개발부문",
               "from": ["상품개발부문", "패션/뷰티개발부문", "생활용품개발부문"]},
}

def cls(r):
    return "good" if r >= 90 else "warn" if r >= 70 else "bad" if r >= 50 else "crit"

def fmt_rate(r):
    r = round(r, 1)
    return int(r) if r == int(r) else r

def pick_xls():
    if len(sys.argv) > 1:
        return sys.argv[1]
    cands = [f for f in glob.glob("*.xls") if "미이수" in f or "안전보건" in f]
    cands = cands or glob.glob("*.xls")
    if not cands:
        sys.exit("‼ .xls 파일을 찾을 수 없습니다.")
    return max(cands, key=os.path.getmtime)

def date_kr(fn):
    if len(sys.argv) > 2:
        s = sys.argv[2]
    else:
        m = re.search(r"(20\d{6})", fn)
        s = m.group(1) if m else None
    if not s:
        sys.exit("‼ 기준일(YYYYMMDD)을 인자로 주거나 파일명에 포함하세요.")
    d = datetime.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    wk = "월화수목금토일"[d.weekday()]
    return d, f"{d.year}. {d.month}. {d.day}.({wk})", s

# ---------- 1. 미이수 .xls 파싱 ----------
fn = pick_xls()
d, kr, ymd = date_kr(fn)
print(f"입력: {fn}\n기준일: {kr}")
sh = xlrd.open_workbook(fn).sheet_by_index(0)
people = []                       # (사번,성명,직책,본부,부서,팀,개인율)
from collections import defaultdict, Counter
miss = defaultdict(lambda: defaultdict(int))   # 본부 -> 부서 -> 미이수수
for r in range(sh.nrows):
    no = sh.cell_value(r, 1)
    if not isinstance(no, float):
        continue
    sabun = str(sh.cell_value(r, 3)).strip()
    name  = str(sh.cell_value(r, 4)).strip()
    jik   = str(sh.cell_value(r, 5)).strip() or "-"
    hq    = str(sh.cell_value(r, 6)).strip()
    dept  = str(sh.cell_value(r, 7)).strip() or "(부서 미구분)"
    team  = str(sh.cell_value(r, 8)).strip()
    rate  = sh.cell_value(r, 12)
    try: rate = int(float(rate))
    except: rate = rate
    people.append((sabun, name, jik, hq, dept, team, rate))
    miss[hq][dept] += 1
print(f"미이수자: {len(people)}명")

# ---------- 2. 기존 HTML에서 부서 대상(roster) 승계 ----------
html = open(HTML, encoding="utf-8").read()
old = json.loads(re.search(r"var HQDATA=(\{.*?\});", html, re.S).group(1))
tgt = {hq: {x["부서"]: x["대상"] for x in info["depts"]} for hq, info in old.items()}
for hq, m in MERGE.items():
    if hq in tgt and all(x in tgt[hq] for x in m["from"]):
        s = sum(tgt[hq].pop(x) for x in m["from"])
        tgt[hq][m["name"]] = tgt[hq].get(m["name"], 0) + s
        print(f"병합: {hq} {m['from']} -> {m['name']} (대상 {s})")

# ---------- 3. 신규 HQDATA 계산 ----------
HQ = {}
for hq, depts_t in tgt.items():
    for dept, c in miss.get(hq, {}).items():
        if dept not in depts_t:
            depts_t[dept] = c   # ORPHAN: 대상=미이수로 임시 (조직개편 추정)
            print(f"⚠ ORPHAN(신규부서) {hq}/{dept} 미이수{c} → 대상={c}로 처리. 확인 요망.")
    rows = []
    for dept, t in depts_t.items():
        m = miss.get(hq, {}).get(dept, 0)
        dn = t - m
        rate = round(dn / t * 100, 1) if t else 0.0
        rows.append({"부서": dept, "대상": t, "이수": dn, "미이수": m, "율": rate, "cls": cls(rate)})
    rows.sort(key=lambda x: (-x["미이수"], x["율"]))
    T = sum(x["대상"] for x in rows); D = sum(x["이수"] for x in rows); M = sum(x["미이수"] for x in rows)
    R = round(D / T * 100, 1)
    HQ[hq] = {"대상": T, "이수": D, "미이수": M, "율": R, "cls": cls(R), "n": len(rows), "depts": rows}

TT = sum(v["대상"] for v in HQ.values()); TD = sum(v["이수"] for v in HQ.values()); TM = sum(v["미이수"] for v in HQ.values())
TR = round(TD / TT * 100, 1)
NDEPT = sum(v["n"] for v in HQ.values())
assert TM == len(people), f"미이수 합 불일치 {TM} != {len(people)}"
print(f"전사: 대상{TT} 이수{TD} 미이수{TM} 율{TR} / 본부{len(HQ)} 부서{NDEPT}")

chart = sorted(HQ.items(), key=lambda x: x[1]["율"])   # 율 asc
cc = Counter(v["cls"] for v in HQ.values())

# ---------- 4. 다운로드용 3시트 XLSX 재생성 ----------
wbk = openpyxl.Workbook()
s1 = wbk.active; s1.title = "본부별 집계"
s1.append(["본부", "대상", "이수", "미이수", "이수율(%)"])
for hq, v in chart:
    s1.append([hq, v["대상"], v["이수"], v["미이수"], fmt_rate(v["율"])])
s2 = wbk.create_sheet("부서별 집계")
s2.append(["본부", "부서", "대상", "이수", "미이수", "이수율(%)"])
for hq, v in chart:
    for x in sorted(v["depts"], key=lambda x: x["율"]):
        s2.append([hq, x["부서"], x["대상"], x["이수"], x["미이수"], fmt_rate(x["율"])])
s3 = wbk.create_sheet("미이수자 명단")
s3.append(["사번", "성명", "직책", "본부", "부서", "팀", "개인이수율(%)"])
for p in people:
    s3.append(list(p))
buf = io.BytesIO(); wbk.save(buf)
b64 = base64.b64encode(buf.getvalue()).decode()

# ---------- 5. HTML 조각 생성 ----------
def chip(c, lbl, n):
    return f'<span class="sd-chip"><i class="dot {c}"></i>{lbl} <b>{n}</b></span>'
seg = "".join(f'<div class="sd-seg {c}" style="flex:{cc[c]}"></div>' for c in ["good","warn","bad","crit"] if cc[c])
statdist = (f'<div class="statdist"><div class="sd-label">본부 상태 분포</div>'
            f'<div class="sd-bar">{seg}</div><div class="sd-legend">'
            f'{chip("good","우수",cc["good"])}{chip("warn","보통",cc["warn"])}'
            f'{chip("bad","미흡",cc["bad"])}{chip("crit","집중",cc["crit"])}</div></div>')

bars = []
for i, (hq, v) in enumerate(chart):
    dl = f"{i*0.07:.2f}"
    bars.append(
        f'<button class="hbar-row" onclick="openModal(\'{hq}\')">'
        f'<span class="hbar-label">{hq}</span>'
        f'<div class="hbar-track"><div class="hbar-fill {v["cls"]}" data-w="{v["율"]}" style="width:0%;transition-delay:{dl}s"></div></div>'
        f'<span class="hbar-val cu-bar {v["cls"]}" data-to="{v["율"]}" data-delay="{dl}">0.0%</span>'
        f'<span class="hbar-go">›</span>'
        f'<span class="hbar-tip">미이수 <b>{v["미이수"]}</b>명 · 대상 {v["대상"]}명</span></button>')
hbar = '<div class="hbar-chart">' + "".join(bars) + "</div>"

alert_items = []
for hq, v in HQ.items():
    for x in v["depts"]:
        if x["미이수"] >= 5 and x["율"] < 50:
            alert_items.append((hq, x["부서"], x["미이수"], x["율"]))
alert_items.sort(key=lambda x: (x[3], -x[2]))
al_li = "".join(
    f'<li><span class="wname">{hq} · {dp}</span>'
    f'<span class="wstat">미이수 <b>{m}</b>명 <span class="wrate">({r}%)</span></span></li>'
    for hq, dp, m, r in alert_items)
alert_ul = f"<ul>{al_li}</ul>"

# ---------- 6. HTML 치환 ----------
def sub1(pat, repl, s, n=1):
    s2, c = re.subn(pat, lambda m: repl, s, count=n)
    assert c == n, f"치환 실패({c}/{n}): {pat[:60]}"
    return s2

# 기준일 (헤더 meta, foot)
html = html.replace("2026. 6. 4.(목)", kr)
# hero ring/수치
html = sub1(r'data-pct="59\.6"', f'data-pct="{TR}"', html)
html = sub1(r'class="cu-rate" data-to="59\.6"', f'class="cu-rate" data-to="{TR}"', html)
html = sub1(r'class="cu" data-to="529"', f'class="cu" data-to="{TD}"', html)
html = sub1(r'class="miss cu" data-to="358"', f'class="miss cu" data-to="{TM}"', html)
# hero-meta 부서 수
html = sub1(r'교육 대상 <b>887명</b> · 관리 본부 12개 · 부서 96개',
            f'교육 대상 <b>{TT}명</b> · 관리 본부 {len(HQ)}개 · 부서 {NDEPT}개', html)
# avgnote 전사평균 + CSS 평균선 위치
html = sub1(r'전사 평균 <b>59\.6%</b>', f'전사 평균 <b>{TR}%</b>', html)
html = sub1(r'\.hbar-track::after\{content:"";position:absolute;top:0;bottom:0;left:59\.6%',
            f'.hbar-track::after{{content:"";position:absolute;top:0;bottom:0;left:{TR}%', html)
# statdist / hbar-chart 블록 교체 (각각 단일 라인 → greedy 매칭으로 라인 전체 치환)
html = sub1(r'<div class="statdist">.*</div>', statdist, html)
html = sub1(r'<div class="hbar-chart">.*</div>', hbar, html)
# alert ul (단일 라인)
html = sub1(r'<ul><li><span class="wname">.*</ul>', alert_ul, html)
# dlmeta
html = sub1(r'미이수자 명단 358명 · 96개 부서 상세',
            f'미이수자 명단 {TM}명 · {NDEPT}개 부서 상세', html)
# 다운로드 파일명
html = sub1(r'var XLSX_NAME="26상반기_안전보건교육_미이수현황_20260604\.xlsx";',
            f'var XLSX_NAME="26상반기_안전보건교육_미이수현황_{ymd}.xlsx";', html)
# XLSX_B64
html = sub1(r'var XLSX_B64="[^"]+";', f'var XLSX_B64="{b64}";', html)
# HQDATA
html = sub1(r'var HQDATA=\{.*?\};',
            "var HQDATA=" + json.dumps(HQ, ensure_ascii=False) + ";", html)

open(HTML, "w", encoding="utf-8").write(html)
print(f"\n✅ {HTML} 갱신 완료")
print(f"   최우선 관리대상 {len(alert_items)}곳 / 상태분포 우수{cc['good']} 보통{cc['warn']} 미흡{cc['bad']} 집중{cc['crit']}")
