# -*- coding: utf-8 -*-
"""
comparator.py  v1.0
단계별 FP 산정 결과물 대조 엔진

한쪽 단계의 FP 산정서 여러 개와 다른 단계의 FP 산정서 여러 개를 읽어
단위프로세스를 짝지어 주고, 짝이 없는 것과 값이 달라진 것을 찾아낸다.

  A 그룹 : 예) 설계단계 FP 산정서 3개
  B 그룹 : 예) 종료단계 FP 산정서 5개

짝짓기 결과는 다섯 가지다.
  일치     같은 기능이 양쪽에 있고 값도 같다
  변경     같은 기능이 양쪽에 있으나 FP유형·복잡도·가중치 등이 다르다
  유사     이름이 완전히 같지는 않으나 비슷해서 같은 기능으로 본 것 (사람이 확인해야 함)
  A만 있음 뒤 단계에서 빠졌다 (누락 의심)
  B만 있음 앞 단계에 없던 것이 생겼다 (추가)

양식 사양과 FP 계산(FPR/FPV)은 calculator_ui.py 의 것을 그대로 쓴다.

단독 실행:
  python comparator.py A1.xlsm A2.xlsm -- B1.xlsm B2.xlsm
"""

import os
import re
import sys
import difflib

import openpyxl

import calculator_ui as U

VERSION = "1.0"

# 짝짓기 기준
KEY_FIELDS = {
    "proc": (["proc"], "단위프로세스명"),
    "app_proc": (["app", "proc"], "애플리케이션 + 단위프로세스명"),
    "app_biz_proc": (["app", "biz", "proc"], "애플리케이션 + 세부업무 + 단위프로세스명"),
    "proc_type": (["proc", "type"], "단위프로세스명 + FP유형"),
}

# 값 비교 항목
DIFF_FIELDS = [("type", "FP유형"), ("cx", "복잡도"), ("wt", "가중치"),
               ("ftr", "RET/FTR"), ("det", "DET"), ("dev", "개발유형"),
               ("app", "애플리케이션명"), ("biz", "세부업무명")]

DEFAULT_OPTS = {"key": "app_proc", "threshold": 0.85, "fuzzy": True,
                "ignore_space": True, "ignore_bracket": True,
                "ignore_case": True,
                "diff": ["type", "cx", "wt"]}


# ---------------------------------------------------------------- 읽기

def _num(v):
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _weight(method, ttype, ftr, det):
    """양식의 산정방법에 맞춰 복잡도와 가중치를 구한다."""
    t = (ttype or "").upper()
    if not t:
        return "", None
    if method == "간이법":
        return "", U.SIMPLE_WEIGHT.get(t)
    f, d = _num(ftr), _num(det)
    if f is None or d is None:
        return "", None
    g = U.fp_rating(t, f, d)
    if g is None or g == 0:
        return ("0" if g == 0 else ""), None
    return g, U.fp_value(t, g)


def load_rows(path, form_key=None):
    """산정양식 한 개를 읽어 기능 목록을 돌려준다."""
    form_key = form_key or U.guess_form(path)
    if form_key not in U.FORMS or not U.is_ready(form_key):
        raise ValueError("다룰 수 없는 양식입니다: %s" % (form_key or "판단 불가"))
    spec = U.FORMS[form_key]["fp"]
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = None
    for n in spec["sheets"]:
        if n in wb.sheetnames:
            ws = wb[n]
            break
    if ws is None:
        raise ValueError("FP 시트를 찾지 못했습니다: %s" % os.path.basename(path))
    hrow = U.find_header_row(ws, spec["anchor"])
    if hrow is None:
        raise ValueError("표 머리글을 찾지 못했습니다: %s" % os.path.basename(path))

    cx = {c["k"]: c["x"] for c in spec["cols"]}
    method = spec.get("method", "상세법")
    if spec.get("method_dv"):
        mc = U.find_method_cell(ws, spec["method_dv"])
        if mc:
            method = U._s(ws[mc].value) or method
    last = U.dv_last_row(ws, cx["type"]) or ws.max_row

    name = os.path.basename(path)
    rows = []
    for r in range(hrow + 1, min(ws.max_row, last) + 1):
        rec = {k: U._s(ws.cell(r, c).value) for k, c in cx.items() if k in
               ("app", "biz", "proc", "desc", "dev", "chg", "type",
                "ftr", "det", "remark")}
        if not any(rec.get(k) for k in ("app", "biz", "proc", "desc", "type")):
            continue
        rec["type"] = rec.get("type", "").upper()
        g, w = _weight(method, rec["type"], rec.get("ftr"), rec.get("det"))
        rec["cx"] = g
        rec["wt"] = w
        rec["src"] = name
        rec["path"] = path
        rec["no"] = r
        rec["form"] = form_key
        rows.append(rec)
    wb.close()
    return {"path": path, "name": name, "form": form_key,
            "formName": U.FORMS[form_key].get("name", form_key),
            "sheet": ws.title, "headerRow": hrow, "method": method,
            "rows": rows}


def load_group(files, label):
    """한 단계(그룹)의 파일 여러 개를 읽어 합친다."""
    out = {"label": label, "files": [], "rows": [], "errors": []}
    for f in files:
        path = f.get("path") if isinstance(f, dict) else f
        form = f.get("form") if isinstance(f, dict) else None
        try:
            d = load_rows(path, form)
        except Exception as e:
            out["errors"].append({"path": path, "msg": str(e)})
            continue
        out["files"].append({"path": path, "name": d["name"], "form": d["form"],
                             "formName": d["formName"], "sheet": d["sheet"],
                             "method": d["method"], "count": len(d["rows"])})
        out["rows"].extend(d["rows"])
    return out


# ---------------------------------------------------------------- 짝짓기

_BRACKET = re.compile(r"[\(\)\[\]\{\}（）［］【】<>《》]")


def norm(s, opts):
    t = U._s(s)
    if opts.get("ignore_bracket", True):
        t = _BRACKET.sub(" ", t)
    if opts.get("ignore_space", True):
        t = re.sub(r"[\s\u3000_\-·:/]+", "", t)
    else:
        t = re.sub(r"\s+", " ", t).strip()
    if opts.get("ignore_case", True):
        t = t.upper()
    return t


def row_key(row, opts):
    fields = KEY_FIELDS.get(opts.get("key", "app_proc"), KEY_FIELDS["app_proc"])[0]
    return "|".join(norm(row.get(f), opts) for f in fields)


def row_text(row, opts):
    """유사도 비교에 쓸 글자열."""
    return norm(" ".join(U._s(row.get(f)) for f in ("app", "biz", "proc")), opts)


def diff_of(a, b, fields):
    out = []
    for k, label in DIFF_FIELDS:
        if k not in fields:
            continue
        va, vb = a.get(k), b.get(k)
        if k in ("wt", "ftr", "det", "chg"):
            na, nb = _num(va), _num(vb)
            same = (na == nb)
        else:
            same = (U._s(va).upper() == U._s(vb).upper())
        if not same:
            out.append({"field": k, "label": label,
                        "a": "" if va is None else U._s(va),
                        "b": "" if vb is None else U._s(vb)})
    return out


def compare(group_a, group_b, opts=None):
    o = dict(DEFAULT_OPTS)
    o.update(opts or {})
    diff_fields = set(o.get("diff") or [])

    A, B = group_a["rows"], group_b["rows"]
    for i, r in enumerate(A):
        r["_i"] = i
        r["_k"] = row_key(r, o)
    for i, r in enumerate(B):
        r["_i"] = i
        r["_k"] = row_key(r, o)

    # 같은 그룹 안의 중복 표시
    def dup_map(rows):
        seen = {}
        for r in rows:
            seen.setdefault(r["_k"], []).append(r)
        return {k: v for k, v in seen.items() if len(v) > 1 and k.strip("|")}
    dup_a, dup_b = dup_map(A), dup_map(B)

    idx_b = {}
    for r in B:
        idx_b.setdefault(r["_k"], []).append(r)

    pairs = []
    used_b = set()
    left_a = []
    for r in A:
        cand = [x for x in idx_b.get(r["_k"], []) if x["_i"] not in used_b]
        if not r["_k"].strip("|") or not cand:
            left_a.append(r)
            continue
        m = cand[0]
        used_b.add(m["_i"])
        d = diff_of(r, m, diff_fields)
        pairs.append({"status": "변경" if d else "일치", "score": 1.0,
                      "a": r, "b": m, "diffs": d})

    left_b = [x for x in B if x["_i"] not in used_b]

    # 이름이 비슷한 것 찾기
    if o.get("fuzzy") and left_a and left_b:
        th = float(o.get("threshold", 0.85))
        texts_b = [(x, row_text(x, o)) for x in left_b]
        still_a, taken = [], set()
        for r in left_a:
            ta = row_text(r, o)
            best, best_s = None, 0.0
            if ta:
                for x, tb in texts_b:
                    if x["_i"] in taken or not tb:
                        continue
                    s = difflib.SequenceMatcher(None, ta, tb).ratio()
                    if s > best_s:
                        best, best_s = x, s
            if best is not None and best_s >= th:
                taken.add(best["_i"])
                pairs.append({"status": "유사", "score": round(best_s, 4),
                              "a": r, "b": best, "diffs": diff_of(r, best, diff_fields)})
            else:
                still_a.append(r)
        left_a = still_a
        left_b = [x for x in left_b if x["_i"] not in taken]

    for r in left_a:
        pairs.append({"status": "A만 있음", "score": 0, "a": r, "b": None, "diffs": []})
    for r in left_b:
        pairs.append({"status": "B만 있음", "score": 0, "a": None, "b": r, "diffs": []})

    order = {"A만 있음": 0, "B만 있음": 1, "변경": 2, "유사": 3, "일치": 4}
    pairs.sort(key=lambda p: (order.get(p["status"], 9),
                              U._s((p["a"] or p["b"]).get("app")),
                              U._s((p["a"] or p["b"]).get("proc"))))

    return {"opts": o,
            "a": summarize(group_a), "b": summarize(group_b),
            "pairs": [strip_pair(p) for p in pairs],
            "stats": stats_of(pairs, dup_a, dup_b),
            "keyLabel": KEY_FIELDS.get(o["key"], KEY_FIELDS["app_proc"])[1],
            "dupA": [{"key": k, "rows": [brief(x) for x in v]} for k, v in dup_a.items()],
            "dupB": [{"key": k, "rows": [brief(x) for x in v]} for k, v in dup_b.items()]}


def brief(r):
    return {"src": r.get("src"), "no": r.get("no"), "app": r.get("app"),
            "biz": r.get("biz"), "proc": r.get("proc"), "type": r.get("type"),
            "cx": r.get("cx"), "wt": r.get("wt"), "ftr": r.get("ftr"),
            "det": r.get("det"), "dev": r.get("dev"), "desc": r.get("desc")}


def strip_pair(p):
    return {"status": p["status"], "score": p["score"], "diffs": p["diffs"],
            "a": brief(p["a"]) if p["a"] else None,
            "b": brief(p["b"]) if p["b"] else None}


def summarize(g):
    by = {}
    fp = 0.0
    for r in g["rows"]:
        t = r.get("type") or "(없음)"
        b = by.setdefault(t, {"n": 0, "fp": 0.0})
        b["n"] += 1
        if r.get("wt"):
            b["fp"] += r["wt"]
            fp += r["wt"]
    return {"label": g["label"], "files": g["files"], "errors": g["errors"],
            "count": len(g["rows"]), "fp": round(fp, 2), "byType": by}


def stats_of(pairs, dup_a, dup_b):
    s = {"일치": 0, "변경": 0, "유사": 0, "A만 있음": 0, "B만 있음": 0}
    fp_a = fp_b = 0.0
    for p in pairs:
        s[p["status"]] = s.get(p["status"], 0) + 1
        if p["a"] and p["a"].get("wt"):
            fp_a += p["a"]["wt"]
        if p["b"] and p["b"].get("wt"):
            fp_b += p["b"]["wt"]
    s["중복A"] = sum(len(v) for v in dup_a.values())
    s["중복B"] = sum(len(v) for v in dup_b.values())
    s["fpA"] = round(fp_a, 2)
    s["fpB"] = round(fp_b, 2)
    s["fpDiff"] = round(fp_b - fp_a, 2)
    return s


# ---------------------------------------------------------------- 내보내기

def export_xlsx(result, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "대조결과"
    a, b = result["a"]["label"], result["b"]["label"]
    head = ["상태", "유사도", "차이 항목",
            "%s 파일" % a, "행", "애플리케이션", "세부업무", "단위프로세스명",
            "FP유형", "복잡도", "가중치",
            "%s 파일" % b, "행", "애플리케이션", "세부업무", "단위프로세스명",
            "FP유형", "복잡도", "가중치"]
    ws.append(head)
    for p in result["pairs"]:
        ra, rb = p["a"], p["b"]
        d = ", ".join("%s %s→%s" % (x["label"], x["a"] or "-", x["b"] or "-")
                      for x in p["diffs"])
        row = [p["status"], p["score"] if p["status"] == "유사" else "", d]
        for r in (ra, rb):
            if r:
                row += [r["src"], r["no"], r["app"], r["biz"], r["proc"],
                        r["type"], r["cx"], r["wt"]]
            else:
                row += ["", "", "", "", "", "", "", ""]
        ws.append(row)
    ws.freeze_panes = "A2"
    widths = [10, 8, 34, 26, 6, 16, 14, 30, 8, 8, 8, 26, 6, 16, 14, 30, 8, 8, 8]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    s = wb.create_sheet("요약")
    st = result["stats"]
    s.append(["대조 기준", result["keyLabel"]])
    s.append(["유사도 임계값", result["opts"]["threshold"]])
    s.append([])
    s.append(["구분", a, b])
    s.append(["파일 수", len(result["a"]["files"]), len(result["b"]["files"])])
    s.append(["기능 수", result["a"]["count"], result["b"]["count"]])
    s.append(["기능점수", result["a"]["fp"], result["b"]["fp"]])
    s.append([])
    s.append(["상태", "건수"])
    for k in ("일치", "변경", "유사", "A만 있음", "B만 있음"):
        s.append([k.replace("A", a).replace("B", b), st.get(k, 0)])
    s.append([])
    s.append(["기능점수 차이 (%s − %s)" % (b, a), st["fpDiff"]])
    s.column_dimensions["A"].width = 30
    s.column_dimensions["B"].width = 18
    s.column_dimensions["C"].width = 18

    types = sorted(set(list(result["a"]["byType"]) + list(result["b"]["byType"])))
    t = wb.create_sheet("유형별")
    t.append(["FP유형", "%s 건수" % a, "%s FP" % a, "%s 건수" % b, "%s FP" % b, "FP 차이"])
    for ty in types:
        x = result["a"]["byType"].get(ty, {"n": 0, "fp": 0})
        y = result["b"]["byType"].get(ty, {"n": 0, "fp": 0})
        t.append([ty, x["n"], round(x["fp"], 2), y["n"], round(y["fp"], 2),
                  round(y["fp"] - x["fp"], 2)])
    for c in "ABCDEF":
        t.column_dimensions[c].width = 14

    wb.save(path)
    return path


# ---------------------------------------------------------------- 단독 실행

def _cli():
    args = sys.argv[1:]
    if "--" not in args:
        print(__doc__)
        return 1
    i = args.index("--")
    fa, fb = args[:i], args[i + 1:]
    if not fa or not fb:
        print("양쪽 파일을 모두 지정하십시오.")
        return 1
    ga = load_group(fa, "A")
    gb = load_group(fb, "B")
    r = compare(ga, gb)
    st = r["stats"]
    print("A %d개 파일 / 기능 %d건 / %.1f FP" % (len(ga["files"]), r["a"]["count"], r["a"]["fp"]))
    print("B %d개 파일 / 기능 %d건 / %.1f FP" % (len(gb["files"]), r["b"]["count"], r["b"]["fp"]))
    print("일치 %d · 변경 %d · 유사 %d · A만 %d · B만 %d · FP차이 %+.1f"
          % (st["일치"], st["변경"], st["유사"], st["A만 있음"], st["B만 있음"], st["fpDiff"]))
    for p in r["pairs"]:
        if p["status"] in ("A만 있음", "B만 있음"):
            r0 = p["a"] or p["b"]
            print("  [%s] %s / %s (%s %d행)"
                  % (p["status"], r0["app"], r0["proc"], r0["src"], r0["no"]))
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
