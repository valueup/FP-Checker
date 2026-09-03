# -*- coding: utf-8 -*-
"""
comparator_json_ui.py  v1.0
설계단계 JSON 과 종료단계 JSON 을 받아 기능점수를 대조하는 화면.

  왼쪽(A)  기준이 된다. 이 자료의 파일·행 차례를 그대로 지킨다.
  오른쪽(B) 왼쪽에 맞춰 붙인다. 짝이 없으면 빈칸으로 둔다.
  비고     설계단계에만 있음 / 일치 / 변경 / 종료단계에만 있음

엑셀 원장을 직접 읽는 comparator_ui.py 와 짝이 되는 화면이다.
맞추는 일은 comparator.py 가 다 한다. 이 파일이 하는 일은
JSON 을 읽어 comparator.py 가 아는 모양(그룹)으로 바꿔 넣고, 결과를 보여 주는 것뿐이다.
그래서 1:N 묶음, 수동 짝짓기 저장, 엑셀 내보내기가 그대로 따라온다.

  python comparator_json_ui.py [설계단계.json 종료단계.json]
"""

import os
import re
import sys
import json
import time
import threading
import configparser

# ── 폴더 구조 대응 (comparator_ui.py 와 같은 방식) ──────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
_PKGS = ("Calculator", "Comparator", "DocParser")
for _p in [os.path.join(os.path.dirname(_HERE), _n) for _n in _PKGS] + [_HERE]:
    if os.path.isdir(_p):
        if _p in sys.path:
            sys.path.remove(_p)
        sys.path.insert(0, _p)


def project_root():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(_HERE) if os.path.basename(_HERE) in _PKGS else _HERE


try:
    from flask import Flask, jsonify, request, Response
except ImportError:
    print("flask 가 필요합니다.  pip install flask")
    raise

import comparator as CMP
import calculator_ui as U      # 크롬 창 띄우기·빈 포트 찾기를 함께 쓴다

try:
    import mapping as MAPPING
except Exception:
    MAPPING = None

VERSION = "1.1"
CMP_MIN = "2.1"          # status_text 와 opts["text"] 가 들어간 판
INI_PATH = os.path.join(project_root(), "comparator_json.ini")

# 이 화면이 읽을 수 있는 JSON 형식
SCHEMAS = ("fpset/1.0",)


def status_text(status, a_lab, b_lab):
    """'A만 있음' 을 그룹 이름으로 바꾼다.

    원래 comparator.py 가 하는 일이다. 엔진이 옛 판이라 그 함수가 없더라도
    화면은 뜨도록 여기에 같은 것을 둔다.
    """
    fn = getattr(CMP, "status_text", None)
    if fn:
        return fn(status, a_lab, b_lab)
    if status == "A만 있음":
        return "%s에만 있음" % a_lab
    if status == "B만 있음":
        return "%s에만 있음" % b_lab
    return status


def engine_check():
    """엔진이 이 화면에 필요한 것을 갖췄는지 본다."""
    v = str(getattr(CMP, "VERSION", "0"))
    need = []
    if not hasattr(CMP, "status_text"):
        need.append("status_text()")
    if "text" not in getattr(CMP, "DEFAULT_OPTS", {}):
        need.append('DEFAULT_OPTS["text"]')
    if v < CMP_MIN or need:
        msg = ("comparator.py 가 v%s 입니다. 이 화면은 v%s 이상이 필요합니다." % (v, CMP_MIN))
        if need:
            msg += " 없는 것: " + ", ".join(need) + "."
        msg += " 새 comparator.py 로 바꾸십시오. 지금은 일부 기능이 빠진 채로 돕니다."
        return msg
    return ""

app = Flask(__name__)

STATE = {"a": None, "b": None, "ga": None, "gb": None,
         "result": None, "rows": None, "store": None, "storePath": ""}

# 화면이 4초마다 신호를 보낸다. 신호가 끊기면 창을 닫은 것으로 보고 프로그램도 끈다.
LAST_PING = [0.0]


# ---------------------------------------------------------------- 설정 파일

def load_ini():
    cp = configparser.ConfigParser()
    cp["main"] = {"dir": project_root(), "a": "", "b": "",
                  "label_a": "설계단계", "label_b": "종료단계"}
    try:
        cp.read(INI_PATH, encoding="utf-8")
    except Exception:
        pass
    if "main" not in cp:
        cp["main"] = {}
    return cp


def save_ini(cp):
    try:
        with open(INI_PATH, "w", encoding="utf-8") as f:
            cp.write(f)
    except Exception:
        pass


def out_dir():
    d = os.path.join(project_root(), "출력")
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------- JSON 읽기

def peek_json(path):
    """이 화면이 다룰 수 있는 파일인지, 무엇이 들어 있는지만 훑는다."""
    try:
        with open(path, encoding="utf-8") as f:
            head = f.read(600)
    except Exception:
        return None
    m = re.search(r'"schema"\s*:\s*"([^"]+)"', head)
    if not m or m.group(1) not in SCHEMAS:
        return None
    d = {"schema": m.group(1)}
    for k in ("id", "stage", "asOf"):
        mm = re.search(r'"%s"\s*:\s*"([^"]*)"' % k, head)
        if mm:
            d[k] = mm.group(1)
    return d


def read_json(path):
    with open(path, encoding="utf-8") as f:
        d = json.load(f)
    if d.get("schema") not in SCHEMAS:
        raise ValueError("다룰 수 없는 형식입니다: %s" % d.get("schema"))
    return d


def to_group(doc, label, path):
    """JSON 을 comparator.py 가 아는 그룹 모양으로 바꾼다.

    comparator.py 는 원래 엑셀 원장을 읽어 이 모양을 만든다. 자리만 같으면
    맞추는 일도 엑셀로 내보내는 일도 손댈 것 없이 그대로 돈다.

    src·no 는 **원본 엑셀 원장의 파일이름과 행번호**를 그대로 쓴다.
    수동 짝짓기를 저장할 때 이것이 열쇠가 되므로, JSON 을 다시 만들어도
    같은 행을 가리킨다.
    """
    rows = []
    for it in doc.get("items", []):
        if it.get("kind") != "fp":
            continue
        raw, nm = it.get("raw", {}), it.get("norm", {})
        codes = it.get("codes", {})
        rows.append({
            "app": raw.get("app", ""), "biz": raw.get("biz", ""),
            "proc": raw.get("proc", ""), "desc": raw.get("desc", ""),
            "dev": raw.get("dev", ""), "chg": raw.get("chg"),
            "type": (raw.get("type") or "").upper(),
            "ftr": raw.get("ftr"), "det": raw.get("det"),
            "cx": raw.get("cx", ""), "wt": raw.get("wt"),
            "remark": raw.get("remark", ""),
            # 정규화 값. 유사도 비교에 이것을 쓰면 이름 짓는 버릇 차이를 덜 탄다.
            "domain": nm.get("domain", ""), "screen": nm.get("screen", ""),
            "action": nm.get("action", ""),
            "ui": codes.get("screenIds") or [], "tb": codes.get("tableNames") or [],
            "src": it["ref"].get("file", ""), "path": path,
            "no": it["ref"].get("row") or 0,
            "form": doc["dataset"].get("id", ""), "uid": it.get("uid", "")})

    files = []
    for s in doc.get("sources", []):
        if s.get("kind") != "fp":
            continue
        files.append({"path": s.get("path", path), "name": s.get("file", ""),
                      "form": doc["dataset"].get("id", ""),
                      "formName": doc["dataset"].get("stage", ""),
                      "sheet": s.get("sheet", ""),
                      "method": s.get("method", ""),
                      "methodSrc": "JSON",
                      "count": s.get("counts", {}).get("items", 0)})
    return {"label": label, "files": files, "rows": rows, "errors": []}


def info_of(doc, path):
    ds = doc.get("dataset", {})
    c = doc.get("index", {}).get("counts", {})
    fan = doc.get("index", {}).get("fanout", {})
    return {"file": os.path.basename(path), "path": path,
            "id": ds.get("id", ""), "stage": ds.get("stage", ""),
            "asOf": ds.get("asOf", ""), "note": ds.get("note", ""),
            "counts": c,
            "sources": [{"kind": s["kind"], "file": s["file"],
                         "method": s.get("method", ""),
                         "n": s.get("counts", {}).get("items", 0)}
                        for s in doc.get("sources", [])],
            "usableKeys": [k for k, v in fan.items() if v.get("usable")],
            "crowdedKeys": [{"k": k, "max": v.get("max")}
                            for k, v in fan.items() if not v.get("usable")]}


# ---------------------------------------------------------------- 화면에 낼 줄

def cell(r):
    if r is None:
        return None
    return {"src": r.get("src"), "no": r.get("no"), "app": r.get("app"),
            "biz": r.get("biz"), "proc": r.get("proc"),
            "domain": r.get("domain"), "screen": r.get("screen"),
            "action": r.get("action"), "type": r.get("type"),
            "ret": r.get("ftr"), "det": r.get("det"), "wt": r.get("wt"),
            "ui": (r.get("ui") or [""])[0]}


def make_rows(result, ga, gb):
    """묶음을 줄로 편다. 왼쪽 파일·행 차례를 지키고, 왼쪽이 없는 것은 뒤로 보낸다."""
    ia = {(r["src"], r["no"]): r for r in ga["rows"]}
    ib = {(r["src"], r["no"]): r for r in gb["rows"]}
    oa = {(r["src"], r["no"]): i for i, r in enumerate(ga["rows"])}
    ob = {(r["src"], r["no"]): i for i, r in enumerate(gb["rows"])}
    a_lab, b_lab = result["a"]["label"], result["b"]["label"]

    groups = sorted(
        result["groups"],
        key=lambda g: (0, oa.get((g["a"][0]["src"], g["a"][0]["no"]), 9 ** 9))
        if g["a"] else (1, ob.get((g["b"][0]["src"], g["b"][0]["no"]), 9 ** 9)))

    out = []
    for g in groups:
        na, nb = len(g["a"]), len(g["b"])
        for k in range(max(na, nb, 1)):
            ra = ia.get((g["a"][k]["src"], g["a"][k]["no"])) if k < na else None
            rb = ib.get((g["b"][k]["src"], g["b"][k]["no"])) if k < nb else None
            out.append({
                "gid": g["gid"], "status": g["status"],
                "label": status_text(g["status"], a_lab, b_lab),
                "by": g.get("byLabel", ""), "score": g.get("score"),
                "basis": g.get("basis", ""), "note": g.get("note", ""),
                "size": [na, nb], "head": (k == 0),
                "fpDiff": round(g["fpB"] - g["fpA"], 2) if k == 0 else None,
                "diffs": g.get("diffs") if k == 0 else [],
                "a": cell(ra), "b": cell(rb)})
    return out


def tally(result, ga, gb):
    """집계. 상태별·FP유형별·분야별·짝지은 경로별."""
    import collections
    a_lab, b_lab = result["a"]["label"], result["b"]["label"]
    st, stfp = collections.Counter(), collections.defaultdict(lambda: [0.0, 0.0])
    route = collections.Counter()
    for g in result["groups"]:
        st[g["status"]] += 1
        stfp[g["status"]][0] += g["fpA"]
        stfp[g["status"]][1] += g["fpB"]
        route[g.get("byLabel") or g["by"]] += 1

    def piv(rows, field):
        d = collections.defaultdict(lambda: [0, 0.0])
        for r in rows:
            k = r.get(field) or "(없음)"
            d[k][0] += 1
            d[k][1] += r.get("wt") or 0
        return d

    def merge(x, y):
        return [{"k": k,
                 "nA": x.get(k, [0, 0.0])[0], "fpA": round(x.get(k, [0, 0.0])[1], 1),
                 "nB": y.get(k, [0, 0.0])[0], "fpB": round(y.get(k, [0, 0.0])[1], 1),
                 "nDiff": y.get(k, [0, 0.0])[0] - x.get(k, [0, 0.0])[0],
                 "fpDiff": round(y.get(k, [0, 0.0])[1] - x.get(k, [0, 0.0])[1], 1)}
                for k in sorted(set(x) | set(y))]

    A, B = ga["rows"], gb["rows"]
    fa = round(sum(r.get("wt") or 0 for r in A), 1)
    fb = round(sum(r.get("wt") or 0 for r in B), 1)
    ma = sum(len(g["a"]) for g in result["groups"] if g["a"] and g["b"])
    mb = sum(len(g["b"]) for g in result["groups"] if g["a"] and g["b"])
    return {
        "status": [{"k": k, "label": status_text(k, a_lab, b_lab), "n": st[k],
                    "fpA": round(stfp[k][0], 1), "fpB": round(stfp[k][1], 1)}
                   for k in CMP.STATUSES if st[k]],
        "byType": merge(piv(A, "type"), piv(B, "type")),
        "byDomain": merge(piv(A, "domain"), piv(B, "domain")),
        "byRoute": [{"k": k, "n": v} for k, v in route.most_common()],
        "total": {"nA": len(A), "nB": len(B), "fpA": fa, "fpB": fb,
                  "fpDiff": round(fb - fa, 1),
                  "groups": len(result["groups"]),
                  "matchedA": ma, "matchedB": mb,
                  "rateA": round(ma / max(1, len(A)) * 100, 1),
                  "rateB": round(mb / max(1, len(B)) * 100, 1)}}


# ---------------------------------------------------------------- API

@app.after_request
def no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/")
def index():
    return Response(HTML, mimetype="text/html")


@app.route("/api/init")
def api_init():
    cp = load_ini()
    m = cp["main"]
    bad = engine_check()
    return jsonify({"ok": True, "version": VERSION,
                    "engine": getattr(CMP, "VERSION", "?"),
                    "warn": bad, "dir": m.get("dir", project_root()),
                    "last": {"a": m.get("a", ""), "b": m.get("b", "")},
                    "labels": {"a": m.get("label_a", "설계단계"),
                               "b": m.get("label_b", "종료단계")},
                    "opts": CMP.DEFAULT_OPTS, "outDir": out_dir(),
                    "mapping": MAPPING is not None})


@app.route("/api/browse", methods=["POST"])
def api_browse():
    """폴더를 훑어 하위 폴더와 읽을 수 있는 JSON 을 돌려준다."""
    d = (request.json or {}).get("dir") or project_root()
    d = os.path.abspath(os.path.expanduser(d.strip().strip('"')))
    if os.path.isfile(d):
        d = os.path.dirname(d)
    if not os.path.isdir(d):
        return jsonify({"ok": False, "msg": "폴더가 없습니다: %s" % d})
    dirs, files = [], []
    try:
        for n in sorted(os.listdir(d)):
            p = os.path.join(d, n)
            if os.path.isdir(p):
                dirs.append({"name": n, "path": p})
            elif n.lower().endswith(".json"):
                info = peek_json(p)
                files.append({"name": n, "path": p,
                              "size": os.path.getsize(p), "ok": info is not None,
                              "stage": (info or {}).get("stage", ""),
                              "asOf": (info or {}).get("asOf", "")})
    except PermissionError:
        return jsonify({"ok": False, "msg": "폴더를 열 수 없습니다."})
    up = os.path.dirname(d)
    return jsonify({"ok": True, "dir": d, "up": up if up != d else "",
                    "dirs": dirs, "files": files})


@app.route("/api/load", methods=["POST"])
def api_load():
    body = request.json or {}
    side = "b" if body.get("side") == "b" else "a"
    path = (body.get("path") or "").strip().strip('"')
    if not path or not os.path.exists(path):
        return jsonify({"ok": False, "msg": "파일을 찾을 수 없습니다."})
    try:
        doc = read_json(path)
    except Exception as e:
        return jsonify({"ok": False, "msg": "읽기 실패: %s" % e})
    STATE[side] = doc
    STATE["g" + side] = None
    STATE["result"] = STATE["rows"] = None
    cp = load_ini()
    cp["main"]["dir"] = os.path.dirname(path)
    cp["main"][side] = path
    save_ini(cp)
    return jsonify({"ok": True, "side": side, "info": info_of(doc, path)})


@app.route("/api/compare", methods=["POST"])
def api_compare():
    if not STATE["a"] or not STATE["b"]:
        return jsonify({"ok": False, "msg": "양쪽 JSON 을 모두 고르십시오."})
    body = request.json or {}
    labels = body.get("labels") or {}
    la = labels.get("a") or STATE["a"]["dataset"].get("stage") or "A"
    lb = labels.get("b") or STATE["b"]["dataset"].get("stage") or "B"

    ga = to_group(STATE["a"], la, STATE["a"].get("_path", ""))
    gb = to_group(STATE["b"], lb, STATE["b"].get("_path", ""))
    STATE["ga"], STATE["gb"] = ga, gb

    opts = dict(CMP.DEFAULT_OPTS)
    opts.update(body.get("opts") or {})
    warn_engine = ""
    if body.get("useNorm", True):
        # 정규화한 화면명·동작으로 견준다. 이름 짓는 버릇 차이를 덜 탄다.
        if "text" in getattr(CMP, "DEFAULT_OPTS", {}):
            opts["text"] = ["screen", "action"]
        else:
            # 옛 엔진은 애플리케이션명·세부업무명·단위프로세스명으로 고정돼 있다.
            # 그 자리에 정규화 값을 넣어 같은 효과를 낸다. 화면에 보이는 값은 그대로다.
            for g in (ga, gb):
                for r in g["rows"]:
                    r["app"], r["biz"] = "", ""
                    r["proc"] = r.get("screen", "") + " " + r.get("action", "")
            warn_engine = ("옛 comparator.py 라 유사도 열 선택을 쓸 수 없어 "
                           "임시 방법으로 견줬습니다. 새 엔진으로 바꾸십시오.")

    store = None
    mp = (body.get("mapping") or "").strip()
    if MAPPING and mp:
        store = MAPPING.Store.load(mp)
        store.labels = {"a": la, "b": lb}
        STATE["store"], STATE["storePath"] = store, mp

    try:
        res = CMP.compare(ga, gb, opts, store=store)
    except Exception as e:
        return jsonify({"ok": False, "msg": "대조 실패: %s" % e})

    STATE["result"] = res
    STATE["rows"] = make_rows(res, ga, gb)
    cp = load_ini()
    cp["main"]["label_a"], cp["main"]["label_b"] = la, lb
    save_ini(cp)
    warns = list(res.get("warns", []))
    if warn_engine:
        warns.insert(0, warn_engine)
    return jsonify({"ok": True, "labels": {"a": la, "b": lb},
                    "stats": tally(res, ga, gb), "rows": len(STATE["rows"]),
                    "note": res.get("dupNote", ""), "warns": warns,
                    "opts": {"threshold": opts.get("threshold"),
                             "text": opts.get("text"),
                             "many": opts.get("many")}})


@app.route("/api/rows", methods=["POST"])
def api_rows():
    rows = STATE["rows"]
    if rows is None:
        return jsonify({"ok": False, "msg": "먼저 대조를 실행하십시오."})
    body = request.json or {}
    want = set(body.get("status") or [])
    q = (body.get("q") or "").strip().lower()
    off, n = int(body.get("offset") or 0), min(int(body.get("limit") or 300), 1000)
    sel = rows
    if want:
        sel = [x for x in sel if x["status"] in want]
    if q:
        def hit(x):
            for s in ("a", "b"):
                v = x[s]
                if v and q in " ".join(str(v.get(k) or "") for k in
                                       ("src", "app", "biz", "proc", "domain",
                                        "screen", "action", "type", "ui")).lower():
                    return True
            return False
        sel = [x for x in sel if hit(x)]
    return jsonify({"ok": True, "total": len(sel), "rows": sel[off:off + n]})


@app.route("/api/export", methods=["POST"])
def api_export():
    if not STATE["result"]:
        return jsonify({"ok": False, "msg": "먼저 대조를 실행하십시오."})
    name = (request.json or {}).get("name") or ("FP대조_%s.xlsx"
                                                % time.strftime("%Y%m%d_%H%M%S"))
    path = os.path.join(out_dir(), name)
    try:
        CMP.export_xlsx(STATE["result"], path)     # 엔진의 엑셀 출력을 그대로 쓴다
    except Exception as e:
        return jsonify({"ok": False, "msg": "저장 실패: %s" % e})
    return jsonify({"ok": True, "path": path})


@app.route("/api/savejson", methods=["POST"])
def api_savejson():
    if not STATE["result"]:
        return jsonify({"ok": False, "msg": "먼저 대조를 실행하십시오."})
    name = "대조결과_%s.json" % time.strftime("%Y%m%d_%H%M%S")
    path = os.path.join(out_dir(), name)
    try:
        CMP.save_json(STATE["result"], path)
    except Exception as e:
        return jsonify({"ok": False, "msg": "저장 실패: %s" % e})
    return jsonify({"ok": True, "path": path})


@app.route("/api/alive", methods=["POST", "GET"])
def api_alive():
    LAST_PING[0] = time.time()
    U.CONNECTED[0] = True
    return jsonify({"ok": True})


@app.route("/api/shutdown", methods=["POST"])
def api_shutdown():
    threading.Timer(0.4, lambda: os._exit(0)).start()
    return jsonify({"ok": True})


def watchdog():
    """창을 닫으면 프로그램도 끈다. 신호가 20초 끊기면 닫힌 것으로 본다."""
    started = time.time()
    while True:
        time.sleep(5)
        if U.CONNECTED[0]:
            if time.time() - LAST_PING[0] > 20:
                os._exit(0)
        elif time.time() - started > 300:
            os._exit(0)


# ---------------------------------------------------------------- 화면

HTML = r"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8">
<title>FP 대조 (JSON)</title>
<style>
*{box-sizing:border-box}
body{font-family:"맑은 고딕",Malgun Gothic,system-ui,sans-serif;font-size:13px;
     margin:0;background:#f4f5f7;color:#222}
header{background:#37474F;color:#fff;padding:10px 16px;display:flex;
       align-items:center;gap:12px}
header h1{font-size:15px;margin:0;font-weight:600}
header .sp{flex:1}
button{font-family:inherit;font-size:12px;padding:5px 12px;border:1px solid #bbb;
       background:#fff;border-radius:4px;cursor:pointer}
button:hover{background:#eef2f6}
button.primary{background:#1565C0;color:#fff;border-color:#1565C0}
button.primary:hover{background:#0d47a1}
button:disabled{opacity:.45;cursor:default}
.wrap{padding:12px 16px}
.card{background:#fff;border:1px solid #dcdfe3;border-radius:6px;padding:12px;
      margin-bottom:12px}
.card h2{font-size:13px;margin:0 0 10px;font-weight:600;color:#37474F}
.two{display:grid;grid-template-columns:1fr 1fr;gap:12px}
.pane{border:1px solid #dcdfe3;border-radius:6px;padding:10px}
.pane.a{background:#f2f7fc;border-color:#c5d9ea}
.pane.b{background:#fdf4ef;border-color:#eed4c3}
.pane h3{margin:0 0 8px;font-size:12px;font-weight:600;display:flex;
         align-items:center;gap:6px}
.pane .who{font-size:11px;color:#667;font-weight:400}
.row{display:flex;gap:6px;align-items:center;margin-bottom:6px}
input[type=text]{font-family:inherit;font-size:12px;padding:5px 7px;
                 border:1px solid #c3c7cc;border-radius:4px;width:100%}
.meta{font-size:11px;color:#445;line-height:1.6;background:#fff;
      border:1px solid #e2e6ea;border-radius:4px;padding:7px;margin-top:6px}
.meta b{color:#222}
table{border-collapse:collapse;width:100%;font-size:12px;background:#fff}
th,td{border:1px solid #dfe2e6;padding:3px 6px;white-space:nowrap;
      overflow:hidden;text-overflow:ellipsis}
th{background:#eceff1;font-weight:600;position:sticky;top:0;z-index:2}
th.a,td.a{background:#f2f7fc}
th.b,td.b{background:#fdf4ef}
.right{text-align:right}.center{text-align:center}
.scroll{max-height:60vh;overflow:auto;border:1px solid #dcdfe3;border-radius:6px}
.tag{display:inline-block;padding:1px 7px;border-radius:9px;font-size:11px;
     font-weight:600;white-space:nowrap}
.t-일치{background:#E2EFDA;color:#375623}
.t-변경{background:#FFF2CC;color:#8a6d00}
.t-A{background:#FDECEA;color:#C62828}
.t-B{background:#E3F2FD;color:#1565C0}
.t-미검토{background:#F2F2F2;color:#666}
.dim{color:#8a9099;font-size:11px}
.bar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:8px}
.chip{padding:3px 10px;border:1px solid #c3c7cc;border-radius:14px;cursor:pointer;
      background:#fff;font-size:12px;user-select:none}
.chip.on{background:#37474F;color:#fff;border-color:#37474F}
.tabs{display:flex;gap:4px;margin-bottom:8px}
.tab{padding:5px 14px;border:1px solid #dcdfe3;border-bottom:none;
     border-radius:5px 5px 0 0;background:#eceff1;cursor:pointer;font-size:12px}
.tab.on{background:#fff;font-weight:600}
.msg{padding:8px 10px;border-radius:4px;font-size:12px;margin-bottom:8px}
.msg.warn{background:#FFF8E1;border:1px solid #ffe082;color:#7a5c00}
.msg.err{background:#FDECEA;border:1px solid #f2b8b5;color:#B71C1C}
.num{font-variant-numeric:tabular-nums}
.picker{position:fixed;inset:0;background:rgba(0,0,0,.35);display:none;
        align-items:center;justify-content:center;z-index:50}
.picker .box{background:#fff;border-radius:8px;width:min(760px,92vw);
             max-height:80vh;display:flex;flex-direction:column;overflow:hidden}
.picker .hd{padding:10px 14px;border-bottom:1px solid #e2e6ea;display:flex;
            gap:8px;align-items:center}
.picker .bd{overflow:auto;padding:6px 0}
.picker .it{padding:6px 14px;cursor:pointer;display:flex;gap:10px;align-items:center}
.picker .it:hover{background:#eef2f6}
.picker .it.no{opacity:.42;cursor:default}
.picker .it .nm{flex:1;overflow:hidden;text-overflow:ellipsis}
</style></head><body>

<header>
  <h1>FP 대조 (JSON)</h1>
  <span class="dim" id="ver" style="color:#b0bec5"></span>
  <span class="sp"></span>
  <button id="btnCmp" class="primary" disabled>대조 실행</button>
  <button id="btnXls" disabled>엑셀로 내보내기</button>
  <button id="btnJs" disabled>결과 JSON 저장</button>
  <button id="btnEnd">닫기</button>
</header>

<div class="wrap">

<div class="card">
  <h2>① 대조할 JSON 두 개를 고르십시오</h2>
  <div class="two">
    <div class="pane a">
      <h3>왼쪽 <span class="who">기준이 됩니다. 이 자료의 행 차례를 그대로 지킵니다</span></h3>
      <div class="row">
        <input type="text" id="pathA" placeholder="설계단계 JSON 경로">
        <button data-pick="a">찾아보기</button>
        <button data-load="a">읽기</button>
      </div>
      <div class="row"><span class="dim" style="width:64px">표시 이름</span>
        <input type="text" id="labA" value="설계단계" style="max-width:180px"></div>
      <div class="meta" id="metaA">아직 읽지 않았습니다.</div>
    </div>
    <div class="pane b">
      <h3>오른쪽 <span class="who">왼쪽에 맞춰 붙입니다</span></h3>
      <div class="row">
        <input type="text" id="pathB" placeholder="종료단계 JSON 경로">
        <button data-pick="b">찾아보기</button>
        <button data-load="b">읽기</button>
      </div>
      <div class="row"><span class="dim" style="width:64px">표시 이름</span>
        <input type="text" id="labB" value="종료단계" style="max-width:180px"></div>
      <div class="meta" id="metaB">아직 읽지 않았습니다.</div>
    </div>
  </div>
  <div class="bar" style="margin-top:10px">
    <label><input type="checkbox" id="useNorm" checked> 정규화한 화면명·동작으로 견주기</label>
    <span class="dim">|</span>
    <label>유사도 기준 <input type="text" id="th" value="0.72" style="width:56px"></label>
    <label><input type="checkbox" id="many" checked> 왼쪽 여러 행이 오른쪽 한 행에 붙는 것 허용</label>
    <span class="dim">|</span>
    <label style="flex:1;display:flex;gap:6px;align-items:center">수동 짝짓기 파일
      <input type="text" id="mp" placeholder="(비우면 쓰지 않음)"></label>
  </div>
</div>

<div id="out" style="display:none">
  <div class="card">
    <h2>② 집계</h2>
    <div id="warn"></div>
    <div class="tabs">
      <div class="tab on" data-t="s">상태별</div>
      <div class="tab" data-t="t">FP유형별</div>
      <div class="tab" data-t="d">분야별</div>
      <div class="tab" data-t="r">짝지은 경로별</div>
    </div>
    <div id="sum"></div>
  </div>

  <div class="card">
    <h2>③ 대조표 <span class="dim" id="cnt"></span></h2>
    <div class="bar" id="chips"></div>
    <div class="bar">
      <input type="text" id="q" style="max-width:300px"
             placeholder="화면명·업무명·파일명으로 찾기">
      <button id="btnMore">더 보기</button>
    </div>
    <div class="scroll"><table id="tbl"></table></div>
  </div>
</div>
</div>

<div class="picker" id="pick">
  <div class="box">
    <div class="hd">
      <b id="pkTitle">파일 고르기</b>
      <input type="text" id="pkDir" style="flex:1">
      <button id="pkGo">이동</button>
      <button id="pkUp">상위</button>
      <button id="pkX">닫기</button>
    </div>
    <div class="bd" id="pkList"></div>
  </div>
</div>

<script>
var R=null,OFF=0,LIM=300,ROWS=[],SIDE='a',LAB={a:'설계단계',b:'종료단계'};
function j(u,b){return fetch(u,{method:b?'POST':'GET',
  headers:{'Content-Type':'application/json'},body:b?JSON.stringify(b):null})
  .then(function(r){return r.json();});}
function esc(t){return String(t==null?'':t).replace(/[&<>"]/g,function(c){
  return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function n1(v){return v==null||v===''?'':(Math.round(v*10)/10).toLocaleString();}

j('/api/init').then(function(d){
  document.getElementById('ver').textContent='v'+d.version+' · 엔진 v'+d.engine;
  document.getElementById('pkDir').value=d.dir;
  if(d.last.a)document.getElementById('pathA').value=d.last.a;
  if(d.last.b)document.getElementById('pathB').value=d.last.b;
  document.getElementById('labA').value=d.labels.a;
  document.getElementById('labB').value=d.labels.b;
  document.getElementById('th').value=d.opts.threshold;
  if(d.warn){var e=document.createElement('div');e.className='msg err';
    e.style.margin='12px 16px 0';e.textContent=d.warn;
    document.body.insertBefore(e,document.querySelector('.wrap'));}
});

/* ---- 파일 고르기 ---- */
document.querySelectorAll('[data-pick]').forEach(function(b){
  b.onclick=function(){SIDE=b.dataset.pick;
    document.getElementById('pkTitle').textContent=
      (SIDE==='a'?'왼쪽(기준)':'오른쪽')+' JSON 고르기';
    document.getElementById('pick').style.display='flex';browse();};});
document.getElementById('pkX').onclick=function(){
  document.getElementById('pick').style.display='none';};
document.getElementById('pkGo').onclick=function(){browse();};
document.getElementById('pkUp').onclick=function(){
  browse(document.getElementById('pkDir').dataset.up||'');};
function browse(dir){
  j('/api/browse',{dir:dir||document.getElementById('pkDir').value}).then(function(d){
    if(!d.ok){alert(d.msg);return;}
    var e=document.getElementById('pkDir');e.value=d.dir;e.dataset.up=d.up;
    var h=d.dirs.map(function(x){
      return '<div class="it" data-d="'+esc(x.path)+'">📁 <span class="nm">'+
             esc(x.name)+'</span></div>';}).join('');
    h+=d.files.map(function(x){
      return '<div class="it'+(x.ok?'':' no')+'" '+(x.ok?'data-f="'+esc(x.path)+'"':'')+'>'+
        (x.ok?'📄':'⛔')+' <span class="nm">'+esc(x.name)+'</span>'+
        '<span class="dim">'+(x.stage?esc(x.stage)+' '+esc(x.asOf)+' · ':'')+
        (x.size/1e6).toFixed(1)+'MB'+(x.ok?'':' · 읽을 수 없는 형식')+'</span></div>';
      }).join('');
    document.getElementById('pkList').innerHTML=h||'<div class="it no">비어 있습니다</div>';
    document.getElementById('pkList').querySelectorAll('[data-d]').forEach(function(x){
      x.onclick=function(){browse(x.dataset.d);};});
    document.getElementById('pkList').querySelectorAll('[data-f]').forEach(function(x){
      x.onclick=function(){
        document.getElementById('path'+SIDE.toUpperCase()).value=x.dataset.f;
        document.getElementById('pick').style.display='none';load(SIDE);};});
  });
}

/* ---- 읽기 ---- */
document.querySelectorAll('[data-load]').forEach(function(b){
  b.onclick=function(){load(b.dataset.load);};});
function load(side){
  var S=side.toUpperCase();
  j('/api/load',{side:side,path:document.getElementById('path'+S).value})
   .then(function(d){
    var el=document.getElementById('meta'+S);
    if(!d.ok){el.innerHTML='<span style="color:#B71C1C">'+esc(d.msg)+'</span>';return;}
    var i=d.info,c=i.counts;
    if(i.stage)document.getElementById('lab'+S).value=i.stage;
    var h='<b>'+esc(i.id)+'</b>';
    if(i.asOf)h+=' · '+esc(i.asOf);
    h+='<br>기능점수 <b>'+c.fp+'행</b> · <b>'+n1(c.fpTotal)+'FP</b>'+
       ' · 화면 '+c.screens+'개 · 테이블 '+c.tables+'개';
    h+='<br><span class="dim">'+i.sources.filter(function(s){return s.kind==='fp';})
       .map(function(s){return esc(s.file)+' ('+s.n+'행'+(s.method?', '+esc(s.method):'')+')';})
       .join('<br>')+'</span>';
    if(i.crowdedKeys.length)h+='<br><span class="dim">몰리는 열쇠: '+
       i.crowdedKeys.map(function(k){return esc(k.k)+'('+k.max+')';}).join(', ')+'</span>';
    el.innerHTML=h;
    ready();
  });
}
function ready(){
  var ok=document.getElementById('metaA').textContent.indexOf('아직')<0 &&
         document.getElementById('metaB').textContent.indexOf('아직')<0;
  document.getElementById('btnCmp').disabled=!ok;
}

/* ---- 대조 ---- */
document.getElementById('btnCmp').onclick=function(){
  var me=this;me.disabled=true;me.textContent='맞추는 중…';
  LAB={a:document.getElementById('labA').value||'A',
       b:document.getElementById('labB').value||'B'};
  j('/api/compare',{labels:LAB,useNorm:document.getElementById('useNorm').checked,
     mapping:document.getElementById('mp').value,
     opts:{threshold:parseFloat(document.getElementById('th').value)||0.72,
           many:document.getElementById('many').checked}})
   .then(function(d){
    me.disabled=false;me.textContent='대조 실행';
    if(!d.ok){alert(d.msg);return;}
    R=d;LAB=d.labels;
    document.getElementById('out').style.display='';
    document.getElementById('btnXls').disabled=false;
    document.getElementById('btnJs').disabled=false;
    var w='';
    if(d.note)w+='<div class="msg warn">'+esc(d.note)+'</div>';
    (d.warns||[]).forEach(function(x){w+='<div class="msg warn">'+esc(x)+'</div>';});
    document.getElementById('warn').innerHTML=w;
    chips();tab('s');fetchRows(true);
  });
};
function chips(){
  var order=['일치','변경','A만 있음','B만 있음','미검토'];
  document.getElementById('chips').innerHTML=order.map(function(s){
    var t=s.replace('A만 있음',LAB.a+'에만 있음').replace('B만 있음',LAB.b+'에만 있음');
    return '<span class="chip on" data-s="'+esc(s)+'">'+esc(t)+'</span>';}).join('');
  document.querySelectorAll('.chip').forEach(function(c){
    c.onclick=function(){c.classList.toggle('on');fetchRows(true);};});
}
document.querySelectorAll('.tab').forEach(function(t){
  t.onclick=function(){
    document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('on');});
    t.classList.add('on');tab(t.dataset.t);};});

function tag(st){
  var c=({'일치':'t-일치','변경':'t-변경','미검토':'t-미검토'})[st]||
        (st.indexOf('A만')===0?'t-A':'t-B');
  var t=st.replace('A만 있음',LAB.a+'에만 있음').replace('B만 있음',LAB.b+'에만 있음');
  return '<span class="tag '+c+'">'+esc(t)+'</span>';
}
function tab(k){
  var s=R.stats,t=s.total,h='';
  h+='<table style="margin-bottom:10px"><tr><th></th><th class="a">'+esc(LAB.a)+
     ' (기준)</th><th class="b">'+esc(LAB.b)+'</th><th>차이</th></tr>'+
     '<tr><th>기능 수</th><td class="right a num">'+t.nA+'</td><td class="right b num">'+
     t.nB+'</td><td class="right num">'+(t.nB-t.nA>0?'+':'')+(t.nB-t.nA)+'</td></tr>'+
     '<tr><th>기능점수</th><td class="right a num">'+n1(t.fpA)+'</td><td class="right b num">'+
     n1(t.fpB)+'</td><td class="right num"><b>'+(t.fpDiff>0?'+':'')+n1(t.fpDiff)+
     '</b></td></tr>'+
     '<tr><th>짝지은 행</th><td class="right a num">'+t.matchedA+' ('+t.rateA+'%)</td>'+
     '<td class="right b num">'+t.matchedB+' ('+t.rateB+'%)</td><td></td></tr></table>';
  if(k==='s'){
    h+='<table><tr><th>비고</th><th class="right">묶음</th><th class="right a">'+
       esc(LAB.a)+' FP</th><th class="right b">'+esc(LAB.b)+' FP</th></tr>';
    s.status.forEach(function(x){h+='<tr><td>'+tag(x.k)+'</td><td class="right num">'+
      x.n+'</td><td class="right a num">'+n1(x.fpA)+'</td><td class="right b num">'+
      n1(x.fpB)+'</td></tr>';});
    h+='</table>';
  }else if(k==='r'){
    h+='<table><tr><th>짝지은 경로</th><th class="right">묶음</th></tr>';
    s.byRoute.forEach(function(x){h+='<tr><td>'+esc(x.k)+'</td><td class="right num">'+
      x.n+'</td></tr>';});
    h+='</table>';
  }else{
    var d=(k==='t'?s.byType:s.byDomain),nm=(k==='t'?'FP유형':'분야');
    h+='<table><tr><th>'+nm+'</th><th class="right a">'+esc(LAB.a)+' 건</th>'+
       '<th class="right a">'+esc(LAB.a)+' FP</th><th class="right b">'+esc(LAB.b)+
       ' 건</th><th class="right b">'+esc(LAB.b)+' FP</th><th class="right">건 차이</th>'+
       '<th class="right">FP 차이</th></tr>';
    d.forEach(function(x){h+='<tr><td>'+esc(x.k)+'</td><td class="right a num">'+x.nA+
      '</td><td class="right a num">'+n1(x.fpA)+'</td><td class="right b num">'+x.nB+
      '</td><td class="right b num">'+n1(x.fpB)+'</td><td class="right num">'+
      (x.nDiff>0?'+':'')+x.nDiff+'</td><td class="right num"><b>'+
      (x.fpDiff>0?'+':'')+n1(x.fpDiff)+'</b></td></tr>';});
    h+='</table>';
  }
  document.getElementById('sum').innerHTML=h;
}

document.getElementById('q').oninput=function(){
  clearTimeout(window._t);window._t=setTimeout(function(){fetchRows(true);},250);};
document.getElementById('btnMore').onclick=function(){fetchRows(false);};
function want(){return Array.prototype.filter.call(
  document.querySelectorAll('.chip'),function(c){return c.classList.contains('on');})
  .map(function(c){return c.dataset.s;});}
function fetchRows(reset){
  if(reset){OFF=0;ROWS=[];}
  j('/api/rows',{status:want(),q:document.getElementById('q').value,
                 offset:OFF,limit:LIM}).then(function(d){
    if(!d.ok){alert(d.msg);return;}
    ROWS=ROWS.concat(d.rows);OFF+=d.rows.length;
    document.getElementById('cnt').textContent='— '+ROWS.length.toLocaleString()+
      ' / '+d.total.toLocaleString()+'줄';
    document.getElementById('btnMore').disabled=(OFF>=d.total);
    draw();
  });
}
function draw(){
  var h='<tr><th colspan="4" class="center">판정</th>'+
        '<th colspan="9" class="center a">'+esc(LAB.a)+' (기준)</th>'+
        '<th colspan="9" class="center b">'+esc(LAB.b)+'</th>'+
        '<th class="center">차이</th></tr><tr>'+
        '<th style="width:118px">비고</th><th style="width:88px">묶음</th>'+
        '<th style="width:120px">짝지은 근거</th><th style="width:52px">유사도</th>';
  ['a','b'].forEach(function(s){
    ['파일','행','분야','화면명','동작','FP유형','RET','DET','가중치'].forEach(
      function(x,i){var w=[150,44,54,178,58,54,42,42,52][i];
        h+='<th class="'+s+'" style="width:'+w+'px">'+x+'</th>';});});
  h+='<th style="width:62px">FP 차이</th></tr>';
  ROWS.forEach(function(r){
    h+='<tr><td class="center">'+(r.head?tag(r.status):'')+'</td>'+
       '<td class="center dim">'+(r.head?esc(r.gid)+
         (r.size[0]>1||r.size[1]>1?' ('+r.size[0]+':'+r.size[1]+')':''):'')+'</td>'+
       '<td class="dim" title="'+esc(r.basis||'')+'">'+(r.head?esc(r.by):'')+'</td>'+
       '<td class="right dim num">'+(r.head&&r.score?n1(r.score*1000)/1000:'')+'</td>';
    ['a','b'].forEach(function(s){
      var v=r[s];
      if(!v){h+='<td class="'+s+'" colspan="9"></td>';return;}
      h+='<td class="'+s+'" title="'+esc(v.src)+'">'+esc(v.src)+'</td>'+
         '<td class="'+s+' right num">'+esc(v.no)+'</td>'+
         '<td class="'+s+'">'+esc(v.domain)+'</td>'+
         '<td class="'+s+'" title="'+esc(v.proc||'')+' / '+esc(v.biz||'')+'">'+
           esc(v.screen||v.proc)+'</td>'+
         '<td class="'+s+' center">'+esc(v.action)+'</td>'+
         '<td class="'+s+' center">'+esc(v.type)+'</td>'+
         '<td class="'+s+' right num">'+esc(v.ret)+'</td>'+
         '<td class="'+s+' right num">'+esc(v.det)+'</td>'+
         '<td class="'+s+' right num">'+n1(v.wt)+'</td>';});
    h+='<td class="right num">'+(r.head&&r.fpDiff?((r.fpDiff>0?'+':'')+n1(r.fpDiff)):'')+
       '</td></tr>';});
  document.getElementById('tbl').innerHTML=h;
}

document.getElementById('btnXls').onclick=function(){
  var me=this;me.disabled=true;me.textContent='만드는 중…';
  j('/api/export',{}).then(function(d){me.disabled=false;
    me.textContent='엑셀로 내보내기';
    alert(d.ok?('저장했습니다\n'+d.path):d.msg);});};
document.getElementById('btnJs').onclick=function(){
  j('/api/savejson',{}).then(function(d){
    alert(d.ok?('저장했습니다\n'+d.path):d.msg);});};
document.getElementById('btnEnd').onclick=function(){
  if(confirm('닫습니까?'))j('/api/shutdown',{}).then(function(){
    document.body.innerHTML='<p style="padding:30px">닫았습니다. 창을 닫으십시오.</p>';});};
/* 창이 살아 있음을 알린다. 이 신호가 끊기면 프로그램이 스스로 끝난다. */
setInterval(function(){fetch('/api/alive',{method:'POST'});},4000);
fetch('/api/alive',{method:'POST'});
</script></body></html>
"""


# ---------------------------------------------------------------- 실행

def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    no_browser = "--no-browser" in sys.argv
    cp = load_ini()
    for side, i in (("a", 0), ("b", 1)):
        if len(args) > i and os.path.exists(args[i]):
            try:
                STATE[side] = read_json(args[i])
                cp["main"][side] = args[i]
                cp["main"]["dir"] = os.path.dirname(args[i])
                print("%s ← %s" % (side.upper(), os.path.basename(args[i])))
            except Exception as e:
                print("읽기 실패 %s : %s" % (args[i], e))
    save_ini(cp)

    port = U.pick_port()
    url = "http://127.0.0.1:%d/" % port
    print("=" * 60)
    print("FP 대조 (JSON) v%s" % VERSION)
    print("  comparator.py    : v%s" % getattr(CMP, "VERSION", "?"))
    print("  calculator_ui.py : v%s" % getattr(U, "VERSION", "?"))
    print("  mapping.py       : %s" % ("있음" if MAPPING else "없음"))
    bad = engine_check()
    if bad:
        print("!" * 60)
        print(bad)
        print("!" * 60)
    print("화면 주소:", url)
    print("창이 뜨지 않으면 브라우저에 위 주소를 직접 입력하십시오.")
    print("=" * 60)

    threading.Thread(target=watchdog, daemon=True).start()
    if not no_browser:
        # 크롬을 --app 으로 띄워 별도 창으로 만든다. 다른 화면들과 같은 방식이며,
        # 프로필 폴더를 따로 두어 이미 열려 있는 크롬 탭에 끼어들지 않게 한다.
        threading.Thread(
            target=lambda: (time.sleep(1.0),
                            U.open_browser(url, ".chrome_comparator_json")),
            daemon=True).start()

    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app.run(host="127.0.0.1", port=port, debug=False,
            use_reloader=False, threaded=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
