# -*- coding: utf-8 -*-
"""
comparator_ui.py  v1.0
단계별 FP 산정 결과물 대조 화면

  comparator.py       대조 엔진 (읽기·짝짓기·차이 판정·엑셀 내보내기)
  comparator_ui.py    화면과 웹서버 (이 파일)
  calculator_ui.py    양식 사양과 FP 계산을 빌려 쓴다

실행
  python comparator_ui.py [--no-browser]
"""

import os
import sys
import time
import subprocess
import threading

from flask import Flask, request, jsonify, Response

# ── 폴더 구조 대응 ────────────────────────────────────────────
# FP-Checker/Calculator, FP-Checker/Comparator 처럼 나뉘어 있어도 서로 불러온다.
# 프로젝트 루트 자체는 넣지 않는다. 루트에 예전 사본이 남아 있으면
# 같은 이름이 먼저 읽혀 옛 코드가 도는 일이 생기기 때문이다.
_HERE = os.path.dirname(os.path.abspath(__file__))
_PKGS = ("Calculator", "Comparator", "DocParser")
for _p in [os.path.join(os.path.dirname(_HERE), _n) for _n in _PKGS] + [_HERE]:
    if os.path.isdir(_p):
        if _p in sys.path:
            sys.path.remove(_p)
        sys.path.insert(0, _p)


def project_root():
    """설정 파일을 둘 위치. 하위 폴더로 나뉘어 있으면 그 위(프로젝트 루트)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(_HERE) if os.path.basename(_HERE) in _PKGS else _HERE


import calculator_ui as U
import comparator as CMP

APP_NAME = "FP 산정 결과물 대조"
VERSION = "1.0"

INI_PATH = os.path.join(project_root(), "comparator.ini")

NEED_CMP = ("KEY_FIELDS", "DIFF_FIELDS", "DEFAULT_OPTS", "load_group",
            "compare", "build", "save_json", "load_json", "Cancelled")
NEED_U = ("FORMS", "is_ready", "form_list", "guess_form")
CMP_MIN = "1.1"          # 이 화면이 요구하는 comparator.py 최소 판


class _NoCancel(Exception):
    """엔진이 예전 것이라 Cancelled 가 없을 때 자리를 메운다."""


def cancelled_exc():
    """취소 예외를 쓰는 시점에 가져온다(불러오는 시점에 건드리지 않는다)."""
    return getattr(CMP, "Cancelled", _NoCancel)


def _ver_tuple(v):
    out = []
    for p in str(v).split("."):
        out.append(int(p) if p.isdigit() else 0)
    return tuple(out)


def module_check():
    """짝이 맞는 파일을 물고 있는지 본다. 같은 이름의 예전 사본이 있으면 여기서 걸린다."""
    miss = [n for n in NEED_CMP if not hasattr(CMP, n)]
    miss += [n for n in NEED_U if not hasattr(U, n)]
    if not miss and _ver_tuple(getattr(CMP, "VERSION", "0")) < _ver_tuple(CMP_MIN):
        miss.append("comparator.py 판 %s 이상 필요 (지금 %s)"
                    % (CMP_MIN, getattr(CMP, "VERSION", "?")))
    info = {"comparator": getattr(CMP, "__file__", "?"),
            "calculator_ui": getattr(U, "__file__", "?"),
            "comparatorVersion": getattr(CMP, "VERSION", "?"),
            "calculatorVersion": getattr(U, "VERSION", "?"),
            "missing": miss}
    if miss:
        info["msg"] = ("엔진 파일이 화면과 짝이 맞지 않습니다. 없는 항목: %s\n"
                       "  comparator.py    : %s\n"
                       "  calculator_ui.py : %s\n"
                       "같은 이름의 예전 파일이 다른 폴더에 남아 있는지 확인하십시오."
                       % (", ".join(miss), info["comparator"], info["calculator_ui"]))
    return info

app = Flask(__name__)
LAST_PING = [time.time()]
RESULT = {"data": None}
JOB = {"running": False, "stage": "", "done": 0, "total": 0, "detail": "",
       "error": None, "cancel": False, "started": 0.0, "elapsed": 0.0,
       "finished": False, "version": 0, "json": ""}
PARTIAL = {"data": None}
ALLOWED = set()          # 이번 화면에서 읽은 파일만 열어 준다


def remember(data):
    for side in ("a", "b"):
        for f in (data.get(side) or {}).get("files", []) or []:
            if f.get("path"):
                ALLOWED.add(os.path.abspath(f["path"]))


def out_dir():
    d = os.path.join(project_root(), "out")
    return d if os.path.isdir(d) else project_root()


def load_ini():
    import configparser
    cp = configparser.ConfigParser()
    cp.read(INI_PATH, encoding="utf-8")
    if "main" not in cp:
        cp["main"] = {}
    return cp


def save_ini(cp):
    try:
        with open(INI_PATH, "w", encoding="utf-8") as f:
            cp.write(f)
    except Exception:
        pass


@app.after_request
def no_cache(resp):
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/")
def index():
    return Response(HTML, mimetype="text/html; charset=utf-8")


@app.route("/api/init")
def api_init():
    chk = module_check()
    if chk["missing"]:
        return jsonify({"ok": False, "msg": chk["msg"], "modules": chk})
    try:
        return jsonify(_init_payload(chk))
    except Exception as e:
        return jsonify({"ok": False, "msg": "시작 정보를 만들지 못했습니다: %s" % e,
                        "modules": chk})


def _init_payload(chk):
    cp = load_ini()
    forms = [f for f in U.form_list(list(U.FORMS)) if f["ready"]]
    keys = [{"k": k, "label": v[1]} for k, v in CMP.KEY_FIELDS.items()]
    dups = [{"k": k, "label": v} for k, v in CMP.DUP_KEYS.items()] + \
           [{"k": k, "label": v[1]} for k, v in CMP.KEY_FIELDS.items()]
    diffs = [{"k": k, "label": lb} for k, lb in CMP.DIFF_FIELDS]
    return {"ok": True, "app": APP_NAME, "version": VERSION, "forms": forms,
            "keys": keys, "dups": dups, "diffs": diffs, "opts": CMP.DEFAULT_OPTS,
            "modules": chk,
            "lastDirs": {"a": last_dir("a"), "b": last_dir("b")},
            "labels": [cp["main"].get("label_a", "설계단계"),
                       cp["main"].get("label_b", "종료단계")]}


def last_dir(side):
    """그룹별로 마지막에 본 폴더. 없으면 반대쪽 것, 그것도 없으면 현재 폴더."""
    m = load_ini()["main"]
    other = "b" if side == "a" else "a"
    for k in ("last_dir_%s" % side, "last_dir_%s" % other, "last_dir"):
        v = m.get(k, "")
        if v and os.path.isdir(v):
            return v
    return os.getcwd()


@app.route("/api/list", methods=["POST"])
def api_list():
    body = request.json or {}
    side = body.get("side") if body.get("side") in ("a", "b") else "a"
    d = body.get("dir", "").strip().strip('"')
    if not d:
        d = last_dir(side)
    if not os.path.isdir(d):
        return jsonify({"ok": False, "msg": "폴더가 아닙니다: %s" % d})
    dirs, files = [], []
    try:
        for n in sorted(os.listdir(d)):
            if n.startswith("."):
                continue
            full = os.path.join(d, n)
            if os.path.isdir(full):
                dirs.append(n)
            elif os.path.splitext(n)[1].lower() in (".xlsx", ".xlsm"):
                files.append(n)
    except Exception as e:
        return jsonify({"ok": False, "msg": "폴더를 읽을 수 없습니다: %s" % e})
    cp = load_ini()
    cp["main"]["last_dir_%s" % side] = os.path.abspath(d)
    save_ini(cp)
    return jsonify({"ok": True, "side": side, "dir": os.path.abspath(d),
                    "parent": os.path.dirname(os.path.abspath(d)),
                    "dirs": dirs[:300], "files": files[:300], "sep": os.sep})


@app.route("/api/browse", methods=["POST"])
def api_browse():
    body = request.json or {}
    mode = body.get("mode", "open")
    side = body.get("side") if body.get("side") in ("a", "b") else "a"
    start = last_dir(side)
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ft = [("엑셀 파일", "*.xlsm *.xlsx"), ("모든 파일", "*.*")]
        if mode == "json":
            p = filedialog.askopenfilename(
                filetypes=[("대조 결과 JSON", "*.json"), ("모든 파일", "*.*")],
                initialdir=out_dir())
            root.destroy()
            return jsonify({"ok": True, "paths": [p] if p else []})
        if mode == "save":
            p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=ft,
                                             initialfile="FP대조결과.xlsx")
            out = [p] if p else []
        else:
            out = list(filedialog.askopenfilenames(filetypes=ft, initialdir=start))
            if out:
                cp = load_ini()
                cp["main"]["last_dir_%s" % side] = os.path.dirname(out[0])
                save_ini(cp)
        root.destroy()
        return jsonify({"ok": True, "paths": out})
    except Exception as e:
        return jsonify({"ok": False, "msg": "파일 선택 창을 열 수 없습니다(%s). "
                                            "폴더에서 고르기를 쓰십시오." % e})


@app.route("/api/guess", methods=["POST"])
def api_guess():
    """양식만 찍는 것이 아니라 실제로 열어 보고 무엇으로 읽히는지 돌려준다."""
    out = {}
    for p in (request.json or {}).get("paths", []):
        form = None
        try:
            form = U.guess_form(p)
            info = CMP.peek(p, form)
            info["ok"] = True
            out[p] = info
        except Exception as e:
            out[p] = {"ok": False, "form": form, "msg": str(e)}
    return jsonify({"ok": True, "info": out,
                    "forms": {k: v.get("form") for k, v in out.items()}})


def _run_compare(body):
    """배경에서 대조를 돌린다.
    진행 상황은 JOB 에, 지금까지 나온 결과는 PARTIAL 과 JSON 파일에 적는다."""
    def progress(stage, done, total, detail):
        JOB.update({"stage": stage, "done": done, "total": total,
                    "detail": detail,
                    "elapsed": round(time.time() - JOB["started"], 1)})

    def cancelled():
        return JOB["cancel"]

    def keep(res):
        PARTIAL["data"] = res
        JOB["version"] += 1
        try:
            CMP.save_json(res, JOB["json"])
        except Exception:
            pass

    def empty(side):
        return {"label": body.get(side, {}).get("label") or side.upper(),
                "files": [], "rows": [], "errors": []}

    state = {"ga": None, "gb": None}
    CANCELLED = cancelled_exc()

    def snap_files(side):
        """파일 하나를 읽을 때마다 지금까지 읽은 것으로 스냅숏을 만든다."""
        def fn(out):
            state[side] = out          # 읽는 중인 그룹을 그대로 들여다본다
            keep(CMP.build(body.get("opts") or CMP.DEFAULT_OPTS,
                           state["ga"] or empty("a"), state["gb"] or empty("b"),
                           [], {}, {}, False))
        return fn

    try:
        ga_in, gb_in = body.get("a", {}), body.get("b", {})
        JOB["json"] = os.path.join(
            out_dir(), "대조_%s.json" % time.strftime("%Y%m%d_%H%M%S"))
        ga = CMP.load_group(ga_in["files"], ga_in.get("label") or "A",
                            progress=progress, cancel=cancelled,
                            on_file=snap_files("ga"))
        state["ga"] = ga
        gb = CMP.load_group(gb_in["files"], gb_in.get("label") or "B",
                            progress=progress, cancel=cancelled,
                            on_file=snap_files("gb"))
        state["gb"] = gb
        res = CMP.compare(ga, gb, body.get("opts") or {},
                          progress=progress, cancel=cancelled, on_stage=keep)
        RESULT["data"] = res
        remember(res)
        keep(res)
        cp = load_ini()
        cp["main"]["label_a"] = ga["label"]
        cp["main"]["label_b"] = gb["label"]
        save_ini(cp)
        JOB["stage"] = "끝"
    except CANCELLED:
        JOB["error"] = "취소했습니다. 지금까지 분석한 것은 그대로 볼 수 있습니다."
        if PARTIAL["data"]:
            RESULT["data"] = PARTIAL["data"]
    except AttributeError as e:
        JOB["error"] = ("엔진 파일이 화면과 짝이 맞지 않습니다 (%s). "
                        "comparator.py 를 함께 바꾸십시오." % e)
    except Exception as e:
        JOB["error"] = "대조 실패: %s" % e
        if PARTIAL["data"]:
            RESULT["data"] = PARTIAL["data"]
    finally:
        JOB["elapsed"] = round(time.time() - JOB["started"], 1)
        JOB["running"] = False
        JOB["finished"] = True


@app.route("/api/compare", methods=["POST"])
def api_compare():
    body = request.json or {}
    chk = module_check()
    if chk["missing"]:
        return jsonify({"ok": False, "msg": chk["msg"], "modules": chk})
    if not body.get("a", {}).get("files") or not body.get("b", {}).get("files"):
        return jsonify({"ok": False, "msg": "양쪽 그룹에 파일을 넣으십시오."})
    if JOB["running"]:
        return jsonify({"ok": False, "msg": "이미 대조가 돌고 있습니다."})
    n = len(body["a"]["files"]) + len(body["b"]["files"])
    JOB.update({"running": True, "stage": "여는 중", "done": 0, "total": n,
                "detail": "", "error": None, "cancel": False,
                "started": time.time(), "elapsed": 0.0, "finished": False,
                "version": 0, "json": ""})
    RESULT["data"] = None
    PARTIAL["data"] = None
    threading.Thread(target=_run_compare, args=(body,), daemon=True).start()
    return jsonify({"ok": True, "started": True})


@app.route("/api/progress")
def api_progress():
    out = {k: JOB[k] for k in ("running", "stage", "done", "total", "detail",
                               "error", "elapsed", "finished", "version", "json")}
    try:
        seen = int(request.args.get("v", -1))
    except ValueError:
        seen = -1
    if PARTIAL["data"] is not None and seen != JOB["version"]:
        out["data"] = PARTIAL["data"]      # 지금까지 나온 것까지
    return jsonify(out)


@app.route("/api/loadjson", methods=["POST"])
def api_loadjson():
    path = (request.json or {}).get("path", "").strip().strip('"')
    if not path or not os.path.exists(path):
        return jsonify({"ok": False, "msg": "파일을 찾을 수 없습니다."})
    try:
        data = CMP.load_json(path)
    except Exception as e:
        return jsonify({"ok": False, "msg": "읽기 실패: %s" % e})
    RESULT["data"] = data
    remember(data)
    return jsonify({"ok": True, "data": data, "name": os.path.basename(path)})


@app.route("/api/cancel", methods=["POST"])
def api_cancel():
    JOB["cancel"] = True
    return jsonify({"ok": True})


@app.route("/api/export", methods=["POST"])
def api_export():
    if not RESULT["data"]:
        return jsonify({"ok": False, "msg": "먼저 대조를 실행하십시오."})
    path = (request.json or {}).get("path", "").strip().strip('"')
    if not path:
        return jsonify({"ok": False, "msg": "저장 경로가 없습니다."})
    try:
        CMP.export_xlsx(RESULT["data"], path)
    except Exception as e:
        return jsonify({"ok": False, "msg": "저장 실패: %s" % e})
    return jsonify({"ok": True, "path": path, "name": os.path.basename(path)})


@app.route("/api/openfile", methods=["POST"])
def api_openfile():
    """대조에 쓴 엑셀 파일을 기본 프로그램으로 연다."""
    path = (request.json or {}).get("path", "").strip().strip('"')
    full = os.path.abspath(path) if path else ""
    if not full or full not in ALLOWED:
        return jsonify({"ok": False, "msg": "이번 대조에 쓴 파일만 열 수 있습니다."})
    if not os.path.exists(full):
        return jsonify({"ok": False, "msg": "파일이 없어졌습니다: %s" % full})
    try:
        if os.name == "nt":
            os.startfile(full)                                   # noqa: S606
        elif sys.platform == "darwin":
            subprocess.Popen(["open", full])
        else:
            subprocess.Popen(["xdg-open", full])
    except Exception as e:
        return jsonify({"ok": False, "msg": "열지 못했습니다: %s" % e})
    return jsonify({"ok": True, "name": os.path.basename(full)})


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
<title>FP 산정 결과물 대조</title>
<style>
*{box-sizing:border-box}
body{margin:0;font-family:"맑은 고딕","Malgun Gothic",sans-serif;font-size:13px;color:#222;background:#f4f5f7}
#bar{display:flex;align-items:center;gap:6px;padding:8px 10px;background:#37474f;color:#fff}
#bar b{font-size:14px;margin-right:8px}
#state{flex:1;color:#cfd8dc;font-size:12px}
button{font-family:inherit;font-size:12px;padding:5px 10px;border:1px solid #bbb;background:#fff;border-radius:3px;cursor:pointer}
button:hover{background:#eef3f8}
button.main{background:#37474f;color:#fff;border-color:#37474f;padding:7px 18px;font-size:13px}
button.main:hover{background:#455a64}
#tabs{display:flex;gap:2px;padding:8px 10px 0}
.tab{padding:7px 16px;border:1px solid #ccc;border-bottom:none;background:#e4e7ea;border-radius:4px 4px 0 0;cursor:pointer}
.tab.on{background:#fff;font-weight:700}
.pane{display:none;background:#fff;border:1px solid #ccc;margin:0 10px 10px;padding:12px}
.pane.on{display:block}
h3{font-size:13px;margin:0 0 8px;padding-left:6px;border-left:4px solid #37474f}
.cols{display:flex;gap:12px}
.col{flex:1;border:1px solid #ddd;border-radius:4px;padding:10px;background:#fafbfc}
.col h4{margin:0 0 8px;font-size:13px}
.col input.nm{width:100%;padding:5px;border:1px solid #bbb;border-radius:3px;font-family:inherit}
.files{margin-top:8px;min-height:60px}
.file{display:flex;gap:6px;align-items:center;padding:4px 6px;border:1px solid #e0e0e0;
  background:#fff;border-radius:3px;margin:3px 0;font-size:12px}
.file .n{flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.file select{font-size:11px;padding:2px;border:1px solid #ccc}
.file .x{color:#c0392b;cursor:pointer;font-weight:700;padding:0 4px}
.fbox{margin:3px 0}
.fbox .file{margin:0}
.sub{font-size:11px;color:#666;padding:2px 6px 0 8px}
.sub.warn{color:#c0392b}
table{border-collapse:collapse;width:100%}
th,td{border:1px solid #d7d7d7;padding:4px 6px;font-size:12px;vertical-align:top}
th{background:#eceff1;font-weight:700;text-align:center}
#wrap{max-height:calc(100vh - 300px);overflow:auto;border:1px solid #ddd;margin-top:8px}
#wrap th{position:sticky;top:0;z-index:2}
td.a{background:#f7fbff}
td.b{background:#fffdf5}
.st{font-weight:700;text-align:center;white-space:nowrap}
.st-일치{color:#2e7d32}
.st-변경{color:#ef6c00}
.st-유사{color:#6a1b9a}
.st-A{color:#c62828}
.st-B{color:#1565c0}
.st-미확인{color:#888}
.opt{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin:8px 0}
.opt label{display:inline-flex;align-items:center;gap:4px}
select,input[type=text],input[type=number]{padding:4px 6px;border:1px solid #bbb;border-radius:3px;font-family:inherit;font-size:12px}
.sum{display:flex;gap:12px;flex-wrap:wrap}
.card{border:1px solid #dde3e8;background:#f7f9fb;border-radius:4px;padding:10px 14px;min-width:150px}
.card .v{font-size:20px;font-weight:700;color:#12507b}
.card .l{font-size:12px;color:#555}
#prog{display:none;align-items:center;gap:10px;padding:6px 12px;background:#fff8e1;
  border-bottom:1px solid #ffe082;font-size:12px}
#prog.on{display:flex}
.pstage{font-weight:700;white-space:nowrap}
.pbar{display:inline-block;width:220px;height:10px;background:#eceff1;border-radius:5px;overflow:hidden}
.pbar>span{display:block;height:100%;width:0;background:#37474f;transition:width .2s}
.pdetail{flex:1;color:#555;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
#pgTime{color:#666;white-space:nowrap}
#partialNote{display:none;margin:0 10px 8px;padding:6px 10px;background:#fff8e1;
  border:1px solid #ffe082;border-radius:3px;font-size:12px}
#partialNote.on{display:block}
.msg{position:fixed;right:16px;bottom:16px;background:#333;color:#fff;padding:9px 14px;border-radius:4px;opacity:0;transition:.25s;z-index:99}
.msg.on{opacity:.95}
.rec{display:block;margin:2px 0;color:#12507b;text-decoration:none;font-size:12px}
.flink{color:#12507b;text-decoration:none;border-bottom:1px dotted #90a4ae;cursor:pointer}
.flink:hover{background:#eaf3fb;text-decoration:none}
small{color:#666}
.warn{color:#c0392b}
.picker{display:none;border:1px solid #ccc;border-radius:3px;padding:8px;background:#fff;margin-top:8px}
.picker .dirbar{display:flex;gap:4px;margin:0 0 6px}
.picker .dirbar input{flex:1;min-width:0}
.entries{max-height:200px;overflow:auto;border:1px solid #eee;padding:6px;background:#fff}
.entries label{display:block;font-size:12px;padding:1px 0}
</style></head><body>

<div id="bar">
  <b>FP 산정 결과물 대조</b>
  <span id="state">파일을 넣고 대조를 실행하십시오.</span>
  <button onclick="loadJson()">저장된 결과 열기</button>
  <button onclick="exportXlsx()">엑셀로 내보내기</button>
  <button onclick="quit()">종료</button>
</div>

<div id="prog">
  <span class="pstage" id="pgStage">준비 중</span>
  <span class="pbar"><span id="pgFill"></span></span>
  <span class="pdetail" id="pgDetail"></span>
  <span id="pgTime"></span>
  <button onclick="cancelRun()">취소</button>
</div>
<div id="tabs">
  <div class="tab on" data-p="p1" onclick="tab('p1')">① 대상 지정</div>
  <div class="tab" data-p="p2" onclick="tab('p2')">② 요약</div>
  <div class="tab" data-p="p3" onclick="tab('p3')">③ 매핑</div>
  <div class="tab" data-p="p4" onclick="tab('p4')">④ 중복 점검</div>
</div>

<div id="partialNote"></div>

<div class="pane on" id="p1">
  <h3>대조할 산출물</h3>
  <div class="cols">
    <div class="col">
      <h4>A 그룹 <small>(앞 단계)</small></h4>
      <input class="nm" id="labA" value="">
      <p style="margin:8px 0 0"><button onclick="addFiles('a')">파일 추가</button>
        <button onclick="togglePicker('a')">폴더에서 고르기</button></p>
      <div class="picker" id="pk_a">
        <div class="dirbar"><input type="text" id="dir_a" placeholder="폴더 경로">
          <button onclick="listDir('a',document.getElementById('dir_a').value)">이동</button></div>
        <div class="entries" id="en_a"></div>
        <p style="margin:6px 0 0"><button onclick="addChecked('a')">선택한 파일 넣기</button>
          <button onclick="checkAll('a')">모두 선택</button>
          <button onclick="togglePicker('a')">닫기</button></p>
      </div>
      <div class="files" id="fa"></div>
    </div>
    <div class="col">
      <h4>B 그룹 <small>(뒤 단계)</small></h4>
      <input class="nm" id="labB" value="">
      <p style="margin:8px 0 0"><button onclick="addFiles('b')">파일 추가</button>
        <button onclick="togglePicker('b')">폴더에서 고르기</button></p>
      <div class="picker" id="pk_b">
        <div class="dirbar"><input type="text" id="dir_b" placeholder="폴더 경로">
          <button onclick="listDir('b',document.getElementById('dir_b').value)">이동</button></div>
        <div class="entries" id="en_b"></div>
        <p style="margin:6px 0 0"><button onclick="addChecked('b')">선택한 파일 넣기</button>
          <button onclick="checkAll('b')">모두 선택</button>
          <button onclick="togglePicker('b')">닫기</button></p>
      </div>
      <div class="files" id="fb"></div>
    </div>
  </div>

  <h3 style="margin-top:18px">대조 기준</h3>
  <div class="opt">
    <label>짝짓기 기준 <select id="key"></select></label>
    <label>중복 판정 기준 <select id="dup"></select></label>
    <label><input type="checkbox" id="fuzzy" checked> 이름이 비슷하면 같은 기능으로 봄</label>
    <label>유사도 <input type="number" id="th" value="0.85" min="0.5" max="1" step="0.01" style="width:70px"></label>
    <label><input type="checkbox" id="igsp" checked> 공백·기호 무시</label>
    <label><input type="checkbox" id="igbr" checked> 괄호 무시</label>
  </div>
  <div class="opt">차이로 볼 항목 <span id="diffs"></span></div>
  <p><button class="main" onclick="run()">대조 실행</button></p>
  <p><small id="mods"></small></p>
</div>

<div class="pane" id="p2"><div id="v2"><p>대조를 실행하면 요약이 나옵니다.</p></div></div>

<div class="pane" id="p3">
  <div class="opt">
    <span>상태</span>
    <label><input type="checkbox" class="fs" value="A만 있음" checked> A만 있음</label>
    <label><input type="checkbox" class="fs" value="B만 있음" checked> B만 있음</label>
    <label><input type="checkbox" class="fs" value="변경" checked> 변경</label>
    <label><input type="checkbox" class="fs" value="유사" checked> 유사</label>
    <label><input type="checkbox" class="fs" value="일치"> 일치</label>
    <label><input type="checkbox" class="fs" value="미확인" checked> 미확인</label>
    <span>검색 <input type="text" id="q" placeholder="기능명·파일명"></span>
    <span id="cnt" style="margin-left:auto"></span>
  </div>
  <div id="wrap"><table id="tb"></table></div>
</div>

<div class="pane" id="p4"><div id="v4"><p>대조를 실행하면 그룹 안의 중복이 나옵니다.</p></div></div>



<div class="msg" id="msg"></div>

<script>
var FORMS=[], KEYS=[], DIFFS=[], OPTS={}, G={a:[],b:[]}, R=null;

function toast(t){var m=document.getElementById('msg');m.textContent=t;m.classList.add('on');
  clearTimeout(m._t);m._t=setTimeout(function(){m.classList.remove('on')},2400);}
function esc(s){return (s===null||s===undefined)?'':String(s)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;');}
function fx(v,d){if(v===null||v===undefined||isNaN(v))return '-';return Number(v).toFixed(d);}
function tab(p){document.querySelectorAll('.tab').forEach(function(t){t.classList.toggle('on',t.dataset.p===p);});
  document.querySelectorAll('.pane').forEach(function(t){t.classList.toggle('on',t.id===p);});}
function base(p){var i=Math.max(p.lastIndexOf('/'),p.lastIndexOf('\\'));return i<0?p:p.substring(i+1);}

/* ---- 파일 목록 ---- */
function paintFiles(){
  ['a','b'].forEach(function(s){
    var h='';
    G[s].forEach(function(f,i){
      var pick;
      if(FORMS.length){
        var o='';
        FORMS.forEach(function(x){o+='<option value="'+x.key+'"'+(f.form===x.key?' selected':'')+'>'+
          esc(x.code+' '+x.name)+'</option>';});
        pick='<select onchange="G[\''+s+'\']['+i+'].form=this.value">'+o+'</select>';
      }else{
        pick='<span class="warn" style="font-size:11px">'+esc(f.form||'양식 미확인')+'</span>';
      }
      var inf=f.info||{}, sub='';
      if(inf.ok===false){
        sub='<div class="sub warn">읽지 못했습니다 : '+esc(inf.msg||'')+'</div>';
      }else if(inf.method){
        sub='<div class="sub"><b>'+esc(inf.method)+'</b>'+
            (inf.methodSrc?' <span style="color:#999">('+esc(inf.methodSrc)+')</span>':'')+
            ' · 기능 '+inf.count+'건'+(inf.fp!==undefined?' '+inf.fp+'FP':'')+
            ' · 시트 '+esc(inf.sheet)+' '+inf.headerRow+'행'+
            (inf.cols&&inf.cols.length?' · 열 '+esc(inf.cols.join(", ")):'')+'</div>';
      }
      h+='<div class="fbox"><div class="file"><span class="n" title="'+esc(f.path)+'">'+
         esc(base(f.path))+'</span>'+pick+
         '<span class="x" onclick="delFile(\''+s+'\','+i+')">×</span></div>'+sub+'</div>';
    });
    if(!G[s].length) h='<small>파일이 없습니다.</small>';
    document.getElementById(s==='a'?'fa':'fb').innerHTML=h;
  });
}
function delFile(s,i){G[s].splice(i,1);paintFiles();}
function addPaths(s,paths){
  if(!paths.length)return;
  fetch('/api/guess',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({paths:paths})}).then(function(r){return r.json();}).then(function(j){
    var skipped=0, bad=0;
    paths.forEach(function(p){
      if(G[s].some(function(f){return f.path===p;})){skipped++;return;}
      var info=(j.info&&j.info[p])||{};
      var f=info.form||j.forms[p];
      if(!f){ f=FORMS.length?FORMS[0].key:''; }
      if(info.ok===false) bad++;
      G[s].push({path:p,form:f,info:info});
    });
    paintFiles();
    toast(paths.length-skipped+'개 넣었습니다.'+(skipped?' (중복 '+skipped+'개 제외)':'')+
          (bad?' · 읽지 못한 파일 '+bad+'개':''));
  });
}
function addFiles(s){
  fetch('/api/browse',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mode:'open',side:s})}).then(function(r){return r.json();}).then(function(j){
    if(!j.ok){toast(j.msg);togglePicker(s);return;}
    addPaths(s,j.paths||[]);});
}
function togglePicker(s){
  var el=document.getElementById('pk_'+s);
  var open=(el.style.display==='block');
  el.style.display=open?'none':'block';
  if(!open) listDir(s,document.getElementById('dir_'+s).value);
}
function listDir(s,d){
  fetch('/api/list',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({dir:d,side:s})}).then(function(r){return r.json();}).then(function(j){
    var box=document.getElementById('en_'+s);
    if(!j.ok){box.innerHTML='<span class="warn">'+esc(j.msg)+'</span>';return;}
    document.getElementById('dir_'+s).value=j.dir;
    var h='';
    if(j.parent&&j.parent!==j.dir)
      h+='<a class="rec" href="#" onclick="listDir(\''+s+'\','+JSON.stringify(j.parent)+');return false;">.. 상위 폴더</a>';
    j.dirs.forEach(function(n){
      h+='<a class="rec" href="#" onclick="listDir(\''+s+'\','+JSON.stringify(j.dir+j.sep+n)+');return false;">[폴더] '+esc(n)+'</a>';});
    j.files.forEach(function(n){
      h+='<label><input type="checkbox" class="pf_'+s+'" value="'+esc(j.dir+j.sep+n)+'"> '+esc(n)+'</label>';});
    if(!j.dirs.length&&!j.files.length) h='<small>이 폴더에 엑셀 파일이 없습니다.</small>';
    box.innerHTML=h;});
}
function checkAll(s){
  var cs=document.querySelectorAll('.pf_'+s);
  var any=false; cs.forEach(function(c){if(!c.checked)any=true;});
  cs.forEach(function(c){c.checked=any;});
}
function addChecked(s){
  var ps=[];
  document.querySelectorAll('.pf_'+s).forEach(function(c){if(c.checked)ps.push(c.value);});
  if(!ps.length){toast('파일을 고르십시오.');return;}
  addPaths(s,ps);
  document.querySelectorAll('.pf_'+s).forEach(function(c){c.checked=false;});
}

/* ---- 대조 실행 ---- */
function opts(){
  var d=[];
  document.querySelectorAll('.df').forEach(function(c){if(c.checked)d.push(c.value);});
  return {key:document.getElementById('key').value,
          dup:document.getElementById('dup').value,
          fuzzy:document.getElementById('fuzzy').checked,
          threshold:parseFloat(document.getElementById('th').value)||0.85,
          ignore_space:document.getElementById('igsp').checked,
          ignore_bracket:document.getElementById('igbr').checked,
          diff:d};
}
var POLL=null, VER=-1, RUNNING=false;
function showProg(on){document.getElementById('prog').classList.toggle('on',!!on);RUNNING=!!on;}
function note(){
  var el=document.getElementById('partialNote');
  if(R&&R.done===false){
    el.className='on';
    el.innerHTML='<b>진행 중입니다.</b> 지금까지 분석한 것까지만 보입니다. '+
      '"미확인"은 아직 짝을 찾는 중인 기능입니다.'+
      (JSONPATH?' &nbsp;중간 결과 저장 위치 : '+esc(JSONPATH):'');
  }else if(R&&JSONPATH){
    el.className='on';
    el.innerHTML='결과를 저장했습니다 : '+esc(JSONPATH);
  }else{ el.className=''; }
}
var JSONPATH='';
function run(){
  if(!G.a.length||!G.b.length){toast('양쪽 그룹에 파일을 넣으십시오.');return;}
  R=null; VER=-1; JSONPATH='';
  document.getElementById('pgStage').textContent='여는 중';
  document.getElementById('pgFill').style.width='0%';
  document.getElementById('pgDetail').textContent='';
  document.getElementById('pgTime').textContent='';
  showProg(true);
  fetch('/api/compare',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({a:{label:document.getElementById('labA').value,files:G.a},
                         b:{label:document.getElementById('labB').value,files:G.b},
                         opts:opts()})}).then(function(r){return r.json();}).then(function(j){
    if(!j.ok){showProg(false);alert(j.msg);return;}
    POLL=setInterval(poll,500); poll();});
}
function poll(){
  fetch('/api/progress?v='+VER).then(function(r){return r.json();}).then(function(j){
    var pct = j.total>0 ? Math.round(j.done*100/j.total) : 0;
    document.getElementById('pgStage').textContent=j.stage||'진행 중';
    document.getElementById('pgFill').style.width=pct+'%';
    document.getElementById('pgDetail').textContent=
      (j.total>0?(j.done+' / '+j.total+'  '):'')+(j.detail||'');
    document.getElementById('pgTime').textContent='경과 '+(j.elapsed||0)+'초';
    if(j.json) JSONPATH=j.json;
    if(j.data){                       // 지금까지 나온 결과로 화면을 채운다
      VER=j.version; R=j.data; paint();
    }
    if(!j.finished) return;
    clearInterval(POLL);POLL=null;showProg(false);
    if(j.error){ note(); alert(j.error); return; }
    note();
    toast('대조를 마쳤습니다. '+(j.elapsed||0)+'초');
    if(document.querySelector('.tab.on').dataset.p==='p1') tab('p2');
  }).catch(function(){});
}
function cancelRun(){
  document.getElementById('pgStage').textContent='취소하는 중입니다...';
  fetch('/api/cancel',{method:'POST'});
}
function loadJson(){
  fetch('/api/browse',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mode:'json'})}).then(function(r){return r.json();}).then(function(j){
    var p=(j.ok&&j.paths&&j.paths.length)?j.paths[0]:prompt('불러올 JSON 경로','');
    if(!p)return;
    fetch('/api/loadjson',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({path:p})}).then(function(r){return r.json();}).then(function(k){
      if(!k.ok){alert(k.msg);return;}
      R=k.data; JSONPATH=p; paint(); tab('p2'); toast('불러왔습니다: '+k.name);});});
}

/* ---- 결과 ---- */
function stClass(s){return s==='A만 있음'?'st-A':(s==='B만 있음'?'st-B':'st-'+s);}
function paint(){
  buildFileIndex();
  var st=R.stats, a=R.a, b=R.b;
  document.getElementById('state').textContent=
    a.label+' '+a.count+'건 / '+b.label+' '+b.count+'건 · 짝 없음 '+
    (st['A만 있음']+st['B만 있음'])+'건';
  var h='<div class="sum">';
  [['일치',st['일치']],['변경',st['변경']],['유사',st['유사']],
   [a.label+'에만 있음',st['A만 있음']],[b.label+'에만 있음',st['B만 있음']]].forEach(function(x){
    h+='<div class="card"><div class="v">'+x[1]+'</div><div class="l">'+esc(x[0])+'</div></div>';});
  h+='</div>';
  h+='<h3 style="margin-top:16px">기능점수</h3><table style="width:auto">'+
     '<tr><th></th><th>'+esc(a.label)+'</th><th>'+esc(b.label)+'</th><th>차이</th></tr>'+
     '<tr><th>파일 수</th><td class="right">'+a.files.length+'</td><td>'+b.files.length+'</td><td></td></tr>'+
     '<tr><th>기능 수</th><td>'+a.count+'</td><td>'+b.count+'</td><td>'+(b.count-a.count>0?'+':'')+(b.count-a.count)+'</td></tr>'+
     '<tr><th>기능점수</th><td>'+fx(a.fp,1)+'</td><td>'+fx(b.fp,1)+'</td><td>'+(st.fpDiff>0?'+':'')+fx(st.fpDiff,1)+'</td></tr></table>';
  var types={};
  Object.keys(a.byType).forEach(function(t){types[t]=1;});
  Object.keys(b.byType).forEach(function(t){types[t]=1;});
  h+='<h3 style="margin-top:16px">FP유형별</h3><table style="width:auto"><tr><th>유형</th>'+
     '<th>'+esc(a.label)+' 건수</th><th>FP</th><th>'+esc(b.label)+' 건수</th><th>FP</th><th>FP 차이</th></tr>';
  Object.keys(types).sort().forEach(function(t){
    var x=a.byType[t]||{n:0,fp:0}, y=b.byType[t]||{n:0,fp:0};
    h+='<tr><th>'+esc(t)+'</th><td>'+x.n+'</td><td>'+fx(x.fp,1)+'</td><td>'+y.n+'</td><td>'+fx(y.fp,1)+
       '</td><td>'+((y.fp-x.fp)>0?'+':'')+fx(y.fp-x.fp,1)+'</td></tr>';});
  h+='</table>';
  h+='<h3 style="margin-top:16px">읽은 파일</h3><table><tr><th>그룹</th><th>파일</th><th>양식</th><th>시트</th><th>산정방법</th><th>기능 수</th></tr>';
  [[a,a.label,'a'],[b,b.label,'b']].forEach(function(g){
    g[0].files.forEach(function(f){
      h+='<tr><td>'+esc(g[1])+'</td><td>'+fileLink(g[2],{src:f.name,no:0})+'</td><td>'+esc(f.formName)+
         '</td><td>'+esc(f.sheet)+'</td><td>'+esc(f.method)+
         (f.methodSrc?' <small>('+esc(f.methodSrc)+')</small>':'')+
         '</td><td class="right">'+f.count+'</td></tr>';});
    (g[0].errors||[]).forEach(function(e){
      h+='<tr><td>'+esc(g[1])+'</td><td colspan="5" class="warn">'+esc(e.path)+' — '+esc(e.msg)+'</td></tr>';});
  });
  h+='</table><p><small>짝짓기 기준 : '+esc(R.keyLabel)+
     ' &nbsp;|&nbsp; 유사도 임계값 '+R.opts.threshold+'</small></p>';
  document.getElementById('v2').innerHTML=h;
  paintDup(); paintTable(); note();
}
function paintDup(){
  var h='';
  if(R.dupLabel) h+='<p><small>중복 판정 기준 : <b>'+esc(R.dupLabel)+'</b></small></p>';
  if(R.dupNote) h+='<p class="warn" style="background:#fff8e1;border:1px solid #ffe082;'+
    'border-radius:3px;padding:8px 10px">'+esc(R.dupNote)+'</p>';
  [[R.dupA,R.a.label,'a'],[R.dupB,R.b.label,'b']].forEach(function(g){
    h+='<h3 style="margin-top:14px">'+esc(g[1])+' 안의 중복 '+g[0].length+'건</h3>';
    if(!g[0].length){h+='<p><small>같은 이름의 기능이 두 번 이상 나오지 않습니다.</small></p>';return;}
    h+='<table><tr><th style="width:170px">파일</th><th style="width:50px">행</th>'+
       '<th style="width:110px">애플리케이션</th><th style="width:110px">세부업무</th>'+
       '<th style="width:200px">단위프로세스명</th><th>단위프로세스 설명</th>'+
       '<th style="width:60px">FP유형</th><th style="width:60px">가중치</th></tr>';
    g[0].forEach(function(d){
      d.rows.forEach(function(r){
        h+='<tr><td>'+fileLink(g[2],r)+'</td><td class="right">'+r.no+'</td><td>'+esc(r.app)+'</td><td>'+esc(r.biz)+
           '</td><td>'+esc(r.proc)+'</td><td>'+esc((r.desc||'').replace(/\n/g," / "))+
           '</td><td>'+esc(r.type)+'</td><td class="right">'+(r.wt===null?'':r.wt)+'</td></tr>';});
      h+='<tr><td colspan="8" style="background:#fafafa"></td></tr>';});
    h+='</table>';});
  document.getElementById('v4').innerHTML=h;
}
var FIDX={a:{},b:{}};
function buildFileIndex(){
  FIDX={a:{},b:{}};
  ['a','b'].forEach(function(k){
    ((R[k]||{}).files||[]).forEach(function(f){ FIDX[k][f.name]=f.path; });
  });
}
function openFile(side,src,no){
  var p=(FIDX[side]||{})[src];
  if(!p){toast('파일 위치를 알 수 없습니다.');return;}
  fetch('/api/openfile',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({path:p})}).then(function(r){return r.json();}).then(function(j){
    if(!j.ok){toast(j.msg);return;}
    toast(j.name+' 을(를) 엽니다.'+(no?' '+no+'행을 보십시오.':''));});
}
function jstr(s){return "'"+String(s===null||s===undefined?'':s)
  .replace(/\\/g,'\\\\').replace(/'/g,"\\'")+"'";}
function fileLink(side,r){
  var path=(FIDX[side]||{})[r.src]||r.src;
  var call="openFile('"+side+"',"+jstr(r.src)+","+(r.no||0)+")";
  return '<a href="#" class="flink" title="'+esc(path)+'" onclick="'+esc(call)+
         ';return false;">'+esc(r.src)+'</a>';
}
function dupIndex(){
  var idx={};
  [['a',R.dupA,R.a.label],['b',R.dupB,R.b.label]].forEach(function(g){
    (g[1]||[]).forEach(function(grp){
      var rows=grp.rows||[];
      rows.forEach(function(r){
        var others=rows.filter(function(x){return x.no!==r.no;}).map(function(x){return x.no;});
        idx[g[0]+'|'+r.src+'|'+r.no]={label:g[2],n:rows.length,others:others};
      });
    });
  });
  return idx;
}
function noteLines(p,idx){
  var out=p.diffs.map(function(x){return esc(x.label)+' '+esc(x.a||'-')+'\u2192'+esc(x.b||'-');});
  if(p.status==='\uc720\uc0ac'&&p.basis)
    out.push('<span style="color:#6a1b9a">\uc720\uc0ac \uadfc\uac70 : '+esc(p.basis)+' ('+fx(p.score,3)+')</span>');
  ['a','b'].forEach(function(side){
    var r=p[side]; if(!r)return;
    var d=idx[side+'|'+r.src+'|'+r.no];
    if(d) out.push('<span style="color:#b06000">'+esc(d.label)+' \uc911\ubcf5 '+d.n+
      '\uac74 (\uac19\uc740 \ub0b4\uc6a9 '+d.others.slice(0,8).join(', ')+
      (d.others.length>8?'\u2026':'')+'\ud589)</span>');
  });
  return out;
}
function paintTable(){
  if(!R)return;
  var DIDX=dupIndex();
  var on={}; document.querySelectorAll('.fs').forEach(function(c){on[c.value]=c.checked;});
  var q=document.getElementById('q').value.trim().toLowerCase();
  var h=['<tr><th rowspan="2" style="width:70px">상태</th><th rowspan="2" style="width:52px">유사도</th>',
    '<th colspan="5">'+esc(R.a.label)+'</th><th colspan="5">'+esc(R.b.label)+'</th>',
    '<th rowspan="2" style="width:230px">달라진 항목 · 비고</th></tr>',
    '<tr><th style="width:120px">파일·행</th><th style="width:110px">애플리케이션</th><th style="width:100px">세부업무</th><th>단위프로세스명</th><th style="width:90px">유형·복잡도</th>',
    '<th style="width:120px">파일·행</th><th style="width:110px">애플리케이션</th><th style="width:100px">세부업무</th><th>단위프로세스명</th><th style="width:90px">유형·복잡도</th></tr>'];
  var n=0;
  R.pairs.forEach(function(p){
    if(!on[p.status])return;
    if(q){var s=JSON.stringify([p.a,p.b]).toLowerCase(); if(s.indexOf(q)<0)return;}
    n++;
    function cell(r,cls){
      if(!r) return '<td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td>';
      var side=(cls==='a')?'a':'b';
      return '<td class="'+cls+'">'+fileLink(side,r)+' '+r.no+'행</td>'+
             '<td class="'+cls+'">'+esc(r.app)+'</td><td class="'+cls+'">'+esc(r.biz)+'</td>'+
             '<td class="'+cls+'">'+esc(r.proc)+'</td>'+
             '<td class="'+cls+'">'+esc(r.type)+(r.cx?' / '+esc(r.cx):'')+(r.wt!==null?' / '+r.wt:'')+'</td>';
    }
    var d=noteLines(p,DIDX).join('<br>');
    h.push('<tr><td class="st '+stClass(p.status)+'">'+p.status.replace('A',R.a.label).replace('B',R.b.label)+'</td>',
      '<td class="right">'+(p.status==='유사'?fx(p.score,3):'')+'</td>',
      cell(p.a,'a'), cell(p.b,'b'),
      '<td>'+d+'</td></tr>');
  });
  document.getElementById('tb').innerHTML=h.join('');
  document.getElementById('cnt').textContent='표시 '+n+' / 전체 '+R.pairs.length+'건';
}
function exportXlsx(){
  if(!R){toast('먼저 대조를 실행하십시오.');return;}
  fetch('/api/browse',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mode:'save'})}).then(function(r){return r.json();}).then(function(j){
    var p=(j.ok&&j.paths&&j.paths.length)?j.paths[0]:prompt('저장할 파일 경로를 입력하십시오.','FP대조결과.xlsx');
    if(!p)return;
    fetch('/api/export',{method:'POST',headers:{'Content-Type':'application/json'},
      body:JSON.stringify({path:p})}).then(function(r){return r.json();}).then(function(k){
      if(!k.ok){alert(k.msg);return;} toast('저장했습니다: '+k.name);});});
}
function quit(){if(!confirm('프로그램을 종료합니다.'))return;
  fetch('/api/shutdown',{method:'POST'});setTimeout(function(){window.close();},300);}

document.querySelectorAll('.fs').forEach(function(c){c.addEventListener('change',paintTable);});
document.getElementById('q').addEventListener('input',paintTable);
setInterval(function(){fetch('/api/alive',{method:'POST'});},4000);

fetch('/api/init').then(function(r){return r.json();}).then(function(j){
  if(j.modules) showModules(j.modules);
  if(j.ok===false||!j.forms||!j.keys){
    alert('시작 정보를 읽지 못했습니다.\n\n'+(j.msg||'서버 응답이 비어 있습니다.'));
    return;
  }
  FORMS=j.forms; KEYS=j.keys; DIFFS=j.diffs; OPTS=j.opts;
  document.getElementById('labA').value=j.labels[0];
  document.getElementById('labB').value=j.labels[1];
  var ld=j.lastDirs||{};
  document.getElementById('dir_a').value=ld.a||'';
  document.getElementById('dir_b').value=ld.b||'';
  var h='';KEYS.forEach(function(k){h+='<option value="'+k.k+'"'+(k.k===OPTS.key?' selected':'')+'>'+esc(k.label)+'</option>';});
  document.getElementById('key').innerHTML=h;
  h='';(j.dups||[]).forEach(function(k){h+='<option value="'+k.k+'"'+(k.k===OPTS.dup?' selected':'')+'>'+esc(k.label)+'</option>';});
  document.getElementById('dup').innerHTML=h;
  h='';DIFFS.forEach(function(d){
    h+='<label style="margin-right:12px"><input type="checkbox" class="df" value="'+d.k+'"'+
       (OPTS.diff.indexOf(d.k)>=0?' checked':'')+'> '+esc(d.label)+'</label>';});
  document.getElementById('diffs').innerHTML=h;
  document.getElementById('th').value=OPTS.threshold;
  paintFiles();
}).catch(function(e){
  alert('시작 정보를 읽지 못했습니다.\n서버가 500 을 냈을 수 있습니다. 실행한 콘솔 창을 확인하십시오.\n\n'+e);
});
function showModules(m){
  var el=document.getElementById('mods');
  if(!el)return;
  el.innerHTML='엔진 comparator.py '+esc(String(m.comparatorVersion))+' · '+esc(m.comparator)+
    '<br>양식 calculator_ui.py '+esc(String(m.calculatorVersion))+' · '+esc(m.calculator_ui)+
    (m.missing&&m.missing.length?'<br><span class="warn">짝이 맞지 않는 항목: '+esc(m.missing.join(", "))+'</span>':'');
}
</script></body></html>
"""


def main():
    U.setup_output()
    no_browser = "--no-browser" in sys.argv
    port = U.pick_port()
    url = "http://127.0.0.1:%d/" % port
    chk = module_check()
    print("=" * 60)
    print(APP_NAME, "v" + VERSION)
    print("  comparator.py    :", chk["comparator"], "v" + str(chk["comparatorVersion"]))
    print("  calculator_ui.py :", chk["calculator_ui"], "v" + str(chk["calculatorVersion"]))
    if chk["missing"]:
        print("!" * 60)
        print(chk["msg"])
        print("!" * 60)
    print("화면 주소:", url)
    print("창이 뜨지 않으면 브라우저에 위 주소를 직접 입력하십시오.")
    print("=" * 60)

    threading.Thread(target=watchdog, daemon=True).start()
    if not no_browser:
        threading.Thread(
            target=lambda: (time.sleep(1.0),
                            U.open_browser(url, ".chrome_comparator")),
            daemon=True).start()

    import logging
    logging.getLogger("werkzeug").setLevel(logging.ERROR)
    app.run(host="127.0.0.1", port=port, debug=False,
            use_reloader=False, threaded=True)


if __name__ == "__main__":
    main()