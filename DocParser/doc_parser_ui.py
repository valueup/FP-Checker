# -*- coding: utf-8 -*-
"""doc_parser_ui — 산출물이 제대로 읽히는지 확인하는 화면.

FP-Checker 전체를 돌리지 않고 doc_parser 만 돌려 본다.
  · 화면 목록 — UI설계서에서 뽑은 화면과 그 속성·처리상세
  · 테이블 목록 — 테이블정의서에서 뽑은 테이블과 논리파일 그룹
  · 연결 — 화면 속성표의 테이블이 테이블정의서에 실제로 있는지
  · 문제 — 읽는 중에 걸린 것

읽기만 한다. 아무것도 고치지 않는다.
설정과 캐시는 FP-Checker 와 같은 것을 쓴다. 포트만 따로 잡아 함께 띄울 수 있다.

실행:  python DocParser/doc_parser_ui.py
"""

import os
import sys
import threading
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]          # FP-Checker
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from flask import Flask, jsonify, render_template_string, request  # noqa: E402

import common                                                       # noqa: E402
from DocParser import doc_parser                                    # noqa: E402
from DocParser.doc_parser import (build_groups, build_mapping,      # noqa: E402
                                  load_screens, load_tables,
                                  mapping_csv, screen_events)

APP_TITLE = "DocParser 산출물 확인"
VERSION = "v1.0"

app = Flask(__name__)
app.config["JSON_AS_ASCII"] = False

DATA = {"screens": {}, "tables": {}, "groups": {}, "t2g": {},
        "problems": [], "ui_dir": "", "tbl_dir": "",
        "files": [], "saved": "", "saved_mb": 0.0}

JOB = {"running": False, "done": False, "error": "",
       "stage": "대기", "files": [], "logs": []}
LOCK = threading.Lock()


class Cache:
    """doc_parser 가 쓸 캐시 어댑터. common 의 파일 캐시를 그대로 쓴다."""

    @staticmethod
    def get(kind, path):
        return common.cache_get(kind, path)

    @staticmethod
    def put(kind, path, data):
        return common.cache_put(kind, path, data)


# ======================================================================
# 읽기
# ======================================================================
def note(msg):
    with LOCK:
        JOB["logs"].append(f"{time.strftime('%H:%M:%S')}  {msg}")
    common.log("확인", msg)


def stage(s):
    with LOCK:
        JOB["stage"] = s


def on_file(rep, i, total):
    with LOCK:
        JOB["files"].append(rep)
    tail = f" · {rep['note']}" if rep["note"] else ""
    cached = " (캐시)" if rep["cached"] else ""
    note(f"{rep['name']} — {rep['count']}개{cached}"
         f" · {rep['ms'] / 1000:.1f}초{tail}")


def read_all(ui_dir, tbl_dir, use_cache):
    cache = Cache if use_cache else None
    try:
        stage("UI설계서 읽는 중")
        screens, sp, rep_ui = load_screens(
            ui_dir, cache, on_file,
            lambda i, c, t: stage(f"UI설계서 {i + 1}번째 · 슬라이드 {c}/{t}"))

        stage("테이블정의서 읽는 중")
        tables, tp, rep_tbl = load_tables(tbl_dir, cache, on_file)

        stage("논리파일 그룹 구성")
        groups, t2g = build_groups(tables)

        DATA.update({"screens": screens, "tables": tables, "groups": groups,
                     "t2g": t2g, "problems": sp + tp,
                     "ui_dir": ui_dir, "tbl_dir": tbl_dir,
                     "files": rep_ui + rep_tbl})
        common.cfg.set_many({("path", "ui_dir"): ui_dir,
                             ("path", "tbl_dir"): tbl_dir,
                             ("view", "use_cache"): "1" if use_cache else "0"})
        cols = sum(len(t["columns"]) for t in tables.values())
        note(f"화면 {len(screens)}개 · 테이블 {len(tables)}개 · "
             f"컬럼 {cols}개 · 논리파일 그룹 {len(groups)}개")
        if sp + tp:
            note(f"확인할 것 {len(sp + tp)}건. '문제' 탭을 보십시오.")

        stage("저장 중")
        save_result()
        stage("완료")
    except Exception as e:
        with LOCK:
            JOB["error"] = str(e)
        note(f"오류: {e}")
        common.log("확인", "실패", exc=True)
        stage("중단")
    finally:
        with LOCK:
            JOB["running"], JOB["done"] = False, True


def save_result(pretty=None):
    """읽은 결과를 out 폴더에 JSON 으로 적는다.

    comparator·calculator·reporter 는 산출물을 다시 읽지 않고 이 파일만 읽는다.
    """
    if not DATA["screens"] and not DATA["tables"]:
        raise RuntimeError("저장할 것이 없습니다. 먼저 산출물을 읽으십시오.")
    if pretty is None:
        pretty = common.cfg.get_bool("view", "json_pretty", False)
    folder = common.out_dir()
    path = folder / doc_parser.make_name()
    data = doc_parser.dump(
        DATA["screens"], DATA["tables"], DATA["groups"], DATA["t2g"],
        DATA["problems"],
        {"ui_dir": DATA["ui_dir"], "tbl_dir": DATA["tbl_dir"],
         "files": DATA["files"]})
    doc_parser.save(path, data, pretty)
    mb = path.stat().st_size / 1024 / 1024
    DATA["saved"], DATA["saved_mb"] = str(path), round(mb, 2)
    note(f"저장: {path.name} ({mb:.2f} MB)")

    keep = common.cfg.get_int("general", "keep_out_files", 0)
    for gone in doc_parser.prune(folder, keep):
        note(f"오래된 파일 삭제: {os.path.basename(gone)}")
    return DATA["saved"]


# ======================================================================
# 조회
# ======================================================================
def screen_rows():
    known = set(DATA["tables"])
    out = []
    for s in DATA["screens"].values():
        miss = [t for t in s["tables"] if t not in known]
        out.append({"id": s["id"], "name": s["name"], "file": s["file"],
                    "events": " / ".join(screen_events(s)),
                    "slides": (f"{min(s['slides'])}~{max(s['slides'])}"
                               if s["slides"] else ""),
                    "attrs": len(s["attrs"]), "funcs": len(s["funcs"]),
                    "tables": len(s["tables"]), "miss": len(miss)})
    return sorted(out, key=lambda r: r["id"])


def table_rows():
    used = {t for s in DATA["screens"].values() for t in s["tables"]}
    out = []
    for t in DATA["tables"].values():
        gk = DATA["t2g"].get(t["name"], "")
        g = DATA["groups"].get(gk, {})
        out.append({"name": t["name"], "kor": t["kor"], "system": t["system"],
                    "file": t["file"], "group": gk,
                    "group_n": len(g.get("tables", [])),
                    "cols": len(t["columns"]),
                    "used": 1 if t["name"] in used else 0})
    return sorted(out, key=lambda r: r["name"])


def mapping(view):
    """화면-테이블 매핑을 세 가지 기준으로 낸다.

    physical  화면 × 물리 테이블 한 줄
    logical   화면 × 논리파일 그룹 한 줄. 같은 그룹의 물리 표를 한 줄로 묶는다
    summary   화면 한 줄. 물리 표 수와 FTR 후보를 나란히 둔다
    """
    pairs, summary = build_mapping(DATA["screens"], DATA["tables"],
                                   DATA["groups"], DATA["t2g"])
    if view == "physical":
        return pairs
    if view == "summary":
        return summary

    fold = {}
    for r in pairs:
        if not r["table"]:
            continue
        key = (r["screen"], r["group"])
        g = fold.get(key)
        if g is None:
            g = fold[key] = {
                "screen": r["screen"], "screen_name": r["screen_name"],
                "group": r["group"], "group_n": r["group_n"],
                "group_det": r["group_det"], "tables": [], "attrs": 0,
                "missing": 0}
        g["tables"].append(r["table"])
        g["attrs"] += r["attrs"]
        if not r["known"]:
            g["missing"] += 1
    out = []
    for g in fold.values():
        g["tables"] = sorted(g["tables"])
        g["table_n"] = len(g["tables"])
        g["tables_text"] = ", ".join(g["tables"])
        out.append(g)
    return sorted(out, key=lambda r: (r["screen"], r["group"]))


def link_report():
    """화면 속성표의 테이블과 테이블정의서를 맞춰 본다."""
    known = set(DATA["tables"])
    missing, used = {}, set()
    for s in DATA["screens"].values():
        for t in s["tables"]:
            if t in known:
                used.add(t)
            else:
                missing.setdefault(t, []).append(s["id"])
    unused = sorted(known - used)
    return {
        "missing": [{"name": t, "n": len(ids), "screens": ", ".join(sorted(ids)[:8])}
                    for t, ids in sorted(missing.items(), key=lambda x: -len(x[1]))],
        "unused": [{"name": t, "kor": DATA["tables"][t]["kor"],
                    "file": DATA["tables"][t]["file"]} for t in unused],
        "used_n": len(used),
    }


# ======================================================================
# 라우트
# ======================================================================
@app.route("/api/config")
def api_config():
    return jsonify({"ui_dir": common.cfg.get("path", "ui_dir", ""),
                    "tbl_dir": common.cfg.get("path", "tbl_dir", ""),
                    "use_cache": common.cfg.get_bool("view", "use_cache", True),
                    "version": VERSION})


@app.route("/api/read", methods=["POST"])
def api_read():
    d = request.get_json(force=True, silent=True) or {}
    with LOCK:
        if JOB["running"]:
            return jsonify({"error": "이미 읽는 중입니다."}), 409
        JOB.update({"running": True, "done": False, "error": "",
                    "stage": "시작", "files": [], "logs": []})
    threading.Thread(target=read_all,
                     args=(d.get("ui_dir", "").strip(),
                           d.get("tbl_dir", "").strip(),
                           bool(d.get("use_cache", True))),
                     daemon=True).start()
    return jsonify({"ok": True})


@app.route("/api/status")
def api_status():
    with LOCK:
        j = {k: JOB[k] for k in ("running", "done", "error", "stage", "files")}
        j["logs"] = JOB["logs"][-200:]
    j["saved"] = DATA["saved"]
    j["saved_mb"] = DATA["saved_mb"]
    j["counts"] = {
        "screens": len(DATA["screens"]), "tables": len(DATA["tables"]),
        "groups": len(DATA["groups"]),
        "cols": sum(len(t["columns"]) for t in DATA["tables"].values()),
        "problems": len(DATA["problems"])}
    return jsonify(j)


@app.route("/api/screens")
def api_screens():
    return jsonify({"rows": screen_rows()})


@app.route("/api/screens/<path:sid>")
def api_screen(sid):
    s = DATA["screens"].get(sid)
    if not s:
        return jsonify({"error": "없는 화면ID"}), 404
    tables = DATA["tables"]
    return jsonify({
        "id": s["id"], "name": s["name"], "file": s["file"],
        "slides": s["slides"], "funcs": s["funcs"],
        "attrs": [{"name": a["name"], "table": a["table"],
                   "known": 1 if a["table"] in tables else 0} for a in s["attrs"]],
        "tables": [{"name": t, "kor": tables.get(t, {}).get("kor", ""),
                    "known": 1 if t in tables else 0} for t in s["tables"]]})


@app.route("/api/tables")
def api_tables():
    return jsonify({"rows": table_rows()})


@app.route("/api/tables/<path:name>")
def api_table(name):
    t = DATA["tables"].get(name)
    if not t:
        return jsonify({"error": "없는 테이블"}), 404
    gk = DATA["t2g"].get(name, "")
    g = DATA["groups"].get(gk, {})
    users = sorted(s["id"] for s in DATA["screens"].values()
                   if name in s["tables"])
    return jsonify({**t, "group": gk, "group_tables": g.get("tables", []),
                    "ret": g.get("ret"), "det": g.get("det"),
                    "det_no_audit": g.get("det_no_audit"), "screens": users})


@app.route("/api/link")
def api_link():
    return jsonify(link_report())


@app.route("/api/problems")
def api_problems():
    return jsonify({"rows": DATA["problems"]})


@app.route("/api/mapping")
def api_mapping():
    """세 기준을 한 번에 돌려준다. 화면에서 셋을 함께 보여 준다."""
    view = request.args.get("view", "all")
    if view == "all":
        return jsonify({"summary": mapping("summary"),
                        "logical": mapping("logical"),
                        "physical": mapping("physical")})
    if view not in ("physical", "logical", "summary"):
        return jsonify({"error": "알 수 없는 기준"}), 400
    return jsonify({"rows": mapping(view)})


@app.route("/api/mapping.csv")
def api_mapping_csv():
    from flask import Response
    pairs, _ = build_mapping(DATA["screens"], DATA["tables"],
                             DATA["groups"], DATA["t2g"])
    return Response(
        mapping_csv(pairs), mimetype="text/csv; charset=utf-8",
        headers={"Content-Disposition":
                 "attachment; filename=screen_table_mapping.csv"})


@app.route("/api/save", methods=["POST"])
def api_save():
    d = request.get_json(force=True, silent=True) or {}
    try:
        path = save_result(bool(d.get("pretty", False)))
    except Exception as e:
        return jsonify({"error": str(e)}), 400
    return jsonify({"path": path, "mb": DATA["saved_mb"]})


@app.route("/api/cache/clear", methods=["POST"])
def api_cache_clear():
    return jsonify({"n": common.cache_clear()})


@app.route("/")
def index():
    return render_template_string(PAGE, title=APP_TITLE)


# ======================================================================
# 화면
# ======================================================================
PAGE = r"""<!doctype html>
<html lang="ko"><head><meta charset="utf-8"><title>{{ title }}</title>
<style>
*{box-sizing:border-box}
body{margin:0;font:13px "맑은 고딕","Malgun Gothic",sans-serif;color:#222;background:#f4f5f7}
.wrap{padding:10px 12px}
.card{background:#fff;border:1px solid #dcdfe4;border-radius:6px;margin-bottom:10px}
.card>h3{margin:0;padding:7px 12px;font-size:12px;font-weight:600;color:#555;
  background:#f7f8fa;border-bottom:1px solid #e6e8ec;border-radius:6px 6px 0 0}
.card>.body{padding:10px 12px}
.row{display:grid;grid-template-columns:120px 1fr auto;gap:6px;align-items:center;margin-bottom:5px}
.row label{color:#555}
input[type=text]{width:100%;padding:5px 8px;border:1px solid #cfd4da;border-radius:4px;
  font:12px Consolas,"D2Coding",monospace}
button{padding:5px 12px;border:1px solid #c3c8cf;background:#fff;border-radius:4px;
  cursor:pointer;font:inherit}
button:hover{background:#f0f2f5}
button.pri{background:#3f5f9e;border-color:#3f5f9e;color:#fff;font-weight:600;padding:6px 20px}
button.pri:hover{background:#33507f}
button:disabled{opacity:.5;cursor:default}
.hint{color:#888;font-size:11px}
.sum{display:flex;gap:20px;flex-wrap:wrap;align-items:baseline}
.sum div{font-size:12px;color:#666}
.sum b{font-size:17px;color:#2f4a72;margin-left:5px}
.tabs{display:flex;gap:2px;border-bottom:1px solid #dcdfe4;padding:0 12px;background:#f7f8fa;
  border-radius:6px 6px 0 0}
.tabs button{border:none;background:none;padding:8px 14px;border-bottom:2px solid transparent;
  border-radius:0;color:#666}
.tabs button.on{color:#2f4a72;font-weight:600;border-bottom-color:#3f5f9e}
table{border-collapse:collapse;width:100%;font-size:12px}
th,td{border-bottom:1px solid #eceef1;padding:4px 6px;text-align:left;white-space:nowrap;
  overflow:hidden;text-overflow:ellipsis}
thead th{position:sticky;top:0;background:#f2f4f7;z-index:2;font-weight:600;color:#555}
tbody tr:hover{background:#dce9fb;cursor:pointer}
tbody tr.sel{background:#cfe0f7}
.scroll{max-height:44vh;overflow:auto;border:1px solid #e2e5ea;border-radius:4px}
.bad{color:#b03030;font-weight:600}
table.grid{table-layout:fixed}
tbody tr.miss td{background:#fdf0f0}
tbody tr.miss:hover td{background:#f7dcdc}
tbody tr.miss td:first-child{box-shadow:inset 3px 0 0 #d05a5a}
.dim{color:#aaa}
.split{display:grid;grid-template-columns:1fr 400px;gap:10px}
pre{margin:0;font:11px Consolas,"D2Coding",monospace;white-space:pre-wrap;
  max-height:110px;overflow:auto;color:#666}
.pill{display:inline-block;padding:1px 7px;border-radius:9px;font-size:11px;
  background:#eef1f5;color:#555;margin:0 4px 3px 0}
.side{padding:9px}
.side h4{margin:10px 0 4px;font-size:12px;color:#444}
button.cp{padding:0 4px;border:1px solid #d3d8df;background:#fff;border-radius:3px;
  font-size:11px;line-height:16px;color:#7a828c;cursor:pointer}
button.cp:hover{background:#eef2f7;color:#2f4a72;border-color:#b9c6da}
td.cpc{padding:2px 4px;text-align:center;overflow:visible}
#toast{position:fixed;left:50%;bottom:26px;transform:translateX(-50%);
  background:#2f3a4a;color:#fff;padding:7px 16px;border-radius:4px;font-size:12px;
  opacity:0;transition:opacity .18s;pointer-events:none;z-index:99}
#toast.on{opacity:.94}
.map3{display:grid;grid-template-columns:minmax(420px,1fr) minmax(560px,1.35fr);gap:9px}
.mapcol{display:flex;flex-direction:column;gap:9px}
.mapbox h5{margin:0 0 4px;font-size:12px;font-weight:600;color:#444}
.mapbox h5 .hint{font-weight:400;margin-left:4px}
@media (max-width:1100px){.map3{grid-template-columns:1fr}}
.saved{display:flex;align-items:center;gap:10px;margin-top:8px;padding:6px 9px;
  background:#f2f6fb;border:1px solid #dbe4f0;border-radius:4px;font-size:12px}
.saved span{font-family:Consolas,"D2Coding",monospace;color:#2f4a72;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1}
</style></head><body><div class="wrap">

<div class="card"><h3>산출물 폴더</h3><div class="body">
  <div class="row"><label>UI설계서</label>
    <input type="text" id="ui_dir" placeholder="*UI설계서*.pptx 가 있는 폴더"><span></span></div>
  <div class="row"><label>테이블정의서</label>
    <input type="text" id="tbl_dir" placeholder="*테이블정의서*.xlsx 가 있는 폴더"><span></span></div>
  <div class="row"><label></label>
    <label class="hint"><input type="checkbox" id="use_cache" checked>
      이미 읽은 파일은 캐시에서 가져오기</label>
    <span><button id="clear">캐시 비우기</button>
      <button class="pri" id="go">읽기</button></span></div>
  <div class="hint" id="msg"></div>
</div></div>

<div class="card"><h3>진행</h3><div class="body">
  <div class="sum">
    <div>화면<b id="c_screens">0</b></div>
    <div>테이블<b id="c_tables">0</b></div>
    <div>컬럼<b id="c_cols">0</b></div>
    <div>논리파일 그룹<b id="c_groups">0</b></div>
    <div>문제<b id="c_problems">0</b></div>
    <div class="hint" id="stage" style="margin-left:auto"></div>
  </div>
  <div class="saved" id="saved" style="display:none">
    <span id="saved_path"></span>
    <label class="hint"><input type="checkbox" id="pretty"> 읽기 쉽게(줄바꿈)</label>
    <button id="resave">다시 저장</button>
  </div>
  <div class="scroll" style="max-height:130px;margin-top:8px"><table>
    <thead><tr><th style="width:80px">구분</th><th>파일</th>
      <th style="width:70px">건수</th><th style="width:70px">시간</th>
      <th style="width:150px">비고</th></tr></thead>
    <tbody id="files"></tbody></table></div>
  <pre id="logs" style="margin-top:6px"></pre>
</div></div>

<div class="card">
  <div class="tabs">
    <button data-t="screens" class="on">화면 목록</button>
    <button data-t="tables">테이블 목록</button>
    <button data-t="mapping">매핑</button>
    <button data-t="link">연결</button>
    <button data-t="problems">문제</button>
  </div>
  <div class="body">
    <div class="row" style="grid-template-columns:1fr auto;margin-bottom:8px">
      <input type="text" id="q" placeholder="이름·ID 로 거르기">
      <span><span class="hint" id="cnt"></span>
        <span id="copybar" style="display:none;margin-left:10px">
          <button id="copyall">보이는 것 모두 복사</button>
        </span>
        <span id="mapbar" style="display:none;margin-left:10px">
          <button id="clearsel">선택 해제</button>
          <button id="csv">CSV 내려받기</button>
        </span></span></div>
    <div id="pane"><div class="hint">폴더를 지정하고 읽기를 누르십시오.</div></div>
  </div>
</div>

</div>
<div id="toast"></div>
<script>
const $=s=>document.querySelector(s);

function toast(msg){
  const t=$("#toast"); t.textContent=msg; t.classList.add("on");
  clearTimeout(toast.t); toast.t=setTimeout(()=>t.classList.remove("on"),1400);
}

async function copyText(text,label){
  if(!text){toast("복사할 것이 없습니다.");return;}
  try{
    await navigator.clipboard.writeText(text);
  }catch(e){
    // 클립보드 권한이 막힌 환경을 위한 대비책
    const ta=document.createElement("textarea");
    ta.value=text; ta.style.position="fixed"; ta.style.opacity="0";
    document.body.appendChild(ta); ta.select();
    try{document.execCommand("copy");}catch(e2){toast("복사하지 못했습니다.");}
    document.body.removeChild(ta);
  }
  toast(label);
}

// 복사 단추는 어느 표에 있든 여기서 받는다. 행 선택과 겹치지 않게 막는다
document.addEventListener("click",e=>{
  const b=e.target.closest("button.cp");
  if(!b)return;
  e.stopPropagation();
  copyText(b.dataset.c, b.dataset.c+" 복사");
},true);
const esc=s=>String(s==null?"":s).replace(/[&<>"]/g,c=>
  ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
let TAB="screens", CACHE={}, LOADED=false;

async function jget(u){const r=await fetch(u);if(!r.ok)throw new Error(r.status);return r.json();}
async function jpost(u,b){return fetch(u,{method:"POST",
  headers:{"Content-Type":"application/json"},body:JSON.stringify(b||{})});}

(async function boot(){
  const c=await jget("/api/config");
  $("#ui_dir").value=c.ui_dir||"";
  $("#tbl_dir").value=c.tbl_dir||"";
  $("#use_cache").checked=!!c.use_cache;
  document.title=document.title+" "+c.version;
  poll();
})();

$("#go").onclick=async()=>{
  $("#go").disabled=true; $("#msg").textContent="";
  CACHE={}; LOADED=false;
  const r=await jpost("/api/read",{ui_dir:$("#ui_dir").value,
    tbl_dir:$("#tbl_dir").value,use_cache:$("#use_cache").checked});
  if(!r.ok){$("#msg").innerHTML='<span class="bad">'+
    esc((await r.json()).error)+"</span>"; $("#go").disabled=false;}
};

$("#resave").onclick=async()=>{
  $("#resave").disabled=true;
  const r=await jpost("/api/save",{pretty:$("#pretty").checked});
  const j=await r.json();
  $("#msg").innerHTML=r.ok?"":('<span class="bad">'+esc(j.error)+"</span>");
  $("#resave").disabled=false;
};

$("#clear").onclick=async()=>{
  const j=await(await jpost("/api/cache/clear")).json();
  $("#msg").textContent=j.n+"개를 지웠습니다.";
};

async function poll(){
  try{
    const j=await jget("/api/status");
    for(const k of ["screens","tables","cols","groups","problems"])
      $("#c_"+k).textContent=j.counts[k];
    $("#stage").textContent=j.stage;
    if(j.saved){
      $("#saved").style.display="flex";
      $("#saved_path").textContent=j.saved+"  ("+j.saved_mb+" MB)";
    }
    $("#files").innerHTML=j.files.map(f=>`<tr><td>${esc(f.kind)}</td>
      <td>${esc(f.name)}</td><td>${f.count}</td>
      <td>${(f.ms/1000).toFixed(1)}초</td>
      <td>${f.cached?'<span class="pill">캐시</span>':""}${esc(f.note)}</td></tr>`).join("");
    $("#logs").textContent=(j.logs||[]).join("\n");
    $("#logs").scrollTop=$("#logs").scrollHeight;
    $("#go").disabled=j.running;
    if(j.error)$("#msg").innerHTML='<span class="bad">'+esc(j.error)+"</span>";
    if(j.done&&!j.running&&!LOADED){LOADED=true;show(TAB);}
  }catch(e){}
  fetch("/api/_beat").catch(()=>{});
  setTimeout(poll,900);
}

document.querySelectorAll(".tabs button").forEach(b=>b.onclick=()=>{
  document.querySelectorAll(".tabs button").forEach(x=>x.classList.remove("on"));
  b.classList.add("on"); TAB=b.dataset.t; $("#q").value="";
  $("#mapbar").style.display=(TAB==="mapping"?"inline":"none");
  $("#copybar").style.display=(TAB==="screens"||TAB==="tables"?"inline":"none");
  show(TAB);
});
$("#q").oninput=()=>render();
$("#csv").onclick=()=>{location.href="/api/mapping.csv";};
$("#copyall").onclick=()=>{
  const D=CACHE[TAB];
  if(!D){toast("먼저 산출물을 읽으십시오.");return;}
  const key=TAB==="screens"?"id":"name";
  const keys=TAB==="screens"?["id","name","file","events"]
                            :["name","kor","group","file"];
  const list=filt(D.rows,keys).map(r=>r[key]);
  copyText(list.join("\n"), `${list.length}개를 복사했습니다.`);
};
$("#clearsel").onclick=()=>{SEL={screen:null,group:null};render();};
let SEL={screen:null,group:null};

async function show(t){
  if(!LOADED){$("#pane").innerHTML='<div class="hint">먼저 산출물을 읽으십시오.</div>';return;}
  if(!CACHE[t]){
    $("#pane").innerHTML='<div class="hint">불러오는 중…</div>';
    const url={screens:"/api/screens",tables:"/api/tables",
               link:"/api/link",problems:"/api/problems",
               mapping:"/api/mapping"}[t];
    CACHE[t]=await jget(url);
  }
  render();
}

function filt(rows,keys){
  const q=$("#q").value.trim().toLowerCase();
  if(!q)return rows;
  return rows.filter(r=>keys.some(k=>String(r[k]||"").toLowerCase().includes(q)));
}

function frame(head,rows,body){
  return `<div class="split"><div class="scroll"><table class="grid">
    <thead><tr>${head}</tr></thead><tbody id="rows">${rows.map(body).join("")}</tbody>
    </table></div><div id="side" class="scroll"><div class="side hint">
    행을 누르면 자세히 봅니다.</div></div></div>`;
}

function bind(fn){
  document.querySelectorAll("#rows tr").forEach(tr=>tr.onclick=()=>{
    document.querySelectorAll("#rows tr").forEach(x=>x.classList.remove("sel"));
    tr.classList.add("sel"); fn(tr.dataset.k);
  });
}

function render(){
  const D=CACHE[TAB];
  if(!D)return;

  if(TAB==="screens"){
    const rows=filt(D.rows,["id","name","file","events"]);
    $("#cnt").textContent=rows.length+" / "+D.rows.length+" 화면";
    $("#pane").innerHTML=frame(
      `<th style="width:30px"></th>
       <th style="width:104px">UI ID</th><th style="width:120px">화면명</th>
       <th>단위프로세스명</th>
       <th style="width:64px">슬라이드</th><th style="width:48px">속성</th>
       <th style="width:48px">기능</th><th style="width:44px">표</th>
       <th style="width:62px">미확인표</th>`,
      rows, r=>`<tr data-k="${esc(r.id)}" class="${r.miss?"miss":""}">
        <td class="cpc"><button class="cp" data-c="${esc(r.id)}"
          title="UI ID 복사">⧉</button></td>
        <td title="${esc(r.id)}">${esc(r.id)}</td>
        <td title="${esc(r.name)}">${esc(r.name)}</td>
        <td title="${esc(r.events)}" class="${r.events?"":"dim"}">${esc(r.events)||"—"}</td>
        <td>${esc(r.slides)}</td><td>${r.attrs}</td>
        <td>${r.funcs}</td><td>${r.tables}</td>
        <td class="${r.miss?"bad":"dim"}">${r.miss||"—"}</td></tr>`);
    bind(detailScreen);
  }

  if(TAB==="tables"){
    const rows=filt(D.rows,["name","kor","group","file"]);
    $("#cnt").textContent=rows.length+" / "+D.rows.length+" 테이블";
    $("#pane").innerHTML=frame(
      `<th style="width:30px"></th>
       <th style="width:150px">테이블명</th><th>한글명</th>
       <th style="width:108px">논리파일 그룹</th><th style="width:52px">그룹표</th>
       <th style="width:48px">컬럼</th><th style="width:62px">화면참조</th>`,
      rows, r=>`<tr data-k="${esc(r.name)}">
        <td class="cpc"><button class="cp" data-c="${esc(r.name)}"
          title="테이블명 복사">⧉</button></td>
        <td title="${esc(r.name)}">${esc(r.name)}</td>
        <td title="${esc(r.kor)}">${esc(r.kor)}</td>
        <td title="${esc(r.group)}">${esc(r.group)}</td><td>${r.group_n}</td>
        <td>${r.cols}</td>
        <td class="${r.used?"":"dim"}">${r.used?"있음":"없음"}</td></tr>`);
    bind(detailTable);
  }

  if(TAB==="mapping"){
    const q=$("#q").value.trim().toLowerCase();
    const hit=(r,keys)=>!q||keys.some(k=>String(r[k]||"").toLowerCase().includes(q));

    const sum=D.summary.filter(r=>hit(r,["screen","name","file","events"]));
    const log=D.logical.filter(r=>
      (!SEL.screen||r.screen===SEL.screen)&&hit(r,["screen","screen_name","group","tables_text"]));
    const phy=D.physical.filter(r=>
      (!SEL.screen||r.screen===SEL.screen)&&
      (!SEL.group||r.group===SEL.group)&&
      hit(r,["screen","screen_name","events","table","kor","group"]));

    const gap=D.summary.filter(r=>r.ftr<r.tables).length;
    $("#cnt").textContent=`화면 ${sum.length} · 논리 ${log.length} · 물리 ${phy.length}`
      +(SEL.screen?` · 선택 ${SEL.screen}${SEL.group?" / "+SEL.group:""}`:"");

    $("#pane").innerHTML=`
      <div class="hint" style="margin-bottom:6px">
        FTR 은 논리파일 단위로 셉니다. 물리 표 수와 FTR 후보가 다른 화면이 ${gap}개입니다.
        정의서에 없는 표는 FTR 에 세지 않았습니다.
        화면 행을 누르면 오른쪽이 그 화면만, 논리 행을 누르면 물리 표가 그 그룹만 남습니다.</div>
      <div class="map3">
        <div class="mapbox">
          <h5>화면 요약 <span class="hint">${sum.length}</span></h5>
          <div class="scroll" style="max-height:52vh"><table class="grid">
          <thead><tr><th style="width:100px">화면ID</th>
          <th style="width:110px">화면명</th><th>단위프로세스명</th>
          <th style="width:52px">물리표</th><th style="width:62px">FTR 후보</th>
          <th style="width:62px">DET 상한</th><th style="width:58px">미확인</th></tr></thead>
          <tbody id="m_sum">${sum.map(r=>`
            <tr data-s="${esc(r.screen)}"
                class="${r.missing?"miss":""} ${SEL.screen===r.screen?"sel":""}">
              <td title="${esc(r.screen)}">${esc(r.screen)}</td>
              <td title="${esc(r.name)}">${esc(r.name)}</td>
              <td title="${esc(r.events)}" class="${r.events?"":"dim"}">
                ${esc(r.events)||"—"}</td>
              <td>${r.tables}</td>
              <td class="${r.ftr<r.tables?"bad":""}">${r.ftr}</td>
              <td>${r.det_max}</td>
              <td class="${r.missing?"bad":"dim"}">${r.missing||"—"}</td></tr>`).join("")
            ||'<tr><td colspan="7" class="dim">없음</td></tr>'}</tbody></table></div>
        </div>
        <div class="mapcol">
          <div class="mapbox">
            <h5>논리 기준 <span class="hint">${log.length}</span></h5>
            <div class="scroll" style="max-height:24vh"><table class="grid">
            <thead><tr><th style="width:100px">화면ID</th>
            <th style="width:104px">논리파일 그룹</th><th style="width:52px">쓴표수</th>
            <th>쓴 물리 테이블</th><th style="width:56px">그룹DET</th>
            <th style="width:60px">참조속성</th><th style="width:56px">미확인</th></tr></thead>
            <tbody id="m_log">${log.map(r=>`
              <tr data-s="${esc(r.screen)}" data-g="${esc(r.group)}"
                  class="${r.missing?"miss":""} ${SEL.group===r.group&&SEL.screen===r.screen?"sel":""}">
                <td title="${esc(r.screen)}">${esc(r.screen)}</td>
                <td title="${esc(r.group)}">${esc(r.group)||'<span class="dim">그룹없음</span>'}</td>
                <td>${r.table_n}</td>
                <td title="${esc(r.tables_text)}">${esc(r.tables_text)}</td>
                <td>${r.group_det==null?"":r.group_det}</td><td>${r.attrs}</td>
                <td class="${r.missing?"bad":"dim"}">${r.missing||"—"}</td></tr>`).join("")
              ||'<tr><td colspan="7" class="dim">없음</td></tr>'}</tbody></table></div>
          </div>
          <div class="mapbox">
            <h5>물리 기준 <span class="hint">${phy.length}</span></h5>
            <div class="scroll" style="max-height:24vh"><table class="grid">
            <thead><tr><th style="width:100px">화면ID</th>
            <th style="width:140px">물리 테이블</th><th>한글명</th>
            <th style="width:56px">업무</th><th style="width:56px">정의서</th>
            <th style="width:104px">논리파일 그룹</th><th style="width:52px">컬럼</th>
            <th style="width:60px">참조속성</th></tr></thead>
            <tbody>${phy.map(r=>`
              <tr class="${r.table&&!r.known?"miss":""}">
                <td title="${esc(r.screen)}">${esc(r.screen)}</td>
                <td class="${r.known?"":"bad"}" title="${esc(r.table)}">
                  ${esc(r.table)||'<span class="dim">표 없음</span>'}</td>
                <td title="${esc(r.kor)}">${esc(r.kor)}</td>
                <td>${esc(r.system)}</td>
                <td class="${r.known?"":"bad"}">${r.table?(r.known?"있음":"없음"):"—"}</td>
                <td title="${esc(r.group)}">${esc(r.group)||'<span class="dim">—</span>'}</td>
                <td>${r.cols==null?"":r.cols}</td><td>${r.attrs}</td></tr>`).join("")
              ||'<tr><td colspan="8" class="dim">없음</td></tr>'}</tbody></table></div>
          </div>
        </div>
      </div>`;

    document.querySelectorAll("#m_sum tr[data-s]").forEach(tr=>tr.onclick=()=>{
      SEL={screen:SEL.screen===tr.dataset.s?null:tr.dataset.s,group:null};render();});
    document.querySelectorAll("#m_log tr[data-s]").forEach(tr=>tr.onclick=()=>{
      const same=SEL.screen===tr.dataset.s&&SEL.group===tr.dataset.g;
      SEL=same?{screen:null,group:null}:{screen:tr.dataset.s,group:tr.dataset.g};
      render();});
  }

  if(TAB==="link"){
    const m=filt(D.missing,["name","screens"]), u=filt(D.unused,["name","kor"]);
    $("#cnt").textContent=`화면이 쓰는 테이블 ${D.used_n}개`;
    $("#pane").innerHTML=`
      <div class="hint" style="margin-bottom:6px">
        오류 목록이 아닙니다. 산출물 사이의 어긋남만 보여 줍니다.
        표기 오류일 수도 있고, 배치나 인터페이스로만 쓰는 테이블일 수도 있습니다.</div>
      <div class="split" style="grid-template-columns:1fr 1fr">
        <div><b style="font-size:12px">화면에는 있으나 테이블정의서에 없음
          (${D.missing.length})</b>
          <div class="scroll" style="margin-top:4px"><table>
          <thead><tr><th style="width:170px">테이블명</th><th style="width:55px">화면수</th>
          <th>나온 화면</th></tr></thead><tbody>
          ${m.map(r=>`<tr><td class="bad">${esc(r.name)}</td><td>${r.n}</td>
            <td>${esc(r.screens)}</td></tr>`).join("")
            ||'<tr><td colspan="3" class="dim">없음</td></tr>'}
          </tbody></table></div></div>
        <div><b style="font-size:12px">테이블정의서에는 있으나 화면에 안 나옴
          (${D.unused.length})</b>
          <div class="scroll" style="margin-top:4px"><table>
          <thead><tr><th style="width:170px">테이블명</th><th>한글명</th></tr></thead>
          <tbody>${u.map(r=>`<tr><td>${esc(r.name)}</td>
            <td>${esc(r.kor)}</td></tr>`).join("")
            ||'<tr><td colspan="2" class="dim">없음</td></tr>'}
          </tbody></table></div></div></div>`;
  }

  if(TAB==="problems"){
    const rows=D.rows.filter(r=>r.toLowerCase()
      .includes($("#q").value.trim().toLowerCase()));
    $("#cnt").textContent=rows.length+" / "+D.rows.length+" 건";
    $("#pane").innerHTML=`<div class="scroll"><table><tbody>
      ${rows.map(r=>`<tr><td>${esc(r)}</td></tr>`).join("")
        ||'<tr><td class="dim">없음</td></tr>'}</tbody></table></div>`;
  }
}

async function detailScreen(k){
  const d=await jget("/api/screens/"+encodeURIComponent(k));
  $("#side").innerHTML=`<div class="side">
    <b>${esc(d.id)}</b> · ${esc(d.name)}
    <div class="hint">${esc(d.file)} · 슬라이드 ${d.slides.join(", ")}</div>
    <h4>참조 테이블 ${d.tables.length}</h4><table><tbody>
      ${d.tables.map(t=>`<tr><td class="${t.known?"":"bad"}">${esc(t.name)}</td>
        <td>${esc(t.kor)||(t.known?"":"정의서에 없음")}</td></tr>`).join("")
        ||'<tr><td class="dim">없음</td></tr>'}</tbody></table>
    <h4>화면 속성 ${d.attrs.length}</h4><table><tbody>
      ${d.attrs.map(a=>`<tr><td>${esc(a.name)}</td>
        <td class="${a.known?"":"dim"}">${esc(a.table)}</td></tr>`).join("")}
      </tbody></table>
    <h4>처리상세 ${d.funcs.length}</h4><table><tbody>
      ${d.funcs.map(f=>`<tr><td style="width:110px">${esc(f.event)}</td>
        <td style="white-space:normal">${esc(f.detail)}</td></tr>`).join("")}
      </tbody></table></div>`;
}

async function detailTable(k){
  const d=await jget("/api/tables/"+encodeURIComponent(k));
  $("#side").innerHTML=`<div class="side">
    <b>${esc(d.name)}</b> · ${esc(d.kor)}
    <div class="hint">${esc(d.system)} · ${esc(d.file)}</div>
    <h4>논리파일 그룹</h4>
    <div><span class="pill">${esc(d.group)}</span>
      <span class="pill">RET ${d.ret}</span>
      <span class="pill">DET ${d.det}</span>
      <span class="pill">DET(감사제외) ${d.det_no_audit}</span></div>
    <div class="hint" style="margin-top:3px">그룹 단위 값입니다.
      한 논리파일로 볼지는 감정인이 판단합니다.</div>
    <h4>그룹 구성 테이블 ${d.group_tables.length}</h4><table><tbody>
      ${d.group_tables.map(t=>`<tr><td>${esc(t)}</td></tr>`).join("")}</tbody></table>
    <h4>컬럼 ${d.columns.length}</h4><pre>${esc(d.columns.join(", "))}</pre>
    <h4>이 표를 쓰는 화면 ${d.screens.length}</h4><table><tbody>
      ${d.screens.map(s=>`<tr><td>${esc(s)}</td></tr>`).join("")
        ||'<tr><td class="dim">없음</td></tr>'}</tbody></table></div>`;
}
</script></body></html>"""


def main():
    common.setup_output()
    common.install_excepthooks()
    common.log("시작", f"{APP_TITLE} {VERSION}")
    common.serve(app, port_key="doc_parser_ui", profile=".chrome_doc_parser",
                 title=APP_TITLE, busy=lambda: JOB["running"])


if __name__ == "__main__":
    main()