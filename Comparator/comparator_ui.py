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
import time
import threading

from flask import Flask, request, jsonify, Response

import calculator_ui as U
import comparator as CMP

APP_NAME = "FP 산정 결과물 대조"
VERSION = "1.0"

INI_PATH = os.path.join(U.app_dir(), "comparator.ini")

app = Flask(__name__)
LAST_PING = [time.time()]
RESULT = {"data": None}


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
    cp = load_ini()
    forms = [f for f in U.form_list(list(U.FORMS)) if f["ready"]]
    keys = [{"k": k, "label": v[1]} for k, v in CMP.KEY_FIELDS.items()]
    diffs = [{"k": k, "label": lb} for k, lb in CMP.DIFF_FIELDS]
    return jsonify({"app": APP_NAME, "version": VERSION, "forms": forms,
                    "keys": keys, "diffs": diffs, "opts": CMP.DEFAULT_OPTS,
                    "lastDir": cp["main"].get("last_dir", ""),
                    "labels": [cp["main"].get("label_a", "설계단계"),
                               cp["main"].get("label_b", "종료단계")]})


@app.route("/api/list", methods=["POST"])
def api_list():
    d = (request.json or {}).get("dir", "").strip().strip('"')
    if not d:
        d = load_ini()["main"].get("last_dir", "") or os.getcwd()
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
    cp["main"]["last_dir"] = os.path.abspath(d)
    save_ini(cp)
    return jsonify({"ok": True, "dir": os.path.abspath(d),
                    "parent": os.path.dirname(os.path.abspath(d)),
                    "dirs": dirs[:300], "files": files[:300], "sep": os.sep})


@app.route("/api/browse", methods=["POST"])
def api_browse():
    mode = (request.json or {}).get("mode", "open")
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        ft = [("엑셀 파일", "*.xlsm *.xlsx"), ("모든 파일", "*.*")]
        if mode == "save":
            p = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=ft,
                                             initialfile="FP대조결과.xlsx")
            out = [p] if p else []
        else:
            out = list(filedialog.askopenfilenames(filetypes=ft))
        root.destroy()
        return jsonify({"ok": True, "paths": out})
    except Exception as e:
        return jsonify({"ok": False, "msg": "파일 선택 창을 열 수 없습니다(%s). "
                                            "폴더에서 고르기를 쓰십시오." % e})


@app.route("/api/guess", methods=["POST"])
def api_guess():
    out = {}
    for p in (request.json or {}).get("paths", []):
        try:
            out[p] = U.guess_form(p)
        except Exception:
            out[p] = None
    return jsonify({"ok": True, "forms": out})


@app.route("/api/compare", methods=["POST"])
def api_compare():
    body = request.json or {}
    ga_in = body.get("a", {})
    gb_in = body.get("b", {})
    if not ga_in.get("files") or not gb_in.get("files"):
        return jsonify({"ok": False, "msg": "양쪽 그룹에 파일을 넣으십시오."})
    try:
        ga = CMP.load_group(ga_in["files"], ga_in.get("label") or "A")
        gb = CMP.load_group(gb_in["files"], gb_in.get("label") or "B")
        res = CMP.compare(ga, gb, body.get("opts") or {})
    except Exception as e:
        return jsonify({"ok": False, "msg": "대조 실패: %s" % e})
    RESULT["data"] = res
    cp = load_ini()
    cp["main"]["label_a"] = ga["label"]
    cp["main"]["label_b"] = gb["label"]
    save_ini(cp)
    return jsonify({"ok": True, "data": res})


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
.opt{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin:8px 0}
.opt label{display:inline-flex;align-items:center;gap:4px}
select,input[type=text],input[type=number]{padding:4px 6px;border:1px solid #bbb;border-radius:3px;font-family:inherit;font-size:12px}
.sum{display:flex;gap:12px;flex-wrap:wrap}
.card{border:1px solid #dde3e8;background:#f7f9fb;border-radius:4px;padding:10px 14px;min-width:150px}
.card .v{font-size:20px;font-weight:700;color:#12507b}
.card .l{font-size:12px;color:#555}
.msg{position:fixed;right:16px;bottom:16px;background:#333;color:#fff;padding:9px 14px;border-radius:4px;opacity:0;transition:.25s;z-index:99}
.msg.on{opacity:.95}
.rec{display:block;margin:2px 0;color:#12507b;text-decoration:none;font-size:12px}
small{color:#666}
.warn{color:#c0392b}
#picker{display:none;border:1px solid #ccc;padding:8px;background:#fff;margin-top:8px}
#entries{max-height:220px;overflow:auto;border:1px solid #eee;padding:6px}
</style></head><body>

<div id="bar">
  <b>FP 산정 결과물 대조</b>
  <span id="state">파일을 넣고 대조를 실행하십시오.</span>
  <button onclick="exportXlsx()">엑셀로 내보내기</button>
  <button onclick="quit()">종료</button>
</div>

<div id="tabs">
  <div class="tab on" data-p="p1" onclick="tab('p1')">① 대상 지정</div>
  <div class="tab" data-p="p2" onclick="tab('p2')">② 요약</div>
  <div class="tab" data-p="p3" onclick="tab('p3')">③ 매핑</div>
  <div class="tab" data-p="p4" onclick="tab('p4')">④ 중복 점검</div>
</div>

<div class="pane on" id="p1">
  <h3>대조할 산출물</h3>
  <div class="cols">
    <div class="col">
      <h4>A 그룹 <small>(앞 단계)</small></h4>
      <input class="nm" id="labA" value="설계단계">
      <p style="margin:8px 0 0"><button onclick="addFiles('a')">파일 추가</button>
        <button onclick="openPicker('a')">폴더에서 고르기</button></p>
      <div class="files" id="fa"></div>
    </div>
    <div class="col">
      <h4>B 그룹 <small>(뒤 단계)</small></h4>
      <input class="nm" id="labB" value="종료단계">
      <p style="margin:8px 0 0"><button onclick="addFiles('b')">파일 추가</button>
        <button onclick="openPicker('b')">폴더에서 고르기</button></p>
      <div class="files" id="fb"></div>
    </div>
  </div>

  <div id="picker">
    <p style="margin:0 0 6px"><input type="text" id="dir" style="width:70%">
      <button onclick="listDir(document.getElementById('dir').value)">이동</button>
      <span id="pickSide" style="margin-left:8px;font-weight:700"></span></p>
    <div id="entries"></div>
    <p style="margin:6px 0 0"><button onclick="addChecked()">선택한 파일 넣기</button>
      <button onclick="document.getElementById('picker').style.display='none'">닫기</button></p>
  </div>

  <h3 style="margin-top:18px">대조 기준</h3>
  <div class="opt">
    <label>짝짓기 기준 <select id="key"></select></label>
    <label><input type="checkbox" id="fuzzy" checked> 이름이 비슷하면 같은 기능으로 봄</label>
    <label>유사도 <input type="number" id="th" value="0.85" min="0.5" max="1" step="0.01" style="width:70px"></label>
    <label><input type="checkbox" id="igsp" checked> 공백·기호 무시</label>
    <label><input type="checkbox" id="igbr" checked> 괄호 무시</label>
  </div>
  <div class="opt">차이로 볼 항목 <span id="diffs"></span></div>
  <p><button class="main" onclick="run()">대조 실행</button> <span id="runMsg"></span></p>
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
    <span>검색 <input type="text" id="q" placeholder="기능명·파일명"></span>
    <span id="cnt" style="margin-left:auto"></span>
  </div>
  <div id="wrap"><table id="tb"></table></div>
</div>

<div class="pane" id="p4"><div id="v4"><p>대조를 실행하면 그룹 안의 중복이 나옵니다.</p></div></div>

<div class="msg" id="msg"></div>

<script>
var FORMS=[], KEYS=[], DIFFS=[], OPTS={}, G={a:[],b:[]}, R=null, PICK='a';

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
      var o='';
      FORMS.forEach(function(x){o+='<option value="'+x.key+'"'+(f.form===x.key?' selected':'')+'>'+esc(x.code+' '+x.name)+'</option>';});
      h+='<div class="file"><span class="n" title="'+esc(f.path)+'">'+esc(base(f.path))+'</span>'+
         '<select onchange="G[\''+s+'\']['+i+'].form=this.value">'+o+'</select>'+
         '<span class="x" onclick="delFile(\''+s+'\','+i+')">×</span></div>';
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
    var skipped=0;
    paths.forEach(function(p){
      if(G[s].some(function(f){return f.path===p;})){skipped++;return;}
      var f=j.forms[p];
      if(!f){ f=FORMS.length?FORMS[0].key:''; }
      G[s].push({path:p,form:f});
    });
    paintFiles();
    toast(paths.length-skipped+'개 넣었습니다.'+(skipped?' (중복 '+skipped+'개 제외)':''));
  });
}
function addFiles(s){
  fetch('/api/browse',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({mode:'open'})}).then(function(r){return r.json();}).then(function(j){
    if(!j.ok){toast(j.msg);openPicker(s);return;}
    addPaths(s,j.paths||[]);});
}
function openPicker(s){PICK=s;
  document.getElementById('pickSide').textContent='→ '+(s==='a'?document.getElementById('labA').value:document.getElementById('labB').value)+' 로 넣기';
  document.getElementById('picker').style.display='block';
  listDir(document.getElementById('dir').value);}
function listDir(d){
  fetch('/api/list',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({dir:d})}).then(function(r){return r.json();}).then(function(j){
    if(!j.ok){document.getElementById('entries').innerHTML='<span class="warn">'+esc(j.msg)+'</span>';return;}
    document.getElementById('dir').value=j.dir;
    var h='';
    if(j.parent&&j.parent!==j.dir)
      h+='<a class="rec" href="#" onclick="listDir('+JSON.stringify(j.parent)+');return false;">.. 상위 폴더</a>';
    j.dirs.forEach(function(n){
      h+='<a class="rec" href="#" onclick="listDir('+JSON.stringify(j.dir+j.sep+n)+');return false;">[폴더] '+esc(n)+'</a>';});
    j.files.forEach(function(n){
      h+='<label style="display:block"><input type="checkbox" class="pf" value="'+esc(j.dir+j.sep+n)+'"> '+esc(n)+'</label>';});
    if(!j.dirs.length&&!j.files.length) h='<small>이 폴더에 엑셀 파일이 없습니다.</small>';
    document.getElementById('entries').innerHTML=h;});
}
function addChecked(){
  var ps=[];
  document.querySelectorAll('.pf').forEach(function(c){if(c.checked)ps.push(c.value);});
  if(!ps.length){toast('파일을 고르십시오.');return;}
  addPaths(PICK,ps);
  document.querySelectorAll('.pf').forEach(function(c){c.checked=false;});
}

/* ---- 대조 실행 ---- */
function opts(){
  var d=[];
  document.querySelectorAll('.df').forEach(function(c){if(c.checked)d.push(c.value);});
  return {key:document.getElementById('key').value,
          fuzzy:document.getElementById('fuzzy').checked,
          threshold:parseFloat(document.getElementById('th').value)||0.85,
          ignore_space:document.getElementById('igsp').checked,
          ignore_bracket:document.getElementById('igbr').checked,
          diff:d};
}
function run(){
  if(!G.a.length||!G.b.length){toast('양쪽 그룹에 파일을 넣으십시오.');return;}
  document.getElementById('runMsg').textContent=' 읽는 중입니다...';
  fetch('/api/compare',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({a:{label:document.getElementById('labA').value,files:G.a},
                         b:{label:document.getElementById('labB').value,files:G.b},
                         opts:opts()})}).then(function(r){return r.json();}).then(function(j){
    document.getElementById('runMsg').textContent='';
    if(!j.ok){alert(j.msg);return;}
    R=j.data; paint(); tab('p2');});
}

/* ---- 결과 ---- */
function stClass(s){return s==='A만 있음'?'st-A':(s==='B만 있음'?'st-B':'st-'+s);}
function paint(){
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
  [[a,a.label],[b,b.label]].forEach(function(g){
    g[0].files.forEach(function(f){
      h+='<tr><td>'+esc(g[1])+'</td><td>'+esc(f.name)+'</td><td>'+esc(f.formName)+
         '</td><td>'+esc(f.sheet)+'</td><td>'+esc(f.method)+'</td><td class="right">'+f.count+'</td></tr>';});
    (g[0].errors||[]).forEach(function(e){
      h+='<tr><td>'+esc(g[1])+'</td><td colspan="5" class="warn">'+esc(e.path)+' — '+esc(e.msg)+'</td></tr>';});
  });
  h+='</table><p><small>짝짓기 기준 : '+esc(R.keyLabel)+
     ' &nbsp;|&nbsp; 유사도 임계값 '+R.opts.threshold+'</small></p>';
  document.getElementById('v2').innerHTML=h;
  paintDup(); paintTable();
}
function paintDup(){
  var h='';
  [[R.dupA,R.a.label],[R.dupB,R.b.label]].forEach(function(g){
    h+='<h3 style="margin-top:14px">'+esc(g[1])+' 안의 중복 '+g[0].length+'건</h3>';
    if(!g[0].length){h+='<p><small>같은 이름의 기능이 두 번 이상 나오지 않습니다.</small></p>';return;}
    h+='<table><tr><th>파일</th><th>행</th><th>애플리케이션</th><th>세부업무</th><th>단위프로세스명</th><th>FP유형</th><th>가중치</th></tr>';
    g[0].forEach(function(d){
      d.rows.forEach(function(r){
        h+='<tr><td>'+esc(r.src)+'</td><td class="right">'+r.no+'</td><td>'+esc(r.app)+'</td><td>'+esc(r.biz)+
           '</td><td>'+esc(r.proc)+'</td><td>'+esc(r.type)+'</td><td class="right">'+(r.wt===null?'':r.wt)+'</td></tr>';});
      h+='<tr><td colspan="7" style="background:#fafafa"></td></tr>';});
    h+='</table>';});
  document.getElementById('v4').innerHTML=h;
}
function paintTable(){
  if(!R)return;
  var on={}; document.querySelectorAll('.fs').forEach(function(c){on[c.value]=c.checked;});
  var q=document.getElementById('q').value.trim().toLowerCase();
  var h=['<tr><th rowspan="2" style="width:70px">상태</th><th rowspan="2" style="width:52px">유사도</th>',
    '<th colspan="5">'+esc(R.a.label)+'</th><th colspan="5">'+esc(R.b.label)+'</th>',
    '<th rowspan="2" style="width:180px">차이</th></tr>',
    '<tr><th style="width:120px">파일·행</th><th style="width:110px">애플리케이션</th><th style="width:100px">세부업무</th><th>단위프로세스명</th><th style="width:90px">유형·복잡도</th>',
    '<th style="width:120px">파일·행</th><th style="width:110px">애플리케이션</th><th style="width:100px">세부업무</th><th>단위프로세스명</th><th style="width:90px">유형·복잡도</th></tr>'];
  var n=0;
  R.pairs.forEach(function(p){
    if(!on[p.status])return;
    if(q){var s=JSON.stringify([p.a,p.b]).toLowerCase(); if(s.indexOf(q)<0)return;}
    n++;
    function cell(r,cls){
      if(!r) return '<td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td><td class="'+cls+'"></td>';
      return '<td class="'+cls+'" title="'+esc(r.src)+'">'+esc(r.src.length>16?r.src.substring(0,14)+'…':r.src)+' '+r.no+'행</td>'+
             '<td class="'+cls+'">'+esc(r.app)+'</td><td class="'+cls+'">'+esc(r.biz)+'</td>'+
             '<td class="'+cls+'">'+esc(r.proc)+'</td>'+
             '<td class="'+cls+'">'+esc(r.type)+(r.cx?' / '+esc(r.cx):'')+(r.wt!==null?' / '+r.wt:'')+'</td>';
    }
    var d=p.diffs.map(function(x){return esc(x.label)+' '+esc(x.a||'-')+'→'+esc(x.b||'-');}).join('<br>');
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
  FORMS=j.forms; KEYS=j.keys; DIFFS=j.diffs; OPTS=j.opts;
  document.getElementById('labA').value=j.labels[0];
  document.getElementById('labB').value=j.labels[1];
  document.getElementById('dir').value=j.lastDir||'';
  var h='';KEYS.forEach(function(k){h+='<option value="'+k.k+'"'+(k.k===OPTS.key?' selected':'')+'>'+esc(k.label)+'</option>';});
  document.getElementById('key').innerHTML=h;
  h='';DIFFS.forEach(function(d){
    h+='<label style="margin-right:12px"><input type="checkbox" class="df" value="'+d.k+'"'+
       (OPTS.diff.indexOf(d.k)>=0?' checked':'')+'> '+esc(d.label)+'</label>';});
  document.getElementById('diffs').innerHTML=h;
  document.getElementById('th').value=OPTS.threshold;
  paintFiles();
});
</script></body></html>
"""


def main():
    U.setup_output()
    no_browser = "--no-browser" in os.sys.argv
    port = U.pick_port()
    url = "http://127.0.0.1:%d/" % port
    print("=" * 60)
    print(APP_NAME, "v" + VERSION)
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
