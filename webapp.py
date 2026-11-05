# -*- coding: utf-8 -*-
"""区域车辆录入工具 · 本地网页版

双击后在本机起一个很小的服务, 用系统浏览器操作:
  http://127.0.0.1:8787

入库 / 末销 / 看板 / 校验 四个页面, 与命令行同一套逻辑。
页面关闭 30 秒后进程自动退出, 不常驻。
"""
import contextlib
import datetime as dt
import io
import json
import os
import re
import sys
import threading
import time
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, unquote, urlparse

PORT = 8787
VERSION = '1.7'


def workspace_dir():
    """绿色版: app 同级有 data/inbox 就用它; 安装到 /Applications 时用
    ~/Documents/车辆录入工具, 首次启动把原项目的数据/配置迁过去。"""
    if getattr(sys, 'frozen', False):
        exe = os.path.abspath(sys.executable)
        app_dir = os.path.dirname(os.path.dirname(os.path.dirname(exe)))
        if os.path.isdir(os.path.join(app_dir, 'data', 'inbox')):
            return app_dir
        home = os.path.expanduser('~/Documents/车辆录入工具')
        bootstrap(home)
        return home
    return os.path.dirname(os.path.abspath(__file__))


def bootstrap(dst):
    """首次启动: 若本机已有项目目录，把库存末销数据盘点与配置复制过来。"""
    import shutil
    if os.path.isdir(os.path.join(dst, 'data', 'inbox')):
        return
    src = os.path.expanduser('~/Documents/工作/末销数据传输')
    if not os.path.isdir(src):
        return
    try:
        for sub in ('data', 'config'):
            s = os.path.join(src, sub)
            if os.path.isdir(s):
                shutil.copytree(s, os.path.join(dst, sub), dirs_exist_ok=True)
    except Exception:
        pass


WS = workspace_dir()
for sub in ('data/inbox', 'data/ledger', 'config/countries', 'out'):
    os.makedirs(os.path.join(WS, sub), exist_ok=True)

import vehicle_flow as VF          # noqa: E402
import dashboard as DB             # noqa: E402

VF.BASE = WS
VF.INBOX = os.path.join(WS, 'data', 'inbox')
VF.LEDGER_DIR = os.path.join(WS, 'data', 'ledger')
VF.CONFIG_DIR = os.path.join(WS, 'config', 'countries')
VF.OUT_DIR = os.path.join(WS, 'out')
DB.DASH_DIR = os.path.join(WS, 'out', 'dashboard')
DB.TARGETS_PATH = os.path.join(WS, 'config', 'targets.json')

LAST_PING = [time.time()]


# ---------------------------------------------------------------- 工具

def _norm_path(p):
    """归一化路径: 解析符号链接/联接点 + (Windows) 短名转长名 + 大小写折叠。
    任何一步失败都回退到上一步的结果，绝不抛异常。"""
    try:
        p = os.path.realpath(os.path.abspath(p))
    except (OSError, ValueError):
        p = os.path.abspath(p)
    if os.name == 'nt':
        try:
            import ctypes
            from ctypes import wintypes
            GetLongPathNameW = ctypes.windll.kernel32.GetLongPathNameW
            GetLongPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
            GetLongPathNameW.restype = wintypes.DWORD
            buf = ctypes.create_unicode_buffer(1024)
            if GetLongPathNameW(p, buf, 1024):
                p = buf.value
        except Exception:
            pass
    return os.path.normcase(p)


def inside_ws(path):
    """只允许访问工作目录内的文件（兼容 Windows 大小写/短名/联接点/OneDrive）。"""
    try:
        p = _norm_path(path)
        ws = _norm_path(WS)
        return (p == ws or p.startswith(ws + os.sep)) and os.path.isfile(path)
    except (TypeError, ValueError):
        return False


def rel_files(files):
    return [{'name': os.path.basename(f), 'path': f,
             'rel': os.path.relpath(f, WS)} for f in files if inside_ws(f)]


def run_capture(func, args=None):
    buf = io.StringIO()
    ok = True
    try:
        with contextlib.redirect_stdout(buf):
            rc = (func(args) if args is not None else func())
        ok = rc in (0, None)
    except SystemExit as e:
        ok = e.code in (0, None)
        buf.write('[中止] %s\n' % (e.code if e.code else ''))
    except Exception:
        ok = False
        buf.write(traceback.format_exc())
    return ok, buf.getvalue()


def parse_multipart(body, boundary):
    """极简 multipart/form-data 解析, 避免依赖已废弃的 cgi 模块。"""
    out = {}
    sep = b'--' + boundary
    for part in body.split(sep):
        part = part.strip(b'\r\n')
        if not part or part == b'--':
            continue
        head, _, data = part.partition(b'\r\n\r\n')
        if not _:
            continue
        data = data[:-2] if data.endswith(b'\r\n') else data
        name, filename = None, None
        for line in head.split(b'\r\n'):
            if b'name="' in line:
                name = line.split(b'name="')[1].split(b'"')[0].decode('utf-8', 'ignore')
            if b'filename="' in line:
                filename = line.split(b'filename="')[1].split(b'"')[0].decode('utf-8', 'ignore')
        if name is None:
            continue
        val = (filename, data) if filename else data.decode('utf-8', 'ignore')
        if name in out:                     # 同名多文件: 累积成列表
            out[name] = out[name] + [val] if isinstance(out[name], list) else [out[name], val]
        else:
            out[name] = val
    return out


# ---------------------------------------------------------------- HTML

PAGE = r"""<!doctype html>
<html lang="zh"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>区域车辆录入工具</title>
<style>
:root{--acc:#1f6feb;--bg:#f6f7f9;--line:#e5e7eb;--sub:#6b7280;--ok:#16a34a;--bad:#dc2626}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:#111827;
  font:14px/1.6 -apple-system,"PingFang SC","Microsoft YaHei",system-ui,sans-serif}
header{background:#fff;border-bottom:1px solid var(--line);padding:18px 0 0}
.wrap{max-width:940px;margin:0 auto;padding:0 22px}
h1{font-size:20px;margin:0 0 2px;font-weight:650;letter-spacing:.2px}
.sub{color:var(--sub);font-size:12.5px}
.snap{display:flex;gap:8px;align-items:center;margin:14px 0 0;padding-bottom:14px}
.snap select{flex:1}
nav{display:flex;gap:4px;max-width:940px;margin:0 auto;padding:0 22px}
nav button{background:none;border:0;border-bottom:2px solid transparent;
  padding:10px 14px;font-size:14px;color:var(--sub);cursor:pointer;font-family:inherit}
nav button.on{color:var(--acc);border-bottom-color:var(--acc);font-weight:600}
main{max-width:940px;margin:14px auto 40px;padding:0 22px}
.card{background:#fff;border:1px solid var(--line);border-radius:12px;padding:18px 20px;
  margin-bottom:14px}
h2{font-size:15px;margin:0 0 4px;font-weight:650}
label{display:block;font-size:12.5px;color:var(--sub);margin:14px 0 5px}
input[type=text],select,textarea{width:100%;padding:8px 10px;border:1px solid var(--line);
  border-radius:8px;font:inherit;background:#fff;outline:none}
input:focus,select:focus,textarea:focus{border-color:var(--acc)}
textarea{resize:vertical;min-height:96px;font-family:Menlo,Consolas,monospace;font-size:13px}
.row{display:flex;gap:10px;flex-wrap:wrap;align-items:center}
.row>*{flex:0 0 auto}
.row .grow{flex:1}
.btn{display:inline-flex;align-items:center;gap:6px;background:#fff;border:1px solid var(--line);
  color:#111827;padding:8px 15px;border-radius:8px;font:inherit;cursor:pointer}
.btn:hover{border-color:#cbd5e1;background:#fbfcfd}
.btn.pri{background:var(--acc);border-color:var(--acc);color:#fff;font-weight:600}
.btn.pri:hover{background:#1a5fd0}
.btn:disabled{opacity:.55;cursor:not-allowed}
.radios{display:flex;gap:18px;align-items:center;margin-top:6px}
.radios label{margin:0;color:#111827;font-size:14px;display:flex;gap:6px;align-items:center}
.hint{color:var(--sub);font-size:12.5px;margin-top:10px;line-height:1.7}
.log{background:#0f172a;color:#e6edf7;border-radius:10px;padding:14px 16px;
  font-family:Menlo,Consolas,monospace;font-size:12.5px;line-height:1.7;
  height:300px;overflow:auto;white-space:pre-wrap;word-break:break-all}
.log .g{color:#4ade80}.log .r{color:#f87171}
.files{margin-top:12px}
.files a{display:inline-flex;align-items:center;gap:6px;background:#eff6ff;color:#1d4ed8;
  border:1px solid #dbeafe;border-radius:8px;padding:7px 12px;margin:4px 6px 0 0;
  text-decoration:none;font-size:13px}
.files a:hover{background:#e0edff}
.tag{display:inline-block;font-size:11.5px;padding:2px 8px;border-radius:20px;
  background:#ecfdf3;color:#15803d;margin-left:8px}
.tag.no{background:#fef2f2;color:#b91c1c}
.hidden{display:none}
.snapnote{background:#f0f6ff;border:1px solid #dbe7fb;border-radius:8px;padding:8px 12px;margin:14px 0 8px;font-size:12.5px;color:#334155;line-height:1.6}
.snapnote b{color:#0f172a}
.snaprow{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin:0 0 4px}
.snaprow .lbl{font-size:13px;color:var(--sub);white-space:nowrap}
.snaprow select{flex:1;min-width:240px}
.snaprow .hint{font-size:12px;color:var(--sub);width:100%;margin-top:2px}
.dashresult{margin-top:16px;padding:16px;border:1px dashed #c7d7f5;border-radius:10px;background:#f8fbff}
.dashresult .big{font-size:15px;font-weight:600;margin-bottom:8px;color:#0f172a}
footer{color:var(--sub);font-size:12px;text-align:center;padding:0 0 26px}
</style></head>
<body>
<header><div class="wrap">
  <h1>区域车辆录入工具</h1>
  <div class="sub">入库 · 末销 · 看板 · 校验　—　一个本地工具，数据不外传</div>
  <div class="snapnote">ℹ️ 先上传平台导出的车辆台账 Excel 作为<b>库存末销数据盘点</b>，后续入库、末销、看板均基于此快照生成。支持各平台模板，自动识别 VIN / 进口商 / 经销商等列。</div>
  <div class="snaprow">
    <span class="lbl">库存末销数据盘点</span>
    <select id="snap" onchange="setSnap(this.value)"></select>
    <button class="btn" onclick="document.getElementById('snapfile').click()">上传快照</button>
    <button class="btn" onclick="loadState()">刷新</button>
    <span class="hint">平台导出的车辆台账（.xlsx / .xls），新 VIN 不在快照里属正常</span>
  </div>
  <input type="file" id="snapfile" accept=".xlsx,.xls" style="display:none" onchange="uploadSnap(this)">
</div>
<nav>
  <button class="on" data-t="intake">入库</button>
  <button data-t="sale">末销</button>
  <button data-t="dash">看板</button>
  <button data-t="check">校验</button>
</nav></header>

<main>
  <section id="p-intake" class="card">
    <h2>入库分店</h2>
    <div class="sub">传入新 VIN，按各店现有库存与网点数自动分配，直接落到「经销商已入库」。</div>
    <div class="hint" style="background:#f8fafc;border:1px solid #eef2f7;border-radius:8px;
      padding:10px 12px;margin:12px 0 0">
      库存末销数据盘点 = 平台导出的当前车辆台账（在库 / 在途 / 已末销），用来算每家店现在有多少库存、
      按网点权重把新车分下去。<b>新 VIN 本来就不在快照里</b>，不需要能查到。
    </div>
    <label>进口商 / 国家</label>
    <select id="inimp"></select>
    <label>VIN 文件（可多选：txt / csv / xlsx / xls）</label>
    <input type="file" id="vinfile" accept=".txt,.csv,.xlsx,.xls" multiple>
    <label>或直接粘贴 VIN（一行可多个，VIN 后可跟车型）</label>
    <textarea id="vintext" placeholder="LDP45BXXXXXXXXXXX  S32&#10;LDP43AXXXXXXXXXXX  S31"></textarea>
    <div class="row" style="margin-top:14px">
      <div><label style="margin:0 0 5px">入库日期</label>
        <input type="text" id="indate" style="width:150px"></div>
      <div><label style="margin:0 0 5px">批次标签（留空则从 PL 文件名自动提取）</label>
        <input type="text" id="intag" style="width:150px"></div>
    </div>
    <div class="row" style="margin-top:18px">
      <button class="btn" onclick="doIntake(true)">预览分配</button>
      <button class="btn pri" onclick="doIntake(false)">生成入库 Excel</button>
    </div>
  </section>

  <section id="p-sale" class="card hidden">
    <h2>末销</h2>
    <div class="sub">两种方式任选其一：按数量生成录入，或上传门店真实末销表识别出库。</div>

    <label>进口商 / 国家</label>
    <select id="imp"></select>

    <label>方式</label>
    <div class="radios">
      <label><input type="radio" name="smode" value="gen" checked onclick="syncSaleMode()"> 按数量生成</label>
      <label><input type="radio" name="smode" value="real" onclick="syncSaleMode()"> 上传真实末销表识别</label>
    </div>

    <div id="sale-gen">
      <label>统计月</label>
      <select id="month"></select>
      <label>末销台数</label>
      <div class="radios">
        <label>我报量 <input type="text" id="count" style="width:80px;margin-left:4px" placeholder="台数"></label>
      </div>
      <div class="hint">超过可执行上限会自动封顶，不会清空任何门店；每店至少保留 1 台。<br>
        末销日期自动均匀铺到当月工作日（周末除外），避免日销峰谷。</div>
      <div class="row" style="margin-top:18px">
        <button class="btn" onclick="doSale(true)">预览方案</button>
        <button class="btn pri" onclick="doSale(false)">生成末销 Excel</button>
      </div>
    </div>

    <div id="sale-real" class="hidden">
      <label>末销表文件（xlsx / xls / csv，可多选）</label>
      <input type="file" id="salefile" accept=".xlsx,.xls,.csv" multiple>
      <div class="hint">表头自动识别：VIN / 车架号、车型 / 商品名称、经销商编码 / 门店、末销时间 / 出库日期、数量。<br>
        带 VIN 的行直接按快照核对状态；只给「经销商 + 数量」的行自动从该店在库车里 FIFO 取。</div>
      <label>默认末销日期（表里没日期时用）</label>
      <input type="date" id="sfdate" style="width:170px">
      <div class="row" style="margin-top:18px">
        <button class="btn" onclick="doSaleFile(true)">预览识别结果</button>
        <button class="btn pri" onclick="doSaleFile(false)">生成末销 Excel</button>
      </div>
    </div>
  </section>

  <section id="p-dash" class="card hidden">
    <h2>区域看板</h2>
    <div class="sub">末销 / 存销比可视化，单文件 HTML，生成后可下载或直接打开。</div>
    <label>统计月</label>
    <div class="row">
      <select id="dmonth"></select>
      <button class="btn pri" onclick="doDash()">生成看板</button>
    </div>
    <div id="dashresult" class="dashresult hidden">
      <div class="big">✅ 看板已生成</div>
      <div id="dashlinks"></div>
    </div>
  </section>

  <section id="p-check" class="card hidden">
    <h2>文件校验</h2>
    <div class="sub">复核已生成的录入文件是否符合平台 8 条导入规则。</div>
    <label>输出目录中的 Excel</label>
    <select id="chkfile"></select>
    <div class="row" style="margin-top:18px">
      <button class="btn pri" onclick="doCheck()">开始校验</button>
    </div>
  </section>

  <div class="card">
    <div class="row" style="justify-content:space-between">
      <h2>运行日志 <span id="stat" class="tag hidden"></span></h2>
      <button class="btn" onclick="document.getElementById('log').textContent=''">清空</button>
    </div>
    <div class="log" id="log">就绪。</div>
    <div class="files" id="files"></div>
  </div>
  <footer><span id="ws"></span> · <a href="#" style="color:#1f6feb;text-decoration:none"
    onclick="fetch('/api/reveal');return false">打开数据目录</a></footer>
</main>

<script>
const $ = s => document.querySelector(s);
let STATE = {snapshot:'', importers:[], outputs:[]};

function log(t, cls){ const el=$('#log'); el.innerHTML +=
  `<span class="${cls||''}">${escapeHtml(t)}</span>`; el.scrollTop = el.scrollHeight; }
function escapeHtml(s){ return s.replace(/&/g,'&amp;').replace(/</g,'&lt;'); }
function mark(ok){ const t=$('#stat'); t.className='tag'+(ok?'':' no');
  t.textContent = ok?'完成':'有错误'; t.classList.remove('hidden'); }

async function loadState(){
  const r = await fetch('/api/state').then(x=>x.json());
  STATE = r;
  $('#snap').innerHTML = r.snapshots.map(s=>`<option value="${s}" ${s===r.snapshot?'selected':''}>${s}</option>`).join('')
    || '<option value="">请先上传库存末销数据盘点（平台导出的车辆台账 .xlsx/.xls）</option>';
  $('#imp').innerHTML = r.importers.map(i=>`<option value="${i.code}">${i.code}　${i.name}</option>`).join('');
  $('#inimp').innerHTML = r.importers.map(i=>`<option value="${i.code}">${i.code}　${i.name}</option>`).join('');
  $('#chkfile').innerHTML = r.outputs.map(o=>`<option value="${o.path}">${o.rel}</option>`).join('');
  $('#ws').textContent = '工作目录：' + r.ws + '　·　v' + (r.version || '?');
  const today = new Date().toISOString().slice(0,10);
  if(!$('#indate').value) $('#indate').value = today;
}
function showFiles(files){
  $('#files').innerHTML = files.map(f=>{
    const isHtml = f.name.endsWith('.html');
    const href = isHtml ? `/view?p=${encodeURIComponent(f.path)}` : `/download?p=${encodeURIComponent(f.path)}`;
    return `<a href="${href}" target="_blank">${isHtml?'打开':'下载'} ${f.name}</a>`;
  }).join('');
}
async function post(url, body, isForm){
  log('\n> ' + url + '\n');
  const opts = isForm ? {method:'POST', body}
                      : {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)};
  const t0 = performance.now();
  const r = await fetch(url, opts).then(x=>x.json()).catch(e=>({ok:false,log:String(e),files:[]}));
  log(r.log || '', r.ok ? 'g' : 'r');
  mark(r.ok);
  showFiles(r.files || []);
  return r;
}
function snapshotArg(){ const v=$('#snap').value; return v ? v : null; }

async function uploadSnap(inp){
  if(!inp.files.length) return;
  const fd = new FormData();
  fd.append('snapshot', inp.files[0]);
  log('\n> 上传库存末销数据盘点\n');
  const r = await fetch('/api/upload_snapshot', {method:'POST', body:fd}).then(x=>x.json());
  log(r.log || '', r.ok ? 'g' : 'r');
  mark(r.ok);
  await loadState();
  if(r.path) $('#snap').value = r.path;
  inp.value = '';
}
async function doIntake(preview){
  const fd = new FormData();
  const files = $('#vinfile').files;
  if(!files.length && !$('#vintext').value.trim()){ alert('请选择 VIN 文件或粘贴 VIN'); return; }
  for(const f of files) fd.append('vinfile', f);
  fd.append('vintext', $('#vintext').value);
  fd.append('importer', $('#inimp').value);
  fd.append('snapshot', snapshotArg() || '');
  fd.append('date', $('#indate').value);
  fd.append('tag', $('#intag').value);
  fd.append('preview', preview ? '1' : '0');
  await post('/api/intake', fd, true);
}
async function doSale(preview){
  await post('/api/sale', {importer:$('#imp').value, month:$('#month').value,
    count:$('#count').value, preview, snapshot:snapshotArg()});
}
async function doSaleFile(preview){
  const fd = new FormData();
  const files = $('#salefile').files;
  if(!files.length){ alert('请选择真实末销表'); return; }
  for(const f of files) fd.append('salefile', f);
  fd.append('date', $('#sfdate').value);
  fd.append('snapshot', snapshotArg() || '');
  fd.append('preview', preview ? '1' : '0');
  await post('/api/sale-file', fd, true);
}
async function setSnap(v){
  fetch('/api/set_snapshot', {method:'POST', headers:{'Content-Type':'application/json'},
    body:JSON.stringify({path:v})}).then(x=>x.json()).then(r=>{
      if(r.log){log('\n'+r.log+'\n', r.ok?'g':'r');}
    }).catch(()=>{});
}
async function doDash(){
  const r = await post('/api/dashboard', {month:$('#dmonth').value, snapshot:snapshotArg()});
  const box = $('#dashresult'), links = $('#dashlinks');
  if(r.ok && r.files && r.files.length){
    links.innerHTML = r.files.map(f=>{
      const open = `/view?p=${encodeURIComponent(f.path)}`;
      const dl = `/download?p=${encodeURIComponent(f.path)}`;
      return `<a href="${open}" target="_blank" class="btn pri" style="text-decoration:none;margin-right:10px">📊 在浏览器打开看板</a>`
           + `<a href="${dl}" class="btn" style="text-decoration:none">⬇ 下载 HTML</a>`;
    }).join('');
    box.classList.remove('hidden');
  } else {
    links.innerHTML = '<span style="color:#b91c1c;font-size:13px">看板已生成，但未获取到打开链接。请把运行日志末尾「看板诊断」那一行发给开发者定位。</span>';
    box.classList.remove('hidden');
  }
}
async function doCheck(){
  if(!$('#chkfile').value){ alert('输出目录里还没有可校验的 Excel'); return; }
  await post('/api/check', {path:$('#chkfile').value});
}

function fillMonths(){
  const now = new Date();
  for(const id of ['month','dmonth']){
    const sel = $('#'+id);
    if(!sel || sel.options.length) continue;
    let opts = '<option value="">自动（取末销数据最新月）</option>';
    for(let i=-12;i<=3;i++){
      const d = new Date(now.getFullYear(), now.getMonth()+i, 1);
      const v = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
      opts += '<option value="'+v+'">'+v+'</option>';
    }
    sel.innerHTML = opts;
  }
  const sf = $('#sfdate');
  if(sf && !sf.value){
    const d = now;
    sf.value = d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
  }
}
function syncSaleMode(){
  const m = document.querySelector('input[name=smode]:checked').value;
  $('#sale-gen').classList.toggle('hidden', m !== 'gen');
  $('#sale-real').classList.toggle('hidden', m !== 'real');
}
function syncMode(){
  const c = document.getElementById('count');
  if(c) c.disabled = false;
}

document.querySelectorAll('nav button').forEach(b=>{
  b.onclick = () => {
    document.querySelectorAll('nav button').forEach(x=>x.classList.remove('on'));
    b.classList.add('on');
    ['intake','sale','dash','check'].forEach(t=>
      $('#p-'+t).classList.toggle('hidden', t !== b.dataset.t));
  };
});

fillMonths();
syncSaleMode();
syncMode();
setInterval(()=>fetch('/api/ping').catch(()=>{}), 5000);
loadState();
</script>
</body></html>
"""


# ---------------------------------------------------------------- 服务

class Handler(BaseHTTPRequestHandler):

    server_version = 'VehicleFlow/1.0'

    def log_message(self, fmt, *a):
        pass

    def _send(self, code, body, ctype='application/json; charset=utf-8', extra=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        elif isinstance(body, str):
            body = body.encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(body)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    # ---------- GET

    def do_GET(self):
        u = urlparse(self.path)
        q = parse_qs(u.query)
        if u.path in ('/', '/index.html'):
            return self._send(200, PAGE, 'text/html; charset=utf-8')
        if u.path == '/api/state':
            return self._send(200, state())
        if u.path == '/api/ping':
            LAST_PING[0] = time.time()
            return self._send(200, {'ok': True})
        if u.path == '/api/quit':
            LAST_PING[0] = time.time()
            return self._send(200, {'ok': True})
        if u.path == '/api/reveal':
            os.system('open "%s"' % WS)
            return self._send(200, {'ok': True})
        if u.path == '/download':
            return self._file(q.get('p', [''])[0], 'application/octet-stream')
        if u.path == '/view':
            return self._file(q.get('p', [''])[0], 'text/html; charset=utf-8')
        return self._send(404, {'ok': False, 'log': 'not found'})

    def _file(self, path, ctype):
        path = unquote(path)
        if not inside_ws(path):
            return self._send(403, {'ok': False, 'log': 'forbidden'})
        with open(path, 'rb') as f:
            data = f.read()
        self.send_response(200)
        self.send_header('Content-Type', ctype)
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Content-Disposition',
                         'inline; filename="%s"' % quote(os.path.basename(path)))
        self.end_headers()
        self.wfile.write(data)

    # ---------- POST

    def do_POST(self):
        LAST_PING[0] = time.time()
        u = urlparse(self.path)
        ctype = self.headers.get('Content-Type', '')
        try:
            if 'multipart/form-data' in ctype:
                boundary = ctype.split('boundary=')[1].encode('utf-8')
                body = parse_multipart(self.rfile.read(int(self.headers['Content-Length'])), boundary)
            else:
                raw = self.rfile.read(int(self.headers.get('Content-Length', 0)))
                body = json.loads(raw or b'{}')
        except Exception as e:
            return self._send(400, {'ok': False, 'log': '请求解析失败: %s' % e})

        if u.path == '/api/intake':
            return self._send(200, api_intake(body))
        if u.path == '/api/upload_snapshot':
            return self._send(200, api_upload_snapshot(body))
        if u.path == '/api/set_snapshot':
            return self._send(200, api_set_snapshot(body))
        if u.path == '/api/sale':
            return self._send(200, api_sale(body))
        if u.path == '/api/sale-file':
            return self._send(200, api_sale_file(body))
        if u.path == '/api/dashboard':
            return self._send(200, api_dashboard(body))
        if u.path == '/api/check':
            return self._send(200, api_check(body))
        return self._send(404, {'ok': False, 'log': 'not found'})


def _current_snapshot():
    cp = os.path.join(VF.BASE, 'data', '.current_snapshot')
    if os.path.exists(cp):
        try:
            with open(cp, encoding='utf-8') as f:
                p = f.read().strip()
            if p and os.path.exists(p):
                return p
        except Exception:
            pass
    return ''


def state():
    import glob
    snaps = sorted(glob.glob(os.path.join(VF.INBOX, '*.xlsx')) +
                   glob.glob(os.path.join(VF.INBOX, '*.xls')), key=os.path.getmtime)
    cur = _current_snapshot()
    if cur and cur not in snaps:
        snaps.append(cur)
    importers = []
    for p in sorted(glob.glob(os.path.join(VF.CONFIG_DIR, '*.json'))):
        try:
            with open(p, encoding='utf-8') as f:
                cfg = json.load(f)
            code = os.path.basename(p)[:-5]
            name = cfg.get('country', code)
            importers.append({'code': code, 'name': name})
        except Exception:
            code = os.path.basename(p)[:-5]
            importers.append({'code': code, 'name': code})
    return {
        'version': VERSION,
        'ws': WS,
        'snapshot': cur or (snaps[-1] if snaps else ''),
        'snapshots': snaps,
        'importers': importers,
        'outputs': rel_files(sorted(glob.glob(os.path.join(VF.OUT_DIR, '**', '*.xlsx'),
                                         recursive=True), key=os.path.getmtime, reverse=True)),
    }


def out_files():
    import glob
    return set(glob.glob(os.path.join(VF.OUT_DIR, '**', '*.*'), recursive=True))


def _save_upload(form, field, subdir='uploads'):
    items = form.get(field)
    if not items:
        return []
    if not isinstance(items, list):
        items = [items]
    dst = os.path.join(VF.INBOX, subdir)
    paths = []
    for item in items:
        if not isinstance(item, tuple):
            continue
        fname, data = item
        if not fname or not data:
            continue
        os.makedirs(dst, exist_ok=True)
        path = os.path.join(dst, '%s_%s' % (dt.datetime.now().strftime('%H%M%S'),
                                           os.path.basename(fname)))
        with open(path, 'wb') as f:
            f.write(data)
        paths.append(path)
    return paths


def api_upload_snapshot(form):
    paths = _save_upload(form, 'snapshot', subdir='')
    if not paths:
        return {'ok': False, 'log': '没有收到文件'}
    p = paths[0]
    # 记为当前快照
    os.makedirs(os.path.join(VF.BASE, 'data'), exist_ok=True)
    with open(os.path.join(VF.BASE, 'data', '.current_snapshot'), 'w', encoding='utf-8') as f:
        f.write(p)
    log = ['已保存快照：%s' % os.path.basename(p)]
    # 扫描快照，为每个进口商生成主数据（进口商/国家/经销商下拉从此有值）
    try:
        rows = VF.load_snapshot(p)
        # 展示识别到的列: 跨平台模板时让使用者知道哪些列被认出、哪些没认到
        try:
            fh = VF.snapshot_field_headers(p)
            order = ['vin', 'imp', 'imp_name', 'dealer', 'dealer_name', 'model',
                     'status', 'sold_at', 'dealer_in', 'imp_in', 'transit',
                     'sold_done', 'dealer_done', 'imp_done', 'country']
            parts = ['%s→%s' % (k, fh[k]) for k in order if k in fh]
            log.append('识别到 %d 列：%s' % (len(fh), '，'.join(parts)))
            missing = [k for k in ('imp', 'dealer', 'sold_at', 'transit') if k not in fh]
            if missing:
                log.append('⚠ 未识别到列：%s（看板/末销可能不完整，可检查列名）' % '、'.join(missing))
        except Exception:
            pass
        imps = sorted({r['imp'] for r in rows if r['imp']})
        ok_imps = []
        for imp in imps:
            try:
                VF.load_config(imp, rows)
                ok_imps.append(imp)
            except Exception as e:
                log.append('  进口商 %s 主数据生成失败：%s' % (imp, e))
        if ok_imps:
            log.append('识别到 %d 个进口商：%s（已生成主数据，可在 config/countries/ 编辑）'
                       % (len(ok_imps), ' / '.join(ok_imps)))
        else:
            log.append('⚠ 快照里没读到进口商编码，请确认是平台导出的车辆台账')
    except Exception as e:
        log.append('快照解析失败：%s' % e)
        log.append('请确认上传的是平台导出的车辆台账（含 VIN / 进口商编码 / 末销时间 / 在途 等列）')
        return {'ok': False, 'log': '\n'.join(log), 'path': p}
    return {'ok': True, 'log': '\n'.join(log), 'path': p}


def api_set_snapshot(b):
    p = (b.get('path') or '').strip()
    if p and os.path.exists(p) and inside_ws(p):
        os.makedirs(os.path.join(VF.BASE, 'data'), exist_ok=True)
        with open(os.path.join(VF.BASE, 'data', '.current_snapshot'), 'w', encoding='utf-8') as f:
            f.write(p)
        return {'ok': True, 'log': '当前快照已切换为：%s' % os.path.basename(p)}
    return {'ok': False, 'log': '快照路径无效：%s' % p}


def api_intake(form):
    from argparse import Namespace
    before = out_files()
    files = _save_upload(form, 'vinfile')
    text = (form.get('vintext') or '').strip()
    if not files and not text:
        return {'ok': False, 'log': '请选择 VIN 文件或粘贴 VIN', 'files': []}
    # 批次标签: 优先用户输入, 否则从第一个 PL 文件名提取
    tag = (form.get('tag') or '').strip()
    if not tag and files:
        base = os.path.basename(files[0])
        stem, _ = os.path.splitext(base)
        # 常见前缀: PL Attachment / 批次标签 / 截短到 24 字符
        stem = stem.replace('PL Attachment ', '').replace('PL附件 ', '')
        tag = stem[:24] if stem else None
    args = Namespace(snapshot=(form.get('snapshot') or None) or None,
                     vin_file=files or None, vin=(text or None),
                     importer=(form.get('importer') or '').strip() or None,
                     date=(form.get('date') or None) or None, importer_date=None,
                     tag=tag or None, out=None,
                     preview=form.get('preview') == '1')
    ok, log = run_capture(VF.cmd_intake, args)
    return {'ok': ok, 'log': log, 'files': rel_files(sorted(out_files() - before))}


def api_sale(b):
    from argparse import Namespace
    count = (b.get('count') or '').strip()
    if not count:
        return {'ok': False, 'log': '请填写报量台数', 'files': []}
    before = out_files()
    args = Namespace(snapshot=b.get('snapshot') or None,
                     importer=b.get('importer'),
                     count=count, month=(b.get('month') or '').strip() or None,
                     source=None, carflow=None, date=None, tag=None, out=None,
                     preview=bool(b.get('preview')))
    ok, log = run_capture(VF.cmd_sale, args)
    return {'ok': ok, 'log': log, 'files': rel_files(sorted(out_files() - before))}


def api_sale_file(form):
    from argparse import Namespace
    files = _save_upload(form, 'salefile')
    if not files:
        return {'ok': False, 'log': '没有收到末销表文件', 'files': []}
    before = out_files()
    args = Namespace(file=files[0], snapshot=(form.get('snapshot') or None) or None,
                     date=(form.get('date') or '').strip() or None,
                     tag=None, out=None, preview=form.get('preview') == '1')
    ok, log = run_capture(VF.cmd_sale_file, args)
    return {'ok': ok, 'log': log, 'files': rel_files(sorted(out_files() - before))}


def api_dashboard(b):
    from argparse import Namespace
    args = Namespace(snapshot=b.get('snapshot') or None, carflow=None,
                     month=(b.get('month') or '').strip() or None, out=None)
    ok, log = run_capture(DB.cmd_dashboard, args)
    files = []
    diag = []
    if ok:
        # 看板是覆盖写同名文件, 不能用「前后文件差集」(已存在时差集为空),
        # 直接从日志取本次生成的路径, 拿不到再兜底取 out 下最新 html
        m = re.search(r'看板已生成:\s*(.+)', log)
        if m:
            p = m.group(1).strip()
            diag.append('匹配路径=%r exists=%s inside_ws=%s'
                        % (p, os.path.exists(p), inside_ws(p)))
            if os.path.exists(p):
                files = [p]
        if not files:
            import glob
            htmls = sorted(glob.glob(os.path.join(VF.OUT_DIR, '**', '*.html'),
                                     recursive=True), key=os.path.getmtime, reverse=True)
            files = htmls[:1]
            diag.append('兜底glob=%r' % files)
    result = rel_files(files)
    diag.append('WS=%r' % WS)
    diag.append('rel_files=%d个' % len(result))
    log = log + '\n[看板诊断] ' + ' | '.join(diag)
    return {'ok': ok, 'log': log, 'files': result}


def api_check(b):
    p = b.get('path') or ''
    if not inside_ws(p):
        return {'ok': False, 'log': '文件不在工作目录内', 'files': []}
    ok, log = run_capture(lambda: VF.check_file(p))
    return {'ok': ok, 'log': log, 'files': []}


def watchdog(srv):
    while True:
        time.sleep(5)
        if time.time() - LAST_PING[0] > 40:
            srv.shutdown()
            return


def main():
    port = PORT
    for _ in range(20):
        try:
            srv = ThreadingHTTPServer(('127.0.0.1', port), Handler)
            break
        except OSError:
            port += 1
    else:
        sys.exit('没有可用端口')
    url = 'http://127.0.0.1:%d/' % port
    print('已启动: %s' % url)
    if '--no-open' not in sys.argv:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    threading.Thread(target=watchdog, args=(srv,), daemon=True).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == '__main__':
    if '--selftest' in sys.argv:
        import urllib.request
        srv = ThreadingHTTPServer(('127.0.0.1', 0), Handler)
        port = srv.server_address[1]
        threading.Thread(target=srv.serve_forever, daemon=True).start()
        time.sleep(0.3)
        d = json.loads(urllib.request.urlopen(
            'http://127.0.0.1:%d/api/state' % port, timeout=5).read())
        print('selftest ok: ws=%s importers=%d snapshots=%d outputs=%d' %
              (d['ws'], len(d['importers']), len(d['snapshots']), len(d['outputs'])))
        srv.shutdown()
        sys.exit(0)
    main()
