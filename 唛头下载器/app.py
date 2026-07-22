# -*- coding: utf-8 -*-
"""唛头PDF批量下载器
上传含 waybill_num 列的 Excel，自动调用物流商API下载唛头PDF并按单号重命名。

凭证读取优先级：
  1. 环境变量 CLIENT_ID / CLIENT_SECRET（用于 Vercel 等公网部署，凭证不进代码）
  2. 同目录 env.doc 文件（用于本地运行）

设计为无状态：处理结果的 zip 直接随响应返回给浏览器，不写服务器磁盘，
因此可直接部署到 Vercel 这类 serverless 平台。本地运行时会额外把PDF存到「唛头PDF」文件夹。
"""
import base64
import io
import os
import re
import socket
import zipfile
from pathlib import Path

import requests
from flask import Flask, request, render_template_string, jsonify
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
API_URL = "http://api.ywcd56.com/openapi/getOutMarkPdf"
API_HOST = "http://api.ywcd56.com"
OUT_DIR = BASE_DIR / "唛头PDF"
# serverless 平台（如 Vercel）磁盘只读，禁用本地保存
IS_SERVERLESS = bool(os.environ.get("VERCEL") or os.environ.get("SERVERLESS"))

app = Flask(__name__)


def lan_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


def load_env():
    """读取凭证：环境变量优先，其次 env.doc 文件"""
    cid = os.environ.get("CLIENT_ID", "").strip()
    sec = os.environ.get("CLIENT_SECRET", "").strip()
    if cid and sec:
        return cid, sec
    creds = {}
    env_path = BASE_DIR / "env.doc"
    if env_path.exists():
        for line in env_path.read_text(encoding="utf-8", errors="ignore").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, _, v = line.partition("=")
            creds[k.strip().upper()] = v.strip()
    return cid or creds.get("CLIENT_ID", ""), sec or creds.get("CLIENT_SECRET", "")


def extract_waybills(file_bytes):
    """从Excel提取物流单号：优先找表头 waybill_num（或含'单号'的列），否则取第一列"""
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    rows = [[("" if c is None else str(c).strip()) for c in row]
            for row in ws.iter_rows(values_only=True)]
    wb.close()
    if not rows:
        return []
    col, start = 0, 0
    header = rows[0]
    for i, h in enumerate(header):
        if h.lower() == "waybill_num" or "单号" in h:
            col, start = i, 1
            break
    nums, seen = [], set()
    for row in rows[start:]:
        if col < len(row):
            v = row[col]
            if v and re.match(r"^[A-Za-z0-9\-_]+$", v) and v not in seen:
                seen.add(v)
                nums.append(v)
    return nums


def fetch_pdf(waybill, api_key, token):
    """调用API获取唛头PDF，返回 (pdf_bytes, error_msg)"""
    headers = {"x-api-key": api_key, "token": token}
    try:
        r = requests.get(API_URL, params={"waybill_num": waybill},
                         headers=headers, timeout=30)
    except requests.RequestException as e:
        return None, f"请求失败: {e}"
    ct = r.headers.get("Content-Type", "")
    if "pdf" in ct or r.content[:4] == b"%PDF":
        return r.content, None
    try:
        data = r.json()
    except ValueError:
        return None, f"响应无法解析 (HTTP {r.status_code})"
    if data.get("code") not in (0, 200, "0", "200"):
        msg = data.get("msg", "")
        return None, f"无唛头数据{f'（{msg}）' if msg else ''}"
    path = data.get("path") or (data.get("data") or {}).get("path") if isinstance(data.get("data"), dict) else data.get("path")
    if not path:
        return None, "接口未返回PDF地址"
    url = path if path.startswith("http") else API_HOST + ("" if path.startswith("/") else "/") + path
    try:
        pr = requests.get(url, headers=headers, timeout=60)
        pr.raise_for_status()
    except requests.RequestException as e:
        return None, f"下载PDF失败: {e}"
    if pr.content[:4] != b"%PDF":
        return None, "下载内容不是PDF"
    return pr.content, None


PAGE = """<!doctype html><html lang="zh"><head><meta charset="utf-8">
<title>唛头PDF批量下载器</title>
<style>
body{font-family:-apple-system,"PingFang SC",sans-serif;max-width:640px;margin:40px auto;padding:0 20px;color:#333}
h1{font-size:22px}.card{border:1px solid #ddd;border-radius:10px;padding:24px;margin-top:16px}
input[type=file]{margin:12px 0}
button{background:#2563eb;color:#fff;border:0;border-radius:8px;padding:10px 24px;font-size:15px;cursor:pointer}
button:disabled{background:#aaa}
#log{white-space:pre-wrap;font-size:13px;background:#f6f6f6;border-radius:8px;padding:12px;margin-top:16px;display:none}
.warn{color:#b45309}.ok{color:#15803d}.err{color:#b91c1c}
.tip{background:#fef9c3;border:1px solid #fde047;border-radius:8px;padding:8px 12px;font-size:13px;color:#854d0e}
a.dl{display:inline-block;margin-top:12px;color:#2563eb;font-weight:600}
</style></head><body>
<h1>唛头PDF批量下载器</h1>
<div class="card">
{% if not has_creds %}<p class="warn">⚠️ 尚未配置凭证：请设置环境变量 CLIENT_ID / CLIENT_SECRET（部署平台）或填写 env.doc（本地）。</p>{% endif %}
<p>上传包含物流单号（waybill_num）的 Excel 文件（.xlsx）：</p>
<p class="tip">💡 建议每次查询单量小于 <b>60 单</b>；单量较大时请自行分批上传，以免请求超时。</p>
<form id="f">
<input type="file" name="file" accept=".xlsx,.xlsm" required>
<br><button type="submit" id="btn">开始下载</button>
</form>
<div id="log"></div>
</div>
<script>
const f=document.getElementById('f'),log=document.getElementById('log'),btn=document.getElementById('btn');
f.addEventListener('submit',async e=>{
  e.preventDefault();btn.disabled=true;log.style.display='block';log.textContent='处理中，请稍候…';
  const fd=new FormData(f);
  try{
    const r=await fetch('/process',{method:'POST',body:fd});
    const d=await r.json();
    log.innerHTML=d.log.map(l=>{
      const cls=l.startsWith('✔')?'ok':(l.startsWith('✘')?'err':(l.startsWith('⚠')?'warn':''));
      return cls?`<span class="${cls}">${l}</span>`:l;
    }).join('\\n');
    if(d.zip_b64){
      const bin=atob(d.zip_b64),arr=new Uint8Array(bin.length);
      for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
      const url=URL.createObjectURL(new Blob([arr],{type:'application/zip'}));
      const a=document.createElement('a');a.href=url;a.className='dl';
      a.download='唛头PDF.zip';a.textContent='📦 下载全部PDF (zip)';
      log.appendChild(document.createElement('br'));log.appendChild(a);
    }
  }catch(err){log.textContent='出错: '+err;}
  btn.disabled=false;
});
</script></body></html>"""


@app.route("/")
def index():
    api_key, secret = load_env()
    return render_template_string(PAGE, has_creds=bool(api_key and secret))


@app.route("/process", methods=["POST"])
def process():
    api_key, secret = load_env()
    log = []
    if not (api_key and secret):
        return jsonify(log=["✘ 未配置凭证：请设置环境变量 CLIENT_ID / CLIENT_SECRET（或本地 env.doc）"], zip_b64=None)
    file = request.files.get("file")
    if not file:
        return jsonify(log=["✘ 未收到文件"], zip_b64=None)
    try:
        waybills = extract_waybills(file.read())
    except Exception as e:
        return jsonify(log=[f"✘ Excel解析失败: {e}"], zip_b64=None)
    if not waybills:
        return jsonify(log=["✘ 未在Excel中找到物流单号"], zip_b64=None)
    log.append(f"共读取 {len(waybills)} 个单号")
    if len(waybills) > 60:
        log.append(f"⚠️ 本次 {len(waybills)} 单已超过建议上限（60单），若长时间无响应或超时，请分批（每批≤60单）重试")
    if not IS_SERVERLESS:
        OUT_DIR.mkdir(exist_ok=True)
    ok, results, failed = 0, [], []
    for wb in waybills:
        pdf, err = fetch_pdf(wb, api_key, secret)
        if pdf:
            if not IS_SERVERLESS:
                (OUT_DIR / f"{wb}.pdf").write_bytes(pdf)
            results.append((f"{wb}.pdf", pdf))
            log.append(f"✔ {wb}.pdf 已下载")
            ok += 1
        else:
            failed.append(wb)
            log.append(f"✘ {wb} {err}")
    log.append(f"\n完成：成功 {ok} / {len(waybills)}")
    if failed:
        log.append(f"✘ 失败单号（{len(failed)}个）: {'、'.join(failed)}")
    zip_b64 = None
    if results:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in results:
                zf.writestr(name, data)
        zip_b64 = base64.b64encode(buf.getvalue()).decode()
    return jsonify(log=log, zip_b64=zip_b64)


if __name__ == "__main__":
    ip = lan_ip()
    print("唛头PDF批量下载器已启动")
    print(f"  本机访问:   http://127.0.0.1:8756")
    print(f"  局域网访问: http://{ip}:8756  （发给同一网络的同事即可）")
    app.run(host="0.0.0.0", port=8756, threaded=True)
