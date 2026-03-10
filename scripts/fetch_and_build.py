"""
fetch_and_build.py
厚生労働省「医療用医薬品供給状況報告」ページから
最新のExcelを自動取得し、HTMLビューアを生成する。
"""

import re
import os
import json
import hashlib
import requests
import openpyxl
from datetime import datetime, timezone, timedelta
from bs4 import BeautifulSoup
from pathlib import Path

# ── 設定 ──────────────────────────────────────────────────────────────────────
MHLW_URL   = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
MHLW_BASE  = "https://www.mhlw.go.jp"
DOCS_DIR   = Path("docs")
CACHE_FILE = Path("scripts/.last_hash")   # 前回のExcelのハッシュ値
JST        = timezone(timedelta(hours=9))

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; DrugSupplyBot/1.0; +https://github.com)"
}

# ── Step 1: 最新ExcelのURLを取得 ──────────────────────────────────────────────
def get_latest_excel_url() -> tuple[str, str]:
    """ページをスクレイピングして最新.xlsxのURLと日付文字列を返す"""
    print("[1/4] 厚労省サイトからExcel URLを取得中...")
    res = requests.get(MHLW_URL, headers=HEADERS, timeout=30)
    res.raise_for_status()
    res.encoding = "utf-8"

    soup = BeautifulSoup(res.text, "lxml")
    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.endswith(".xlsx") and "iyakuhin" in href:
            full_url = href if href.startswith("http") else MHLW_BASE + href
            # リンクテキストから日付を抽出（例: 令和８年２月27日現在）
            label = a.get_text(strip=True)
            print(f"  URL: {full_url}")
            print(f"  ラベル: {label}")
            return full_url, label

    raise RuntimeError("ExcelファイルのURLが見つかりませんでした")

# ── Step 2: Excelをダウンロード（変更がなければスキップ） ─────────────────────
def download_excel(url: str) -> tuple[bytes, bool]:
    """Excelをダウンロード。前回と同じなら (data, False) を返す"""
    print("[2/4] Excelをダウンロード中...")
    res = requests.get(url, headers=HEADERS, timeout=60)
    res.raise_for_status()
    data = res.content

    current_hash = hashlib.md5(data).hexdigest()
    last_hash = CACHE_FILE.read_text().strip() if CACHE_FILE.exists() else ""

    if current_hash == last_hash:
        print("  前回と同じファイル。更新をスキップします。")
        return data, False

    CACHE_FILE.parent.mkdir(exist_ok=True)
    CACHE_FILE.write_text(current_hash)
    print(f"  ダウンロード完了 ({len(data)//1024} KB), hash={current_hash[:8]}")
    return data, True

# ── Step 3: Excelを解析してデータ抽出 ─────────────────────────────────────────
def parse_excel(data: bytes) -> list[list]:
    """
    Excelを解析し、行データのリストを返す。
    各行は以下の順の配列:
    [drugType, atc, generic, dose, yj, name, maker, productType,
     basic, secure, listedDate, supply, supplyUpdated, reason,
     resolveOutlook, resolveDate, volume, volumeDate, volumeAmount,
     otherUpdated, isNew]
    """
    print("[3/4] Excelを解析中...")
    import io
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.active
    raw_rows = list(ws.iter_rows(values_only=True))

    def cell(row, i):
        v = row[i] if i < len(row) else None
        if v is None: return ""
        s = str(v).strip()
        return "" if s in ("None", "") else s

    def date_cell(row, i):
        v = row[i] if i < len(row) else None
        if v is None: return ""
        s = str(v)
        return s.split(" ")[0] if "00:00:00" in s else s.strip()

    result = []
    # ヘッダー行は index=1、データは index=2 以降
    for row in raw_rows[2:]:
        name = cell(row, 5)
        if not name:
            continue
        result.append([
            cell(row, 0),  # drugType
            cell(row, 1),  # atc
            cell(row, 2),  # generic
            cell(row, 3),  # dose
            cell(row, 4),  # yj
            name,          # name
            cell(row, 6),  # maker
            cell(row, 7),  # productType
            cell(row, 8),  # basic
            cell(row, 9),  # secure
            date_cell(row, 10),  # listedDate
            cell(row, 11), # supply
            date_cell(row, 12),  # supplyUpdated
            cell(row, 13), # reason
            cell(row, 14), # resolveOutlook
            cell(row, 15), # resolveDate
            cell(row, 16), # volume
            date_cell(row, 17),  # volumeDate
            cell(row, 18), # volumeAmount
            date_cell(row, 19),  # otherUpdated
            cell(row, 20), # isNew
        ])

    print(f"  {len(result)} 品目を取得")
    wb.close()
    return result

# ── Step 4: HTMLを生成 ────────────────────────────────────────────────────────
def build_html(drug_data: list, label: str, generated_at: str) -> str:
    """データを埋め込んだスタンドアロンHTMLを返す"""
    print("[4/4] HTMLを生成中...")
    json_str = json.dumps(drug_data, ensure_ascii=False, separators=(',', ':'))

    # 集計
    total   = len(drug_data)
    alert   = sum(1 for r in drug_data if r[11] and "通常出荷" not in r[11])
    stop    = sum(1 for r in drug_data if r[11] == "⑤供給停止")
    limited = sum(1 for r in drug_data if r[11] and "限定出荷" in r[11])
    new_n   = sum(1 for r in drug_data if r[20] and r[20].strip())
    no_re   = sum(1 for r in drug_data if r[11] and "通常出荷" not in r[11] and r[14] == "ウ． 未定")

    # f文字列内で {label} の後に ":" が続くとPythonがformat specifierと誤認するため
    # JavaScriptのconst D行は文字列結合で別途生成する
    js_d_line = 'const D={"label":"' + label + '","generated":"' + generated_at + '"};'

    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>医療用医薬品供給状況ビューア（{label}）</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:'Noto Sans JP','Hiragino Kaku Gothic ProN','Yu Gothic',sans-serif;background:#f0f4f8;color:#0f172a;font-size:13px}}
header{{background:linear-gradient(135deg,#0f172a 0%,#1e3a5f 100%);color:#fff;padding:14px 24px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 16px rgba(0,0,0,.3);position:sticky;top:0;z-index:100}}
.sub{{font-size:9px;letter-spacing:3px;color:#94a3b8;text-transform:uppercase;margin-bottom:3px}}
h1{{font-size:17px;font-weight:900;letter-spacing:-.3px}}
.hright{{text-align:right;font-size:10px;color:#94a3b8;line-height:1.6}}
.hright .date{{color:#7dd3fc;font-size:12px;font-weight:700}}
.wrap{{padding:16px 24px}}
.summary{{display:flex;gap:9px;margin-bottom:13px;flex-wrap:wrap}}
.card{{flex:1;min-width:88px;background:#fff;border-radius:12px;padding:10px 13px;border:2px solid;cursor:pointer;transition:all .15s;text-align:left}}
.card:hover{{transform:translateY(-1px);box-shadow:0 4px 12px rgba(0,0,0,.12)}}
.card .num{{font-size:22px;font-weight:900;line-height:1}}
.card .lbl{{font-size:9px;font-weight:700;margin-top:3px;opacity:.85}}
.filters{{display:flex;gap:8px;margin-bottom:11px;flex-wrap:wrap;align-items:center}}
.filters input,.filters select{{padding:7px 10px;border-radius:8px;border:1.5px solid #e2e8f0;font-size:12px;outline:none;background:#fff;font-family:inherit}}
.filters input{{flex:2;min-width:200px}}
.filters select{{min-width:120px}}
.chk{{display:flex;align-items:center;gap:5px;font-size:12px;color:#475569;cursor:pointer;background:#fff;border:1.5px solid #e2e8f0;border-radius:7px;padding:6px 11px;user-select:none;white-space:nowrap}}
.chk input{{accent-color:#0369a1}}
.rbtn{{padding:7px 13px;background:#f1f5f9;border:1.5px solid #e2e8f0;border-radius:7px;font-size:12px;cursor:pointer;color:#64748b;font-family:inherit}}
.ci{{font-size:11px;color:#94a3b8;margin-left:auto;white-space:nowrap}}
.tw{{background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,.07);overflow:hidden}}
.ts{{overflow-x:auto}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead tr{{background:#f8fafc}}
th{{padding:9px 11px;text-align:left;font-weight:700;color:#64748b;font-size:10px;border-bottom:2px solid #e2e8f0;white-space:nowrap;cursor:pointer;user-select:none}}
th:hover{{background:#f1f5f9}}
tbody tr{{border-bottom:1px solid #f1f5f9;transition:background .1s;cursor:pointer}}
tbody tr:nth-child(even){{background:#fafafa}}
tbody tr:hover{{background:#eff6ff!important}}
td{{padding:9px 11px;vertical-align:top}}
.badge{{display:inline-block;padding:2px 9px;border-radius:20px;font-size:9px;font-weight:900;white-space:nowrap;border:1.5px solid}}
.tag{{display:inline-block;padding:1px 7px;border-radius:5px;font-size:9px;font-weight:700}}
.nb{{background:#dc2626;color:#fff;padding:1px 7px;border-radius:20px;font-size:9px;font-weight:900}}
.pager{{display:flex;align-items:center;justify-content:center;gap:7px;padding:12px;border-top:1px solid #f1f5f9;flex-wrap:wrap}}
.pb{{padding:5px 10px;font-size:12px;border-radius:7px;border:1.5px solid #e2e8f0;background:#fff;color:#334155;cursor:pointer;font-weight:700;font-family:inherit}}
.pb:disabled{{color:#d1d5db;cursor:default}}
.pi{{font-size:12px;color:#64748b;padding:0 6px}}
.fbar{{padding:9px 14px;border-top:1px solid #f1f5f9;display:flex;justify-content:space-between;align-items:center;font-size:10px;color:#94a3b8;flex-wrap:wrap;gap:6px}}
.mo{{display:none;position:fixed;inset:0;background:rgba(15,23,42,.65);z-index:1000;align-items:center;justify-content:center;padding:20px}}
.mo.open{{display:flex}}
.mb{{background:#fff;border-radius:16px;padding:24px 26px;max-width:640px;width:100%;max-height:88vh;overflow-y:auto;box-shadow:0 24px 64px rgba(0,0,0,.3);animation:mIn .2s}}
@keyframes mIn{{from{{opacity:0;transform:scale(.96)}}to{{opacity:1;transform:scale(1)}}}}
.mh{{display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:14px}}
.mc{{font-size:20px;color:#94a3b8;background:none;border:none;cursor:pointer;padding:0 4px;line-height:1}}
.ig{{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:13px}}
.ic{{background:#f8fafc;border-radius:9px;padding:9px 12px;border:1px solid #e2e8f0}}
.ic .l{{font-size:9px;color:#94a3b8;margin-bottom:3px;font-weight:700}}
.ic .v{{font-size:12px;font-weight:800;line-height:1.4}}
.ir{{margin-bottom:9px;padding:10px 12px;border-radius:9px;border:1px solid}}
.ir .l{{font-size:9px;font-weight:800;margin-bottom:4px}}
.ir .v{{font-size:12px;line-height:1.7}}
.db{{margin-top:13px;padding:11px 13px;background:#f8fafc;border-radius:9px;border:1px solid #e2e8f0}}
.db .dt{{font-size:10px;font-weight:800;color:#64748b;margin-bottom:7px}}
.dr{{display:flex;gap:7px;align-items:center;margin-bottom:4px}}
.empty{{text-align:center;padding:44px;color:#94a3b8}}
@media(max-width:640px){{.ig{{grid-template-columns:1fr}}.summary{{gap:6px}}.card .num{{font-size:18px}}}}
</style>
</head>
<body>

<header>
  <div>
    <div class="sub">厚生労働省 / MHLW Official Data</div>
    <h1>医療用医薬品 供給状況ビューア</h1>
  </div>
  <div class="hright">
    <div>医薬品医療機器等法 第18条の４に基づく届出データ</div>
    <div class="date">{label}</div>
    <div style="font-size:9px;margin-top:1px">自動更新: {generated_at}</div>
  </div>
</header>

<div class="wrap">
  <div class="summary">
    <button class="card" onclick="fc()" style="color:#334155;background:#f8fafc;border-color:#e2e8f0">
      <div class="num">{total:,}</div><div class="lbl">全登録品目</div></button>
    <button class="card" onclick="tc('newOnly')" style="color:#0369a1;background:#e0f2fe;border-color:#7dd3fc">
      <div class="num">{new_n}</div><div class="lbl">今回更新</div></button>
    <button class="card" onclick="tc('alertOnly')" style="color:#dc2626;background:#fef2f2;border-color:#fca5a5">
      <div class="num">{alert:,}</div><div class="lbl">要注意計</div></button>
    <button class="card" onclick="sf('限定')" style="color:#b45309;background:#fffbeb;border-color:#fcd34d">
      <div class="num">{limited:,}</div><div class="lbl">限定出荷</div></button>
    <button class="card" onclick="sf('⑤供給停止')" style="color:#dc2626;background:#fef2f2;border-color:#fca5a5">
      <div class="num">{stop:,}</div><div class="lbl">供給停止</div></button>
    <button class="card" style="color:#7c3aed;background:#f5f3ff;border-color:#c4b5fd;cursor:default">
      <div class="num">{no_re:,}</div><div class="lbl">解消未定</div></button>
  </div>

  <div class="filters">
    <input type="text" id="sq" placeholder="販売名・成分名・製造販売業者・YJコードで検索…" oninput="af()">
    <select id="supF" onchange="af()">
      <option value="">出荷対応：全て</option>
      <option>②限定出荷（自社の事情）</option>
      <option>③限定出荷（他社品の影響）</option>
      <option>④限定出荷（その他）</option>
      <option>⑤供給停止</option>
    </select>
    <select id="atcF" onchange="af()"><option value="">薬効分類：全て</option></select>
    <select id="reaF" onchange="af()">
      <option value="">理由：全て</option>
      <option>１．需要増</option>
      <option>２．原材料調達上の問題</option>
      <option>３．製造トラブル（製造委託を含む）</option>
      <option>４．品質トラブル（製造委託を含む）</option>
      <option>５．行政処分（製造委託を含む）</option>
      <option>８．その他の理由</option>
    </select>
    <label class="chk"><input type="checkbox" id="newOnly" onchange="af()"> 今回更新のみ</label>
    <label class="chk"><input type="checkbox" id="alertOnly" onchange="af()"> 要注意のみ</label>
    <button class="rbtn" onclick="rf()">リセット</button>
    <span class="ci" id="ci"></span>
  </div>

  <div class="tw">
    <div class="ts">
      <table>
        <thead><tr>
          <th onclick="sb(0)">販売名 / 成分名 ⇅</th>
          <th onclick="sb(1)">製造販売業者 ⇅</th>
          <th>薬効分類</th>
          <th onclick="sb(3)">出荷対応 ⇅</th>
          <th>出荷量</th>
          <th>理由</th>
          <th>解消見込み</th>
          <th>更新</th>
        </tr></thead>
        <tbody id="tb"></tbody>
      </table>
    </div>
    <div class="pager" id="pg"></div>
    <div class="fbar">
      <span>出典：厚生労働省「医療用医薬品供給状況報告」| ブラウザ内処理・外部送信なし</span>
      <span>全 {total:,} 品目</span>
    </div>
  </div>
</div>

<div class="mo" id="modal" onclick="if(event.target===this)cm()">
  <div class="mb" id="mb"></div>
</div>

<script>
{js_d_line}
const RAW={json_str};
const SS={{"①通常出荷":{{c:"#15803d",b:"#f0fdf4",bd:"#86efac"}},
  "②限定出荷（自社の事情）":{{c:"#b45309",b:"#fffbeb",bd:"#fcd34d"}},
  "③限定出荷（他社品の影響）":{{c:"#b45309",b:"#fff7ed",bd:"#fdba74"}},
  "④限定出荷（その他）":{{c:"#b45309",b:"#fffbeb",bd:"#fcd34d"}},
  "⑤供給停止":{{c:"#dc2626",b:"#fef2f2",bd:"#fca5a5"}}}};
const VS={{"Aプラス．出荷量増加":{{c:"#0369a1",b:"#e0f2fe"}},
  "A．出荷量通常":{{c:"#15803d",b:"#f0fdf4"}},
  "B．出荷量減少":{{c:"#b45309",b:"#fffbeb"}},
  "C．出荷停止":{{c:"#dc2626",b:"#fef2f2"}},
  "D．薬価削除予定":{{c:"#4b5563",b:"#f3f4f6"}}}};
const SO={{"⑤供給停止":4,"②限定出荷（自社の事情）":3,"③限定出荷（他社品の影響）":3,"④限定出荷（その他）":2,"①通常出荷":0}};

let filtered=[],sc=3,sa=false,pg=1;
const PS=50;
function st(s){{return SS[s]||{{c:"#64748b",b:"#f8fafc",bd:"#e2e8f0"}};}}
function vt(v){{return VS[v]||null;}}
function ia(s){{return s&&!s.includes("通常出荷");}}

// ATCリスト構築
window.addEventListener("DOMContentLoaded",()=>{{
  const atcs=[...new Set(RAW.map(r=>r[1]).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"ja"));
  const sel=document.getElementById("atcF");
  atcs.forEach(a=>{{const o=document.createElement("option");o.value=a;o.textContent=a;sel.appendChild(o);}});
  af();
}});

function af(){{
  const q=document.getElementById("sq").value.toLowerCase();
  const sf=document.getElementById("supF").value;
  const af2=document.getElementById("atcF").value;
  const rf2=document.getElementById("reaF").value;
  const no=document.getElementById("newOnly").checked;
  const ao=document.getElementById("alertOnly").checked;
  filtered=RAW.filter(r=>{{
    if(q&&!((r[5]||"").toLowerCase().includes(q)||(r[2]||"").toLowerCase().includes(q)||(r[6]||"").toLowerCase().includes(q)||(r[4]||"").includes(q)))return false;
    if(sf&&r[11]!==sf)return false;
    if(af2&&r[1]!==af2)return false;
    if(rf2&&r[13]!==rf2)return false;
    if(no&&!(r[20]&&r[20].trim()))return false;
    if(ao&&!ia(r[11]))return false;
    return true;
  }});
  filtered.sort((a,b)=>{{
    if(sc===3){{const pa=SO[a[11]]??0,pb=SO[b[11]]??0;return sa?pa-pb:pb-pa;}}
    const cm={{0:5,1:6,2:1,4:16}};const ci=cm[sc]??sc;
    const va=a[ci]||"",vb=b[ci]||"";
    return sa?va.localeCompare(vb,"ja"):vb.localeCompare(va,"ja");
  }});
  pg=1;
  document.getElementById("ci").textContent=`${{filtered.length.toLocaleString()}} / ${{RAW.length.toLocaleString()}} 件`;
  rt();rp();
}}
function sb(c){{if(sc===c)sa=!sa;else{{sc=c;sa=true;}}af();}}
function tc(id){{document.getElementById(id).checked=!document.getElementById(id).checked;af();}}
function sf(v){{
  if(v==="限定"){{document.getElementById("supF").value="";document.getElementById("alertOnly").checked=true;}}
  else document.getElementById("supF").value=v;
  af();
}}
function fc(){{rf();}}
function rf(){{
  document.getElementById("sq").value="";
  document.getElementById("supF").value="";
  document.getElementById("atcF").value="";
  document.getElementById("reaF").value="";
  document.getElementById("newOnly").checked=false;
  document.getElementById("alertOnly").checked=false;
  af();
}}
function rt(){{
  const s=(pg-1)*PS,rows=filtered.slice(s,s+PS);
  if(!rows.length){{document.getElementById("tb").innerHTML=`<tr><td colspan="8" class="empty">該当する品目がありません</td></tr>`;return;}}
  document.getElementById("tb").innerHTML=rows.map(r=>{{
    const s2=st(r[11]),vs=vt(r[16]),isn=r[20]&&r[20].trim();
    const idx=RAW.indexOf(r);
    return `<tr onclick="sd(${{idx}})">
      <td style="max-width:220px"><div style="font-weight:700;color:#0f172a;line-height:1.3">${{r[5]}}</div>
        ${{r[2]?`<div style="font-size:10px;color:#64748b;margin-top:2px">${{r[2]}}${{r[3]?' / '+r[3]:''}}</div>`:''}}
        ${{r[4]?`<div style="font-size:9px;color:#94a3b8;margin-top:1px">YJ: ${{r[4]}}</div>`:''}}</td>
      <td style="font-size:11px;color:#475569;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r[6]}}">${{r[6]||'―'}}</td>
      <td><span style="background:#f1f5f9;color:#334155;padding:1px 6px;border-radius:4px;font-size:9px">${{r[1]||'―'}}</span></td>
      <td>${{r[11]?`<span class="badge" style="background:${{s2.b}};color:${{s2.c}};border-color:${{s2.bd}}">${{r[11]}}</span>`:'<span style="color:#94a3b8;font-size:10px">―</span>'}}</td>
      <td>${{vs?`<span class="tag" style="background:${{vs.b}};color:${{vs.c}}">${{r[16]}}</span>`:`<span style="font-size:10px;color:#94a3b8">${{r[16]||'―'}}</span>`}}</td>
      <td style="font-size:10px;color:#475569;max-width:150px"><div style="overflow:hidden;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical">${{r[13]==='７．ー'||!r[13]?'―':r[13]}}</div></td>
      <td style="font-size:10px;white-space:nowrap;color:${{r[14]==='ウ． 未定'?'#7c3aed':r[14]==='ア． あり'?'#15803d':'#475569'}}">${{r[14]&&r[14]!=='エ． －'?r[14]:'―'}}${{r[15]?'<br><span style="color:#94a3b8;font-size:9px">'+r[15]+'</span>':''}}</td>
      <td style="text-align:center">${{isn?`<span class="nb">New</span>`:'<span style="color:#d1d5db;font-size:9px">―</span>'}}</td>
    </tr>`;
  }}).join("");
}}
function rp(){{
  const tot=Math.ceil(filtered.length/PS);
  if(tot<=1){{document.getElementById("pg").innerHTML='';return;}}
  let h='';
  h+=`<button class="pb" onclick="gp(1)" ${{pg===1?'disabled':''}}>«</button>`;
  h+=`<button class="pb" onclick="gp(${{pg-1}})" ${{pg===1?'disabled':''}}>‹</button>`;
  let s=Math.max(1,pg-3),e=Math.min(tot,s+6);
  if(e-s<6)s=Math.max(1,e-6);
  for(let p=s;p<=e;p++)h+=`<button class="pb" onclick="gp(${{p}})" style="${{p===pg?'background:#0f172a;color:#fff;border-color:#0f172a':''}}">${{p}}</button>`;
  h+=`<button class="pb" onclick="gp(${{pg+1}})" ${{pg===tot?'disabled':''}}>›</button>`;
  h+=`<button class="pb" onclick="gp(${{tot}})" ${{pg===tot?'disabled':''}}>»</button>`;
  h+=`<span class="pi">${{pg}}/${{tot}}ページ（各${{PS}}件）</span>`;
  document.getElementById("pg").innerHTML=h;
}}
function gp(p){{pg=p;rt();rp();window.scrollTo({{top:0,behavior:'smooth'}});}}

function sd(i){{
  const r=RAW[i];
  const s=st(r[11]),vs=vt(r[16]),isn=r[20]&&r[20].trim();
  document.getElementById("mb").innerHTML=`
    <div class="mh">
      <div style="flex:1;padding-right:12px">
        <div style="display:flex;gap:7px;align-items:center;margin-bottom:5px;flex-wrap:wrap">
          ${{isn?`<span class="nb">今回更新</span>`:''}}
          ${{r[8]?`<span class="tag" style="background:#fce7f3;color:#9d174d">基礎的医薬品</span>`:''}}
          ${{r[9]?`<span class="tag" style="background:#fef3c7;color:#92400e">供給確保医薬品 ${{r[9]}}</span>`:''}}
        </div>
        <h2 style="font-size:15px;font-weight:900;color:#0f172a;line-height:1.3;margin-bottom:4px">${{r[5]}}</h2>
        <div style="color:#64748b;font-size:11px">${{r[2]}}${{r[3]?' / '+r[3]:''}}</div>
        ${{r[4]?`<div style="color:#94a3b8;font-size:10px;margin-top:2px">YJコード: ${{r[4]}}</div>`:''}}
      </div>
      <button class="mc" onclick="cm()">✕</button>
    </div>
    <div class="ig">
      <div class="ic" style="background:${{s.b}};border-color:${{s.bd}}">
        <div class="l">⑫出荷対応の状況</div>
        <div class="v" style="color:${{s.c}}">${{r[11]||'―'}}</div>
        ${{r[12]?`<div style="font-size:9px;color:#94a3b8;margin-top:3px">更新日: ${{r[12]}}</div>`:''}}
      </div>
      <div class="ic" style="background:${{vs?vs.b:'#f8fafc'}}">
        <div class="l">⑰出荷量の状況</div>
        <div class="v" style="color:${{vs?vs.c:'#334155'}}">${{r[16]||'―'}}</div>
        ${{r[17]?`<div style="font-size:9px;color:#94a3b8;margin-top:3px">改善見込: ${{r[17]}}</div>`:''}}
      </div>
      <div class="ic">
        <div class="l">⑦製造販売業者</div>
        <div class="v">${{r[6]||'―'}}</div>
        <div style="font-size:9px;color:#94a3b8;margin-top:2px">${{r[7]||''}} / ${{r[0]||''}}</div>
      </div>
      <div class="ic">
        <div class="l">②薬効分類</div>
        <div class="v">${{r[1]||'―'}}</div>
        ${{r[10]?`<div style="font-size:9px;color:#94a3b8;margin-top:2px">薬価収載: ${{r[10]}}</div>`:''}}
      </div>
    </div>
    ${{r[13]&&r[13]!=='７．ー'?`<div class="ir" style="background:#fffbeb;border-color:#fde68a"><div class="l" style="color:#b45309">⑭出荷停止等の理由</div><div class="v" style="color:#78350f">${{r[13]}}</div></div>`:''}}
    ${{r[14]&&r[14]!=='エ． －'?`<div class="ir" style="background:${{r[14]==='ウ． 未定'?'#f5f3ff':'#f0fdf4'}};border-color:${{r[14]==='ウ． 未定'?'#c4b5fd':'#bbf7d0'}}"><div class="l" style="color:${{r[14]==='ウ． 未定'?'#7c3aed':'#16a34a'}}">⑮解消見込み</div><div class="v" style="color:${{r[14]==='ウ． 未定'?'#4c1d95':'#166534'}}">${{r[14]}}${{r[15]?' ／ '+r[15]:''}}</div></div>`:''}}
    ${{r[18]?`<div class="ir" style="background:#f0f9ff;border-color:#bae6fd"><div class="l" style="color:#0369a1">⑲出荷量の改善見込み量</div><div class="v" style="color:#075985">${{r[18]}}</div></div>`:''}}
    <div class="db">
      <div class="dt">出荷対応区分の定義</div>
      ${{[["①通常出荷","全ての受注に対応でき、十分な在庫量が確保できている状況"],
         ["②限定出荷（自社の事情）","自社の事情により、全ての受注に対応できない状況"],
         ["③限定出荷（他社品の影響）","他社品の影響等にて、全ての受注に対応できない状況"],
         ["④限定出荷（その他）","その他の理由にて、全ての受注に対応できない状況"],
         ["⑤供給停止","供給を停止している状況"]].map(([k,v])=>{{const ss=st(k);return `<div class="dr"><span class="badge" style="background:${{ss.b}};color:${{ss.c}};border-color:${{ss.bd}}">${{k}}</span><span style="font-size:10px;color:#475569">${{v}}</span></div>`;}}).join('')}}
    </div>
    <div style="margin-top:12px;text-align:right"><button class="pb" onclick="cm()" style="padding:8px 20px">閉じる</button></div>
  `;
  document.getElementById("modal").classList.add("open");
}}
function cm(){{document.getElementById("modal").classList.remove("open");}}
document.addEventListener("keydown",e=>{{if(e.key==="Escape")cm();}});
</script>
</body>
</html>"""
    return html


# ── メイン処理 ────────────────────────────────────────────────────────────────
def main():
    DOCS_DIR.mkdir(exist_ok=True)

    excel_url, label = get_latest_excel_url()
    excel_data, changed = download_excel(excel_url)

    if not changed:
        print("データに変更がないため終了します。")
        return

    drug_data = parse_excel(excel_data)
    generated_at = datetime.now(JST).strftime("%Y-%m-%d %H:%M JST")

    html = build_html(drug_data, label, generated_at)

    # index.html として保存（GitHub Pages のエントリーポイント）
    out_path = DOCS_DIR / "index.html"
    out_path.write_text(html, encoding="utf-8")
    print(f"  → {out_path} ({len(html)//1024} KB)")

    # 更新履歴に日付を記録
    history_path = DOCS_DIR / "update_history.txt"
    with open(history_path, "a", encoding="utf-8") as f:
        f.write(f"{generated_at} | {label} | {len(drug_data)} 品目 | {excel_url}\n")

    print(f"\n✅ 完了: {len(drug_data):,} 品目 → {out_path}")


if __name__ == "__main__":
    main()
