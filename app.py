# -*- coding: utf-8 -*-
"""
Punchlist Dashboard — สถานี PP18 SI YAN (MRT สายสีม่วง, Contract 1) · R2
Streamlit app: อ่าน/แก้ข้อมูลจาก Google Sheet แบบ near real-time + ดูรูปแบบราย จุด
"""
import os, io, base64
from datetime import datetime, timezone, timedelta
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import json
import streamlit.components.v1 as components

# ----------------------------------------------------------------------------
st.set_page_config(page_title="Punchlist PP18 · R2", page_icon="🚇", layout="wide")
APP_DIR = os.path.dirname(os.path.abspath(__file__))

# ---- column names (ตรงกับหัวตารางใน Google Sheet) ----
C_NO="ลำดับ"; C_NICK="ชื่อเรียก"; C_FLOOR="ชั้น"; C_SYS="ระบบ"; C_DWGF="ไฟล์แบบ"; C_DWG="Drawing No."
C_PAGE="หน้า"; C_RED="ข้อความในกรอบแดง"; C_OWNER="ผู้รับผิดชอบ"; C_START="กำหนดเริ่ม"
C_DUE="กำหนดเสร็จ"; C_LOC="ตำแหน่ง/บริเวณ"; C_DETAIL="รายละเอียดงาน"; C_STATUS="สถานะ"
C_DONE="วันที่เสร็จจริง"; C_NOTE="หมายเหตุ"
ALL_COLS=[C_NO,C_NICK,C_FLOOR,C_SYS,C_DWGF,C_DWG,C_PAGE,C_RED,C_OWNER,C_START,C_DUE,
          C_LOC,C_DETAIL,C_STATUS,C_DONE,C_NOTE]

STATUS_ORDER=["เลยกำหนด–รอตรวจสอบ","กำลังดำเนินการ","รอเชื่อม","รอวัสดุ","รอดำเนินการ","จบงาน","ยกเลิก–รวมแผนใหม่"]
STATUS_META={
 "เลยกำหนด–รอตรวจสอบ":{"key":"crit","color":"#d03b3b","icon":"🔴","short":"เลยกำหนด"},
 "กำลังดำเนินการ":{"key":"warn","color":"#fab219","icon":"🟡","short":"กำลังดำเนินการ"},
 "รอเชื่อม":{"key":"conn","color":"#2f7fd1","icon":"🔵","short":"รอเชื่อม"},
 "รอวัสดุ":{"key":"matl","color":"#9b59b6","icon":"🟣","short":"รอวัสดุ"},
 "รอดำเนินการ":{"key":"neut","color":"#8a8a86","icon":"⚪","short":"รอดำเนินการ"},
 "จบงาน":{"key":"done","color":"#1f9d57","icon":"✅","short":"จบงาน"},
 "ยกเลิก–รวมแผนใหม่":{"key":"cancel","color":"#4b5563","icon":"⛔","short":"ยกเลิก"},
 "เสร็จแล้ว":{"key":"done","color":"#1f9d57","icon":"✅","short":"จบงาน"},
}
FLOOR_ORDER=["Multipurpose","Concourse","Upper Platform"]
FLOOR_COLOR={"Multipurpose":"#2a78d6","Concourse":"#eb6834","Upper Platform":"#1baf7a"}
SERIES="#2a78d6"; CRIT="#d03b3b"; INK="#0b0b0b"; MUTED="#898781"
TH_MONTH={1:"ม.ค.",2:"ก.พ.",3:"มี.ค.",4:"เม.ย.",5:"พ.ค.",6:"มิ.ย.",7:"ก.ค.",8:"ส.ค.",9:"ก.ย.",10:"ต.ค.",11:"พ.ย.",12:"ธ.ค."}

def bkk_today():
    return (datetime.now(timezone.utc)+timedelta(hours=7)).date()

def sget(k, default=None):
    try: return st.secrets.get(k, default)
    except Exception: return default

def has_secret(k):
    try: return k in st.secrets
    except Exception: return False

# ----------------------------------------------------------------------------
# Data source: Google Sheet (service account) > public CSV url > local seed
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_ws(wsname=None):
    # รองรับ service account 2 แบบ: วาง JSON ทั้งก้อน (gcp_service_account_json) หรือ TOML table (gcp_service_account)
    info=None
    if has_secret("gcp_service_account_json"):
        import json as _json
        info=_json.loads(str(st.secrets["gcp_service_account_json"]))
    elif has_secret("gcp_service_account"):
        info=dict(st.secrets["gcp_service_account"])
    if info and has_secret("sheet_url"):
        import gspread
        from google.oauth2.service_account import Credentials
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
        creds=Credentials.from_service_account_info(info, scopes=scopes)
        gc=gspread.authorize(creds)
        sh=gc.open_by_url(st.secrets["sheet_url"])
        if wsname:
            try: return sh.worksheet(wsname)
            except Exception: return sh.add_worksheet(title=wsname, rows=200, cols=max(len(ALL_COLS),20))
        _wn=sget("worksheet","Punchlist")
        try: return sh.worksheet(_wn)
        except Exception: return sh.sheet1
    return None

# ---- รูปแทน (upload) : เก็บรูปที่อัปโหลดเป็น base64 (แบ่งหลายเซลล์) ในแท็บซ่อนของ Google Sheet ----
OVERRIDE_WS = "รูปแทน"
CHUNK = 45000        # Google Sheet จำกัด ~50,000 ตัวอักษร/เซลล์
MAXCH = 6            # สูงสุด 6 เซลล์/รูป (~270k = รูปคมชัดได้)
_LASTCOL = chr(ord('A')+MAXCH)   # 'G'

def _to_chunks(b64):
    ch=[b64[i:i+CHUNK] for i in range(0,len(b64),CHUNK)]
    return (ch+[""]*MAXCH)[:MAXCH]

@st.cache_data(ttl=60, show_spinner=False)
def load_overrides():
    """คืน dict {ลำดับจุด: base64jpg} ของรูปที่อัปโหลดทับไว้ (ถ้าไม่มีแท็บ = {})"""
    try:
        ws=get_ws()
        if ws is None: return {}
        try: ows=ws.spreadsheet.worksheet(OVERRIDE_WS)
        except Exception: return {}
        out={}
        for r in ows.get_all_values()[1:]:
            if r and str(r[0]).strip().isdigit():
                b="".join(r[1:1+MAXCH]).strip()
                if b: out[int(r[0])]=b
        return out
    except Exception:
        return {}

def compress_to_b64(data, maxw=1000):
    """ย่อ + บีบ JPEG ให้ base64 พอดี ≤ MAXCH เซลล์ โดยคงคุณภาพให้อ่านแบบได้"""
    from PIL import Image
    img=Image.open(io.BytesIO(data)).convert("RGB")
    w,h=img.size
    if w>maxw: img=img.resize((maxw,int(h*maxw/w)), Image.LANCZOS)
    limit=CHUNK*MAXCH
    for q in (85,78,70,60,50,40):
        buf=io.BytesIO(); img.save(buf,format="JPEG",quality=q,optimize=True)
        b=base64.b64encode(buf.getvalue()).decode()
        if len(b)<=limit: return b
    img=img.resize((800,int(img.size[1]*800/img.size[0])), Image.LANCZOS)
    buf=io.BytesIO(); img.save(buf,format="JPEG",quality=45,optimize=True)
    return base64.b64encode(buf.getvalue()).decode()

def save_override(no, b64):
    ws=get_ws(); sh=ws.spreadsheet
    try: ows=sh.worksheet(OVERRIDE_WS)
    except Exception:
        ows=sh.add_worksheet(title=OVERRIDE_WS, rows=200, cols=MAXCH+1)
        ows.update(values=[["ลำดับ"]+[f"b64_{i+1}" for i in range(MAXCH)]], range_name="A1")
    row=[str(no)]+_to_chunks(b64)     # กว้างคงที่ MAXCH+1 → เขียนทับล้างของเก่าเสมอ
    col=ows.col_values(1); rowidx=None
    for i,v in enumerate(col[1:], start=2):
        if str(v).strip()==str(no): rowidx=i; break
    if rowidx is None:
        ows.append_row(row, value_input_option="RAW")
    else:
        ows.update(values=[row], range_name=f"A{rowidx}:{_LASTCOL}{rowidx}")

# ---- รูปหลักฐานงานเสร็จหน้างาน : เก็บได้หลายรูป/จุด เป็น base64 ในแท็บ "รูปงานเสร็จ" ----
DONE_WS = "รูปงานเสร็จ"
DONE_HEADERS = ["ลำดับ","id","วันที่","หมายเหตุ"] + [f"b64_{i+1}" for i in range(MAXCH)]
DONE_NCOL = len(DONE_HEADERS)
DONE_LASTCOL = chr(ord('A')+DONE_NCOL-1)

def pin_bad(p):
    ep=sget("edit_pin","")
    return ep=="" or str(p)!=str(ep)

def _bkk_stamp():
    return (datetime.now(timezone.utc)+timedelta(hours=7)).strftime("%Y%m%d%H%M%S%f")

@st.cache_data(ttl=60, show_spinner=False)
def load_done_photos():
    """คืน dict {ลำดับจุด: [ {id,date,note,b64}, ... ]} ของรูปงานเสร็จหน้างาน"""
    out={}
    try:
        ws=get_ws()
        if ws is None: return out
        try: dws=ws.spreadsheet.worksheet(DONE_WS)
        except Exception: return out
        for r in dws.get_all_values()[1:]:
            if not r or not str(r[0]).strip().isdigit(): continue
            b="".join(r[4:4+MAXCH]).strip() if len(r)>4 else ""
            if not b: continue
            out.setdefault(int(r[0]),[]).append({
                "id":(r[1] if len(r)>1 else "").strip(),
                "date":(r[2] if len(r)>2 else "").strip(),
                "note":(r[3] if len(r)>3 else "").strip(),
                "b64":b,
            })
    except Exception:
        pass
    return out

def save_done_photo(no, b64, note=""):
    ws=get_ws(); sh=ws.spreadsheet
    try: dws=sh.worksheet(DONE_WS)
    except Exception:
        dws=sh.add_worksheet(title=DONE_WS, rows=400, cols=DONE_NCOL)
        dws.update(values=[DONE_HEADERS], range_name="A1")
    pid=_bkk_stamp()
    row=[str(no),pid,bkk_today().strftime("%Y-%m-%d"),str(note)[:200]]+_to_chunks(b64)
    dws.append_row(row, value_input_option="RAW")
    return pid

def del_done_photo(no, pid):
    ws=get_ws(); sh=ws.spreadsheet
    try: dws=sh.worksheet(DONE_WS)
    except Exception: return
    for i,r in enumerate(dws.get_all_values()[1:], start=2):
        if str(r[0]).strip()==str(no) and len(r)>1 and str(r[1]).strip()==str(pid):
            dws.delete_rows(i); return

# ---- แบบแปลนติดจุด (interactive plan viewer) ----
PLANS_DIR=os.path.join(APP_DIR,"plans")
@st.cache_data(show_spinner=False)
def load_plans_meta():
    import json as _j
    try:
        with open(os.path.join(PLANS_DIR,"plans_meta.json"),encoding="utf-8") as fh:
            return _j.load(fh)
    except Exception:
        return {"plans":[],"points":{}}
@st.cache_data(show_spinner=False)
def plan_img_url(key):
    p=os.path.join(PLANS_DIR,key+".jpg")
    if not os.path.exists(p): return ""
    return "data:image/jpeg;base64,"+base64.b64encode(open(p,"rb").read()).decode()

_PLAN_TPL = r"""<meta charset="utf-8">
<style>
#wrap{position:relative;width:100%;height:__H__px;overflow:hidden;background:#eceae6;border:1px solid #d9d7d2;border-radius:10px;touch-action:none;font-family:-apple-system,Segoe UI,Roboto,'Noto Sans Thai',sans-serif;}
#stage{position:absolute;left:0;top:0;transform-origin:0 0;will-change:transform;}
#plan{display:block;-webkit-user-select:none;user-select:none;pointer-events:none;}
.mk{position:absolute;transform:translate(-50%,-50%);width:30px;height:30px;border-radius:50%;display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:14px;border:2.5px solid #fff;box-shadow:0 1px 5px rgba(0,0,0,.55);cursor:pointer;}
.mk:hover{box-shadow:0 0 0 4px rgba(0,0,0,.22),0 1px 5px rgba(0,0,0,.55);}
#bar{position:absolute;left:10px;top:10px;right:10px;background:rgba(22,22,22,.88);color:#fff;padding:8px 12px;border-radius:8px;font-size:13px;display:none;z-index:5;pointer-events:none;}
#card{position:absolute;right:10px;bottom:10px;width:min(360px,74%);max-height:80%;overflow:auto;background:#fff;border:1px solid #ccc;border-radius:10px;box-shadow:0 5px 20px rgba(0,0,0,.3);padding:14px 16px 16px;font-size:13.5px;line-height:1.55;display:none;z-index:6;}
#card h4{margin:0 0 8px;font-size:15.5px;padding-right:20px;}
#card .x{position:absolute;right:12px;top:9px;cursor:pointer;color:#999;font-size:19px;font-weight:700;}
#ctl{position:absolute;left:10px;bottom:10px;display:flex;gap:6px;z-index:5;}
#ctl button{width:34px;height:34px;border:none;border-radius:8px;background:rgba(22,22,22,.82);color:#fff;font-size:18px;cursor:pointer;line-height:1;}
#hint{position:absolute;right:10px;top:10px;background:rgba(255,255,255,.86);color:#555;padding:4px 8px;border-radius:6px;font-size:11px;z-index:4;}
</style>
<div id="wrap">
 <div id="stage"><img id="plan" src="__IMG__"/></div>
 <div id="bar"></div><div id="card"></div>
 <div id="hint">ลาก = เลื่อน · ล้อเมาส์ = ซูม</div>
 <div id="ctl"><button id="zin">+</button><button id="zout">−</button><button id="zr">⤢</button></div>
</div>
<script>
const PTS=__PTS__;
const wrap=document.getElementById('wrap'),stage=document.getElementById('stage'),plan=document.getElementById('plan'),bar=document.getElementById('bar'),card=document.getElementById('card');
let scale=1,tx=0,ty=0;
function apply(){stage.style.transform='translate('+tx+'px,'+ty+'px) scale('+scale+')';}
function esc(s){return (s==null?'':(''+s)).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
function showCard(p){card.style.display='block';
 card.innerHTML='<span class="x">✕</span><h4>'+p.icon+' จุดที่ '+p.no+(p.nick?' ('+esc(p.nick)+')':'')+' — '+esc(p.status)+'</h4>'+
 (p.nick?'<b>ชื่อเรียก:</b> '+esc(p.nick)+'<br>':'')+
 '<b>ชั้น/ระบบ:</b> '+esc(p.floor)+' · '+esc(p.sys)+'<br><b>Drawing:</b> '+esc(p.dwg)+' (หน้า '+esc(p.page)+')<br>'+
 '<b>ตำแหน่ง:</b> '+esc(p.loc)+'<br><b>รายละเอียดงาน:</b> '+esc(p.detail)+'<br><b>ในกรอบแดง:</b> '+esc(p.red)+'<br>'+
 '<b>ผู้รับผิดชอบ:</b> '+esc(p.owner)+'<br><b>กำหนด:</b> '+esc(p.start)+' → '+esc(p.due)+' ('+esc(p.days)+')<br>'+
 '<b>หมายเหตุ:</b> '+esc(p.note||'–');
 card.querySelector('.x').onclick=function(e){e.stopPropagation();card.style.display='none';};}
function build(){PTS.forEach(function(p){var m=document.createElement('div');m.className='mk';m.style.left=p.x+'%';m.style.top=p.y+'%';m.style.background=p.color;m.textContent=p.no;
 m.onmouseenter=function(){bar.style.display='block';bar.innerHTML='<b>จุดที่ '+p.no+(p.nick?' ('+esc(p.nick)+')':'')+'</b> · '+p.icon+' '+esc(p.status)+' · '+esc(p.owner)+' · ครบ '+esc(p.due)+' ('+esc(p.days)+')';};
 m.onmouseleave=function(){bar.style.display='none';};
 m.onclick=function(e){e.stopPropagation();showCard(p);};stage.appendChild(m);});}
function initFit(){var ww=wrap.clientWidth,wh=wrap.clientHeight,iw=plan.naturalWidth,ih=plan.naturalHeight;if(!iw||!ww){return requestAnimationFrame(initFit);}var dw=ww,dh=ww*ih/iw;stage.style.width=dw+'px';stage.style.height=dh+'px';plan.style.width=dw+'px';plan.style.height=dh+'px';var s=Math.min(1,wh/dh);scale=s;tx=(ww-dw*s)/2;ty=(wh-dh*s)/2;apply();}
plan.onload=initFit;initFit();
function zoomAt(cx,cy,f){var ns=Math.min(9,Math.max(0.4,scale*f)),r=ns/scale;tx=cx-(cx-tx)*r;ty=cy-(cy-ty)*r;scale=ns;apply();}
wrap.addEventListener('wheel',function(e){e.preventDefault();var b=wrap.getBoundingClientRect();zoomAt(e.clientX-b.left,e.clientY-b.top,e.deltaY<0?1.15:1/1.15);},{passive:false});
document.getElementById('zin').onclick=function(){zoomAt(wrap.clientWidth/2,wrap.clientHeight/2,1.3);};
document.getElementById('zout').onclick=function(){zoomAt(wrap.clientWidth/2,wrap.clientHeight/2,1/1.3);};
document.getElementById('zr').onclick=initFit;
var drag=false,px2,py2;
wrap.addEventListener('pointerdown',function(e){if(e.target.classList&&e.target.classList.contains('mk'))return;drag=true;px2=e.clientX;py2=e.clientY;try{wrap.setPointerCapture(e.pointerId);}catch(_){}});
wrap.addEventListener('pointermove',function(e){if(!drag)return;tx+=e.clientX-px2;ty+=e.clientY-py2;px2=e.clientX;py2=e.clientY;apply();});
wrap.addEventListener('pointerup',function(){drag=false;});
wrap.addEventListener('click',function(){card.style.display='none';});
build();
</script>"""
def build_plan_html(img_url, pts, height=640):
    import json as _j
    return _PLAN_TPL.replace("__H__",str(height)).replace("__IMG__",img_url).replace("__PTS__",_j.dumps(pts,ensure_ascii=False))

@st.cache_data(ttl=60, show_spinner=False)
def load_raw():
    """คืน (DataFrame ดิบตามหัวตาราง, โหมด)"""
    try:
        ws=get_ws()
    except Exception as e:
        st.warning(f"เชื่อม Google Sheet ไม่สำเร็จ ({e}). ใช้ข้อมูลตั้งต้นแทน")
        ws=None
    if ws is not None:
        df=pd.DataFrame(ws.get_all_records())
        mode="gsheet"
    else:
        url=sget("public_csv_url","")
        if url:
            df=pd.read_csv(url); mode="csv_url"
        else:
            df=pd.read_csv(os.path.join(APP_DIR,"data","punchlist_seed.csv")); mode="local"
    # ทนทานต่อหัวคอลัมน์เพี้ยน (encoding/แปลง CSV): ถ้าชื่อไม่ตรงแต่จำนวน≥15 ให้จับคู่ตามตำแหน่งที่รู้ลำดับแน่นอน
    if (not all(c in df.columns for c in ALL_COLS)) and df.shape[1] >= len(ALL_COLS):
        df = df.rename(columns={df.columns[i]: ALL_COLS[i] for i in range(len(ALL_COLS))})
    for c in ALL_COLS:
        if c not in df.columns: df[c]=""
    df=df[[c for c in ALL_COLS if c in df.columns]]
    df=df.fillna("")   # กันค่า NaN โผล่เป็น "nan" ในหน้าจอ
    df=df[df[C_NO].astype(str).str.strip()!=""]      # ตัดแถวว่าง
    return df.reset_index(drop=True), mode

def enrich(df):
    df=df.copy()
    df[C_NO]=pd.to_numeric(df[C_NO], errors="coerce")
    today=bkk_today()
    def dleft(s):
        try:
            y,m,d=str(s).split(" ")[0].split("-")[:3]
            from datetime import date as _d
            return (_d(int(y),int(m),int(d))-today).days
        except Exception:
            return None
    df["_days"]=df[C_DUE].map(dleft)
    df["_skey"]=df[C_STATUS].map(lambda s: STATUS_META.get(str(s).strip(),{}).get("key","neut"))
    df["_month"]=df[C_DUE].map(lambda s: str(s)[:7] if str(s)[:4].isdigit() else "")
    return df

def fmt_days(r):
    if r["_skey"]=="done": return "จบงานแล้ว"
    d=r["_days"]
    if d is None or pd.isna(d): return "–"
    d=int(d)
    if r["_skey"]=="crit" or d<0: return f"เลย {abs(d)} วัน"
    return f"{d} วัน"

# ---- ตรวจความต่างก่อนบันทึก (diff) + กล่องยืนยัน ----
def _canon_no(v):
    s="" if v is None else str(v).strip()
    if s.lower() in ("","nan","none"): return ""
    try:
        f=float(s); return str(int(f)) if f==int(f) else str(f)
    except Exception:
        return s

def _cellval(v):
    s="" if v is None else str(v).strip()
    return "" if s.lower()=="nan" else s

def _same(a,b):
    a=_cellval(a); b=_cellval(b)
    if a==b: return True
    try: return float(a)==float(b)
    except Exception: return False

def _s(v):
    if v is None: return ""
    if isinstance(v,float):
        if v!=v: return ""            # NaN
        if v==int(v): return str(int(v))
        return str(v)
    s=str(v); return "" if s.lower()=="nan" else s

def compute_diff(before, after):
    def as_map(df):
        m={}
        for _,r in df.iterrows():
            k=_canon_no(r.get(C_NO,""))
            if k: m[k]=r
        return m
    bm=as_map(before); am=as_map(after)
    bkeys, akeys = set(bm), set(am)
    kf=lambda x: (0,float(x)) if x.replace('.','',1).isdigit() else (1,x)
    added=sorted(akeys-bkeys, key=kf)
    deleted=sorted(bkeys-akeys, key=kf)
    mods=[]
    for k in sorted(akeys & bkeys, key=kf):
        ch=[(c,_cellval(bm[k].get(c,"")),_cellval(am[k].get(c,"")))
            for c in ALL_COLS if not _same(bm[k].get(c,""), am[k].get(c,""))]
        if ch: mods.append((k,ch))
    return {"added":added,"deleted":deleted,"mods":mods,
            "n_add":len(added),"n_del":len(deleted),"n_mod":len(mods)}

@st.dialog("ยืนยันการแก้ไขก่อนบันทึก")
def confirm_save_dialog():
    diff=st.session_state.get("_pending_diff"); out=st.session_state.get("_pending_edited")
    if diff is None or out is None:
        st.write("ไม่มีข้อมูลรอบันทึก"); return
    st.markdown(f"ตรวจการเปลี่ยนแปลงก่อนเขียนกลับ Google Sheet — "
                f"**➕ เพิ่ม {diff['n_add']} · 🗑️ ลบ {diff['n_del']} · ✏️ แก้ไข {diff['n_mod']} แถว**")
    if diff["added"]:
        st.markdown("**➕ เพิ่มจุดใหม่:** "+", ".join(f"จุดที่ {k}" for k in diff["added"]))
    if diff["deleted"]:
        st.markdown("**🗑️ ลบจุด:** "+", ".join(f"จุดที่ {k}" for k in diff["deleted"]))
    if diff["mods"]:
        st.markdown("**✏️ แก้ไขรายจุด:**")
        for i,(k,ch) in enumerate(diff["mods"]):
            if i>=25:
                st.caption(f"… และอีก {len(diff['mods'])-25} แถว"); break
            lines="; ".join(f"{c}: “{ov or '—'}” → “{nv or '—'}”" for c,ov,nv in ch)
            st.markdown(f"- **จุดที่ {k}** — {lines}")
    st.divider()
    c1,c2=st.columns(2)
    if c1.button("✅ ยืนยันบันทึก", type="primary", use_container_width=True):
        try:
            ws=get_ws()
            values=[list(out.columns)]+[[_s(v) for v in row] for row in out.values.tolist()]
            ws.clear(); ws.update(values=values, range_name="A1")
            st.cache_data.clear()
            st.session_state["_saved_rows"]=len(out)
            for kk in ("_pending_diff","_pending_edited","_show_confirm"): st.session_state.pop(kk,None)
            st.rerun()
        except Exception as e:
            st.error(f"บันทึกไม่สำเร็จ: {e}")
    if c2.button("✖ ยกเลิก", use_container_width=True):
        for kk in ("_pending_diff","_pending_edited","_show_confirm"): st.session_state.pop(kk,None)
        st.rerun()

# ----------------------------------------------------------------------------
# Charts
# ----------------------------------------------------------------------------
def hbar(pairs, colors=None, height=None):
    labels=[p[0] for p in pairs][::-1]
    values=[p[1] for p in pairs][::-1]
    if isinstance(colors,list): colors=colors[::-1]
    fig=go.Figure(go.Bar(
        x=values, y=labels, orientation="h",
        marker_color=colors if colors else SERIES,
        text=values, textposition="outside", cliponaxis=False,
        hovertemplate="%{y}: %{x} จุด<extra></extra>",
    ))
    fig.update_layout(
        height=height or (40+28*len(labels)), margin=dict(l=6,r=24,t=6,b=6),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(size=13,color=INK), showlegend=False,
        xaxis=dict(visible=False), yaxis=dict(tickfont=dict(size=13)),
    )
    return fig

def month_stacked(df):
    months=sorted([m for m in df["_month"].unique() if m])
    lab=[f"{TH_MONTH.get(int(m[5:7]),m[5:7])} {m[:4]}" for m in months]
    crit=[int(((df["_month"]==m)&(df["_skey"]=="crit")).sum()) for m in months]
    rest=[int((df["_month"]==m).sum())-c for m,c in zip(months,crit)]
    fig=go.Figure()
    fig.add_bar(y=lab[::-1], x=crit[::-1], orientation="h", name="เลยกำหนด",
                marker_color=CRIT, text=[c or "" for c in crit[::-1]], textposition="inside")
    fig.add_bar(y=lab[::-1], x=rest[::-1], orientation="h", name="ตามแผน (ยังไม่เลยกำหนด)",
                marker_color=SERIES, text=rest[::-1], textposition="inside")
    fig.update_layout(barmode="stack", height=40+34*len(lab),
        margin=dict(l=6,r=16,t=6,b=6), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(size=13,color=INK), legend=dict(orientation="h",y=-0.25,x=0),
        xaxis=dict(visible=False), yaxis=dict(tickfont=dict(size=13)))
    return fig

# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
# ---- เมนูหลัก (hub) : เข้าเว็บเจอเมนูเลือกงานก่อน แล้วลิงก์เข้าแต่ละส่วน ----
def _goto(v):
    st.session_state["view"]=v; st.rerun()

MENU_APPS=[
 {"icon":"📋","title":"Punchlist PP18 SI YAN","desc":"ติดตามงานตามกรอบสีแดง 75 จุด — ตาราง · รูปแบบ · แบบแปลนติดจุด · แก้ไขข้อมูล","view":"punchlist","tag":"พร้อมใช้","ready":True},
 {"icon":"🚒","title":"Punchlist FP · Multipurpose","desc":"งานระบบดับเพลิง (FP) ชั้น Multipurpose — ตาราง · แบบแปลนติดจุด · แก้ไขข้อมูล (ชุดแยก แก้ไขได้เอง)","view":"punchlist_fp","tag":"พร้อมใช้","ready":True},
 {"icon":"🧰","title":"งานคงเหลือ (Remaining Work)","desc":"แผนงานระบบที่ยังเหลือ — ไทม์ไลน์ Gantt · ตารางกรอง · แบบแปลนโซน 5 ชั้น (ต้องใส่ PIN)","view":"remaining","tag":"กำลังปรับปรุง","ready":True},
 {"icon":"📝","title":"บันทึกผลงานประจำวัน (Actual)","desc":"บันทึกความยาวท่อติดตั้งจริงรายวัน — ระบบ/โซน/หย่อม · Progress รายวัน · สรุปสะสม · Export (ต้องใส่ PIN)","view":"actual","tag":"พร้อมใช้","ready":True},
 {"icon":"➕","title":"เพิ่มงานถัดไป…","desc":"ช่องสำหรับงานส่วนใหม่ในอนาคต (เพิ่มการ์ดในเมนูได้เรื่อยๆ)","view":None,"tag":"เร็วๆ นี้","ready":False},
]

def render_menu():
    st.markdown("""<style>
    .hubic{font-size:32px;line-height:1;} .hubt{font-weight:700;font-size:17px;margin:8px 0 5px;color:#1a1a1a;}
    .hubd{color:#6b6862;font-size:13px;line-height:1.5;min-height:58px;}
    .hubtag{display:inline-block;font-size:11px;font-weight:600;padding:2px 10px;border-radius:20px;background:#e9f1ff;color:#2a78d6;}
    .hubtag.soon{background:#f0efec;color:#9a978f;}
    .hubtag.wip{background:#fdf3d9;color:#c88a00;}
    </style>""",unsafe_allow_html=True)
    st.markdown("<div style='padding:6px 2px 2px'><span style='font-size:30px;font-weight:800'>🚇 MRT PP18 · ศูนย์รวมงาน</span>"
                "<div style='color:#6b6862;margin-top:4px'>สายสีม่วง (Contract 1) — SI YAN · เลือกงานที่ต้องการเปิด</div></div>",unsafe_allow_html=True)
    st.write("")
    cols=st.columns(3, gap="medium")
    for i,a in enumerate(MENU_APPS):
        with cols[i%3]:
            with st.container(border=True):
                tg=("hubtag soon" if not a["ready"]
                    else "hubtag wip" if a.get("tag")=="กำลังปรับปรุง" else "hubtag")
                st.markdown(f"<div class='hubic'>{a['icon']}</div><div class='hubt'>{a['title']}</div>"
                            f"<div class='hubd'>{a['desc']}</div><span class='{tg}'>{a['tag']}</span>",unsafe_allow_html=True)
                st.write("")
                if a["ready"] and a["view"]:
                    if st.button("เปิด →", key=f"hub_{i}", use_container_width=True,
                                 type=("primary" if a["view"]=="punchlist" else "secondary")):
                        _goto(a["view"])
                else:
                    st.button("เร็วๆ นี้", key=f"hub_{i}", use_container_width=True, disabled=True)
    st.write(""); st.divider()
    st.caption("ศูนย์รวมเครื่องมือติดตามงาน MRT PP18 · เพิ่มงานส่วนใหม่ได้โดยเพิ่มการ์ดในเมนูนี้")

def render_placeholder(v):
    if st.columns([1,4])[0].button("← เมนูหลัก", key="ph_back"): _goto("menu")
    st.markdown("## 🧰 งานส่วนใหม่ (ตัวอย่าง)")
    st.info("นี่คือหน้าตัวอย่างของ **งานส่วนใหม่** — โครงหน้าและการลิงก์จากเมนูพร้อมแล้ว เหลือแค่ใส่เนื้อหาจริง")
    st.markdown("- บอกผมได้เลยว่าส่วนนี้เป็นงานอะไร (เช่น อีกสถานี, Punchlist อีกชุด, ตารางตรวจงาน, รายงานอื่น)\n"
                "- ผมทำให้เหมือน Punchlist PP18 (ตาราง/รูป/แบบแปลน/แก้ไข) หรือรูปแบบใหม่ตามต้องการก็ได้")
    st.caption("เมนูหลักลิงก์ได้หลายงาน — เพิ่มการ์ดใหม่เมื่อไหร่ก็ได้")

# ---- งานคงเหลือ (Remaining Work) : หน้าใหม่ ใส่ PIN ก่อนเข้า + ป้ายกำลังปรับปรุง ----
@st.cache_data(show_spinner=False)
def load_remaining_html():
    try:
        with open(os.path.join(APP_DIR,"remaining_work.html"),encoding="utf-8") as fh:
            return fh.read()
    except Exception:
        return ""

def render_remaining():
    if st.columns([1,5])[0].button("← เมนูหลัก", key="rem_back"):
        st.session_state["view"]="menu"; st.session_state.pop("_rem_ok",None); st.rerun()
    st.markdown("### 🧰 งานคงเหลือ (Remaining Work) — PP18 SI YAN")
    # ---- PIN gate (เข้าหน้านี้ต้องใส่ PIN; ใช้ remaining_pin ถ้าตั้งไว้ ไม่งั้นใช้ PIN เดียวกับแก้ไข) ----
    gate=str(sget("remaining_pin","") or sget("edit_pin","") or "2569")
    if not st.session_state.get("_rem_ok"):
        st.info("🔒 หน้านี้อยู่ระหว่างปรับปรุง — ใส่ PIN เพื่อเข้าดู")
        cp=st.columns([2,1,3])
        pv=cp[0].text_input("PIN", type="password", key="rem_pin",
                            label_visibility="collapsed", placeholder="ใส่ PIN เพื่อเข้า")
        if cp[1].button("เข้า", use_container_width=True, type="primary"):
            if str(pv)==gate: st.session_state["_rem_ok"]=True; st.rerun()
            else: st.error("PIN ไม่ถูกต้อง")
        st.caption("PIN เดียวกับที่ใช้แก้ไข Punchlist (หรือกำหนดแยกได้ที่ Secrets: remaining_pin)")
        st.stop()
    # ---- ป้ายกำลังปรับปรุง ----
    st.warning("🚧 หน้านี้ **กำลังปรับปรุง** — กำลังเพิ่ม: คลิกโยงแบบแปลน↔ไทม์ไลน์ และแก้ไขผ่าน Google Sheet · "
               "ตอนนี้แสดงข้อมูลชุดตัวอย่างจากไฟล์ Remaining work list R1 (อ่านอย่างเดียว)")
    html=load_remaining_html()
    if not html:
        st.error("ยังไม่พบไฟล์ remaining_work.html ในโปรเจกต์"); st.stop()
    import streamlit.components.v1 as components
    components.html(html, height=1500, scrolling=True)

# ---- บันทึกผลงานประจำวัน (Actual) — เชื่อม Google Sheet (แท็บ "บันทึกผลงาน") ----
ACTUAL_WS="บันทึกผลงาน"
ACT_HEADERS=["วันที่","หย่อม","ระบบ","โซน","ประเภทท่อ","ขนาด","ความยาว","ผู้บันทึก","บันทึกเมื่อ"]
ACT_ZONES=["UPF","CC","MPP","RF/GND"]
ACT_ZLAB={"UPF":"Upper Platform (UPF)","CC":"Concourse (CC)","MPP":"Multipurpose (MPP)","RF/GND":"Roof/Ground (RF/GND)"}
ACT_SYS=["ECS","FP","SN"]
ACT_SYSLAB={"ECS":"ECS (น้ำเย็น/คอนเดนเซอร์)","FP":"FP (ดับเพลิง)","SN":"SN (สุขาภิบาล)"}
ACT_SYSCOL={"ECS":"#2a78d6","FP":"#ef4444","SN":"#10b981"}

@st.cache_data(show_spinner=False)
def load_pipe_boq():
    import json as _j
    try: return _j.load(open(os.path.join(APP_DIR,"data","pipe_boq.json"),encoding="utf-8"))
    except Exception: return []

@st.cache_data(show_spinner=False)
def load_hyom_points():
    import json as _j
    try: return _j.load(open(os.path.join(APP_DIR,"data","hyom_points.json"),encoding="utf-8"))
    except Exception: return []

@st.cache_data(ttl=45, show_spinner=False)
def load_actual_log_df():
    try:
        ws=get_ws()
        if ws is None: return pd.DataFrame(columns=ACT_HEADERS)
        try: aws=ws.spreadsheet.worksheet(ACTUAL_WS)
        except Exception: return pd.DataFrame(columns=ACT_HEADERS)
        df=pd.DataFrame(aws.get_all_records())
        if df.empty: return pd.DataFrame(columns=ACT_HEADERS)
        for c in ACT_HEADERS:
            if c not in df.columns: df[c]=""
        df["ความยาว"]=pd.to_numeric(df["ความยาว"], errors="coerce").fillna(0.0)
        df["ขนาด"]=df["ขนาด"].astype(str)
        return df
    except Exception:
        return pd.DataFrame(columns=ACT_HEADERS)

def save_actual_rows(rows):
    ws=get_ws(); sh=ws.spreadsheet
    try: aws=sh.worksheet(ACTUAL_WS)
    except Exception:
        aws=sh.add_worksheet(title=ACTUAL_WS, rows=2000, cols=len(ACT_HEADERS))
        aws.update(values=[ACT_HEADERS], range_name="A1")
    aws.append_rows(rows, value_input_option="USER_ENTERED")

def _iso_thai(s):
    try:
        dd,mm,yy=str(s).split("/"); return "%s-%02d-%02d"%(yy,int(mm),int(dd))
    except Exception: return str(s)


# ===== Actual card: display fragments (styled like the mockup) =====
# ---- shared CSS (from the mockup) ----
ACT_CSS = r"""
.act-wrap *{box-sizing:border-box;font-family:-apple-system,'Segoe UI',Roboto,'Noto Sans Thai',sans-serif;}
.act-kpis{display:flex;gap:11px;margin:6px 0 6px;flex-wrap:wrap;}
.act-kpi{background:#fff;border:1px solid #e5e3de;border-radius:13px;padding:11px 17px;min-width:150px;box-shadow:0 1px 3px rgba(0,0,0,.04);}
.act-kpi .n{font-size:23px;font-weight:800;line-height:1;color:#1a1a1a;} .act-kpi .l{font-size:11.5px;color:#6b6862;margin-top:5px;}
.act-kpi.big .n{color:#1f9d57;} .act-kpi.tdy .n{color:#2a78d6;}
.act-chips{display:flex;gap:9px;flex-wrap:wrap;margin:2px 0 6px;}
.act-chip{background:#fff;border:1px solid #e5e3de;border-radius:10px;padding:8px 14px;font-size:12.5px;min-width:130px;}
.act-chip .r{display:flex;justify-content:space-between;align-items:center;} .act-chip b{font-size:15px;}
.act-dot{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:5px;}
.act-sub{color:#6b6862;font-size:12.5px;margin:2px 0 4px;}
.act-wrap table{width:100%;border-collapse:collapse;font-size:12.5px;}
.act-wrap th{text-align:left;padding:7px 9px;background:#f6f5f3;color:#6b6862;font-size:11px;font-weight:700;border-bottom:2px solid #e5e3de;}
.act-wrap td{padding:6px 9px;border-bottom:1px solid #f0efec;}
.act-wrap .rt{text-align:right;} .act-wrap .num{font-variant-numeric:tabular-nums;}
.act-wrap .tag{font-size:10.5px;font-weight:700;padding:1px 8px;border-radius:20px;color:#fff;}
.act-tt{font-size:12.5px;font-weight:700;margin:8px 0 6px;}
/* native Streamlit tabs -> pill segmented control (match mockup) */
.stTabs [data-baseweb="tab-list"]{background:#e9e8e4;border-radius:11px;padding:4px;gap:3px;display:inline-flex;border-bottom:none;}
.stTabs [data-baseweb="tab-list"] button[data-baseweb="tab"]{border-radius:8px;padding:6px 16px;margin:0;height:auto;color:#6b6862;font-weight:600;}
.stTabs [data-baseweb="tab-list"] button[data-baseweb="tab"]:hover{background:rgba(255,255,255,.5);color:#1a1a1a;}
.stTabs [data-baseweb="tab-list"] button[aria-selected="true"]{background:#fff;color:#1a1a1a;box-shadow:0 1px 3px rgba(0,0,0,.12);}
.stTabs [data-baseweb="tab-highlight"],.stTabs [data-baseweb="tab-border"]{display:none;background:transparent;height:0;}
.stTabs [data-baseweb="tab-list"] button[data-baseweb="tab"] p{font-weight:600;}
"""

# CSS used INSIDE the iframes (self-contained)
_IFRAME_CSS = r"""
*{box-sizing:border-box;margin:0;padding:0;font-family:-apple-system,'Segoe UI',Roboto,'Noto Sans Thai',sans-serif;}
body{background:#fff;color:#1a1a1a;padding:2px 2px 8px;}
.legend{display:flex;gap:15px;font-size:12px;margin:2px 0 8px;flex-wrap:wrap;}
.dot{width:10px;height:10px;border-radius:50%;display:inline-block;margin-right:5px;}
table{width:100%;border-collapse:collapse;font-size:12.5px;}
th{text-align:left;padding:7px 9px;background:#f6f5f3;color:#6b6862;font-size:11px;font-weight:700;border-bottom:2px solid #e5e3de;}
td{padding:6px 9px;border-bottom:1px solid #f0efec;}
.rt{text-align:right;} .num{font-variant-numeric:tabular-nums;}
.tag{font-size:10.5px;font-weight:700;padding:1px 8px;border-radius:20px;color:#fff;}
.drow{cursor:pointer;} .drow:hover td{background:#f4f8ff;} .drow.today td{background:#eef7f0;font-weight:700;}
.ddetail td{background:#fafbfc;font-size:11.5px;color:#555;}
.filters{display:flex;gap:7px;align-items:center;margin-bottom:11px;flex-wrap:wrap;}
.fbtn{padding:6px 13px;border-radius:8px;border:1px solid #d9d7d2;background:#f6f5f3;font-size:12px;font-weight:600;color:#555;cursor:pointer;}
.fbtn.on{background:#2a78d6;border-color:#2a78d6;color:#fff;}
.grp td{background:#fbfbfa;font-weight:800;}
.sec-h{font-size:12.5px;font-weight:700;margin:14px 0 6px;} .muted{color:#8a8a86;font-weight:600;}
input[type=date],select{font-size:12.5px;padding:6px 9px;border:1px solid #d9d7d2;border-radius:8px;font-family:inherit;background:#fff;color:#333;}
.empty{padding:22px;color:#9a978f;font-size:13px;text-align:center;}
"""

# ---- shared JS core: data-prep from log + helpers ----
_JS_CORE = r"""
const DATA=__DATA__;
const SYS=DATA.sysmeta, SYSK=['ECS','FP','SN'], ZLABEL=DATA.zlabel, ZONES=DATA.zones;
const thMon={1:"ม.ค.",2:"ก.พ.",3:"มี.ค.",4:"เม.ย.",5:"พ.ค.",6:"มิ.ย.",7:"ก.ค.",8:"ส.ค.",9:"ก.ย.",10:"ต.ค.",11:"พ.ย.",12:"ธ.ค."};
function fd(iso){const p=iso.split("-").map(Number);return p[2]+" "+thMon[p[1]];}
function isoOf(d){const p=(d||'').split('/');return p.length===3?p[2]+'-'+p[1].padStart(2,'0')+'-'+p[0].padStart(2,'0'):d;}
const LOG=DATA.log.map(e=>({...e,iso:isoOf(e.date),len:(+e.len||0)}));
// daily series (ascending, ISO) from log
function dailySeries(){let m={};LOG.forEach(e=>{if(!m[e.iso])m[e.iso]={date:e.iso,ECS:0,FP:0,SN:0};if(SYS[e.sys])m[e.iso][e.sys]+=e.len;});
 return Object.values(m).sort((a,b)=>a.date<b.date?-1:1);}
const DAILY=dailySeries();
function allDays(){return DAILY;}
// items (cumulative per sys/type/dia/zone) from log
function itemsFromLog(){let m={};LOG.forEach(e=>{const k=e.sys+'|'+e.type+'|'+e.dia+'|'+e.zone;if(!m[k])m[k]={sys:e.sys,type:e.type,dia:String(e.dia),zone:e.zone,act:0};m[k].act+=e.len;});return Object.values(m);}
const ITEMS=itemsFromLog();
function sysCum(s){return ITEMS.filter(i=>i.sys===s).reduce((a,i)=>a+i.act,0);}
function grand(){return ITEMS.reduce((a,i)=>a+i.act,0);}
function ptype(t){return t.replace(' Pipe','').replace('Chilled Water','Chilled').replace('Condenser Water','Condenser');}
"""

# ---- Progress page (chart + daily table + date filter + export) ----
_JS_PROGRESS = r"""
// ---- KPI + chips (context) ----
function renderKPI(){const g=grand();const today=DATA.today;
 const tl=LOG.filter(e=>e.iso===today).reduce((s,e)=>s+e.len,0);
 const nd=DAILY.length; const avg=nd?Math.round(g/nd):0;
 document.getElementById('kpis').innerHTML=
  kpi('big','ติดตั้งสะสม (จากวันนี้)',Math.round(g).toLocaleString()+' ม.')+
  kpi('tdy','ทำได้วันนี้',Math.round(tl).toLocaleString()+' ม.')+
  kpi('','จำนวนวันบันทึก',nd+' วัน')+
  kpi('','เฉลี่ย/วัน',avg.toLocaleString()+' ม.');
 let sc='';for(const s of SYSK){sc+='<div class="schip"><div class="r"><span><span class="dot" style="background:'+SYS[s].c+'"></span>'+s+'</span><b style="color:'+SYS[s].c+'">'+Math.round(sysCum(s)).toLocaleString()+'</b></div><div style="font-size:10.5px;color:#8a8a86;margin-top:2px">เมตรสะสม</div></div>';}
 document.getElementById('syschips').innerHTML=sc;}
function kpi(c,l,n){return '<div class="kpi '+c+'"><div class="n">'+n+'</div><div class="l">'+l+'</div></div>';}
// ---- date range ----
function inRange(iso){const f=document.getElementById('d-from').value,t=document.getElementById('d-to').value;return (!f||iso>=f)&&(!t||iso<=t);}
function quickRange(m){const dates=DAILY.map(d=>d.date);if(!dates.length)return;const today=DATA.today;let from;
 if(m==='all')from=dates[0];
 else if(m==='month')from=today.slice(0,7)+'-01';
 else{const d=new Date(today);d.setDate(d.getDate()-(m-1));from=d.toISOString().slice(0,10);}
 document.getElementById('d-from').value=from;document.getElementById('d-to').value=today;renderChart();renderDTable();}
function rangeLabel(){return (document.getElementById('d-from').value||'all')+'_'+(document.getElementById('d-to').value||'all');}
function csvDL(name,lines){const csv='﻿'+lines.join('\r\n');const blob=new Blob([csv],{type:'text/csv;charset=utf-8'});
 const url=URL.createObjectURL(blob);const a=document.createElement('a');a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);}
function exportSummary(){let cc=0,cum=[];DAILY.forEach(d=>{cc+=d.ECS+d.FP+d.SN;cum.push(cc);});
 let L=['วันที่,ECS (ม.),FP (ม.),SN (ม.),รวมวันนั้น (ม.),สะสม (ม.)'],n=0;
 DAILY.forEach((d,i)=>{if(inRange(d.date)){L.push([d.date,d.ECS,d.FP,d.SN,d.ECS+d.FP+d.SN,cum[i]].join(','));n++;}});
 if(!n){return;} csvDL('actual_daily_summary_'+rangeLabel()+'.csv',L);}
function exportDetail(){let L=['วันที่,หย่อม,ระบบ,โซน,ขนาดท่อ,ความยาว (ม.),ผู้บันทึก'],n=0;
 LOG.slice().sort((a,b)=>a.iso<b.iso?-1:1).forEach(e=>{if(inRange(e.iso)){L.push([e.date,(e.hyom||'—'),e.sys,e.zone,'"'+ptype(e.type)+' · Ø'+e.dia+'mm"',e.len,(e.by||'')].join(','));n++;}});
 if(!n){return;} csvDL('actual_detail_'+rangeLabel()+'.csv',L);}
// ---- chart ----
function renderChart(){
 const cumAll=[];let cc=0;DAILY.forEach(d=>{cc+=d.ECS+d.FP+d.SN;cumAll.push(cc);});
 const idx=DAILY.map((d,i)=>i).filter(i=>inRange(DAILY[i].date));
 const D=idx.map(i=>DAILY[i]);const cum=idx.map(i=>cumAll[i]);const n=D.length;
 const rsum=D.reduce((s,d)=>s+d.ECS+d.FP+d.SN,0);
 document.getElementById('range-sum').textContent=n?('ช่วงที่เลือก: '+n+' วัน · ทำได้ '+Math.round(rsum).toLocaleString()+' ม.'):'';
 if(!n){document.getElementById('chart').innerHTML='<div class="empty">ยังไม่มีข้อมูลในช่วงวันที่ที่เลือก</div>';return;}
 const bw=Math.max(14,Math.min(30,Math.floor(760/n)));const gap=6;const W=Math.max(720,n*(bw+gap)+70);const H=300;
 const mL=44,mR=52,mT=14,mB=42;const pw=W-mL-mR,ph=H-mT-mB;
 const dtot=D.map(d=>d.ECS+d.FP+d.SN);const maxD=Math.max(10,...dtot)*1.15;const maxC=Math.max(10,...cum)*1.05;
 const x=i=>mL+i*(pw/n)+(pw/n-bw)/2;const yD=v=>mT+ph-(v/maxD*ph);const yC=v=>mT+ph-(v/maxC*ph);
 let s='<svg viewBox="0 0 '+W+' '+H+'" width="'+W+'" height="'+H+'" font-family="inherit">';
 for(let g=0;g<=4;g++){const v=maxD*g/4;const yy=yD(v);s+='<line x1="'+mL+'" y1="'+yy+'" x2="'+(W-mR)+'" y2="'+yy+'" stroke="#f0f0f0"/><text x="'+(mL-6)+'" y="'+(yy+3)+'" font-size="9" fill="#999" text-anchor="end">'+Math.round(v)+'</text>';}
 for(let g=0;g<=4;g++){const v=maxC*g/4;const yy=yC(v);s+='<text x="'+(W-mR+6)+'" y="'+(yy+3)+'" font-size="9" fill="#1f9d57" text-anchor="start">'+Math.round(v).toLocaleString()+'</text>';}
 D.forEach((d,i)=>{let yb=mT+ph;const xx=x(i);
   for(const sy of SYSK){const h=d[sy]/maxD*ph;if(h>0){yb-=h;s+='<rect x="'+xx+'" y="'+yb+'" width="'+bw+'" height="'+h+'" fill="'+SYS[sy].c+'" opacity="0.9"><title>'+fd(d.date)+' · '+sy+': '+d[sy]+' ม.</title></rect>';}}
   if(i%Math.ceil(n/12)===0||i===n-1){s+='<text x="'+(xx+bw/2)+'" y="'+(H-mB+14)+'" font-size="8.5" fill="#888" text-anchor="middle" transform="rotate(35 '+(xx+bw/2)+' '+(H-mB+14)+')">'+fd(d.date)+'</text>';}});
 let pts=D.map((d,i)=>({x:x(i)+bw/2,y:yC(cum[i])}));
 let path='M'+pts.map(p=>p.x.toFixed(1)+' '+p.y.toFixed(1)).join(' L');
 s+='<path d="'+path+'" fill="none" stroke="#1f9d57" stroke-width="2.4"/>';
 pts.forEach(p=>{s+='<circle cx="'+p.x+'" cy="'+p.y+'" r="2.6" fill="#1f9d57"/>';});
 const lp=pts[pts.length-1];s+='<text x="'+(lp.x-4)+'" y="'+(lp.y-8)+'" font-size="10" font-weight="700" fill="#1f9d57" text-anchor="end">สะสม '+Math.round(cum[cum.length-1]).toLocaleString()+' ม.</text>';
 s+='</svg>';
 document.getElementById('chart').innerHTML=s;
 document.getElementById('dlegend').innerHTML=SYSK.map(sy=>'<span><span class="dot" style="background:'+SYS[sy].c+'"></span>'+sy+' (ม./วัน)</span>').join('')+'<span><span style="display:inline-block;width:16px;height:3px;background:#1f9d57;vertical-align:3px;margin-right:5px"></span>สะสม (ม.)</span>';
}
// ---- daily table ----
let openDay=null;
function renderDTable(){const asc=DAILY;let cum=[],cc=0;asc.forEach(d=>{cc+=d.ECS+d.FP+d.SN;cum.push(cc);});
 let rows='';let any=false;
 for(let i=asc.length-1;i>=0;i--){const d=asc[i];if(!inRange(d.date))continue;any=true;const tot=d.ECS+d.FP+d.SN;const isT=d.date===DATA.today;
  rows+='<tr class="drow'+(isT?' today':'')+'" onclick="toggleDay(\''+d.date+'\')"><td>'+fd(d.date)+(isT?' (วันนี้)':'')+'</td><td class="rt num">'+d.ECS+'</td><td class="rt num">'+d.FP+'</td><td class="rt num">'+d.SN+'</td><td class="rt num"><b>'+tot+'</b></td><td class="rt num">'+Math.round(cum[i]).toLocaleString()+'</td></tr>';
  if(openDay===d.date){const es=LOG.filter(e=>e.iso===d.date);
   let det=es.length?es.map(e=>'• '+ptype(e.type)+' · Ø'+e.dia+'mm ('+e.zone+(e.hyom&&e.hyom!=='—'?' · '+e.hyom:'')+') '+e.len+' ม.'+(e.by?' — '+e.by:'')).join('<br>'):'ไม่มีรายละเอียด';
   rows+='<tr class="ddetail"><td colspan="6">'+det+'</td></tr>';}}
 document.getElementById('dbody').innerHTML=any?rows:'<tr><td colspan="6" class="empty">ยังไม่มีข้อมูลในช่วงที่เลือก</td></tr>';}
function toggleDay(d){openDay=openDay===d?null:d;renderDTable();}
function initRange(){const ds=DAILY.map(d=>d.date);document.getElementById('d-from').value=ds.length?ds[0]:DATA.today;document.getElementById('d-to').value=DATA.today;}
initRange();renderChart();renderDTable();
"""

# ---- Summary page (by size/zone/หย่อม + drilldown) ----
_JS_SUMMARY = r"""
function renderKPI(){const g=grand();const today=DATA.today;
 const tl=LOG.filter(e=>e.iso===today).reduce((s,e)=>s+e.len,0);
 const nd=DAILY.length;const avg=nd?Math.round(g/nd):0;
 document.getElementById('kpis').innerHTML=
  kpi('big','ติดตั้งสะสม (จากวันนี้)',Math.round(g).toLocaleString()+' ม.')+
  kpi('tdy','ทำได้วันนี้',Math.round(tl).toLocaleString()+' ม.')+
  kpi('','จำนวนวันบันทึก',nd+' วัน')+
  kpi('','เฉลี่ย/วัน',avg.toLocaleString()+' ม.');
 let sc='';for(const s of SYSK){sc+='<div class="schip"><div class="r"><span><span class="dot" style="background:'+SYS[s].c+'"></span>'+s+'</span><b style="color:'+SYS[s].c+'">'+Math.round(sysCum(s)).toLocaleString()+'</b></div><div style="font-size:10.5px;color:#8a8a86;margin-top:2px">เมตรสะสม</div></div>';}
 document.getElementById('syschips').innerHTML=sc;}
function kpi(c,l,n){return '<div class="kpi '+c+'"><div class="n">'+n+'</div><div class="l">'+l+'</div></div>';}
let GB='size',HYSEL='';
function setGB(g){GB=g;HYSEL='';document.getElementById('hy-detail').value='';['size','zone','hyom'].forEach(x=>document.getElementById('gb-'+x).classList.toggle('on',g===x));renderSum();}
function populateHyDetail(){let hs=DATA.hyom.slice().sort((a,b)=>a.no-b.no);
 document.getElementById('hy-detail').innerHTML='<option value="">— ทั้งหมด (ดูสรุป)</option>'+hs.map(p=>'<option value="'+p.no+'">จุดที่ '+p.no+' · '+p.sysN+' · '+p.zone+' · '+p.loc+'</option>').join('');}
function onHyDetail(){HYSEL=document.getElementById('hy-detail').value;renderSum();}
function pointLog(no){return LOG.filter(e=>{const mm=(e.hyom||'').match(/\d+/);return mm&&+mm[0]===+no;});}
function renderHyDetail(){const p=DATA.hyom.find(x=>x.no===+HYSEL);const box=document.getElementById('sumtbl');
 if(!p){box.innerHTML='';return;}
 let rows={};pointLog(p.no).forEach(e=>{const k=e.sys+'|'+e.type+'|'+e.dia+'|'+e.zone;if(!rows[k])rows[k]={sys:e.sys,type:e.type,dia:e.dia,zone:e.zone,cum:0};rows[k].cum+=e.len;});
 let arr=Object.values(rows).sort((a,b)=>a.sys.localeCompare(b.sys)||a.type.localeCompare(b.type)||parseInt(a.dia)-parseInt(b.dia));
 let tot=arr.reduce((s,r)=>s+r.cum,0);let syss=[...new Set(arr.map(r=>r.sys))];
 let head='<div style="background:#f6f8fb;border:1px solid #e3ecf7;border-radius:10px;padding:11px 14px;margin-bottom:12px"><div style="font-size:14.5px;font-weight:800">📍 จุดที่ '+p.no+' · '+p.loc+'</div><div style="font-size:12.5px;color:#2a4a6b;margin-top:4px">โซน <b>'+(ZLABEL[p.zone]||p.zone)+'</b> · ระบบที่พบ <b>'+(syss.join(', ')||'—')+'</b> · ติดตั้งสะสมรวม <b>'+Math.round(tot).toLocaleString()+' ม.</b></div></div>';
 let body=arr.map(r=>'<tr><td><span class="tag" style="background:'+(SYS[r.sys]?SYS[r.sys].c:'#888')+'">'+r.sys+'</span></td><td>'+ptype(r.type)+' · Ø'+r.dia+'mm</td><td>'+r.zone+'</td><td class="rt num">'+Math.round(r.cum).toLocaleString()+'</td></tr>').join('');
 if(!arr.length)body='<tr><td colspan="4" class="empty">ยังไม่มีข้อมูลติดตั้งที่จุดนี้ (นับตั้งแต่วันนี้)</td></tr>';
 box.innerHTML=head+'<table><thead><tr><th>ระบบ</th><th>ขนาดท่อ</th><th>โซน</th><th class="rt">ติดตั้งสะสม (ม.)</th></tr></thead><tbody>'+body+'<tr class="grp"><td colspan="3">รวมจุดนี้</td><td class="rt num"><b>'+Math.round(tot).toLocaleString()+' ม.</b></td></tr></tbody></table>';}
function renderSum(){if(HYSEL){renderHyDetail();return;}
 if(!ITEMS.length){document.getElementById('sumtbl').innerHTML='<div class="empty">ยังไม่มีข้อมูลบันทึก — เริ่มกรอกในแท็บ “กรอกผลงาน” แล้วยอดสะสมจะแสดงที่นี่</div>';return;}
 let html='<table><thead><tr><th>รายการ</th>';
 if(GB==='size'){html+='<th class="rt">UPF</th><th class="rt">CC</th><th class="rt">MPP</th><th class="rt">RF/GND</th><th class="rt">รวม (ม.)</th></tr></thead><tbody>';
  for(const s of SYSK){let sub=ITEMS.filter(i=>i.sys===s);if(!sub.length)continue;
   let m={};sub.forEach(i=>{const k=i.type+'|'+i.dia;if(!m[k])m[k]={type:i.type,dia:i.dia,UPF:0,CC:0,MPP:0,'RF/GND':0};m[k][i.zone]+=i.act;});
   let arr=Object.values(m).sort((a,b)=>a.type.localeCompare(b.type)||parseInt(a.dia)-parseInt(b.dia));
   let st={UPF:0,CC:0,MPP:0,'RF/GND':0};arr.forEach(x=>{['UPF','CC','MPP','RF/GND'].forEach(z=>st[z]+=x[z]);});
   const stot=st.UPF+st.CC+st.MPP+st['RF/GND'];
   html+='<tr class="grp"><td>'+SYS[s].n+'</td><td class="rt num">'+Math.round(st.UPF)+'</td><td class="rt num">'+Math.round(st.CC)+'</td><td class="rt num">'+Math.round(st.MPP)+'</td><td class="rt num">'+Math.round(st['RF/GND'])+'</td><td class="rt num"><b>'+Math.round(stot).toLocaleString()+'</b></td></tr>';
   for(const x of arr){const tot=x.UPF+x.CC+x.MPP+x['RF/GND'];
    html+='<tr><td style="padding-left:20px">'+ptype(x.type)+' · Ø'+x.dia+'mm</td><td class="rt num">'+Math.round(x.UPF)+'</td><td class="rt num">'+Math.round(x.CC)+'</td><td class="rt num">'+Math.round(x.MPP)+'</td><td class="rt num">'+Math.round(x['RF/GND'])+'</td><td class="rt num">'+Math.round(tot)+'</td></tr>';}}
 }else if(GB==='zone'){html+='<th class="rt">ECS</th><th class="rt">FP</th><th class="rt">SN</th><th class="rt">รวม (ม.)</th></tr></thead><tbody>';
  for(const z of ZONES){let e=ITEMS.filter(i=>i.zone===z&&i.sys==='ECS').reduce((a,i)=>a+i.act,0),f=ITEMS.filter(i=>i.zone===z&&i.sys==='FP').reduce((a,i)=>a+i.act,0),n=ITEMS.filter(i=>i.zone===z&&i.sys==='SN').reduce((a,i)=>a+i.act,0);
   if(e+f+n<=0)continue;
   html+='<tr><td><b>'+z+'</b></td><td class="rt num">'+Math.round(e)+'</td><td class="rt num">'+Math.round(f)+'</td><td class="rt num">'+Math.round(n)+'</td><td class="rt num"><b>'+Math.round(e+f+n).toLocaleString()+'</b></td></tr>';}
 }else{html+='<th class="rt">โซน</th><th class="rt">ECS</th><th class="rt">FP</th><th class="rt">SN</th><th class="rt">รวม (ม.)</th></tr></thead><tbody>';
  const cell=v=>v?Math.round(v).toLocaleString():'<span style="color:#cfcdc8">–</span>';
  let plog={},un={ECS:0,FP:0,SN:0};
  LOG.forEach(e=>{const mm=(e.hyom||'').match(/\d+/);if(mm){const no=+mm[0];(plog[no]=plog[no]||{ECS:0,FP:0,SN:0})[e.sys]+=e.len;}else{un[e.sys]+=e.len;}});
  let pts=DATA.hyom.slice().sort((a,b)=>a.no-b.no);let g={ECS:0,FP:0,SN:0},body='';
  for(const p of pts){const pl=plog[p.no];if(!pl)continue;let v={ECS:pl.ECS,FP:pl.FP,SN:pl.SN};
   const tot=v.ECS+v.FP+v.SN;g.ECS+=v.ECS;g.FP+=v.FP;g.SN+=v.SN;
   const multi=[v.ECS,v.FP,v.SN].filter(x=>x>0).length>1;
   body+='<tr'+(multi?' style="background:#fffdf0"':'')+'><td>จุดที่ '+p.no+' · '+p.loc+(multi?' 🔀':'')+'</td><td class="rt">'+p.zone+'</td><td class="rt num">'+cell(v.ECS)+'</td><td class="rt num">'+cell(v.FP)+'</td><td class="rt num">'+cell(v.SN)+'</td><td class="rt num"><b>'+Math.round(tot).toLocaleString()+'</b></td></tr>';}
  if(un.ECS+un.FP+un.SN>0)body+='<tr><td>ไม่ระบุจุด</td><td></td><td class="rt num">'+cell(un.ECS)+'</td><td class="rt num">'+cell(un.FP)+'</td><td class="rt num">'+cell(un.SN)+'</td><td class="rt num"><b>'+Math.round(un.ECS+un.FP+un.SN).toLocaleString()+'</b></td></tr>';
  const gt=g.ECS+g.FP+g.SN+un.ECS+un.FP+un.SN;
  if(!body)body='<tr><td colspan="6" class="empty">ยังไม่มีบันทึกที่ติดหย่อม — เริ่มกรอกโดยเลือกหย่อม แล้วยอดจะแยกรายจุดที่นี่</td></tr>';
  else html+='<tr class="grp"><td>รวมทุกจุด</td><td></td><td class="rt num">'+Math.round(g.ECS+un.ECS).toLocaleString()+'</td><td class="rt num">'+Math.round(g.FP+un.FP).toLocaleString()+'</td><td class="rt num">'+Math.round(g.SN+un.SN).toLocaleString()+'</td><td class="rt num"><b>'+Math.round(gt).toLocaleString()+'</b></td></tr>';
  html+=body;
 }
 html+='</tbody></table>';document.getElementById('sumtbl').innerHTML=html;}
populateHyDetail();renderSum();
"""

# ---- KPI/chips bar shared markup used inside iframes ----
_KPI_BAR = ('<div class="khead" style="margin-bottom:6px">'
            '<div class="kpis" id="kpis" style="display:flex;gap:11px;flex-wrap:wrap;margin:2px 0 6px"></div>'
            '<div class="syschips" id="syschips" style="display:flex;gap:9px;flex-wrap:wrap"></div></div>'
            '<style>.kpi{background:#fff;border:1px solid #e5e3de;border-radius:13px;padding:10px 16px;min-width:150px;box-shadow:0 1px 3px rgba(0,0,0,.04)}'
            '.kpi .n{font-size:22px;font-weight:800;line-height:1;color:#1a1a1a}.kpi .l{font-size:11.5px;color:#6b6862;margin-top:5px}'
            '.kpi.big .n{color:#1f9d57}.kpi.tdy .n{color:#2a78d6}'
            '.schip{background:#fff;border:1px solid #e5e3de;border-radius:10px;padding:7px 13px;font-size:12.5px;min-width:130px}'
            '.schip .r{display:flex;justify-content:space-between;align-items:center}.schip b{font-size:15px}</style>')


def _page(body, js, data):
    dj = json.dumps(data, ensure_ascii=False)
    return ('<!doctype html><html><head><meta charset="utf-8">'
            '<meta name="viewport" content="width=device-width,initial-scale=1">'
            '<style>' + _IFRAME_CSS + '</style></head><body><div class="act-wrap">'
            + body + '</div><script>' + js.replace('__DATA__', dj) + '</script></body></html>')


def progress_page(data):
    body = (
        '<div class="filters">'
        '<span style="font-size:11.5px;color:#6b6862;font-weight:600">ช่วงวันที่:</span>'
        '<input id="d-from" type="date" onchange="renderChart();renderDTable()">'
        '<span style="color:#8a8a86">–</span>'
        '<input id="d-to" type="date" onchange="renderChart();renderDTable()">'
        '<button class="fbtn" onclick="quickRange(7)">7 วันล่าสุด</button>'
        '<button class="fbtn" onclick="quickRange(\'month\')">เดือนนี้</button>'
        '<button class="fbtn" onclick="quickRange(\'all\')">ทั้งหมด</button>'
        '<span style="width:1px;height:22px;background:#e0ded9;margin:0 4px"></span>'
        '<button class="fbtn" onclick="exportSummary()" style="background:#1f9d57;color:#fff;border-color:#1f9d57">⬇ Export สรุปรายวัน</button>'
        '<button class="fbtn" onclick="exportDetail()" style="background:#2a78d6;color:#fff;border-color:#2a78d6">⬇ Export รายละเอียด</button>'
        '<span id="range-sum" style="font-size:11.5px;color:#2a78d6;font-weight:700;margin-left:auto"></span>'
        '</div>'
        '<div class="legend" id="dlegend"></div>'
        '<div id="chart" style="width:100%;overflow-x:auto"></div>'
        '<div class="sec-h">ตารางรายวัน <span class="muted">(แตะแถวเพื่อดูรายละเอียดวันนั้น)</span></div>'
        '<table><thead><tr><th>วันที่</th><th class="rt">ECS</th><th class="rt">FP</th><th class="rt">SN</th><th class="rt">รวมวันนั้น (ม.)</th><th class="rt">สะสม (ม.)</th></tr></thead><tbody id="dbody"></tbody></table>')
    return _page(body, _JS_CORE + _JS_PROGRESS, data)


def summary_page(data):
    body = (
        '<div class="filters"><span style="font-size:11.5px;color:#6b6862;font-weight:600">แยกตาม:</span>'
        '<button class="fbtn on" id="gb-size" onclick="setGB(\'size\')">ขนาดท่อ</button>'
        '<button class="fbtn" id="gb-zone" onclick="setGB(\'zone\')">โซน</button>'
        '<button class="fbtn" id="gb-hyom" onclick="setGB(\'hyom\')">หย่อม (จุด)</button>'
        '<span style="margin-left:12px;font-size:11.5px;color:#6b6862;font-weight:600">🔍 ดูละเอียดรายหย่อม:</span>'
        '<select id="hy-detail" onchange="onHyDetail()" style="min-width:280px"></select></div>'
        '<div id="sumtbl"></div>')
    return _page(body, _JS_CORE + _JS_SUMMARY, data)


def kpi_chips_html(log, hyom, today_iso):
    """Python-computed KPI + chips for the page header (st.markdown)."""
    SYSC = {"ECS": "#2a78d6", "FP": "#ef4444", "SN": "#10b981"}
    total = sum(float(e.get("len", 0) or 0) for e in log)
    today = sum(float(e.get("len", 0) or 0) for e in log if e.get("iso") == today_iso)
    days = len({e.get("iso") for e in log if e.get("iso")})
    avg = round(total / days) if days else 0
    syscum = {s: sum(float(e.get("len", 0) or 0) for e in log if e.get("sys") == s) for s in ["ECS", "FP", "SN"]}

    def kpi(cls, label, val):
        return f'<div class="act-kpi {cls}"><div class="n">{val}</div><div class="l">{label}</div></div>'
    kpis = (kpi("big", "ติดตั้งสะสม (จากวันนี้)", f"{round(total):,} ม.") +
            kpi("tdy", "ทำได้วันนี้", f"{round(today):,} ม.") +
            kpi("", "จำนวนวันบันทึก", f"{days} วัน") +
            kpi("", "เฉลี่ย/วัน", f"{avg:,} ม."))
    chips = ""
    for s in ["ECS", "FP", "SN"]:
        chips += (f'<div class="act-chip"><div class="r"><span><span class="act-dot" style="background:{SYSC[s]}"></span>{s}</span>'
                  f'<b style="color:{SYSC[s]}">{round(syscum[s]):,}</b></div>'
                  f'<div style="font-size:10.5px;color:#8a8a86;margin-top:2px">เมตรสะสม</div></div>')
    return (f'<div class="act-wrap"><div class="act-kpis">{kpis}</div>'
            f'<div class="act-chips">{chips}</div></div>')


def today_rows_html(log, today_iso):
    """Mockup-style 'today's entries' table (read-only) for the entry tab."""
    SYSC = {"ECS": "#2a78d6", "FP": "#ef4444", "SN": "#10b981"}
    rows = [e for e in log if e.get("iso") == today_iso]
    heads = ["วันที่", "ระบบ", "โซน", "หย่อม", "รายการท่อ", "ความยาว (ม.)", "ผู้บันทึก"]
    th = "".join(f'<th{" class=rt" if h=="ความยาว (ม.)" else ""}>{h}</th>' for h in heads)
    if not rows:
        body = '<tr><td colspan="7" style="color:#9a978f;padding:12px;text-align:center">ยังไม่มีบันทึกวันนี้</td></tr>'
    else:
        body = ""
        for e in rows:
            t = str(e.get("type", "")).replace(" Pipe", "")
            badge = f'<span class="tag" style="background:{SYSC.get(e.get("sys"),"#888")}">{e.get("sys","")}</span>'
            hy = e.get("hyom", "—") or "—"
            body += (f'<tr><td>{e.get("date","")}</td><td>{badge}</td><td>{e.get("zone","")}</td>'
                     f'<td>{hy}</td><td>{t} · Ø{e.get("dia","")}mm</td>'
                     f'<td class="rt num">{round(float(e.get("len",0) or 0))}</td>'
                     f'<td>{e.get("by","") or "—"}</td></tr>')
    tot = sum(float(e.get("len", 0) or 0) for e in rows)
    return (f'<div class="act-wrap"><div class="act-tt">รายการที่บันทึกวันนี้ '
            f'<span style="color:#8a8a86;font-weight:600">({len(rows)} รายการ · {round(tot)} ม.)</span></div>'
            f'<table><thead><tr>{th}</tr></thead><tbody>{body}</tbody></table></div>')


def render_actual():
    st.markdown(f"<style>{ACT_CSS}</style>", unsafe_allow_html=True)
    hcol=st.columns([5,1])
    hcol[0].markdown("### 📝 บันทึกผลงานประจำวัน (Actual) — งานท่อ PP18")
    if hcol[1].button("← เมนูหลัก", key="act_back", use_container_width=True):
        st.session_state["view"]="menu"; st.session_state.pop("_act_ok",None); st.rerun()
    st.caption("บันทึกความยาวท่อติดตั้งจริงรายวัน → เก็บลง Google Sheet · นับตั้งแต่วันนี้ · ระบบ ECS / FP / SN")
    gate=str(sget("actual_pin","") or sget("edit_pin","") or "2569")
    if not st.session_state.get("_act_ok"):
        st.info("🔒 ใส่ PIN เพื่อเข้าหน้านี้")
        cp=st.columns([2,1,3])
        pv=cp[0].text_input("PIN", type="password", key="act_pin", label_visibility="collapsed", placeholder="ใส่ PIN เพื่อเข้า")
        if cp[1].button("เข้า", use_container_width=True, type="primary", key="act_enter"):
            if str(pv)==gate: st.session_state["_act_ok"]=True; st.rerun()
            else: st.error("PIN ไม่ถูกต้อง")
        st.caption("PIN เดียวกับที่ใช้แก้ไข Punchlist (หรือกำหนดแยกได้ที่ Secrets: actual_pin)")
        st.stop()

    connected = get_ws() is not None
    boq=load_pipe_boq(); hyom=load_hyom_points(); logdf=load_actual_log_df()
    _as=st.session_state.pop("_act_saved",None)
    if _as is not None:
        st.toast(f"บันทึกผลงานแล้ว {_as} รายการ ✓", icon="✅")
    if not connected:
        st.warning("ยังไม่ได้เชื่อม Google Sheet — กรอกดูได้แต่ยังบันทึกไม่ได้ (ตั้ง service account ก่อน)")

    today_iso=bkk_today().isoformat()
    _lg=[]
    if not logdf.empty:
        for _,r in logdf.iterrows():
            _lg.append({"date":str(r["วันที่"]),"sys":str(r["ระบบ"]),"type":str(r["ประเภทท่อ"]),
                        "dia":str(r["ขนาด"]),"zone":str(r["โซน"]),
                        "hyom":(str(r.get("หย่อม","—")) or "—"),
                        "len":float(r["ความยาว"] or 0),"by":(str(r.get("ผู้บันทึก","")) or ""),
                        "iso":_iso_thai(r["วันที่"])})
    _hy=[]
    for p in hyom:
        try: _no=int(str(p.get("no","")).strip() or 0)
        except Exception: _no=0
        _hy.append({"no":_no,"sysN":p.get("sysN",p.get("sys","")),"zone":p.get("zone",""),"loc":p.get("loc","")})
    _DATA={"log":_lg,"hyom":_hy,"today":today_iso,"zones":ACT_ZONES,"zlabel":ACT_ZLAB,
           "sysmeta":{s:{"c":ACT_SYSCOL[s],"n":ACT_SYSLAB[s]} for s in ACT_SYS}}
    # iframe heights sized to content so the Progress/Summary tabs don't inner-scroll
    _ndays=len({e["iso"] for e in _lg})
    _ptset=set()
    for _e in _lg:
        _dd="".join(ch for ch in str(_e.get("hyom","")) if ch.isdigit())
        if _dd: _ptset.add(_dd)
    _sizerows=len({(e["sys"],e["type"],e["dia"]) for e in _lg})+len({e["sys"] for e in _lg})
    _sum_h=170+max(_sizerows, len(_ptset)+2, 4)*34
    _prog_h=560+max(_ndays,1)*34

    st.markdown(kpi_chips_html(_lg,_hy,today_iso), unsafe_allow_html=True)

    tabE,tabP,tabS=st.tabs(["📝 กรอกผลงาน","📅 Progress รายวัน","📊 สรุปสะสม"])

    with tabE:
        c=st.columns([1.2,3,1.7,1.7,1.3])
        d=c[0].date_input("วันที่", value=bkk_today(), key="act_d")
        hopts=["— ไม่ระบุจุด"]+[f"จุดที่ {p['no']} · {p['sysN']} · {p['zone']} · {p['loc']}" for p in hyom]
        hpick=c[1].selectbox("หย่อมงาน (จุด)", hopts, key="act_h")
        zpick=c[2].selectbox("โซน", [ACT_ZLAB[z] for z in ACT_ZONES], key="act_z")
        spick=c[3].selectbox("ระบบ", [ACT_SYSLAB[s] for s in ACT_SYS], key="act_s")
        by=c[4].text_input("ผู้บันทึก", key="act_b", placeholder="ชื่อ")
        sys=ACT_SYS[[ACT_SYSLAB[s] for s in ACT_SYS].index(spick)]
        zone=ACT_ZONES[[ACT_ZLAB[z] for z in ACT_ZONES].index(zpick)]
        hy_no=""
        if hpick!="— ไม่ระบุจุด":
            try: hy_no="".join(ch for ch in hpick.split("·")[0] if ch.isdigit())
            except Exception: hy_no=""
        cat=sorted({(r["type"],str(r["dia"])) for r in boq if r["sys"]==sys},
                   key=lambda t:(t[0], int("".join(ch for ch in t[1] if ch.isdigit()) or 0)))
        def _cum(t,dd):
            l=0.0
            if not logdf.empty:
                m=(logdf["ระบบ"]==sys)&(logdf["ประเภทท่อ"]==t)&(logdf["ขนาด"].astype(str)==dd)&(logdf["โซน"]==zone)
                l=float(logdf.loc[m,"ความยาว"].sum())
            return l
        edf=pd.DataFrame([{"รายการท่อ":f"{t.replace(' Pipe','')} · Ø{dd}mm","สะสมที่บันทึก (ม.)":round(_cum(t,dd)),
                           "ติดตั้งวันนี้ (ม.)":0,"_type":t,"_dia":dd} for t,dd in cat])
        st.caption(f"กรอกความยาวที่ติดตั้งวันนี้ (เฉพาะที่ทำ) — {ACT_SYSLAB[sys]} · {ACT_ZLAB[zone]} · {('จุด '+hy_no) if hy_no else 'ไม่ระบุจุด'}")
        ed=st.data_editor(edf, key=f"act_ed_{sys}_{zone}", hide_index=True, use_container_width=True, height=int((len(edf)+1)*36),
            column_config={"_type":None,"_dia":None,
                "รายการท่อ":st.column_config.TextColumn("รายการท่อ", disabled=True),
                "สะสมที่บันทึก (ม.)":st.column_config.NumberColumn("สะสมที่บันทึก (ม.)", disabled=True, format="%d"),
                "ติดตั้งวันนี้ (ม.)":st.column_config.NumberColumn("ติดตั้งวันนี้ (ม.)", min_value=0, step=1, format="%d")})
        if st.button("💾 บันทึกผลงานวันนี้", type="primary", key="act_save"):
            if not connected: st.error("ยังไม่ได้เชื่อม Google Sheet")
            else:
                rows=[]; stamp=_bkk_stamp(); dstr=d.strftime("%d/%m/%Y")
                for _,r in ed.iterrows():
                    try: ln=float(r["ติดตั้งวันนี้ (ม.)"])
                    except Exception: ln=0
                    if ln and ln>0:
                        rows.append([dstr, ("จุด "+hy_no) if hy_no else "—", sys, zone, r["_type"], str(r["_dia"]), ln, str(by or ""), stamp])
                if not rows: st.warning("ยังไม่ได้กรอกความยาวสักช่อง")
                else:
                    try:
                        save_actual_rows(rows); st.cache_data.clear()
                        st.session_state["_act_saved"]=len(rows); st.rerun()
                    except Exception as e:
                        st.error(f"บันทึกไม่สำเร็จ: {e}")
        st.markdown(today_rows_html(_lg,today_iso), unsafe_allow_html=True)

    with tabP:
        if not _lg:
            st.info("ยังไม่มีข้อมูลบันทึก — เริ่มกรอกในแท็บ “กรอกผลงาน” แล้วกราฟ/ตารางจะแสดงที่นี่")
        else:
            components.html(progress_page(_DATA), height=_prog_h, scrolling=True)

    with tabS:
        if not _lg:
            st.info("ยังไม่มีข้อมูลบันทึก — เริ่มกรอกในแท็บ “กรอกผลงาน” แล้วสรุปสะสมจะแสดงที่นี่")
        else:
            components.html(summary_page(_DATA), height=_sum_h, scrolling=True)

# ===================== การ์ด Punchlist FP · Multipurpose (ชุดแยก) =====================
FP_WS = "PunchlistFP"

@st.cache_data(ttl=60, show_spinner=False)
def load_raw_fp(wsname):
    ws=get_ws(wsname)
    if ws is None: return pd.DataFrame(columns=ALL_COLS)
    df=pd.DataFrame(ws.get_all_records())
    for c in ALL_COLS:
        if c not in df.columns: df[c]=""
    df=df[[c for c in ALL_COLS if c in df.columns]].fillna("")
    if C_NO in df.columns:
        df=df[df[C_NO].astype(str).str.strip()!=""]
    return df.reset_index(drop=True)

@st.cache_data(show_spinner=False)
def load_plans_fp():
    try:
        with open(os.path.join(APP_DIR,"plans_fp","plans_meta.json"),encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return {"plans":[],"points":{}}

@st.cache_data(show_spinner=False)
def plan_fp_img(key):
    p=os.path.join(APP_DIR,"plans_fp",key+".jpg")
    if not os.path.exists(p): return ""
    return "data:image/jpeg;base64,"+base64.b64encode(open(p,"rb").read()).decode()

def render_punchlist_fp():
    if st.columns([1,5])[0].button("← เมนูหลัก", key="fp_back"):
        st.session_state["view"]="menu"; st.rerun()
    _sv=st.session_state.pop("_fp_saved_rows",None)
    if _sv is not None:
        st.toast(f"บันทึกกลับ Google Sheet สำเร็จ ({_sv} แถว) ✓", icon="✅")
    connected=True
    try:
        raw=load_raw_fp(FP_WS)
    except Exception as e:
        raw=pd.DataFrame(columns=ALL_COLS); connected=False
        st.warning(f"เชื่อม Google Sheet ไม่สำเร็จ ({e})")
    df=enrich(raw)
    left,right=st.columns([4,1])
    with left:
        st.markdown("### 🚒 Punchlist FP · Multipurpose")
        st.caption(f"MRT สายสีม่วง (Contract 1) · งานระบบดับเพลิง (FP) ชั้น Multipurpose · {len(df)} จุด")
    with right:
        if st.button("🔄 รีเฟรชข้อมูล", key="fp_refresh", use_container_width=True):
            st.cache_data.clear(); st.rerun()
        st.caption(("🟢 เชื่อม Google Sheet (แก้ได้)" if connected else "🟠 ยังไม่เชื่อม Sheet")+f" · ณ {bkk_today().strftime('%d/%m/%Y')}")
    tot=len(df)
    scount={s:int((df[C_STATUS].astype(str).str.strip()==s).sum()) for s in STATUS_ORDER}
    kk=st.columns(1+len(STATUS_ORDER))
    kk[0].metric("ทั้งหมด", tot)
    for _i,_s0 in enumerate(STATUS_ORDER):
        _m=STATUS_META[_s0]; kk[_i+1].metric(f"{_m['icon']} {_m['short']}", scount[_s0])
    st.divider()
    if tot>0:
        st.markdown("##### สรุปตามมิติ")
        rr=st.columns(2)
        with rr[0]:
            st.caption("สถานะงาน")
            _sv2=[s for s in STATUS_ORDER if scount[s]>0]
            st.plotly_chart(hbar([(STATUS_META[s]["short"],scount[s]) for s in _sv2], colors=[STATUS_META[s]["color"] for s in _sv2]), use_container_width=True, config={"displayModeBar":False})
        with rr[1]:
            st.caption("ตามระบบงาน")
            _sc=df[C_SYS].replace("","(ไม่ระบุ)").value_counts()
            st.plotly_chart(hbar(list(_sc.items())), use_container_width=True, config={"displayModeBar":False})
        st.divider()
    tabT,tabP,tabE=st.tabs(["📋 ตารางรายการ","🗺️ แบบแปลนติดจุด","✏️ แก้ไขข้อมูล"])
    with tabT:
        st.caption(f"แสดง {len(df)} จุด · คอลัมน์ตรงกับ Google Sheet")
        disp=df[ALL_COLS].copy()
        for c in (C_NO,C_PAGE): disp[c]=pd.to_numeric(disp[c],errors="coerce").astype("Int64")
        st.dataframe(disp, use_container_width=True, hide_index=True, height=320,
            column_config={C_NO:st.column_config.NumberColumn("ลำดับ",format="%d",width="small"),
                C_NICK:st.column_config.TextColumn("รหัส/ชื่อเรียก",width="small"),
                C_PAGE:st.column_config.NumberColumn("หน้า",format="%d",width="small")})
    with tabP:
        meta=load_plans_fp()
        if not meta.get("plans"):
            st.info("ยังไม่มีข้อมูลแบบแปลนติดจุด")
        else:
            def _code(n):
                r=df[df[C_NO].astype(str).str.strip()==str(n)]
                v=str(r.iloc[0][C_NICK]).strip() if len(r) else ""
                return v or str(n)
            opts={f"{p['file']} · หน้า {p['page']}  ({len(p['nos'])} จุด: {', '.join(_code(n) for n in p['nos'])})":p for p in meta["plans"]}
            lab=st.selectbox("เลือกแบบแปลน (เฉพาะหน้าที่มีจุด)", list(opts.keys()), key="fp_plansel")
            plan=opts[lab]
            dmap={int(r[C_NO]):r for _,r in df.iterrows() if str(r[C_NO]).strip() not in ("","nan")}
            pts=[]
            for n in plan["nos"]:
                pos=meta["points"].get(str(n)); row=dmap.get(n)
                if not pos or row is None: continue
                mm=STATUS_META.get(str(row[C_STATUS]).strip(),{})
                pts.append({"no":str(row.get(C_NICK,"") or n),"x":pos["x"],"y":pos["y"],"nick":"",
                    "color":mm.get("color","#8a8a86"),"icon":mm.get("icon","⚪"),
                    "status":str(row[C_STATUS]),"floor":str(row[C_FLOOR]),"sys":str(row[C_SYS]),
                    "dwg":str(row[C_DWG]),"page":str(row[C_PAGE]),"loc":str(row[C_LOC]),
                    "detail":str(row[C_DETAIL]),"red":str(row[C_RED]),"owner":str(row[C_OWNER]),
                    "start":str(row[C_START]),"due":str(row[C_DUE]),"days":fmt_days(row),
                    "note":str(row[C_NOTE] or "")})
            _cc=st.columns([3,1])
            _cc[0].caption("🖱️ ชี้/แตะหมุด = รายละเอียด · ลาก/ล้อเมาส์ = เลื่อน-ซูม · สีหมุด = สถานะ")
            _cnt={}
            for _p in pts: _cnt[_p['status']]=_cnt.get(_p['status'],0)+1
            _cc[1].caption("  ".join(f"{STATUS_META[s]['icon']} {_cnt[s]}" for s in STATUS_ORDER if _cnt.get(s)))
            components.html(build_plan_html(plan_fp_img(plan["key"]), pts, height=640), height=662, scrolling=False)
            st.caption("แบบระบบดับเพลิง (FP) ชั้น Multipurpose — จากไฟล์ FP Multi · ข้อมูลจุดดึงสดจาก Google Sheet")
    with tabE:
        st.caption("แก้ไข / เพิ่ม / ลบจุดได้ (ปุ่ม + ด้านล่าง) แล้วกด บันทึก (ต้องใส่ PIN)")
        edf=st.data_editor(raw, num_rows="dynamic", use_container_width=True, height=320, key="fp_editor",
            column_config={C_STATUS:st.column_config.SelectboxColumn("สถานะ",options=STATUS_ORDER),
                C_FLOOR:st.column_config.SelectboxColumn("ชั้น",options=FLOOR_ORDER),
                C_NICK:st.column_config.TextColumn("รหัส/ชื่อเรียก"),
                C_NO:st.column_config.NumberColumn("ลำดับ",format="%d")})
        cp,cb=st.columns([2,1])
        pin=cp.text_input("PIN บันทึก", type="password", key="fp_pin", label_visibility="collapsed", placeholder="ใส่ PIN เพื่อบันทึก")
        if cb.button("💾 บันทึกกลับ Google Sheet", type="primary", key="fp_save", use_container_width=True):
            if pin_bad(pin): st.error("PIN ไม่ถูกต้อง (หรือยังไม่ได้ตั้ง edit_pin ใน Secrets)")
            else:
                try:
                    out=edf.copy(); out=out[out[C_NO].map(lambda v:_canon_no(v)!="")]
                    ws=get_ws(FP_WS)
                    if ws is None:
                        st.error("ยังไม่ได้เชื่อม Google Sheet")
                    else:
                        values=[list(out.columns)]+[[_s(v) for v in row] for row in out.values.tolist()]
                        ws.clear(); ws.update(values=values, range_name="A1")
                        st.cache_data.clear(); st.session_state["_fp_saved_rows"]=len(out); st.rerun()
                except Exception as e:
                    st.error(f"บันทึกไม่สำเร็จ: {e}")

if "view" not in st.session_state: st.session_state["view"]="menu"
if st.session_state["view"]=="menu":
    render_menu(); st.stop()
if st.session_state["view"]=="remaining":
    render_remaining(); st.stop()
if st.session_state["view"]=="actual":
    render_actual(); st.stop()
if st.session_state["view"]=="punchlist_fp":
    render_punchlist_fp(); st.stop()
if st.session_state["view"]!="punchlist":
    render_placeholder(st.session_state["view"]); st.stop()

# ---- ปุ่มกลับเมนูหลัก (แสดงเฉพาะหน้า Punchlist) ----
if st.columns([1,5])[0].button("← เมนูหลัก", key="pl_back"):
    st.session_state["view"]="menu"; st.rerun()

raw, mode = load_raw()
df = enrich(raw)

# ---- pop-up แจ้งเตือนหลังบันทึกสำเร็จ (แสดงหลัง rerun) ----
_saved = st.session_state.pop("_saved_rows", None)
if _saved is not None:
    st.toast(f"บันทึกกลับ Google Sheet สำเร็จ ({_saved} แถว) ✓", icon="✅")
    try: st.balloons()
    except Exception: pass
_imgsaved = st.session_state.pop("_img_saved", None)
if _imgsaved is not None:
    st.toast(f"อัปเดตรูปของจุดที่ {_imgsaved} แล้ว ✓", icon="🖼️")
_dsaved = st.session_state.pop("_done_saved", None)
if _dsaved is not None:
    _no,_n=_dsaved
    st.toast(f"บันทึกรูปงานเสร็จของจุดที่ {_no} แล้ว ({_n} รูป) ✓", icon="📷")
    try: st.balloons()
    except Exception: pass
_ddel = st.session_state.pop("_done_del", None)
if _ddel is not None:
    st.toast(f"ลบรูปงานเสร็จของจุดที่ {_ddel} แล้ว", icon="🗑️")

# ---- header ----
left,right=st.columns([4,1])
with left:
    st.markdown("### 🚇 Punchlist Dashboard — PP18 SI YAN `R2`")
    st.caption("MRT สายสีม่วง (Contract 1) · ติดตามงานตามกรอบสีแดงในแบบ · 75 จุด")
with right:
    if st.button("🔄 รีเฟรชข้อมูล", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    badge={"gsheet":"🟢 เชื่อม Google Sheet (แก้ได้)","csv_url":"🔵 อ่านจากลิงก์ (ดูอย่างเดียว)","local":"🟠 ข้อมูลตั้งต้น (ยังไม่เชื่อม Sheet)"}
    st.caption(badge.get(mode,mode)+f" · ณ {bkk_today().strftime('%d/%m/%Y')}")

# ---- KPIs ----
tot=len(df)
scount={s:int((df[C_STATUS].astype(str).str.strip()==s).sum()) for s in STATUS_ORDER}
c_soon=int(((~df["_skey"].isin(["crit","done","cancel"])) & (df["_days"].apply(lambda x: x is not None and not pd.isna(x) and 0<=x<=7))).sum())
k=st.columns(1+len(STATUS_ORDER))
k[0].metric("ทั้งหมด", tot)
for _ki,_ks in enumerate(STATUS_ORDER):
    _km=STATUS_META[_ks]; k[_ki+1].metric(f"{_km['icon']} {_km['short']}", scount[_ks])
st.caption(f"⏰ ครบกำหนดใน 7 วัน (ยังไม่จบงาน): {c_soon} จุด")

st.divider()

# ---- charts ----
st.markdown("##### สรุปตามมิติ")
r1=st.columns(2)
with r1[0]:
    st.caption("สถานะงาน")
    pairs=[(STATUS_META[s]["icon"]+" "+STATUS_META[s]["short"], int((df[C_STATUS]==s).sum())) for s in STATUS_ORDER]
    st.plotly_chart(hbar(pairs, colors=[STATUS_META[s]["color"] for s in STATUS_ORDER]), use_container_width=True, config={"displayModeBar":False})
with r1[1]:
    st.caption("ตามชั้น")
    fl=[(f,int((df[C_FLOOR]==f).sum())) for f in FLOOR_ORDER if (df[C_FLOOR]==f).any()]
    st.plotly_chart(hbar(fl, colors=[FLOOR_COLOR.get(f,SERIES) for f,_ in fl]), use_container_width=True, config={"displayModeBar":False})
r2=st.columns(2)
with r2[0]:
    st.caption("ตามระบบงาน")
    sysc=df[C_SYS].value_counts()
    st.plotly_chart(hbar(list(sysc.items())), use_container_width=True, config={"displayModeBar":False})
with r2[1]:
    st.caption("ตามผู้รับผิดชอบ")
    ow=df[C_OWNER].replace("-","(ไม่ระบุ)").value_counts()
    st.plotly_chart(hbar(list(ow.items())), use_container_width=True, config={"displayModeBar":False})
st.caption("กำหนดเสร็จรายเดือน (เลยกำหนด = แดง)")
st.plotly_chart(month_stacked(df), use_container_width=True, config={"displayModeBar":False})

st.divider()

# ---- tabs: table / image viewer / edit ----
tab1,tab2,tab_plan,tab3=st.tabs(["📋 ตารางรายการ","🖼️ ดูรูปแบบรายจุด","🗺️ แบบแปลนติดจุด","✏️ แก้ไขข้อมูล"])

with tab1:
    f=st.columns([3,1.3,1.3,1.3,1.3])
    q=f[0].text_input("ค้นหา", placeholder="ตำแหน่ง / รายละเอียด / Drawing No. / ผู้รับผิดชอบ …", label_visibility="collapsed")
    ff=f[1].selectbox("ชั้น", ["ทุกชั้น"]+FLOOR_ORDER, label_visibility="collapsed")
    fsys=f[2].selectbox("ระบบ", ["ทุกระบบ"]+sorted(df[C_SYS].dropna().unique().tolist()), label_visibility="collapsed")
    fow=f[3].selectbox("ผู้รับผิดชอบ", ["ทุกคน"]+sorted(df[C_OWNER].dropna().unique().tolist()), label_visibility="collapsed")
    fst=f[4].selectbox("สถานะ", ["ทุกสถานะ"]+STATUS_ORDER, label_visibility="collapsed")
    view=df.copy()
    if ff!="ทุกชั้น": view=view[view[C_FLOOR]==ff]
    if fsys!="ทุกระบบ": view=view[view[C_SYS]==fsys]
    if fow!="ทุกคน": view=view[view[C_OWNER]==fow]
    if fst!="ทุกสถานะ": view=view[view[C_STATUS]==fst]
    if q:
        ql=q.lower()
        hay=view[[C_LOC,C_DETAIL,C_DWG,C_RED,C_OWNER,C_FLOOR,C_SYS]].astype(str).agg(" ".join,axis=1).str.lower()
        view=view[hay.str.contains(ql, na=False)]
    st.caption(f"แสดง {len(view)} / {tot} จุด  ·  คอลัมน์ตรงกับ Google Sheet")
    disp = view[ALL_COLS].copy()          # แสดงครบทุกคอลัมน์ เรียงเหมือนใน Google Sheet
    for c in (C_NO, C_PAGE):
        disp[c] = pd.to_numeric(disp[c], errors="coerce").astype("Int64")
    st.dataframe(disp, use_container_width=True, hide_index=True, height=560,
        column_config={
            C_NO:   st.column_config.NumberColumn("ลำดับ", format="%d", width="small"),
            C_NICK: st.column_config.TextColumn("ชื่อเรียก", width="medium", help="ชื่อเล่นที่หน้างานใช้เรียกจุดนี้"),
            C_PAGE: st.column_config.NumberColumn("หน้า", format="%d", width="small"),
        })

with tab2:
    nos=sorted(df[C_NO].dropna().astype(int).tolist())
    ov = load_overrides()
    donep = load_done_photos()
    sel=st.selectbox("เลือกจุด", nos,
        format_func=lambda n:f"จุดที่ {n}"+(f"  📷×{len(donep.get(n,[]))}" if donep.get(n) else ""))
    row=df[df[C_NO]==sel].iloc[0]
    cimg,cinfo=st.columns([2,1])
    with cimg:
        p=os.path.join(APP_DIR,"images",f"{sel}.jpg")
        if sel in ov:
            try:
                st.image(base64.b64decode(ov[sel]), use_container_width=True)
                st.caption("🖼️ รูปที่อัปโหลดทับไว้")
            except Exception:
                if os.path.exists(p): st.image(p, use_container_width=True)
        elif os.path.exists(p):
            st.image(p, use_container_width=True)
        else:
            st.info("ไม่มีรูปแบบสำหรับจุดนี้ (จุดที่เพิ่มใหม่)")
        with st.expander(f"📤 อัปโหลด / เปลี่ยนรูปของจุดที่ {sel}"):
            if mode!="gsheet":
                st.caption("อัปโหลดได้เมื่อเชื่อม Google Sheet (โหมดแก้ไข) — ตอนนี้เป็นโหมดดูอย่างเดียว")
            else:
                up=st.file_uploader("เลือกรูปใหม่ (PNG/JPG)", type=["png","jpg","jpeg"], key=f"up_{sel}")
                cpin,cbtn=st.columns([2,1])
                pin2=cpin.text_input("PIN", type="password", key=f"uppin_{sel}", label_visibility="collapsed", placeholder="ใส่ PIN เพื่อบันทึก")
                if cbtn.button("💾 บันทึกรูป", key=f"upbtn_{sel}", use_container_width=True):
                    if up is None:
                        st.warning("ยังไม่ได้เลือกรูป")
                    elif str(pin2)!=str(sget("edit_pin","")) or sget("edit_pin","")=="":
                        st.error("PIN ไม่ถูกต้อง")
                    else:
                        try:
                            save_override(sel, compress_to_b64(up.getvalue()))
                            st.cache_data.clear()
                            st.session_state["_img_saved"]=sel
                            st.rerun()
                        except Exception as e:
                            st.error(f"อัปโหลดไม่สำเร็จ: {e}")
                st.caption("รูปจะถูกบีบให้เล็กลงเพื่อเก็บในชีต แล้วแทนรูปเดิมของจุดนี้ทันที")
    with cinfo:
        meta=STATUS_META.get(row[C_STATUS],{})
        st.markdown(f"**จุดที่ {sel}** — {meta.get('icon','')} {meta.get('short',row[C_STATUS])}")
        st.markdown(f"""
- **ชั้น / ระบบ:** {row[C_FLOOR]} · {row[C_SYS]}
- **Drawing No.:** {row[C_DWG]} (หน้า {row[C_PAGE]})
- **ตำแหน่ง:** {row[C_LOC]}
- **รายละเอียด:** {row[C_DETAIL]}
- **ข้อความในกรอบแดง:** {row[C_RED]}
- **ผู้รับผิดชอบ:** {row[C_OWNER]}
- **กำหนด:** {row[C_START]} → {row[C_DUE]}  ({fmt_days(row)})
- **หมายเหตุ:** {row[C_NOTE] or '–'}
""")

    # ---- รูปหลักฐานงานเสร็จหน้างาน (เก็บได้หลายรูป/จุด) ----
    st.divider()
    mine = donep.get(int(sel), [])
    st.markdown(f"##### 📷 รูปงานเสร็จหน้างาน — จุดที่ {sel}  ({len(mine)} รูป)")
    if mine:
        gc=st.columns(4)
        for i,ph in enumerate(mine):
            with gc[i%4]:
                try:
                    st.image(base64.b64decode(ph["b64"]), use_container_width=True,
                             caption=f"{ph['date']}"+(f" · {ph['note']}" if ph['note'] else ""))
                except Exception:
                    st.caption("(รูปเสียหาย)")
    else:
        st.caption("ยังไม่มีรูปงานเสร็จของจุดนี้ — อัปโหลดหลักฐานงานที่แก้เสร็จได้ในกล่องด้านล่าง")
    with st.expander(f"➕ เพิ่ม / จัดการรูปงานเสร็จของจุดที่ {sel}"):
        if mode!="gsheet":
            st.caption("อัปโหลดได้เมื่อเชื่อม Google Sheet (โหมดแก้ไข) — ตอนนี้เป็นโหมดดูอย่างเดียว")
        else:
            dpin=st.text_input("PIN", type="password", key=f"dpin_{sel}",
                               placeholder="ใส่ PIN เพื่อเพิ่ม/ลบรูป", label_visibility="collapsed")
            st.markdown("**เพิ่มรูป** (เลือกได้หลายรูปพร้อมกัน)")
            ups=st.file_uploader("รูปงานเสร็จ (PNG/JPG)", type=["png","jpg","jpeg"],
                                 accept_multiple_files=True, key=f"dup_{sel}")
            dnote=st.text_input("หมายเหตุ (ไม่บังคับ)", key=f"dnote_{sel}",
                                placeholder="เช่น เก็บงานเสร็จ 23/07, ผู้ตรวจ …")
            if st.button("💾 บันทึกรูปงานเสร็จ", key=f"dbtn_{sel}", use_container_width=True):
                if not ups:
                    st.warning("ยังไม่ได้เลือกรูป")
                elif pin_bad(dpin):
                    st.error("PIN ไม่ถูกต้อง")
                else:
                    try:
                        n=0
                        for uf in ups:
                            save_done_photo(sel, compress_to_b64(uf.getvalue()), dnote); n+=1
                        st.cache_data.clear()
                        st.session_state["_done_saved"]=(sel,n)
                        st.rerun()
                    except Exception as e:
                        st.error(f"บันทึกไม่สำเร็จ: {e}")
            if mine:
                st.markdown("**ลบรูป**")
                opt={f"{i+1}. {ph['date']}"+(f" · {ph['note']}" if ph['note'] else ""):ph["id"]
                     for i,ph in enumerate(mine)}
                dsel=st.selectbox("เลือกรูปที่จะลบ", list(opt.keys()), key=f"delsel_{sel}")
                if st.button("🗑️ ลบรูปนี้", key=f"delbtn_{sel}"):
                    if pin_bad(dpin):
                        st.error("PIN ไม่ถูกต้อง")
                    else:
                        try:
                            del_done_photo(sel, opt[dsel]); st.cache_data.clear()
                            st.session_state["_done_del"]=sel; st.rerun()
                        except Exception as e:
                            st.error(f"ลบไม่สำเร็จ: {e}")
            st.caption("รูปถูกบีบให้เล็กก่อนเก็บในชีต (แท็บ \"รูปงานเสร็จ\") · เก็บรูปมากอาจทำให้ชีตโหลดช้าลง")

with tab_plan:
    meta=load_plans_meta()
    if not meta.get("plans"):
        st.info("ยังไม่มีข้อมูลแบบแปลนติดจุด")
    else:
        opts={f"{p['file']} · หน้า {p['page']}  ({len(p['nos'])} จุด: {', '.join(map(str,p['nos']))})":p for p in meta["plans"]}
        lab=st.selectbox("เลือกแบบแปลน (เฉพาะหน้าที่มีจุด)", list(opts.keys()), key="plansel")
        plan=opts[lab]
        dmap={int(r[C_NO]):r for _,r in df.iterrows() if str(r[C_NO]).strip() not in ("","nan")}
        pts=[]
        for no in plan["nos"]:
            pos=meta["points"].get(str(no)); row=dmap.get(no)
            if not pos or row is None: continue
            mm=STATUS_META.get(str(row[C_STATUS]).strip(),{})
            pts.append({"no":no,"x":pos["x"],"y":pos["y"],"nick":str(row.get(C_NICK,"") or ""),
                "color":mm.get("color","#8a8a86"),"icon":mm.get("icon","⚪"),
                "status":str(row[C_STATUS]),"floor":str(row[C_FLOOR]),"sys":str(row[C_SYS]),
                "dwg":str(row[C_DWG]),"page":str(row[C_PAGE]),"loc":str(row[C_LOC]),
                "detail":str(row[C_DETAIL]),"red":str(row[C_RED]),"owner":str(row[C_OWNER]),
                "start":str(row[C_START]),"due":str(row[C_DUE]),"days":fmt_days(row),
                "note":str(row[C_NOTE] or "")})
        cc=st.columns([3,1])
        cc[0].caption("🖱️ ชี้หมุด = สรุปเร็ว · แตะหมุด = รายละเอียดเต็ม · ลาก/ล้อเมาส์ = เลื่อน-ซูม · สีหมุด = สถานะ")
        _cnt={}
        for p in pts: _cnt[p['status']]=_cnt.get(p['status'],0)+1
        cc[1].caption("  ".join(f"{STATUS_META[s]['icon']} {_cnt[s]}" for s in STATUS_ORDER if _cnt.get(s)))
        import streamlit.components.v1 as components
        components.html(build_plan_html(plan_img_url(plan["key"]), pts, height=640), height=662, scrolling=False)
        st.caption("รูปแบบเต็มหน้าที่มีจุด (จากไฟล์ Punchlist R1) · ตำแหน่งหมุดคำนวณจากเลขวงกลมแดงในแบบ · ข้อมูลจุดดึงสดจาก Google Sheet")

with tab3:
    st.caption("แก้ไขข้อมูลได้เลย · เพิ่ม/ลบแถวได้ (ปุ่ม + ด้านล่าง / เลือกแถวแล้วลบ) แล้วกด \"บันทึก\" (ต้องใส่ PIN)")
    if mode!="gsheet":
        st.warning("โหมดนี้ยังไม่ได้เชื่อม Google Sheet — แก้ได้แต่ **บันทึกกลับไม่ได้** จนกว่าจะตั้งค่า service account (ดู README)")
    edited=st.data_editor(
        raw, num_rows="dynamic", use_container_width=True, height=460, key="editor",
        column_config={
            C_STATUS: st.column_config.SelectboxColumn("สถานะ", options=STATUS_ORDER),
            C_FLOOR: st.column_config.SelectboxColumn("ชั้น", options=FLOOR_ORDER),
            C_NO: st.column_config.NumberColumn("ลำดับ", format="%d"),
        },
    )
    cpin,cbtn=st.columns([2,1])
    pin=cpin.text_input("PIN สำหรับบันทึก", type="password", label_visibility="collapsed", placeholder="ใส่ PIN เพื่อบันทึก")
    if cbtn.button("💾 บันทึกกลับ Google Sheet", type="primary", use_container_width=True):
        ws=get_ws() if mode=="gsheet" else None
        if ws is None:
            st.error("ยังไม่ได้เชื่อม Google Sheet — ตั้งค่า service account ก่อน (README ข้อ 2)")
        elif pin_bad(pin):
            st.error("PIN ไม่ถูกต้อง (หรือยังไม่ได้ตั้ง edit_pin ใน Secrets)")
        else:
            out=edited.copy()
            out=out[out[C_NO].map(lambda v:_canon_no(v)!="")]
            diff=compute_diff(raw, out)
            if diff["n_add"]==0 and diff["n_del"]==0 and diff["n_mod"]==0:
                st.info("ไม่มีการเปลี่ยนแปลงให้บันทึก")
            else:
                st.session_state["_pending_edited"]=out
                st.session_state["_pending_diff"]=diff
                st.session_state["_show_confirm"]=True
                st.rerun()
    if st.session_state.get("_show_confirm"):
        confirm_save_dialog()

st.divider()
st.caption("ที่มา: 2026-07-24_Punchlist_MRT_PP18_R2.xlsx · คอลัมน์ \"เหลือ (วัน)\" คำนวณสดจากวันปัจจุบัน (เขต Asia/Bangkok) · รูปแบบราย จุดฝังในโปรเจกต์ (โฟลเดอร์ images/)")
