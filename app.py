# -*- coding: utf-8 -*-
"""
Punchlist Dashboard — สถานี PP18 SI YAN (MRT สายสีม่วง, Contract 1) · R1
Streamlit app: อ่าน/แก้ข้อมูลจาก Google Sheet แบบ near real-time + ดูรูปแบบราย จุด
"""
import os, io, base64
from datetime import datetime, timezone, timedelta
import pandas as pd
import streamlit as st
import plotly.graph_objects as go

# ----------------------------------------------------------------------------
st.set_page_config(page_title="Punchlist PP18 · R1", page_icon="🚇", layout="wide")
APP_DIR = os.path.dirname(os.path.abspath(__file__))

# ---- column names (ตรงกับหัวตารางใน Google Sheet) ----
C_NO="ลำดับ"; C_FLOOR="ชั้น"; C_SYS="ระบบ"; C_DWGF="ไฟล์แบบ"; C_DWG="Drawing No."
C_PAGE="หน้า"; C_RED="ข้อความในกรอบแดง"; C_OWNER="ผู้รับผิดชอบ"; C_START="กำหนดเริ่ม"
C_DUE="กำหนดเสร็จ"; C_LOC="ตำแหน่ง/บริเวณ"; C_DETAIL="รายละเอียดงาน"; C_STATUS="สถานะ"
C_DONE="วันที่เสร็จจริง"; C_NOTE="หมายเหตุ"
ALL_COLS=[C_NO,C_FLOOR,C_SYS,C_DWGF,C_DWG,C_PAGE,C_RED,C_OWNER,C_START,C_DUE,
          C_LOC,C_DETAIL,C_STATUS,C_DONE,C_NOTE]

STATUS_ORDER=["เลยกำหนด–รอตรวจสอบ","กำลังดำเนินการ","รอเชื่อม","รอวัสดุ","รอดำเนินการ","จบงาน"]
STATUS_META={
 "เลยกำหนด–รอตรวจสอบ":{"key":"crit","color":"#d03b3b","icon":"🔴","short":"เลยกำหนด"},
 "กำลังดำเนินการ":{"key":"warn","color":"#fab219","icon":"🟡","short":"กำลังดำเนินการ"},
 "รอเชื่อม":{"key":"conn","color":"#2f7fd1","icon":"🔵","short":"รอเชื่อม"},
 "รอวัสดุ":{"key":"matl","color":"#9b59b6","icon":"🟣","short":"รอวัสดุ"},
 "รอดำเนินการ":{"key":"neut","color":"#8a8a86","icon":"⚪","short":"รอดำเนินการ"},
 "จบงาน":{"key":"done","color":"#1f9d57","icon":"✅","short":"จบงาน"},
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
def get_ws():
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
        wsname=sget("worksheet","Punchlist")
        try: return sh.worksheet(wsname)
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
 card.innerHTML='<span class="x">✕</span><h4>'+p.icon+' จุดที่ '+p.no+' — '+esc(p.status)+'</h4>'+
 '<b>ชั้น/ระบบ:</b> '+esc(p.floor)+' · '+esc(p.sys)+'<br><b>Drawing:</b> '+esc(p.dwg)+' (หน้า '+esc(p.page)+')<br>'+
 '<b>ตำแหน่ง:</b> '+esc(p.loc)+'<br><b>รายละเอียดงาน:</b> '+esc(p.detail)+'<br><b>ในกรอบแดง:</b> '+esc(p.red)+'<br>'+
 '<b>ผู้รับผิดชอบ:</b> '+esc(p.owner)+'<br><b>กำหนด:</b> '+esc(p.start)+' → '+esc(p.due)+' ('+esc(p.days)+')<br>'+
 '<b>หมายเหตุ:</b> '+esc(p.note||'–');
 card.querySelector('.x').onclick=function(e){e.stopPropagation();card.style.display='none';};}
function build(){PTS.forEach(function(p){var m=document.createElement('div');m.className='mk';m.style.left=p.x+'%';m.style.top=p.y+'%';m.style.background=p.color;m.textContent=p.no;
 m.onmouseenter=function(){bar.style.display='block';bar.innerHTML='<b>จุดที่ '+p.no+'</b> · '+p.icon+' '+esc(p.status)+' · '+esc(p.owner)+' · ครบ '+esc(p.due)+' ('+esc(p.days)+')';};
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
    st.markdown("### 🚇 Punchlist Dashboard — PP18 SI YAN `R1`")
    st.caption("MRT สายสีม่วง (Contract 1) · ติดตามงานตามกรอบสีแดงในแบบ · 51 จุด")
with right:
    if st.button("🔄 รีเฟรชข้อมูล", use_container_width=True):
        st.cache_data.clear(); st.rerun()
    badge={"gsheet":"🟢 เชื่อม Google Sheet (แก้ได้)","csv_url":"🔵 อ่านจากลิงก์ (ดูอย่างเดียว)","local":"🟠 ข้อมูลตั้งต้น (ยังไม่เชื่อม Sheet)"}
    st.caption(badge.get(mode,mode)+f" · ณ {bkk_today().strftime('%d/%m/%Y')}")

# ---- KPIs ----
tot=len(df)
scount={s:int((df[C_STATUS].astype(str).str.strip()==s).sum()) for s in STATUS_ORDER}
c_soon=int(((~df["_skey"].isin(["crit","done"])) & (df["_days"].apply(lambda x: x is not None and not pd.isna(x) and 0<=x<=7))).sum())
k=st.columns(1+len(STATUS_ORDER))
k[0].metric("ทั้งหมด", tot)
for _i,_s in enumerate(STATUS_ORDER):
    _m=STATUS_META[_s]; k[_i+1].metric(f"{_m['icon']} {_m['short']}", scount[_s])
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
            pts.append({"no":no,"x":pos["x"],"y":pos["y"],
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
st.caption("ที่มา: 2026-07-22_Punchlist_MRT_PP18_R1.xlsx · คอลัมน์ \"เหลือ (วัน)\" คำนวณสดจากวันปัจจุบัน (เขต Asia/Bangkok) · รูปแบบราย จุดฝังในโปรเจกต์ (โฟลเดอร์ images/)")
