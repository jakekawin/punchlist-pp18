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

STATUS_ORDER=["เลยกำหนด–รอตรวจสอบ","กำลังดำเนินการ","รอดำเนินการ"]
STATUS_META={
 "เลยกำหนด–รอตรวจสอบ":{"key":"crit","color":"#d03b3b","icon":"🔴","short":"เลยกำหนด"},
 "กำลังดำเนินการ":{"key":"warn","color":"#fab219","icon":"🟡","short":"กำลังดำเนินการ"},
 "รอดำเนินการ":{"key":"neut","color":"#8a8a86","icon":"⚪","short":"รอดำเนินการ"},
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
    d=r["_days"]
    if d is None or pd.isna(d): return "–"
    d=int(d)
    if r["_skey"]=="crit" or d<0: return f"เลย {abs(d)} วัน"
    return f"{d} วัน"

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
c_crit=int((df["_skey"]=="crit").sum()); c_warn=int((df["_skey"]=="warn").sum()); c_neut=int((df["_skey"]=="neut").sum())
c_soon=int(((df["_skey"]!="crit") & (df["_days"].apply(lambda x: x is not None and not pd.isna(x) and 0<=x<=7))).sum())
k=st.columns(5)
k[0].metric("ทั้งหมด", tot)
k[1].metric("🔴 เลยกำหนด", c_crit)
k[2].metric("🟡 กำลังดำเนินการ", c_warn)
k[3].metric("⚪ รอดำเนินการ", c_neut)
k[4].metric("⏰ ครบกำหนดใน 7 วัน", c_soon)

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
tab1,tab2,tab3=st.tabs(["📋 ตารางรายการ","🖼️ ดูรูปแบบรายจุด","✏️ แก้ไขข้อมูล"])

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
    sel=st.selectbox("เลือกจุด", nos, format_func=lambda n:f"จุดที่ {n}")
    row=df[df[C_NO]==sel].iloc[0]
    ov = load_overrides()
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
        elif str(pin)!=str(sget("edit_pin","")) or sget("edit_pin","")=="":
            st.error("PIN ไม่ถูกต้อง (หรือยังไม่ได้ตั้ง edit_pin ใน Secrets)")
        else:
            try:
                out=edited.copy()
                out=out[out[C_NO].astype(str).str.strip()!=""]
                values=[list(out.columns)]+out.fillna("").astype(str).values.tolist()
                ws.clear(); ws.update(values=values, range_name="A1")
                st.cache_data.clear()
                st.session_state["_saved_rows"]=len(out)   # ให้ pop-up เด้งหลัง rerun
                st.rerun()
            except Exception as e:
                st.error(f"บันทึกไม่สำเร็จ: {e}")

st.divider()
st.caption("ที่มา: 2026-07-22_Punchlist_MRT_PP18_R1.xlsx · คอลัมน์ \"เหลือ (วัน)\" คำนวณสดจากวันปัจจุบัน (เขต Asia/Bangkok) · รูปแบบราย จุดฝังในโปรเจกต์ (โฟลเดอร์ images/)")
