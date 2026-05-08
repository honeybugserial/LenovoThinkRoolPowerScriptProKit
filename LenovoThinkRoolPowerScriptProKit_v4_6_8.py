import sys, csv, time, threading, re, subprocess, platform
try:
    import tomllib          # Python 3.11+
except ImportError:
    try:
        import tomli as tomllib   # pip install tomli
    except ImportError:
        tomllib = None
from datetime import datetime, date
from html import unescape
from io import BytesIO
from pathlib import Path
from collections import defaultdict
from typing import Any, Dict, List, Optional
from urllib.parse import quote

import requests
from PyQt6 import QtWidgets, QtGui, QtCore
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QSizePolicy,
    QPushButton, QLabel, QFileDialog, QFrame, QScrollArea,
    QLineEdit, QComboBox, QTabWidget, QGridLayout,
    QTreeWidget, QTreeWidgetItem, QHeaderView, QAbstractItemView,
    QDialog, QCheckBox, QDialogButtonBox, QMenuBar, QMenu, QProgressBar, QPlainTextEdit
)
from PyQt6.QtGui import QColor, QPalette, QDesktopServices, QCursor, QFont, QAction
import threading
from PyQt6.QtCore import Qt, QUrl, QThread, pyqtSignal, QObject, QMetaObject, QTimer
from PyQt6.QtCore import Q_ARG

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER

# ══════════════════════════════════════════════════════════════════════════════
# USER CONFIGURATION  —  tweak anything here without touching the rest of the code
# ══════════════════════════════════════════════════════════════════════════════

# ── App identity ──────────────────────────────────────────────────────────────
APP_NAME        = "-- LenovoThinkToolPowerScriptProKit --"
APP_VERSION     = "4.6.8"
APP_TITLE       = "-- LenovoThinkToolPowerScriptProKit --"           # appears BEFORE "TOOLKIT" in the header bar
APP_SUBTITLE    = " _for things and stuff"          # appears after "LENOVO" in the header bar
REPORTS_FOLDER  = "Reports"          # subfolder created next to the script

# ── Window ────────────────────────────────────────────────────────────────────
WIN_WIDTH       = 1300               # initial window width
WIN_HEIGHT      = 840                # initial window height
WIN_MIN_WIDTH   = 900                # minimum window width
WIN_MIN_HEIGHT  = 560                # minimum window height

# ── Panels ────────────────────────────────────────────────────────────────────
TOPBAR_HEIGHT       = 54             # header bar with logo

# ── Controls ──────────────────────────────────────────────────────────────────
CTRL_HEIGHT_LG  = 42                 # serial input, detect/lookup buttons
CTRL_HEIGHT_MD  = 38                 # OS combobox, scan button
CTRL_HEIGHT_SM  = 34                 # filter row controls (search, category, severity)
CTRL_HEIGHT_XS  = 36                 # CSV viewer toolbar controls

# ── Fonts ─────────────────────────────────────────────────────────────────────
FONT_LOGO       = 25                 # APP_NAME header
FONT_LABEL      = 14                 # input bar labels, tab bar, dropdowns, inputs
FONT_BUTTON     = 14                 # main buttons (scan, look up, etc.)
FONT_TREE       = 13                 # driver tree rows
FONT_TREE_HDR   = 12                 # driver tree column headers
FONT_PANEL_TITLE= 16                 # product name in left panel
FONT_PANEL_SUB  = 12                 # serial/mtm subtitle line
FONT_SECTION    = 11                 # section labels (WARRANTY INFORMATION, etc.)
FONT_KV_KEY     = 12                 # key column in info panels
FONT_KV_VALUE   = 12                 # value column in info panels
FONT_STATUS     = 12                 # status bar messages
FONT_BADGE      = 11                 # warranty status badge

# ══════════════════════════════════════════════════════════════════════════════

DARK_BG        = "#0f1117"
PANEL_BG       = "#171b26"
CARD_BG        = "#1e2333"
CARD_HOVER     = "#252a3d"
BORDER         = "#2a2f45"
ACCENT         = "#e2231a"
ACCENT_SOFT    = "#ff4d44"
TEXT_PRIMARY   = "#eef0f6"
TEXT_SECONDARY = "#8890aa"
TEXT_DIM       = "#555d78"
SEV_CRITICAL   = "#ff4d44"
SEV_RECOMMEND  = "#3b9eff"
SEV_OPTIONAL   = "#a0aec0"
GREEN          = "#34d058"
YELLOW         = "#f0b429"

SEVERITY_COLORS = {"critical": SEV_CRITICAL, "recommended": SEV_RECOMMEND, "optional": SEV_OPTIONAL}

CATEGORY_ICONS = {
    "Audio":"🔊","BIOS/UEFI":"⚡","Bluetooth and Modem":"📡",
    "Camera and Card Reader":"📷","Diagnostic":"🔬",
    "Display and Video Graphics":"🖥","Fingerprint Reader":"👆",
    "Keyboard and Mouse":"⌨️","Networking: Ethernet":"🔌",
    "Networking: Wireless LAN":"📶","Networking: Wireless WAN":"📲",
    "Patch":"🩹","Power Management":"🔋","Software and Utilities":"🛠",
    "Storage":"💾","ThinkVantage Technology":"🧰",
}

SPEC_ORDER = ["Processor","Memory","Hard Drive","Wireless Network",
              "Graphics","Monitor","Camera","Ports","Included Warranty","End of Service"]

BASE_URL = "https://pcsupport.lenovo.com/us/en"
API_URL  = f"{BASE_URL}/api/v4"
HDRS = {
    "Accept":"application/json, text/plain, */*",
    "Content-Type":"application/json",
    "Origin":BASE_URL,
    "Referer":f"{BASE_URL}/warranty-lookup",
    "User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) thinkfo/3.0",
}

def norm_serial(s):
    return re.sub(r"[^A-Za-z0-9]","",s or "").upper()

def _root(js):
    return js.get("Data") or js.get("data") or {}

def parse_iso_date(s):
    if not s: return None
    try: return datetime.fromisoformat(s.strip().replace("/","-")).date()
    except: return None

def compute_active(start_s, end_s):
    today=date.today(); start=parse_iso_date(start_s); end=parse_iso_date(end_s)
    if not end and not start: return None
    if end and today>end: return False
    if start and today<start: return False
    return True

def product_title_from_name(name):
    name=(name or "").strip()
    name=re.sub(r"\s*-\s*Type\s+[A-Za-z0-9]{4}\s*$","",name,flags=re.I)
    name=re.sub(r"\s+(Laptop|Notebook|Desktop|Workstation|Tablet)\s*$","",name,flags=re.I)
    return name

def detect_local_serial():
    system=platform.system()
    if system=="Windows":
        for cmd in [
            ["powershell","-NoProfile","-Command","(Get-WmiObject Win32_BIOS).SerialNumber"],
            ["powershell","-NoProfile","-Command","(Get-CimInstance Win32_BIOS).SerialNumber"],
            ["wmic","bios","get","SerialNumber","/value"],
        ]:
            try:
                out=subprocess.check_output(cmd,text=True,timeout=8,stderr=subprocess.DEVNULL,
                    creationflags=subprocess.CREATE_NO_WINDOW).strip()
                if not out: continue
                if "=" in out:
                    for line in out.splitlines():
                        if "SerialNumber=" in line:
                            val=line.split("=",1)[1].strip()
                            if val: return val
                else:
                    val=[l.strip() for l in out.splitlines() if l.strip()][-1]
                    if val and val.lower() not in ("serialnumber","none",""): return val
            except: continue
    elif system=="Linux":
        try:
            out=subprocess.check_output(["sudo","dmidecode","-s","system-serial-number"],
                text=True,timeout=5,stderr=subprocess.DEVNULL).strip()
            if out: return out
        except: pass
    elif system=="Darwin":
        try:
            out=subprocess.check_output(["system_profiler","SPHardwareDataType"],text=True,timeout=5,stderr=subprocess.DEVNULL)
            for line in out.splitlines():
                if "Serial Number" in line: return line.split(":",1)[1].strip()
        except: pass
    return None

def get_warranty(serial,timeout=15.0):
    payload={"serialNumber":serial,"country":"us","language":"en"}
    r=requests.post(f"{API_URL}/upsell/redport/getIbaseInfo",headers=HDRS,json=payload,timeout=timeout)
    r.raise_for_status(); return r.json()

def extract_fields(wj):
    root=_root(wj); mi=root.get("machineInfo") or {}; cw=root.get("currentWarranty") or {}
    if not cw:
        bw=root.get("baseWarranties") or []
        cw=bw[0] if isinstance(bw,list) and bw else (bw if isinstance(bw,dict) else {})
    return {
        "productName":mi.get("productName"),"serial":mi.get("serial") or mi.get("serialNumber"),
        "machineType":mi.get("type") or mi.get("machineType"),"product":mi.get("product"),
        "model":mi.get("model"),"shipToCountry":mi.get("shipToCountry"),
        "warrantyStatus":root.get("warrantyStatus"),"planName":cw.get("name"),
        "deliveryType":cw.get("deliveryTypeName") or cw.get("deliveryType"),
        "startDate":cw.get("startDate"),"endDate":cw.get("endDate") or cw.get("EndDate"),
        "subSeries":mi.get("subSeries"),"fullId":mi.get("fullId"),
        "specification":mi.get("specification") or "",
    }

def parse_spec_html(spec_html):
    if not spec_html: return {}
    html=re.sub(r"\s+"," ",unescape(spec_html).replace("Â"," ")); out={}
    for row in re.findall(r"<tr>(.*?)</tr>",html,flags=re.I|re.S):
        cells=[unescape(re.sub(r"<.*?>","",c)).strip()
               for c in re.findall(r"<td[^>]*>(.*?)</td>",row,flags=re.I|re.S)]
        if len(cells)>=2 and cells[0]:
            val=" ".join(cells[1:]).strip().replace("DRR4","DDR4"); out.setdefault(cells[0],val)
    if "Memory" in out and out["Memory"]:
        parts=[p.strip(" ;") for p in re.split(r"[;|]",out["Memory"]) if p.strip()]
        seen,uniq=set(),[]
        for p in parts:
            k=re.sub(r"\s+","",p).lower()
            if k not in seen: seen.add(k); uniq.append(p)
        out["Memory"]="; ".join(uniq)
    return {k:" ".join((v or "").split()) for k,v in out.items()}

def build_product_url(wj):
    root=_root(wj); mi=root.get("machineInfo") or {}; id_path=mi.get("fullId")
    if not id_path:
        parts=[mi.get("group"),mi.get("series"),mi.get("subSeries"),mi.get("type"),mi.get("product"),mi.get("serial")]
        id_path="/".join(p for p in parts if p)
    slug=quote((id_path or "").lower(),safe="/")
    return f"{BASE_URL}/products/{slug}" if slug else ""

def get_drivers(full_id, os_filter=None, timeout=20.0):
    url = f"{API_URL}/downloads/drivers?productId={full_id}"
    if os_filter:
        url += f"&osId={os_filter}"

    referer = f"{BASE_URL}/products/{full_id}/downloads"

    session = requests.Session()

    # Warmup: visit product page to get Akamai session cookies
    session.get(
        f"{BASE_URL}/products/{full_id}",
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        },
        timeout=timeout
    )

    # Now fetch drivers with correct headers and session cookies
    r = session.get(url, headers={
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer": referer,
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0",
        "x-requested-with": "XMLHttpRequest",
        "x-requested-timezone": "America/New_York",
        "DNT": "1",
    }, timeout=timeout)

    r.raise_for_status()
    return r.json()

def _s(v):
    """Safely coerce any API value to a plain string."""
    if v is None: return ""
    if isinstance(v, str): return v
    if isinstance(v, dict):
        # Unix timestamp dict e.g. {"Unix": 1640155860000}
        unix = v.get("Unix") or v.get("unix")
        if unix:
            try: return datetime.fromtimestamp(int(unix)/1000).strftime("%d %b %Y")
            except: pass
        return v.get("Name") or v.get("name") or v.get("Value") or ""
    return str(v)

def parse_drivers(raw_json, os_filter_name=None):
    body=raw_json.get("body") or raw_json.get("Body") or {}
    items=body.get("DownloadItems") or []; rows=[]
    for item in items:
        # Category
        cat_raw=item.get("Category") or {}
        cat=_s(cat_raw.get("Name") if isinstance(cat_raw,dict) else cat_raw) or "Other"
        # Title / version / date / severity
        title=_s(item.get("Title") or item.get("Name") or "")
        si=item.get("SummaryInfo") or {}
        version=_s(si.get("Version") or item.get("Version") or "")
        severity=_s(si.get("Priority") or item.get("SeverityType") or item.get("Severity") or "")
        date_raw=item.get("Date") or ""
        # OS keys for client-side filtering
        os_keys=[k.lower() for k in (item.get("OperatingSystemKeys") or [])]
        if os_filter_name and os_filter_name.lower() not in os_keys and "os independent" not in os_keys:
            continue
        for f in item.get("Files") or []:
            size=_s(f.get("Size") or "")
            date_f=_s(f.get("Date") or date_raw)
            rows.append({
                "Category":    cat,
                "Description": title,
                "File Name":   _s(f.get("Name") or ""),
                "Size":        size,
                "Version":     _s(f.get("Version") or version),
                "Release Date":date_f,
                "Severity":    _s(f.get("Priority") or severity),
                "URL":         _s(f.get("URL") or f.get("Url") or ""),
                "TypeString":  _s(f.get("TypeString") or ""),
            })
    return rows

def extract_os_list(raw_json):
    """Build OS list from AllOperatingSystems if populated, else from OperatingSystemKeys."""
    body=raw_json.get("body") or {}
    all_os=body.get("AllOperatingSystems") or []
    if all_os:
        # Old format: list of {ID, Name} dicts — return as {name, id} pairs
        return [{"name":o.get("Name",""), "id":o.get("ID","")} for o in all_os if o.get("Name")]
    # New format: collect unique names from DownloadItems.OperatingSystemKeys
    items=body.get("DownloadItems") or []
    seen=set(); result=[]
    for item in items:
        for k in (item.get("OperatingSystemKeys") or []):
            if k and k not in seen and k.lower()!="os independent":
                seen.add(k); result.append({"name":k,"id":k})
    return sorted(result, key=lambda x:x["name"])

# ── PDF ───────────────────────────────────────────────────────────────────────
def build_pdf_report(wf,spec,product_url,drivers=None,os_name=""):
    buf=BytesIO()
    doc=SimpleDocTemplate(buf,pagesize=letter,
        leftMargin=0.75*inch,rightMargin=0.75*inch,topMargin=0.75*inch,bottomMargin=0.75*inch)
    LENOVO_RED=colors.HexColor("#e2231a"); DARK=colors.HexColor("#0f1117")
    MID=colors.HexColor("#1e2333"); LIGHT=colors.HexColor("#f4f5f8")
    DIM=colors.HexColor("#8890aa"); C_GREEN=colors.HexColor("#1a8f3a")
    C_RED=colors.HexColor("#c41c14"); C_YELLOW=colors.HexColor("#b07d00")
    WHITE=colors.white; BLACK=colors.HexColor("#0f1117")
    styles=getSampleStyleSheet()
    def sty(n,**kw):
        return ParagraphStyle(n+"_x"+str(abs(hash(str(kw)))),parent=styles.get(n,styles["Normal"]),**kw)
    title_sty  =sty("Normal",fontSize=19,fontName="Helvetica-Bold",textColor=WHITE,leading=23)
    sub_sty    =sty("Normal",fontSize=8.5,fontName="Helvetica",textColor=DIM,leading=13)
    sec_sty    =sty("Normal",fontSize=10.5,fontName="Helvetica-Bold",textColor=LENOVO_RED,leading=14,spaceAfter=3)
    lbl_sty    =sty("Normal",fontSize=8,fontName="Helvetica-Bold",textColor=DIM,leading=12)
    val_sty    =sty("Normal",fontSize=8.5,fontName="Helvetica",textColor=BLACK,leading=13)
    sm_sty     =sty("Normal",fontSize=7.5,fontName="Helvetica",textColor=DIM,leading=11)
    url_sty    =sty("Normal",fontSize=7.5,fontName="Helvetica",textColor=colors.HexColor("#1a56db"),leading=11)
    story=[]
    reptime=datetime.now().strftime("%d %b %Y  %H:%M")
    prod_name=product_title_from_name(wf.get("productName")) or wf.get("product") or "Lenovo Device"
    hdr=Table([[Paragraph(prod_name,title_sty),
                Paragraph(f"Generated {reptime}",sty("Normal",fontSize=8,fontName="Helvetica",textColor=DIM,alignment=TA_RIGHT))]],
              colWidths=["68%","32%"])
    hdr.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),DARK),("TOPPADDING",(0,0),(-1,-1),14),
        ("BOTTOMPADDING",(0,0),(-1,-1),14),("LEFTPADDING",(0,0),(-1,-1),16),("RIGHTPADDING",(0,0),(-1,-1),16),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    story.append(hdr); story.append(Spacer(1,8))
    story.append(Paragraph(f"Serial: {wf.get('serial') or '-'}  ·  MTM: {wf.get('product') or '-'}/{wf.get('model') or '-'}  ·  Type: {wf.get('machineType') or '-'}  ·  Ship-To: {wf.get('shipToCountry') or '-'}",sub_sty))
    story.append(Spacer(1,14))
    active=compute_active(wf.get("startDate"),wf.get("endDate"))
    warr_status=wf.get("warrantyStatus") or ("In Warranty" if active else "Out of Warranty")
    badge_color=C_GREEN if active is True else (C_RED if active is False else C_YELLOW)
    badge_bg=colors.HexColor("#e6f9ec") if active is True else (colors.HexColor("#fde8e8") if active is False else colors.HexColor("#fef9e7"))
    badge_row=Table([[
        Paragraph(warr_status.upper(),sty("Normal",fontSize=10,fontName="Helvetica-Bold",textColor=badge_color,alignment=TA_CENTER)),
        Paragraph(f"<b>Plan:</b> {wf.get('planName') or '-'}  |  <b>Delivery:</b> {wf.get('deliveryType') or '-'}  |  <b>Start:</b> {wf.get('startDate') or '-'}  <b>End:</b> {wf.get('endDate') or '-'}",
                  sty("Normal",fontSize=8.5,fontName="Helvetica",textColor=BLACK,leading=13))
    ]],colWidths=[1.3*inch,None])
    badge_row.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),badge_bg),("BACKGROUND",(1,0),(1,0),LIGHT),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),12),("RIGHTPADDING",(0,0),(-1,-1),12),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LINEAFTER",(0,0),(0,0),1,colors.HexColor("#d0d5e0")),("BOX",(0,0),(-1,-1),1,colors.HexColor("#d0d5e0"))]))
    story.append(badge_row); story.append(Spacer(1,16))
    def kv_tbl(pairs,title):
        rows=[[Paragraph(title,sec_sty),""]]+[[Paragraph(k,lbl_sty),Paragraph(v,val_sty)] for k,v in pairs]
        t=Table(rows,colWidths=[1.35*inch,None])
        t.setStyle(TableStyle([("SPAN",(0,0),(1,0)),("LINEBELOW",(0,0),(-1,0),1.5,LENOVO_RED),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,LIGHT]),("LINEBELOW",(0,1),(-1,-1),0.4,colors.HexColor("#dde1ea")),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),("VALIGN",(0,0),(-1,-1),"TOP")]))
        return t
    warr_pairs=[("Serial",wf.get("serial") or "-"),("Machine Type",wf.get("machineType") or "-"),
        ("MTM Product",wf.get("product") or "-"),("Model",wf.get("model") or "-"),
        ("Ship-To",wf.get("shipToCountry") or "-"),("Warranty Status",warr_status),
        ("Plan",wf.get("planName") or "-"),("Delivery Type",wf.get("deliveryType") or "-"),
        ("Start Date",wf.get("startDate") or "-"),("End Date",wf.get("endDate") or "-")]
    spec_pairs=[(k,spec[k]) for k in SPEC_ORDER if k in spec] or [("—","No spec data")]
    two_col=Table([[kv_tbl(warr_pairs,"WARRANTY INFO"),kv_tbl(spec_pairs,"BUILD SPEC")]],
                  colWidths=["48%","52%"],hAlign="LEFT")
    two_col.setStyle(TableStyle([("LEFTPADDING",(0,0),(-1,-1),0),("RIGHTPADDING",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"TOP")]))
    story.append(two_col)
    if product_url:
        story.append(Spacer(1,12)); story.append(HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#d0d5e0")))
        story.append(Spacer(1,5)); story.append(Paragraph(f'<b>Product Page:</b> <a href="{product_url}">{product_url}</a>',url_sty))
    if drivers:
        story.append(Spacer(1,18))
        os_label=f" — {os_name}" if os_name else ""
        story.append(Paragraph(f"DRIVER DOWNLOADS{os_label}",sec_sty))
        story.append(HRFlowable(width="100%",thickness=1.5,color=LENOVO_RED)); story.append(Spacer(1,6))
        by_cat=defaultdict(list)
        for d in drivers: by_cat[d.get("Category","Other")].append(d)
        for cat in sorted(by_cat.keys()):
            cat_rows=by_cat[cat]
            ch=Table([[Paragraph(cat,sty("Normal",fontSize=9,fontName="Helvetica-Bold",textColor=WHITE))]],colWidths=["100%"])
            ch.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),DARK),("TOPPADDING",(0,0),(-1,-1),5),
                ("BOTTOMPADDING",(0,0),(-1,-1),5),("LEFTPADDING",(0,0),(-1,-1),10)]))
            story.append(ch)
            drv_data=[[Paragraph("<b>File</b>",sm_sty),Paragraph("<b>Version</b>",sm_sty),
                        Paragraph("<b>Size</b>",sm_sty),Paragraph("<b>Sev</b>",sm_sty),Paragraph("<b>URL</b>",sm_sty)]]
            for d in cat_rows:
                url_v=d.get("URL",""); sev=d.get("Severity","")
                sc=colors.HexColor("#c41c14") if sev.lower()=="critical" else (colors.HexColor("#1a56db") if sev.lower()=="recommended" else DIM)
                drv_data.append([Paragraph(d.get("File Name",""),sm_sty),Paragraph(d.get("Version",""),sm_sty),
                    Paragraph(d.get("Size",""),sm_sty),
                    Paragraph(sev,sty("Normal",fontSize=7.5,fontName="Helvetica-Bold",textColor=sc,leading=11)),
                    Paragraph(f'<a href="{url_v}">{url_v}</a>',url_sty) if url_v else Paragraph("-",sm_sty)])
            dt=Table(drv_data,colWidths=[1.8*inch,1.0*inch,0.65*inch,0.65*inch,None])
            dt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,0),colors.HexColor("#e8eaf0")),
                ("ROWBACKGROUNDS",(0,1),(-1,-1),[WHITE,LIGHT]),("LINEBELOW",(0,0),(-1,-1),0.3,colors.HexColor("#dde1ea")),
                ("TOPPADDING",(0,0),(-1,-1),3),("BOTTOMPADDING",(0,0),(-1,-1),3),
                ("LEFTPADDING",(0,0),(-1,-1),5),("RIGHTPADDING",(0,0),(-1,-1),5),("VALIGN",(0,0),(-1,-1),"TOP")]))
            story.append(dt); story.append(Spacer(1,6))
    story.append(Spacer(1,10)); story.append(HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#d0d5e0")))
    story.append(Spacer(1,4))
    story.append(Paragraph(f"{APP_NAME}  ·  {reptime}  ·  pcsupport.lenovo.com",
        sty("Normal",fontSize=7.5,fontName="Helvetica",textColor=DIM,alignment=TA_CENTER)))
    doc.build(story); buf.seek(0); return buf

# ── Workers ───────────────────────────────────────────────────────────────────
class WarrantyWorker(QObject):
    finished=pyqtSignal(dict,dict,str,object)
    error=pyqtSignal(str)
    def __init__(self,serial): super().__init__(); self.serial=serial
    def run(self):
        try:
            raw=get_warranty(self.serial); wf=extract_fields(raw)
            spec=parse_spec_html(wf.get("specification") or ""); url=build_product_url(raw)
            self.finished.emit(wf,spec,url,raw)
        except Exception as e: self.error.emit(str(e))

class DriverFetchWorker(QObject):
    """Fetches drivers from Lenovo API for a specific OS (re-fetches every time)"""
    finished = pyqtSignal(list)
    error = pyqtSignal(str)

    def __init__(self, full_id, os_id=None):
        super().__init__()
        self.full_id = full_id
        self.os_id = os_id

    def run(self):
        try:
            raw = get_drivers(self.full_id, self.os_id)
            drivers = parse_drivers(raw)
            self.finished.emit(drivers)
        except Exception as e:
            self.error.emit(str(e))


def _exe_dir() -> Path:
    """Directory next to the exe — for editable/user files (tools.toml, downloads, reports)."""
    return Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent

def _resource_dir() -> Path:
    """Bundled read-only resources (splash.png) — uses _MEIPASS when frozen."""
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


# ==========================
# SPLASH CONFIGURATION
# ==========================
SPLASH_IMAGE    = "splash.png"
SPLASH_STEPS = [
    {"message": "Detecting Things...",             "pause": 0.7},
    {"message": "Things Detected [1/2]",           "pause": 0.9},
    {"message": "Things Detected [2/2]",           "pause": 0.81},
    {"message": "Do run now. Completed. Starting!", "pause": 1.2},
]
SPLASH_TYPING        = True
SPLASH_TYPING_SPEED  = 25
SPLASH_FADE_IN       = True
SPLASH_FADE_DURATION = 800
SPLASH_ALWAYS_ON_TOP = True


# ==========================
# SPLASH SCREEN
# ==========================
class SplashScreen(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        flags = Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: rgba(15, 15, 22, 0.97);
                border-radius: 22px;
                border: 1px solid rgba(255,255,255,0.1);
            }
            QLabel { background: transparent; border: none; }
        """)
        card_layout = QtWidgets.QVBoxLayout(card)
        card_layout.setContentsMargins(25, 25, 25, 25)
        card_layout.setSpacing(10)

        # Image
        img_path = str(_resource_dir() / SPLASH_IMAGE)
        pix = QtGui.QPixmap(img_path)
        if not pix.isNull():
            img_label = QLabel()
            img_label.setPixmap(pix)
            img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            card_layout.addWidget(img_label)
        else:
            fallback = QLabel("Loading…")
            fallback.setAlignment(Qt.AlignmentFlag.AlignCenter)
            fallback.setStyleSheet("color:white;font-size:24px;font-weight:600;")
            card_layout.addWidget(fallback)

        self.status = QLabel("Starting...")
        self.status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status.setStyleSheet("color:#ffffff;font-size:17px;font-weight:500;background:transparent;border:none;")
        card_layout.addWidget(self.status)

        self.progress = QtWidgets.QProgressBar()
        self.progress.setFixedHeight(16)
        self.progress.setTextVisible(False)
        self.progress.setStyleSheet("""
            QProgressBar { background-color:rgba(35,35,42,0.95); border:1px solid rgba(70,70,80,0.7); border-radius:8px; }
            QProgressBar::chunk { background-color:#e2231a; border-radius:8px; }
        """)
        card_layout.addWidget(self.progress)
        layout.addWidget(card)

        self.adjustSize()
        scr = QtWidgets.QApplication.primaryScreen().availableGeometry()
        self.move((scr.width() - self.width()) // 2, (scr.height() - self.height()) // 2)

        self._fade_started = False
        if SPLASH_FADE_IN:
            self.setWindowOpacity(0.0)

    def showEvent(self, event):
        super().showEvent(event)
        if SPLASH_FADE_IN and not self._fade_started:
            self._fade_started = True
            QTimer.singleShot(0, lambda: self._do_fade())

    def _do_fade(self):
        self.anim = QtCore.QPropertyAnimation(self, b"windowOpacity")
        self.anim.setDuration(SPLASH_FADE_DURATION)
        self.anim.setStartValue(0.0)
        self.anim.setEndValue(1.0)
        self.anim.start()

    def update_status(self, text, progress=None):
        if SPLASH_TYPING:
            self._type_text(text)
        else:
            self.status.setText(text)
        if progress is not None:
            self.progress.setValue(int(progress))
        QtWidgets.QApplication.processEvents()

    def _type_text(self, text):
        self.status.setText("")
        QtWidgets.QApplication.processEvents()
        for ch in text:
            self.status.setText(self.status.text() + ch)
            QtWidgets.QApplication.processEvents()
            QtCore.QThread.msleep(SPLASH_TYPING_SPEED)

    def run_steps(self, on_complete):
        steps = SPLASH_STEPS
        total = len(steps)
        self._on_complete = on_complete
        def _execute():
            for i, step in enumerate(steps, 1):
                progress = int((i / total) * 100) if total > 0 else 50
                self.update_status(step["message"], progress)
                time.sleep(step["pause"])
            QtCore.QMetaObject.invokeMethod(self, "_finish",
                Qt.ConnectionType.QueuedConnection)
        self._thread = threading.Thread(target=_execute, daemon=True)
        self._thread.start()

    @QtCore.pyqtSlot()
    def _finish(self):
        self.close()
        if self._on_complete:
            self._on_complete()


def _get_tools_toml() -> Path:
    """
    Returns path to tools.toml.
    On first run, copies the bundled version from _MEIPASS next to the exe so the user can edit it.
    After that, always uses the editable copy next to the exe.
    """
    user_toml = _exe_dir() / "tools.toml"
    if user_toml.exists():
        return user_toml
    bundled = _resource_dir() / "tools.toml"
    if bundled.exists():
        import shutil as _shutil
        _shutil.copy(bundled, user_toml)
    return user_toml


# ── Stylesheet ────────────────────────────────────────────────────────────────
STYLESHEET=f"""
QMainWindow,QWidget{{background-color:{DARK_BG};color:{TEXT_PRIMARY};font-family:"Segoe UI","SF Pro Display",sans-serif;}}
QTabWidget::pane{{border:none;border-top:1px solid {BORDER};background:{DARK_BG};}}
QTabWidget::tab-bar{{alignment:left;}}
QTabBar{{background:{PANEL_BG};border-bottom:1px solid {BORDER};}}
QTabBar::tab{{background:{PANEL_BG};color:{TEXT_DIM};border:none;border-right:1px solid {BORDER};padding:14px 36px;font-size:13px;font-weight:600;min-width:110px;}}
QTabBar::tab:selected{{color:{TEXT_PRIMARY};background:{DARK_BG};border:2px solid {ACCENT};border-radius:4px;}}
QTabBar::tab:hover:!selected{{color:{TEXT_SECONDARY};background:{CARD_BG};}}
QScrollArea{{border:none;background:transparent;}}
QScrollBar:vertical{{background:{PANEL_BG};width:6px;border-radius:3px;}}
QScrollBar::handle:vertical{{background:{BORDER};border-radius:3px;min-height:30px;}}
QScrollBar::handle:vertical:hover{{background:{TEXT_DIM};}}
QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{{height:0px;}}
QScrollBar:horizontal{{background:{PANEL_BG};height:6px;border-radius:3px;}}
QScrollBar::handle:horizontal{{background:{BORDER};border-radius:3px;min-width:30px;}}
QScrollBar::add-line:horizontal,QScrollBar::sub-line:horizontal{{width:0px;}}
QLineEdit{{background:{PANEL_BG};border:1px solid {BORDER};border-radius:8px;color:{TEXT_PRIMARY};padding:9px 14px;font-size:{FONT_LABEL}px;}}
QLineEdit:focus{{border-color:{ACCENT};}}
QComboBox{{background:{PANEL_BG};border:1px solid {BORDER};border-radius:8px;color:{TEXT_PRIMARY};padding:9px 14px;font-size:{FONT_LABEL}px;min-width:160px;}}
QComboBox:focus{{border-color:{ACCENT};}}
QComboBox::drop-down{{border:none;width:24px;}}
QComboBox QAbstractItemView{{background:{CARD_BG};border:1px solid {BORDER};color:{TEXT_PRIMARY};selection-background-color:{ACCENT};border-radius:8px;padding:4px;}}
QPushButton{{background:{ACCENT};color:white;border:2px solid transparent;border-radius:8px;padding:8px 20px;font-size:{FONT_BUTTON}px;font-weight:600;margin:2px;}}
QPushButton:hover{{background:{ACCENT_SOFT};border-color:transparent;}}
QPushButton:pressed{{background:#c41c14;}}
QPushButton:disabled{{background:{BORDER};color:{TEXT_DIM};border-color:transparent;}}
QPushButton:focus{{outline:none;border:2px solid {ACCENT_SOFT};}}
QPushButton#ghost{{background:transparent;border:2px solid {BORDER};color:{TEXT_SECONDARY};margin:2px;}}
QPushButton#ghost:hover{{border-color:{ACCENT};color:{ACCENT_SOFT};}}
QPushButton#ghost:focus{{outline:none;border:2px solid {ACCENT};color:{ACCENT_SOFT};}}
QPushButton#ghost:disabled{{border-color:{BORDER};color:{TEXT_DIM};}}
QTreeWidget{{background:{CARD_BG};border:1px solid {BORDER};border-radius:8px;color:{TEXT_PRIMARY};font-size:{FONT_TREE}px;outline:none;}}
QTreeWidget::item{{padding:4px 6px;border-bottom:1px solid {BORDER};}}
QTreeWidget::item:selected{{background:{ACCENT}22;color:{TEXT_PRIMARY};}}
QTreeWidget::item:hover{{background:{CARD_HOVER};}}
QHeaderView::section{{background:{PANEL_BG};color:{TEXT_DIM};border:none;border-bottom:1px solid {BORDER};padding:7px 8px;font-size:{FONT_TREE_HDR}px;font-weight:600;}}
"""

# ── Shared driver tree ────────────────────────────────────────────────────────
class DriverTree(QTreeWidget):
    download_requested = pyqtSignal(str)   # emitted when a row download btn is clicked

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setHeaderLabels(["File Name", "Version", "Size", "Date", "Download", ""])
        # Severity column removed — to restore, add back "Severity" at index 4 and shift Download/Select to 5/6

        self.setColumnWidth(0, 540)   # File Name
        self.setColumnWidth(1, 100)   # Version
        self.setColumnWidth(2, 90)    # Size
        self.setColumnWidth(3, 115)   # Date
        # self.setColumnWidth(4, 115) # Severity (hidden)
        self.setColumnWidth(4, 120)   # Download
        self.setColumnWidth(5, 50)    # Select

        self.header().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.header().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.header().setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.header().setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        self.header().setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        self.header().setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)
        self.header().setStretchLastSection(False)

        self.setAlternatingRowColors(False)
        self.setRootIsDecorated(True)
        self.setSortingEnabled(False)
        self.setSelectionMode(QAbstractItemView.SelectionMode.NoSelection)

        # Checkbox style
        self.cb_style = f"""
            QCheckBox {{
                spacing: 10px;
            }}
            QCheckBox::indicator {{
                width: 28px;
                height: 28px;
                border: 2px solid {BORDER};
                border-radius: 7px;
                background-color: #252a3d;
            }}
            QCheckBox::indicator:hover {{
                border-color: {ACCENT_SOFT};
                background-color: #2f3649;
            }}
            QCheckBox::indicator:checked {{
                background-color: {ACCENT};
                border-color: {ACCENT};
                image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='28' height='28' viewBox='0 0 24 24' fill='none' stroke='white' stroke-width='3.5' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpolyline points='20 6 9 17 4 12'%3E%3C/polyline%3E%3C/svg%3E");
            }}
        """

        self._header_cb = QCheckBox()
        self._header_cb.setTristate(True)
        self._header_cb.setStyleSheet(self.cb_style)
        self._header_cb.stateChanged.connect(self._on_header_toggled)

        self._updating = False

    def showEvent(self, event):
        super().showEvent(event)
        self._reposition_header_cb()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._reposition_header_cb()

    def _reposition_header_cb(self):
        """Center the Select All checkbox in the header"""
        rect = self.header().sectionViewportPosition(5)
        w = self.header().sectionSize(5)
        h = self.header().height()
        self._header_cb.setParent(self.header())
        self._header_cb.move(rect + (w - 28) // 2, (h - 28) // 2)
        self._header_cb.show()

    def _on_header_toggled(self, state):
        if self._updating: return
        # Tristate cycles Unchecked->Partial->Checked; skip partial, treat it as checked
        if state == Qt.CheckState.PartiallyChecked.value:
            self._updating = True
            self._header_cb.setCheckState(Qt.CheckState.Checked)
            self._updating = False
            return
        checked = (state == Qt.CheckState.Checked.value)
        # Scroll to top first so Qt has created all item widgets
        self.scrollToTop()
        for i in range(self.topLevelItemCount()):
            cat = self.topLevelItem(i)
            if cat.isHidden(): continue
            for j in range(cat.childCount()):
                row = cat.child(j)
                if row.isHidden(): continue
                # Force the item into view so itemWidget is created
                self.scrollToItem(row)
                cb = self.itemWidget(row, 5)
                if cb:
                    cb.blockSignals(True)
                    cb.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
                    cb.blockSignals(False)
        self.scrollToTop()
        self._sync_header()

    def _sync_header(self):
        total = checked = 0
        for i in range(self.topLevelItemCount()):
            cat = self.topLevelItem(i)
            if cat.isHidden(): continue
            for j in range(cat.childCount()):
                row = cat.child(j)
                if not row.isHidden():
                    total += 1
                    cb = self.itemWidget(row, 5)
                    if cb and cb.checkState() == Qt.CheckState.Checked:
                        checked += 1

        self._updating = True
        if total == 0 or checked == 0:
            self._header_cb.setCheckState(Qt.CheckState.Unchecked)
        elif checked == total:
            self._header_cb.setCheckState(Qt.CheckState.Checked)
        else:
            self._header_cb.setCheckState(Qt.CheckState.PartiallyChecked)
        self._updating = False

    def populate(self, drivers):
        self.clear()

        by_cat = defaultdict(list)
        for d in drivers:
            by_cat[d.get("Category", "Other")].append(d)

        for cat_name in sorted(by_cat.keys()):
            cat_item = QTreeWidgetItem(self)
            cat_item.setExpanded(True)

            icon = CATEGORY_ICONS.get(cat_name, "📦")
            cat_item.setText(0, f"{icon}  {cat_name}")
            cat_item.setForeground(0, QColor(TEXT_SECONDARY))
            font = QFont()
            font.setBold(True)
            cat_item.setFont(0, font)
            # Span category label across all columns so it doesn't truncate
            self.setFirstColumnSpanned(self.topLevelItemCount() - 1, self.rootIndex(), True)

            for d in by_cat[cat_name]:
                row = QTreeWidgetItem(cat_item)

                row.setText(0, d.get("File Name", ""))
                row.setText(1, d.get("Version", ""))
                row.setText(2, d.get("Size", ""))
                row.setText(3, d.get("Release Date", ""))
                # Severity column hidden — to restore: setText(4, sev), shift btn/cb to 5/6
                sev = d.get("Severity", "")
                # row.setText(4, sev)
                # row.setForeground(4, QColor(SEVERITY_COLORS.get(sev.lower(), TEXT_DIM)))
                row.setData(0, Qt.ItemDataRole.UserRole + 1, sev)  # store for filtering

                row.setData(0, Qt.ItemDataRole.UserRole, d.get("URL", ""))

                # Download button
                btn = QPushButton(f"↓ {Path(d.get('URL','')).suffix.upper().lstrip('.') or 'DL'}")
                btn.setFixedSize(90, 28)
                btn.setStyleSheet(f"""
                    QPushButton {{
                        background:{ACCENT}22; 
                        color:{ACCENT_SOFT};
                        border:1px solid {ACCENT}55; 
                        border-radius:7px;
                        font-size:11px; 
                        font-weight:700;
                        padding: 4px 14px;
                    }}
                    QPushButton:hover {{ 
                        background:{ACCENT}; 
                        color:white; 
                        border-color:{ACCENT}; 
                    }}
                """)
                url = d.get("URL", "")
                btn.clicked.connect(lambda _, u=url: self.download_requested.emit(u))
                self.setItemWidget(row, 4, btn)

                # Checkbox - far right
                cb = QCheckBox()
                cb.setStyleSheet(self.cb_style)
                cb.stateChanged.connect(self._sync_header)
                self.setItemWidget(row, 5, cb)

        self._sync_header()
        self._reposition_header_cb()

    def checked_urls(self):
        urls = []
        for i in range(self.topLevelItemCount()):
            cat = self.topLevelItem(i)
            for j in range(cat.childCount()):
                row = cat.child(j)
                if not row.isHidden():
                    cb = self.itemWidget(row, 5)
                    if cb and cb.checkState() == Qt.CheckState.Checked:
                        url = row.data(0, Qt.ItemDataRole.UserRole)
                        if url:
                            urls.append(url)
        return urls

    def apply_filter(self, text, sev_filter, cat_filter="All"):
        text = text.lower()
        for i in range(self.topLevelItemCount()):
            cat_item = self.topLevelItem(i)
            cat_label = " ".join(cat_item.text(0).split()[1:]) if cat_item.text(0) else ""
            cat_match = cat_filter in ("All", "All Categories") or cat_label.lower() == cat_filter.lower()

            if not cat_match:
                cat_item.setHidden(True)
                continue

            visible = 0
            for j in range(cat_item.childCount()):
                row = cat_item.child(j)
                match_text = not text or any(text in (row.text(c) or "").lower() for c in range(4))
                sev_val = row.data(0, Qt.ItemDataRole.UserRole + 1) or ""
                match_sev = sev_filter in ("All", "All Severities") or sev_val.lower() == sev_filter.lower()
                show = match_text and match_sev
                row.setHidden(not show)
                if show:
                    visible += 1

            cat_item.setHidden(visible == 0)

        self._sync_header()
  
# ── Driver CSV tab ─────────────────────────────────────────────────────────────
class DriversTab(QWidget):
    def __init__(self,parent=None):
        super().__init__(parent); self.all_data=[]; self._build()

    def _build(self):
        root=QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)
        # Stats bar
        self.stats_bar=QFrame(); self.stats_bar.setFixedHeight(62)
        self.stats_bar.setStyleSheet(f"background:{DARK_BG};border-bottom:1px solid {BORDER};"); self.stats_bar.hide()
        sb=QHBoxLayout(self.stats_bar); sb.setContentsMargins(28,0,28,0); sb.setSpacing(40)
        self.st_device=self._stat("—","DEVICE"); self.st_total=self._stat("0","TOTAL FILES")
        self.st_crit=self._stat("0","CRITICAL"); self.st_cats=self._stat("0","CATEGORIES")
        for s in [self.st_device,self.st_total,self.st_crit,self.st_cats]: sb.addLayout(s[2])
        sb.addStretch(); root.addWidget(self.stats_bar)
        # Toolbar
        tb=QFrame(); tb.setFixedHeight(54); tb.setStyleSheet(f"background:{PANEL_BG};border-bottom:1px solid {BORDER};")
        tbl=QHBoxLayout(tb); tbl.setContentsMargins(20,0,20,0); tbl.setSpacing(8)
        self.search=QLineEdit(); self.search.setPlaceholderText("🔍  Search..."); self.search.setFixedHeight(36)
        self.search.textChanged.connect(self._filter); tbl.addWidget(self.search,stretch=3)
        self.sev_cb=QComboBox(); self.sev_cb.addItems(["All","Critical","Recommended","Optional"])
        self.sev_cb.setFixedHeight(36); self.sev_cb.currentTextChanged.connect(self._filter); tbl.addWidget(self.sev_cb)
        self.count_lbl=QLabel(""); self.count_lbl.setStyleSheet(f"color:{TEXT_DIM};font-size:12px;background:transparent;"); tbl.addWidget(self.count_lbl)
        tbl.addStretch()
        load_btn=QPushButton("  📂  Load CSV"); load_btn.setFixedHeight(36); load_btn.clicked.connect(self._load_csv); tbl.addWidget(load_btn)
        root.addWidget(tb)
        # Tree wrap
        wrap=QWidget(); wrap.setStyleSheet(f"background:{DARK_BG};")
        wl=QVBoxLayout(wrap); wl.setContentsMargins(16,14,16,14)
        self.empty=QWidget(); el=QVBoxLayout(self.empty); el.setAlignment(Qt.AlignmentFlag.AlignCenter); el.setSpacing(8)
        for txt,style in [("💻",f"font-size:50px;background:transparent;"),("No CSV loaded",f"font-size:17px;font-weight:700;color:{TEXT_SECONDARY};background:transparent;"),("Click  📂 Load CSV  above",f"font-size:12px;color:{TEXT_DIM};background:transparent;")]:
            l=QLabel(txt); l.setStyleSheet(style); l.setAlignment(Qt.AlignmentFlag.AlignCenter); el.addWidget(l)
        wl.addWidget(self.empty)
        self.tree=DriverTree(); wl.addWidget(self.tree); root.addWidget(wrap,stretch=1)

    def _stat(self,v,l):
        layout=QVBoxLayout(); layout.setSpacing(2)
        vl=QLabel(v); vl.setStyleSheet(f"font-size:19px;font-weight:700;color:{TEXT_PRIMARY};background:transparent;")
        ll=QLabel(l); ll.setStyleSheet(f"font-size:10px;color:{TEXT_DIM};letter-spacing:1.2px;background:transparent;")
        layout.addWidget(vl); layout.addWidget(ll); return vl,ll,layout

    def _load_csv(self):
        path,_=QFileDialog.getOpenFileName(self,"Open Driver CSV","","CSV Files (*.csv)")
        if not path: return
        with open(path,newline="",encoding="utf-8") as f: self.all_data=list(csv.DictReader(f))
        if self.all_data: self._populate(path)

    def _populate(self,path):
        self.empty.hide(); self.tree.populate(self.all_data)
        total=len(self.all_data); crits=sum(1 for r in self.all_data if r.get("Severity","").lower()=="critical")
        cats=len({r.get("Category","") for r in self.all_data}); stem=Path(path).stem
        m=re.match(r"^(.+?)_\d{4}-\d{2}-\d{2}$",stem); device=m.group(1).replace("_"," ") if m else stem
        self.st_device[0].setText(device[:26]+("…" if len(device)>26 else ""))
        self.st_total[0].setText(str(total)); self.st_crit[0].setText(str(crits)); self.st_cats[0].setText(str(cats))
        self.stats_bar.show(); self._filter()

    def _filter(self):
        self.tree.apply_filter(self.search.text(),self.sev_cb.currentText())
        total=sum(self.tree.topLevelItem(i).childCount()-sum(self.tree.topLevelItem(i).child(j).isHidden() for j in range(self.tree.topLevelItem(i).childCount())) for i in range(self.tree.topLevelItemCount()))
        self.count_lbl.setText(f"{total} files shown")

# ── Warranty + Live Drivers tab ────────────────────────────────────────────────
class WarrantyTab(QWidget):
    dl_finished = pyqtSignal(str, int, bool)  # dest_path, count, cancelled

    def __init__(self,parent=None):
        super().__init__(parent); self._wf={}; self._spec={}; self._product_url=""
        self._raw_json=None; self._full_id=""; self._drivers=[]
        self._w_thread=self._w_worker=self._d_thread=self._d_worker=None
        self._build(); self._try_detect()
        self.driver_tree.download_requested.connect(self._download_single)
        self.dl_finished.connect(self._on_dl_finished)

    def _build(self):
        root=QVBoxLayout(self); root.setContentsMargins(0,0,0,0); root.setSpacing(0)
        # Input bar
        ib=QFrame(); ib.setStyleSheet(f"background:{PANEL_BG};border-bottom:1px solid {BORDER};")
        il=QHBoxLayout(ib); il.setContentsMargins(20,14,20,14); il.setSpacing(10)
        sn_lbl=QLabel("Serial Number"); sn_lbl.setStyleSheet(f"font-size:{FONT_LABEL}px;font-weight:600;color:{TEXT_SECONDARY};background:transparent;border:none;"); il.addWidget(sn_lbl)
        self.serial_input=QLineEdit(); self.serial_input.setPlaceholderText("e.g. R90KTB18"); self.serial_input.setFixedHeight(CTRL_HEIGHT_LG); self.serial_input.setStyleSheet(f"QLineEdit{{background:{PANEL_BG};border:1px solid {BORDER};border-radius:8px;color:{TEXT_PRIMARY};padding:9px 14px;font-size:{FONT_LABEL}px;margin-bottom:4px;}}QLineEdit:focus{{border-color:{ACCENT};margin-bottom:4px;}}"); self.serial_input.returnPressed.connect(self._lookup); il.addWidget(self.serial_input,stretch=1)
        self.detect_btn=QPushButton("⚡ Detect"); self.detect_btn.setObjectName("ghost"); self.detect_btn.setFixedHeight(CTRL_HEIGHT_LG); self.detect_btn.clicked.connect(self._try_detect); il.addWidget(self.detect_btn)
        self.lookup_btn=QPushButton("Look Up"); self.lookup_btn.setFixedHeight(CTRL_HEIGHT_LG); self.lookup_btn.setFixedWidth(110); self.lookup_btn.clicked.connect(self._lookup); il.addWidget(self.lookup_btn)
        self.status_lbl=QLabel(""); self.status_lbl.setStyleSheet(f"font-size:{FONT_STATUS}px;color:{TEXT_DIM};background:transparent;border:none;"); il.addWidget(self.status_lbl)
        root.addWidget(ib)

        # Tab widget
        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self._on_tab_changed)

        # ── Tab 1: Warranty ──────────────────────────────────────────
        warranty_widget = QWidget(); warranty_widget.setStyleSheet(f"background:{DARK_BG};")
        wl = QVBoxLayout(warranty_widget); wl.setContentsMargins(0,0,0,0); wl.setSpacing(0)
        ls=QScrollArea(); ls.setWidgetResizable(True); ls.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        ls.setStyleSheet(f"QScrollArea{{border:none;background:{DARK_BG};}}")
        self.left_w=QWidget(); self.left_w.setStyleSheet(f"background:{DARK_BG};")
        self.left_lay=QVBoxLayout(self.left_w); self.left_lay.setContentsMargins(24,24,24,24); self.left_lay.setSpacing(12)
        self.empty_left=QWidget(); el=QVBoxLayout(self.empty_left); el.setAlignment(Qt.AlignmentFlag.AlignCenter); el.setSpacing(8)
        for txt,style in [("🛡",f"font-size:42px;background:transparent;"),("Enter a serial number",f"font-size:15px;font-weight:700;color:{TEXT_SECONDARY};background:transparent;"),("Warranty info will appear here",f"font-size:11px;color:{TEXT_DIM};background:transparent;")]:
            l=QLabel(txt); l.setStyleSheet(style); l.setAlignment(Qt.AlignmentFlag.AlignCenter); el.addWidget(l)
        self.left_lay.addWidget(self.empty_left); self.left_lay.addStretch()
        ls.setWidget(self.left_w); wl.addWidget(ls, stretch=1)
        self.tabs.addTab(warranty_widget, "🛡  Warranty")

        # ── Tab 2: Drivers ───────────────────────────────────────────
        driver_widget = QWidget(); driver_widget.setStyleSheet(f"background:{DARK_BG};")
        rl=QVBoxLayout(driver_widget); rl.setContentsMargins(14,14,18,14); rl.setSpacing(8)
        # Driver toolbar: status + search + category
        filter_row=QHBoxLayout(); filter_row.setSpacing(8)
        self.drv_status=QLabel(""); self.drv_status.setStyleSheet(f"font-size:11px;color:{TEXT_DIM};background:transparent;"); filter_row.addWidget(self.drv_status)
        filter_row.addStretch()
        self.drv_count=QLabel(""); self.drv_count.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;background:transparent;"); filter_row.addWidget(self.drv_count)
        self.drv_search=QLineEdit(); self.drv_search.setPlaceholderText("🔍 Filter results..."); self.drv_search.setFixedHeight(CTRL_HEIGHT_SM); self.drv_search.setFixedWidth(200)
        self.drv_search.textChanged.connect(self._filter_drv); filter_row.addWidget(self.drv_search)
        self.cat_cb=QComboBox(); self.cat_cb.addItem("All Categories"); self.cat_cb.setFixedHeight(CTRL_HEIGHT_SM); self.cat_cb.setMinimumWidth(165)
        self.cat_cb.currentTextChanged.connect(self._filter_drv); filter_row.addWidget(self.cat_cb)
        rl.addLayout(filter_row)
        self.driver_tree=DriverTree(); rl.addWidget(self.driver_tree,stretch=1)
        self.tabs.addTab(driver_widget, "⬇  Drivers")

        # ── Tab 3: Tools ─────────────────────────────────────────────
        tools_widget = QWidget(); tools_widget.setStyleSheet(f"background:{DARK_BG};")
        tl = QVBoxLayout(tools_widget); tl.setContentsMargins(0,0,0,0); tl.setSpacing(0)

        # Button area (scrollable)
        tools_scroll = QScrollArea(); tools_scroll.setWidgetResizable(True)
        tools_scroll.setStyleSheet(f"QScrollArea{{border:none;background:{DARK_BG};}}")
        tools_inner = QWidget(); tools_inner.setStyleSheet(f"background:{DARK_BG};")
        self.tools_btn_lay = QGridLayout(tools_inner)
        self.tools_btn_lay.setContentsMargins(24,24,24,12); self.tools_btn_lay.setSpacing(10)
        # Make all 3 columns equal width
        for col in range(3):
            self.tools_btn_lay.setColumnStretch(col, 1)
        tools_scroll.setWidget(tools_inner)
        tl.addWidget(tools_scroll, stretch=1)

        # Output panel
        out_frame = QFrame(); out_frame.setStyleSheet(f"background:{PANEL_BG};border-top:1px solid {BORDER};")
        out_lay = QVBoxLayout(out_frame); out_lay.setContentsMargins(14,8,14,8); out_lay.setSpacing(4)
        out_hdr = QHBoxLayout()
        out_lbl = QLabel("Output"); out_lbl.setStyleSheet(f"font-size:11px;font-weight:700;color:{TEXT_DIM};background:transparent;")
        out_hdr.addWidget(out_lbl); out_hdr.addStretch()
        clear_btn = QPushButton("Clear"); clear_btn.setObjectName("ghost")
        clear_btn.setFixedHeight(36); clear_btn.setFixedWidth(80)
        out_hdr.addWidget(clear_btn)
        out_lay.addLayout(out_hdr)
        self.tools_output = QPlainTextEdit()
        self.tools_output.setReadOnly(True)
        self.tools_output.setFixedHeight(180)
        self.tools_output.setStyleSheet(f"""
            QPlainTextEdit {{
                background:{DARK_BG}; color:{TEXT_PRIMARY};
                border:1px solid {BORDER}; border-radius:6px;
                font-family: Consolas, monospace; font-size:12px;
                padding:6px;
            }}
        """)
        clear_btn.clicked.connect(self.tools_output.clear)
        out_lay.addWidget(self.tools_output)
        tl.addWidget(out_frame)

        self.tabs.addTab(tools_widget, "🔧  Tools")
        self._load_tools()

        root.addWidget(self.tabs, stretch=1)

        # Action bar
        ab=QFrame(); ab.setStyleSheet(f"background:{PANEL_BG};border-top:1px solid {BORDER};")
        al=QHBoxLayout(ab); al.setContentsMargins(20,12,20,12); al.setSpacing(8)
        self.save_pdf_btn=QPushButton("📄  Save PDF Report"); self.save_pdf_btn.setObjectName("ghost"); self.save_pdf_btn.setEnabled(False); self.save_pdf_btn.clicked.connect(self._save_pdf); al.addWidget(self.save_pdf_btn)
        self.prod_btn=QPushButton("🌐  Product Page"); self.prod_btn.setObjectName("ghost"); self.prod_btn.setEnabled(False); self.prod_btn.clicked.connect(lambda:QDesktopServices.openUrl(QUrl(self._product_url))); al.addWidget(self.prod_btn)
        self.drv_page_btn=QPushButton("📦  Driver Page"); self.drv_page_btn.setObjectName("ghost"); self.drv_page_btn.setEnabled(False); self.drv_page_btn.clicked.connect(lambda:QDesktopServices.openUrl(QUrl(self._product_url+"/downloads"))); al.addWidget(self.drv_page_btn)
        self.copy_urls_btn=QPushButton("📋  Copy Selected URLs"); self.copy_urls_btn.setObjectName("ghost"); self.copy_urls_btn.setEnabled(False); self.copy_urls_btn.clicked.connect(self._copy_selected_urls); al.addWidget(self.copy_urls_btn)
        self.dl_selected_btn=QPushButton("⬇  Download Selected"); self.dl_selected_btn.setObjectName("ghost"); self.dl_selected_btn.setEnabled(False); self.dl_selected_btn.clicked.connect(self._download_selected); al.addWidget(self.dl_selected_btn)
        al.addStretch()
        quit_btn=QPushButton("Quit"); quit_btn.setObjectName("ghost"); quit_btn.clicked.connect(QApplication.instance().quit); al.addWidget(quit_btn)
        root.addWidget(ab)

    def _load_tools(self):
        """Read tools.toml and populate the Tools tab with buttons."""
        # Clear existing buttons
        while self.tools_btn_lay.count():
            item = self.tools_btn_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        # Re-apply column stretches
        cols = 3
        for col in range(cols):
            self.tools_btn_lay.setColumnStretch(col, 1)

        toml_path = _get_tools_toml()

        if tomllib is None:
            lbl = QLabel("⚠️  TOML support not available.\nInstall tomli: pip install tomli")
            lbl.setStyleSheet(f"color:{TEXT_DIM};font-size:13px;background:transparent;")
            self.tools_btn_lay.addWidget(lbl)
            return

        if not toml_path.exists():
            # Create a sample tools.toml
            sample = '''# Lenovo Toolkit - Tools Configuration
# Add [[tool]] sections to create buttons in the Tools tab.
#
# Fields:
#   label       = "Button label"           (required)
#   command     = "command to run"         (required)
#   icon        = "🔧"                     (optional, emoji or text)
#   description = "Tooltip text"           (optional)
#   shell       = "powershell" or "cmd"    (optional, default: powershell)
#   confirm     = true                     (optional, shows confirmation dialog)

[[tool]]
label = "Flush DNS"
icon = "🌐"
description = "Clear the DNS resolver cache"
command = "ipconfig /flushdns"
shell = "cmd"

[[tool]]
label = "Clear Temp Files"
icon = "🗑"
description = "Delete temporary files from %TEMP%"
command = "Remove-Item $env:TEMP\\\\* -Recurse -Force -ErrorAction SilentlyContinue; Write-Host 'Temp files cleared.'"
shell = "powershell"
confirm = true

[[tool]]
label = "Network Info"
icon = "📡"
description = "Show IP configuration"
command = "ipconfig /all"
shell = "cmd"

[[tool]]
label = "System Info"
icon = "💻"
description = "Show basic system information"
command = "Get-ComputerInfo | Select-Object CsName, OsName, OsVersion, CsTotalPhysicalMemory | Format-List"
shell = "powershell"
'''
            toml_path.write_text(sample, encoding="utf-8")
            self.tools_output.appendPlainText(f"[tools] Created sample tools.toml at {toml_path}")

        try:
            with open(toml_path, "rb") as f:
                config = tomllib.load(f)
        except Exception as e:
            lbl = QLabel(f"⚠️  Error reading tools.toml:\n{e}")
            lbl.setStyleSheet(f"color:{ACCENT};font-size:12px;background:transparent;")
            self.tools_btn_lay.addWidget(lbl)
            return

        tools = config.get("tool", [])
        if not tools:
            lbl = QLabel("No tools defined in tools.toml yet.")
            lbl.setStyleSheet(f"color:{TEXT_DIM};font-size:13px;background:transparent;")
            self.tools_btn_lay.addWidget(lbl, 0, 0)
            return

        for i, tool in enumerate(tools):
            label   = tool.get("label", "Tool")
            icon    = tool.get("icon", "🔧")
            desc    = tool.get("description", "")

            btn = QPushButton(f"{icon}  {label}")
            btn.setObjectName("ghost")
            btn.setFixedHeight(CTRL_HEIGHT_LG)
            if desc:
                btn.setToolTip(desc)
            btn.clicked.connect(lambda _, t=tool: self._run_tool(t))
            self.tools_btn_lay.addWidget(btn, i // cols, i % cols)

        # Reload button spans full width below
        total_rows = (len(tools) + cols - 1) // cols
        reload_btn = QPushButton("↺  Reload tools.toml")
        reload_btn.setObjectName("ghost")
        reload_btn.setFixedHeight(CTRL_HEIGHT_LG)
        reload_btn.clicked.connect(self._load_tools)
        self.tools_btn_lay.addWidget(reload_btn, total_rows + 1, 0, 1, cols)

    def _run_tool(self, tool: dict):
        label      = tool.get("label", "Tool")
        command    = tool.get("command", "")
        shell      = tool.get("shell", "powershell").lower()
        confirm    = tool.get("confirm", False)
        github     = tool.get("github", "")
        asset_regex= tool.get("asset_regex", "")
        run_after  = tool.get("run_after", False)
        terminal   = tool.get("terminal", False)

        if not command and not github:
            self.tools_output.appendPlainText(f"[{label}] No command or github defined.")
            return

        if confirm:
            dlg = QDialog(self)
            dlg.setWindowTitle("Confirm")
            dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT_PRIMARY};")
            dl = QVBoxLayout(dlg); dl.setContentsMargins(20,20,20,20); dl.setSpacing(14)
            dl.addWidget(QLabel(f"Run: {label}?"))
            btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
            btns.accepted.connect(dlg.accept); btns.rejected.connect(dlg.reject)
            dl.addWidget(btns)
            if dlg.exec() != QDialog.DialogCode.Accepted:
                return

        self.tools_output.appendPlainText(f"\n▶ {label}")
        self.tabs.setCurrentIndex(2)

        def log(msg):
            QMetaObject.invokeMethod(self.tools_output, "appendPlainText",
                Qt.ConnectionType.QueuedConnection, Q_ARG(str, msg))

        if github:
            # GitHub fetch mode
            def run_github():
                try:
                    # Normalize repo
                    m = re.match(r"https://github\.com/([^/]+/[^/]+)", github)
                    repo = m.group(1) if m else github.strip("/")

                    log(f"Fetching latest release from {repo}...")
                    api = f"https://api.github.com/repos/{repo}/releases"
                    r = requests.get(api, timeout=20)
                    r.raise_for_status()

                    asset = None
                    version = None
                    for release in r.json():
                        for a in release.get("assets", []):
                            if not asset_regex or re.search(asset_regex, a["name"]):
                                asset = a
                                version = release.get("tag_name", "")
                                break
                        if asset:
                            break

                    if not asset:
                        log("[ERROR] No matching asset found.")
                        return

                    log(f"Found: {asset['name']} ({version})")

                    out_dir = _exe_dir() / "Downloads" / "Tools"
                    out_dir.mkdir(parents=True, exist_ok=True)
                    out_path = out_dir / asset["name"]

                    log(f"Downloading to {out_path}...")
                    dl = requests.get(asset["browser_download_url"], stream=True, timeout=60, headers={
                        "User-Agent": "Mozilla/5.0"
                    })
                    dl.raise_for_status()
                    total = int(dl.headers.get("content-length", 0))
                    done = 0
                    with open(out_path, "wb") as f:
                        for chunk in dl.iter_content(65536):
                            if chunk:
                                f.write(chunk)
                                done += len(chunk)
                                if total:
                                    pct = int(done / total * 100)
                                    if pct % 25 == 0:
                                        log(f"  {pct}%")

                    log(f"✅ Downloaded: {out_path}")

                    if run_after:
                        log(f"Launching {out_path.name}...")
                        import ctypes
                        ctypes.windll.shell32.ShellExecuteW(None, "runas", str(out_path), None, None, 1)

                except Exception as e:
                    log(f"[ERROR] {e}")

            threading.Thread(target=run_github, daemon=True).start()
            return

        # Normal command mode
        def run():
            try:
                if shell == "powershell":
                    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command]
                else:
                    args = ["cmd", "/c", command]

                if terminal:
                    subprocess.Popen(args, creationflags=subprocess.CREATE_NEW_CONSOLE)
                    log("(opened in terminal window)")
                else:
                    result = subprocess.run(args, capture_output=True, text=True, timeout=60)
                    output = result.stdout.strip() or result.stderr.strip() or "(no output)"
                    log(output)
            except subprocess.TimeoutExpired:
                log("[ERROR] Command timed out after 60s")
            except Exception as e:
                log(f"[ERROR] {e}")

        threading.Thread(target=run, daemon=True).start()

    def _on_tab_changed(self, index):
        if not hasattr(self, 'copy_urls_btn'): return
        is_warranty = (index == 0)
        is_drivers  = (index == 1)
        self.save_pdf_btn.setVisible(is_warranty)
        self.copy_urls_btn.setVisible(is_drivers)
        self.dl_selected_btn.setVisible(is_drivers)
        self.prod_btn.setVisible(is_warranty or is_drivers)
        self.drv_page_btn.setVisible(is_warranty or is_drivers)

    def _try_detect(self):
        from PyQt6.QtWidgets import QMessageBox
        sn = detect_local_serial()
        if sn:
            self.serial_input.setText(sn)
            self.status_lbl.setText(f"⚡ Detected: {sn}")
            msg = QMessageBox(self)
            msg.setWindowTitle("Serial Detected")
            msg.setText(f"<b>{sn}</b> was detected on this machine.")
            msg.setInformativeText("Fetch warranty info and driver URLs now?")
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg.setDefaultButton(QMessageBox.StandardButton.Yes)
            msg.setStyleSheet(f"""
                QMessageBox {{ background: {PANEL_BG}; }}
                QLabel {{ color: {TEXT_PRIMARY}; background: transparent; }}
                QPushButton {{ background: {ACCENT}; color: white; border: none; border-radius: 7px; padding: 7px 18px; font-size: 13px; font-weight: 600; min-width: 80px; }}
                QPushButton:hover {{ background: {ACCENT_SOFT}; }}
                QPushButton[text="No"] {{ background: transparent; border: 1px solid {BORDER}; color: {TEXT_SECONDARY}; }}
            """)
            if msg.exec() == QMessageBox.StandardButton.Yes:
                self._lookup()
        else:
            self.status_lbl.setText("No serial detected — enter manually.")

    def _lookup(self):
        raw=self.serial_input.text().strip()
        if not raw: self.status_lbl.setText("⚠️  Enter a serial number."); return
        serial=norm_serial(raw); self.lookup_btn.setEnabled(False); self.lookup_btn.setText("Looking up…"); self.status_lbl.setText(f"Fetching {serial}…")
        self._w_worker=WarrantyWorker(serial); self._w_thread=QThread()
        self._w_worker.moveToThread(self._w_thread); self._w_thread.started.connect(self._w_worker.run)
        self._w_worker.finished.connect(self._on_warranty); self._w_worker.error.connect(self._on_warranty_err)
        self._w_worker.finished.connect(self._w_thread.quit); self._w_worker.error.connect(self._w_thread.quit)
        self._w_thread.start()

    def _on_warranty(self,wf,spec,product_url,raw_json):
        self._wf=wf; self._spec=spec; self._product_url=product_url; self._raw_json=raw_json
        self._full_id=(_root(raw_json).get("machineInfo") or {}).get("fullId") or ""
        self.lookup_btn.setEnabled(True); self.lookup_btn.setText("Look Up"); self.status_lbl.setText("")
        self._build_left(); self.save_pdf_btn.setEnabled(True)
        self.prod_btn.setEnabled(bool(product_url)); self.drv_page_btn.setEnabled(bool(product_url))
        if self._full_id: self._fetch_drivers()

    def _on_warranty_err(self,msg):
        self.lookup_btn.setEnabled(True); self.lookup_btn.setText("Look Up"); self.status_lbl.setText(f"❌ {msg}")

    def _build_left(self):
        while self.left_lay.count():
            item=self.left_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        wf=self._wf; active=compute_active(wf.get("startDate"),wf.get("endDate"))
        title=product_title_from_name(wf.get("productName")) or wf.get("product") or "Unknown"
        badge_color=GREEN if active is True else (SEV_CRITICAL if active is False else YELLOW)
        status_text=wf.get("warrantyStatus") or ("In Warranty" if active else "Out of Warranty")

        # ── Product header ──
        self.left_lay.addWidget(self._sec_lbl("MACHINE TYPE"))
        pc=QFrame(); pc.setStyleSheet(f"background:{PANEL_BG};border-radius:10px;border:1px solid {BORDER};")
        pcl=QHBoxLayout(pc); pcl.setContentsMargins(16,14,16,14)
        iv=QVBoxLayout(); iv.setSpacing(4)
        nl=QLabel(title); nl.setStyleSheet(f"font-size:{FONT_PANEL_TITLE}px;font-weight:700;color:{TEXT_PRIMARY};background:transparent;border:none;"); nl.setWordWrap(True); iv.addWidget(nl)
        sl=QLabel(f"Serial: {wf.get('serial') or '-'}  ·  MTM: {wf.get('product') or '-'}  ·  Model: {wf.get('model') or '-'}")
        sl.setStyleSheet(f"font-size:{FONT_PANEL_SUB}px;color:{TEXT_SECONDARY};background:transparent;border:none;"); iv.addWidget(sl)
        pcl.addLayout(iv,stretch=1)
        badge=QLabel(status_text.upper()); badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        badge.setStyleSheet(f"font-size:{FONT_BADGE}px;font-weight:700;letter-spacing:0.8px;border-radius:6px;padding:6px 12px;background:{badge_color}22;color:{badge_color};border:1px solid {badge_color}55;")
        pcl.addWidget(badge); self.left_lay.addWidget(pc)

        # ── Warranty Information ──
        root_data = _root(self._raw_json) if self._raw_json else {}
        mi = root_data.get("machineInfo") or {}
        cw = root_data.get("currentWarranty") or {}
        entitlements = root_data.get("activeDeliveryTypeList") or []
        remaining_days = cw.get("remainingDays")
        eos_date = mi.get("eosDate") or "-"
        build_date = mi.get("buildDate") or "-"
        ship_date = mi.get("shipDate") or "-"
        end_color=GREEN if active is True else (SEV_CRITICAL if active is False else TEXT_PRIMARY)
        warr_pairs=[
            ("Machine Type",   wf.get("machineType") or "-"),
            ("Ship-To",        wf.get("shipToCountry") or "-"),
            ("Build Date",     build_date),
            ("Ship Date",      ship_date),
            ("Plan",           wf.get("planName") or "-"),
            ("Delivery Type",  wf.get("deliveryType") or "-"),
            ("Start Date",     wf.get("startDate") or "-"),
            ("End Date",       wf.get("endDate") or "-"),
        ]
        if remaining_days is not None and remaining_days > 0:
            warr_pairs.append(("Days Remaining", f"{remaining_days} days ({cw.get('remainingMonths',0)} months)"))
        warr_pairs.append(("End of Service",  eos_date))
        if entitlements:
            warr_pairs.append(("Entitlements", ", ".join(entitlements)))
        self.left_lay.addWidget(self._sec_lbl("WARRANTY INFORMATION"))
        self.left_lay.addWidget(self._kv(warr_pairs, {"End Date": end_color}))

        # ── Build Configuration ──
        if self._spec:
            self.left_lay.addWidget(self._sec_lbl("BUILD CONFIGURATION"))
            self.left_lay.addWidget(self._kv([(k,self._spec[k]) for k in SPEC_ORDER if k in self._spec]))
        self.left_lay.addStretch()

    def _sec_lbl(self,text):
        l=QLabel(text); l.setStyleSheet(f"font-size:{FONT_SECTION}px;font-weight:700;letter-spacing:1.1px;color:{TEXT_DIM};background:transparent;margin-top:4px;"); return l

    def _kv(self,pairs,overrides=None):
        frame=QFrame(); frame.setStyleSheet(f"background:{PANEL_BG};border-radius:8px;border:1px solid {BORDER};")
        grid=QGridLayout(frame); grid.setContentsMargins(14,10,14,10); grid.setSpacing(7); grid.setColumnStretch(1,1)
        for i,(k,v) in enumerate(pairs):
            kl=QLabel(k); kl.setStyleSheet(f"font-size:{FONT_KV_KEY}px;color:{TEXT_DIM};background:transparent;border:none;")
            vc=(overrides or {}).get(k,TEXT_PRIMARY)
            vl=QLabel(v); vl.setStyleSheet(f"font-size:{FONT_KV_VALUE}px;color:{vc};background:transparent;border:none;"); vl.setWordWrap(True)
            grid.addWidget(kl,i,0); grid.addWidget(vl,i,1)
        return frame

    def _fetch_drivers(self):
        """Fetch all drivers for this product"""
        self.drv_status.setText("⏳ Loading drivers…")
        self._d_worker = DriverFetchWorker(self._full_id)
        self._d_thread = QThread()
        self._d_worker.moveToThread(self._d_thread)
        self._d_thread.started.connect(self._d_worker.run)
        self._d_worker.finished.connect(self._on_drv)
        self._d_worker.error.connect(self._on_drv_err)
        self._d_worker.finished.connect(self._d_thread.quit)
        self._d_worker.error.connect(self._d_thread.quit)
        self._d_thread.finished.connect(self._d_thread.deleteLater)
        self._d_thread.start()

    def _on_drv(self, drivers):
        self._drivers = drivers
        self.driver_tree.populate(drivers)

        # Rebuild category combo
        cats = sorted({d.get("Category", "") for d in drivers if d.get("Category")})
        self.cat_cb.blockSignals(True)
        self.cat_cb.clear()
        self.cat_cb.addItem("All Categories")
        for c in cats:
            self.cat_cb.addItem(c)
        self.cat_cb.blockSignals(False)

        self._update_drv_count()
        self.copy_urls_btn.setEnabled(True)
        self.dl_selected_btn.setEnabled(True)
        self.drv_status.setText(f"✅ {len(drivers)} drivers found")

    def _on_drv_err(self,msg):
        self.drv_status.setText(f"❌ {msg}")

    def _filter_drv(self):
        self.driver_tree.apply_filter(self.drv_search.text(),"All Severities",self.cat_cb.currentText()); self._update_drv_count()

    def _update_drv_count(self):
        total=sum(self.driver_tree.topLevelItem(i).childCount()-sum(self.driver_tree.topLevelItem(i).child(j).isHidden() for j in range(self.driver_tree.topLevelItem(i).childCount())) for i in range(self.driver_tree.topLevelItemCount()))
        self.drv_count.setText(f"{total} files")

    def _copy_selected_urls(self):
        urls = self.driver_tree.checked_urls()
        if not urls:
            self.drv_status.setText("⚠️  No drivers selected — check some boxes first.")
            return
        QApplication.clipboard().setText("\n".join(urls))
        self.drv_status.setText(f"📋  Copied {len(urls)} URL{'s' if len(urls)!=1 else ''} to clipboard.")

    def _download_single(self, url):
        """Download a single file from a row button click."""
        self._run_downloads([url])

    def _download_selected(self):
        urls = self.driver_tree.checked_urls()
        if not urls:
            self.drv_status.setText("⚠️  No drivers selected — check some boxes first.")
            return
        self._run_downloads(urls)

    def _run_downloads(self, urls):
        # Build destination folder
        serial = (self._wf.get("serial") or "UNKNOWN").upper()
        dest   = _exe_dir() / "Downloads" / serial
        dest.mkdir(parents=True, exist_ok=True)

        # Progress dialog
        dlg = QDialog(self)
        dlg.setWindowTitle("Downloading Drivers")
        dlg.setMinimumWidth(480)
        dlg.setStyleSheet(f"background:{DARK_BG};color:{TEXT_PRIMARY};")
        dlg_lay = QVBoxLayout(dlg)
        dlg_lay.setContentsMargins(20,20,20,20)
        dlg_lay.setSpacing(10)

        file_lbl = QLabel("Starting...")
        file_lbl.setStyleSheet(f"color:{TEXT_PRIMARY};font-size:13px;background:transparent;")
        dlg_lay.addWidget(file_lbl)

        bar = QProgressBar()
        bar.setRange(0, 100)
        bar.setTextVisible(True)
        bar.setFixedHeight(18)
        bar.setStyleSheet(f"""
            QProgressBar {{ background:{PANEL_BG}; border:1px solid {BORDER}; border-radius:6px; }}
            QProgressBar::chunk {{ background:{ACCENT}; border-radius:6px; }}
        """)
        dlg_lay.addWidget(bar)

        overall_lbl = QLabel(f"0 / {len(urls)} files")
        overall_lbl.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;background:transparent;")
        dlg_lay.addWidget(overall_lbl)

        self._dl_cancel = False
        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("ghost")
        cancel_btn.clicked.connect(lambda: setattr(self, '_dl_cancel', True))
        dlg_lay.addWidget(cancel_btn)

        dlg.setModal(True)
        dlg.show()

        def do_downloads():
            for i, url in enumerate(urls, 1):
                if self._dl_cancel:
                    break
                fname = Path(url).name or f"driver_{i}"
                out_path = dest / fname
                QMetaObject.invokeMethod(file_lbl, "setText",
                    Qt.ConnectionType.QueuedConnection,
                    Q_ARG(str, f"({i}/{len(urls)}) {fname}"))
                QMetaObject.invokeMethod(bar, "setValue",
                    Qt.ConnectionType.QueuedConnection, Q_ARG(int, 0))
                try:
                    r = requests.get(url, stream=True, timeout=60, headers={
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:150.0) Gecko/20100101 Firefox/150.0"
                    })
                    r.raise_for_status()
                    total = int(r.headers.get("content-length", 0))
                    downloaded = 0
                    with open(out_path, "wb") as f:
                        for chunk in r.iter_content(chunk_size=65536):
                            if self._dl_cancel: break
                            if chunk:
                                f.write(chunk)
                                downloaded += len(chunk)
                                if total:
                                    QMetaObject.invokeMethod(bar, "setValue",
                                        Qt.ConnectionType.QueuedConnection,
                                        Q_ARG(int, int(downloaded / total * 100)))
                    QMetaObject.invokeMethod(overall_lbl, "setText",
                        Qt.ConnectionType.QueuedConnection,
                        Q_ARG(str, f"{i} / {len(urls)} files"))
                except Exception as e:
                    print(f"[download] failed {url}: {e}")

            # Signal main thread via Qt signal (thread-safe)
            self.dl_finished.emit(str(dest), len(urls), self._dl_cancel)
            QMetaObject.invokeMethod(dlg, "close", Qt.ConnectionType.QueuedConnection)

        threading.Thread(target=do_downloads, daemon=True).start()

    def _on_dl_finished(self, dest, count, cancelled):
        if cancelled:
            self.drv_status.setText("⚠️  Download cancelled.")
        else:
            self.drv_status.setText(f"✅  {count} file{'s' if count!=1 else ''} saved to Downloads/")
            QDesktopServices.openUrl(QUrl.fromLocalFile(dest))

    def _session_dir(self):
        """Return (and create) a per-session folder: Reports/SN-YYYYMMDD-HHMMSS/"""
        serial  = (self._wf.get("serial") or "UNKNOWN").upper()
        stamp   = datetime.now().strftime("%Y%m%d-%H%M%S")
        folder  = _exe_dir() / REPORTS_FOLDER / f"{serial}-{stamp}"
        folder.mkdir(parents=True, exist_ok=True)
        return folder

    def _save_pdf(self):
        if not self._wf: return

        # ── Report options dialog ──
        from PyQt6.QtWidgets import QDialog, QCheckBox, QDialogButtonBox
        dlg = QDialog(self)
        dlg.setWindowTitle("Report Options")
        dlg.setFixedWidth(320)
        dlg.setStyleSheet(f"""
            QDialog {{ background:{PANEL_BG}; }}
            QLabel  {{ color:{TEXT_PRIMARY}; background:transparent; border:none; }}
            QCheckBox {{ color:{TEXT_PRIMARY}; background:transparent; spacing:8px; font-size:13px; }}
            QCheckBox::indicator {{ width:16px; height:16px; border-radius:4px; border:1px solid {BORDER}; background:{CARD_BG}; }}
            QCheckBox::indicator:checked {{ background:{ACCENT}; border-color:{ACCENT}; }}
            QPushButton {{ background:{ACCENT}; color:white; border:none; border-radius:7px; padding:8px 20px; font-size:13px; font-weight:600; }}
            QPushButton:hover {{ background:{ACCENT_SOFT}; }}
            QPushButton[text="Cancel"] {{ background:transparent; border:1px solid {BORDER}; color:{TEXT_SECONDARY}; }}
        """)
        dl = QVBoxLayout(dlg); dl.setContentsMargins(22,20,22,20); dl.setSpacing(14)

        hdr = QLabel("What to include in the PDF?")
        hdr.setStyleSheet(f"font-size:13px;font-weight:700;color:{TEXT_PRIMARY};background:transparent;border:none;")
        dl.addWidget(hdr)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"border:none;border-top:1px solid {BORDER};background:transparent;")
        dl.addWidget(sep)

        cb_system  = QCheckBox("System Report  (warranty + build spec)")
        cb_drivers = QCheckBox("Driver List  (all loaded drivers)")
        cb_system.setChecked(True)
        cb_drivers.setChecked(bool(self._drivers))
        cb_drivers.setEnabled(bool(self._drivers))
        if not self._drivers:
            cb_drivers.setStyleSheet(f"color:{TEXT_DIM};background:transparent;spacing:8px;font-size:13px;")
        dl.addWidget(cb_system)
        dl.addWidget(cb_drivers)

        note = QLabel("At least one option must be selected.")
        note.setStyleSheet(f"font-size:11px;color:{TEXT_DIM};background:transparent;border:none;")
        dl.addWidget(note)

        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)
        dl.addWidget(btns)

        # Prevent OK if nothing checked
        def _validate():
            btns.button(QDialogButtonBox.StandardButton.Ok).setEnabled(
                cb_system.isChecked() or cb_drivers.isChecked())
        cb_system.stateChanged.connect(_validate)
        cb_drivers.stateChanged.connect(_validate)

        if dlg.exec() != QDialog.DialogCode.Accepted:
            return

        include_system  = cb_system.isChecked()
        include_drivers = cb_drivers.isChecked()

        # ── Pick save path ──
        serial  = (self._wf.get("serial") or "UNKNOWN").upper()
        machine = product_title_from_name(self._wf.get("productName") or "") or self._wf.get("product") or "machine"
        m       = re.sub(r"[^A-Za-z0-9_\-]+","",machine.strip().replace(" ","_"))
        stamp   = datetime.now().strftime("%Y%m%d-%H%M%S")
        suffix  = "_driversOnly" if (not include_system and include_drivers) else ""
        default = f"{serial}_{m}{suffix}-{stamp}.pdf"

        session_dir = self._session_dir()
        path,_ = QFileDialog.getSaveFileName(self,"Save PDF Report",str(session_dir / default),"PDF Files (*.pdf)")
        if not path: return

        os_name = "Windows 11"
        wf_arg  = self._wf   if include_system  else {}
        spec_arg= self._spec if include_system  else {}
        url_arg = self._product_url if include_system else ""
        drv_arg = self._drivers if include_drivers else None

        buf = build_pdf_report(wf_arg, spec_arg, url_arg, drivers=drv_arg, os_name=os_name)
        Path(path).write_bytes(buf.read())
        self.status_lbl.setText(f"✅ Saved: {Path(path).name}")
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

# ── Main window ───────────────────────────────────────────────────────────────
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setWindowIcon(QtGui.QIcon(str(_resource_dir() / "icon.ico")))
        self.resize(WIN_WIDTH,WIN_HEIGHT)
        self.setMinimumSize(WIN_MIN_WIDTH,WIN_MIN_HEIGHT)
        self.setStyleSheet(STYLESHEET)
        self._build()

    def _build(self):
        # ── Menu bar ──
        menubar = self.menuBar()
        menubar.setStyleSheet(f"""
            QMenuBar {{
                background: {PANEL_BG};
                color: {TEXT_SECONDARY};
                font-size: 13px;
                padding: 2px 4px;
                border-bottom: 1px solid {BORDER};
            }}
            QMenuBar::item {{
                background: transparent;
                padding: 6px 12px;
                border-radius: 4px;
            }}
            QMenuBar::item:selected {{
                background: {CARD_BG};
                color: {TEXT_PRIMARY};
            }}
            QMenu {{
                background: {CARD_BG};
                color: {TEXT_PRIMARY};
                border: 1px solid {BORDER};
                border-radius: 6px;
                padding: 4px;
                font-size: 13px;
            }}
            QMenu::item {{
                padding: 8px 24px 8px 14px;
                border-radius: 4px;
            }}
            QMenu::item:selected {{
                background: {ACCENT};
                color: white;
            }}
            QMenu::separator {{
                height: 1px;
                background: {BORDER};
                margin: 4px 8px;
            }}
        """)

        # File menu
        file_menu = menubar.addMenu("  File  ")

        act_save = QAction("💾  Save PDF Report", self)
        act_save.setShortcut("Ctrl+S")
        act_save.triggered.connect(self._save_report)
        file_menu.addAction(act_save)

        file_menu.addSeparator()

        act_exit = QAction("Exit", self)
        act_exit.setShortcut("Ctrl+Q")
        act_exit.triggered.connect(self.close)
        file_menu.addAction(act_exit)

        # Help menu
        help_menu = menubar.addMenu("  Help  ")

        act_about = QAction(f"About {APP_NAME}", self)
        act_about.triggered.connect(self._show_about)
        help_menu.addAction(act_about)

        act_lenovo = QAction("Lenovo Support Site", self)
        act_lenovo.triggered.connect(lambda: QDesktopServices.openUrl(QUrl("https://pcsupport.lenovo.com")))
        help_menu.addAction(act_lenovo)

        # ── Main content ──
        central = QWidget()
        self.setCentralWidget(central)

        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        topbar = QFrame()
        topbar.setFixedHeight(TOPBAR_HEIGHT)
        topbar.setStyleSheet(f"background:{PANEL_BG};border-bottom:1px solid {BORDER};")

        tbl = QHBoxLayout(topbar)
        tbl.setContentsMargins(24, 0, 24, 0)

        logo = QLabel()
        logo.setTextFormat(Qt.TextFormat.RichText)
        logo.setText(
            f"<span style='font-size:{FONT_LOGO}px;font-weight:800;letter-spacing:2px;color:{TEXT_PRIMARY}'>{APP_TITLE}</span>"
            f"<span style='font-size:{FONT_LOGO}px;font-weight:800;letter-spacing:2px;color:{ACCENT_SOFT}'> {APP_SUBTITLE}</span>"
        )

        tbl.addStretch()
        tbl.addWidget(logo)
        tbl.addStretch()

        root.addWidget(topbar)
        root.addSpacing(3)
        self.main = WarrantyTab()
        root.addWidget(self.main, stretch=1)

    def _save_report(self):
        """Proxy so menu File > Save works even when focus is on main panel."""
        if hasattr(self.main, '_save_pdf'):
            self.main._save_pdf()

    def _show_about(self):
        from PyQt6.QtWidgets import QDialog
        dlg = QDialog(self)
        dlg.setWindowTitle("About")
        dlg.setFixedSize(445, 510)
        dlg.setStyleSheet(f"""
            QDialog {{ background: {PANEL_BG}; }}
            QLabel  {{ background: transparent; border: none; }}
        """)
        lay = QVBoxLayout(dlg); lay.setContentsMargins(28, 28, 28, 24); lay.setSpacing(10)
        title = QLabel()
        title.setTextFormat(Qt.TextFormat.RichText)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        title.setText(
            f"<span style='font-size:20px;font-weight:800;color:{TEXT_PRIMARY}'>{APP_TITLE}</span>"
            #f"<span style='font-size:20px;font-weight:800;color:{ACCENT_SOFT}'>{APP_SUBTITLE}</span>"
        )

        lay.addStretch()
        lay.addWidget(title)
        lay.addStretch()
        ver = QLabel(f"Version {APP_VERSION}")
        ver.setStyleSheet(f"font-size:12px;color:{TEXT_DIM};")
        lay.addWidget(ver)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"border:none;border-top:1px solid {BORDER};")
        lay.addWidget(sep)

        desc = QLabel("Warranty lookup, build configuration, and driver URL harvesting for Lenovo devices.\n\nUses the official Lenovo pcsupport API.\n\nyou're never going to be able to plan your build perfectly in your mind in one shot in the mental. Sometimes you just gotta start building and see what ideas god bestows upon you...\n\nWhat do you call it when you meditate on the toilet while thinking of your trunk layout plans? Medaitaint.")
        desc.setWordWrap(True)
        desc.setStyleSheet(f"font-size:13px;color:{TEXT_SECONDARY};line-height:1.5;")
        lay.addWidget(desc)

        lay.addStretch()

        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(36)
        close_btn.clicked.connect(dlg.accept)
        lay.addWidget(close_btn)

        dlg.exec()

def main():
    app=QApplication(sys.argv); app.setStyle("Fusion")
    p=QPalette()
    p.setColor(QPalette.ColorRole.Window,QColor(DARK_BG)); p.setColor(QPalette.ColorRole.WindowText,QColor(TEXT_PRIMARY))
    p.setColor(QPalette.ColorRole.Base,QColor(PANEL_BG)); p.setColor(QPalette.ColorRole.AlternateBase,QColor(CARD_BG))
    p.setColor(QPalette.ColorRole.Text,QColor(TEXT_PRIMARY)); p.setColor(QPalette.ColorRole.Button,QColor(PANEL_BG))
    p.setColor(QPalette.ColorRole.ButtonText,QColor(TEXT_PRIMARY)); p.setColor(QPalette.ColorRole.Highlight,QColor(ACCENT))
    p.setColor(QPalette.ColorRole.HighlightedText,QColor("#ffffff")); app.setPalette(p)

    def launch_main():
        win = MainWindow()
        app._main_win = win   # keep reference
        win.show()

    splash = SplashScreen()
    app._splash = splash      # keep reference so GC doesn't destroy it
    splash.show()
    splash.run_steps(on_complete=launch_main)
    sys.exit(app.exec())

if __name__=="__main__":
    main()