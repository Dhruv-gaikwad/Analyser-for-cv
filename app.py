"""AI Resume / CV Analyzer - web version.

Same colours, layout and wording as the desktop (Tkinter) app.
All reading / analysis / export logic comes from analyzer_core.py.
"""
import html
import os
import re
import tempfile
from pathlib import Path

import streamlit as st

import analyzer_core as core

# ------------------------------------------------------------
# THEME  (copied from the desktop app's Theme class)
# ------------------------------------------------------------
BG, CARD, TEXT, MUTED, BORDER = "#F4F5FB", "#FFFFFF", "#1E2140", "#6C7293", "#E2E5F3"
PRIMARY, PRIMARY_DARK, PRIMARY_SOFT, ACCENT = "#4F46E5", "#3F38C4", "#EEF0FF", "#06B6D4"
SUCCESS, WARNING, DANGER = "#10B981", "#F59E0B", "#EF4444"
HEAD_FROM, HEAD_TO, TRACK = "#2A2784", "#6D4AE8", "#EDEFF8"

CATEGORY_COLORS = {
    "Programming & Technology": "#4F46E5",
    "Data & Analytics": "#0EA5E9",
    "Digital Marketing": "#EC4899",
    "Finance & Accounting": "#10B981",
    "Business & Management": "#F59E0B",
    "Human Resources": "#8B5CF6",
    "Soft Skills": "#64748B",
}
AVATAR_COLORS = ["#4F46E5", "#0EA5E9", "#EC4899", "#10B981", "#F59E0B", "#8B5CF6", "#EF4444", "#14B8A6"]

st.set_page_config(page_title=core.APP_TITLE, page_icon="📄", layout="wide",
                   initial_sidebar_state="collapsed")

CSS = f"""
<style>
html, body, [class*="css"], .stApp {{
  font-family: "Segoe UI", "Inter", system-ui, -apple-system, Roboto, Arial, sans-serif;
  color: {TEXT};
}}
.stApp {{ background: {BG}; }}
header[data-testid="stHeader"], #MainMenu, footer, [data-testid="stToolbar"],
[data-testid="stDecoration"] {{ display: none !important; }}
.block-container {{ padding: 0 22px 24px 22px !important; max-width: 100% !important; }}

/* ---------- gradient header ---------- */
.hdr {{ margin: 0 -22px 0 -22px; padding: 26px 34px; display:flex; align-items:center; gap:18px;
  background: linear-gradient(90deg, {HEAD_FROM}, {HEAD_TO}); color:#fff; }}
.hdr .logo {{ width:62px; height:62px; border-radius:50%; background:rgba(255,255,255,.16);
  display:flex; align-items:center; justify-content:center; font-size:30px; position:relative; flex:0 0 auto; }}
.hdr .logo .ok {{ position:absolute; right:-2px; bottom:-2px; width:22px; height:22px; border-radius:50%;
  background:{SUCCESS}; border:2px solid #fff; font-size:13px; line-height:18px; text-align:center; font-weight:700; }}
.hdr h1 {{ margin:0; font-size:27px; font-weight:700; color:#fff; letter-spacing:.2px; padding:0; }}
.hdr p {{ margin:4px 0 0 0; font-size:14px; color:#D9DCFF; }}
.hdr .chips {{ margin-left:auto; display:flex; gap:10px; flex-wrap:wrap; justify-content:flex-end; }}
.hdr .chip {{ background:rgba(255,255,255,.16); border-radius:20px; padding:7px 16px; font-size:13px; color:#E4E6FF; }}
.hdr .chip b {{ color:#fff; font-size:14px; margin-right:6px; }}
.strip {{ height:4px; margin: 0 -22px 14px -22px; background:{TRACK}; overflow:hidden; }}
.strip.on i {{ display:block; height:4px; width:35%; background:linear-gradient(90deg,{PRIMARY},{ACCENT});
  animation: slide 1.2s infinite ease-in-out; }}
@keyframes slide {{ 0% {{margin-left:-35%}} 100% {{margin-left:100%}} }}

/* ---------- cards ---------- */
.card {{ background:{CARD}; border:1px solid {BORDER}; border-radius:12px; padding:16px 18px; }}
.card h3 {{ margin:0; font-size:16px; font-weight:700; color:{TEXT}; }}
.card .sub {{ font-size:12px; color:{MUTED}; margin:2px 0 12px 0; }}
.kpi {{ background:{CARD}; border:1px solid {BORDER}; border-radius:12px; padding:14px 18px 14px 20px; position:relative; overflow:hidden; }}
.kpi::before {{ content:""; position:absolute; left:0; top:0; bottom:0; width:5px; background:var(--c); }}
.kpi .cap {{ font-size:11px; font-weight:700; color:{MUTED}; letter-spacing:.8px; text-transform:uppercase; }}
.kpi .val {{ font-size:30px; font-weight:700; color:{TEXT}; line-height:1.25; }}
.kpi .val.s {{ font-size:19px; padding:7px 0; }}

/* ---------- bar charts ---------- */
.bar {{ display:flex; align-items:center; gap:10px; height:30px; }}
.bar .lab {{ width:var(--lw,172px); flex:0 0 var(--lw,172px); font-size:12.5px; color:{TEXT}; white-space:nowrap;
  overflow:hidden; text-overflow:ellipsis; text-align:right; }}
.bar .track {{ flex:1; height:12px; background:{TRACK}; border-radius:6px; overflow:hidden; }}
.bar .fill {{ height:100%; border-radius:6px; background:var(--c,{PRIMARY}); }}
.bar .num {{ width:48px; font-size:12px; font-weight:700; color:{MUTED}; }}

/* ---------- chips ---------- */
.chips2 {{ display:flex; flex-wrap:wrap; gap:7px; }}
.c2 {{ padding:4px 12px; border-radius:14px; font-size:12.5px; background:var(--bg); color:var(--fg); }}
.muted {{ color:{MUTED}; font-size:13px; }}
.dot {{ display:inline-block; width:12px; height:12px; border-radius:50%; background:var(--c); margin-right:7px; vertical-align:-1px; }}

/* ---------- cv list ---------- */
.cvc {{ display:flex; align-items:center; gap:12px; }}
.av {{ width:42px; height:42px; border-radius:50%; flex:0 0 42px; color:#fff; font-weight:700; font-size:14px;
  display:flex; align-items:center; justify-content:center; background:var(--c); }}
.cvc .t {{ font-size:14px; font-weight:700; color:{TEXT}; line-height:1.2; }}
.cvc .s {{ font-size:12px; color:{MUTED}; }}
.pbar {{ height:5px; background:{TRACK}; border-radius:3px; margin-top:6px; overflow:hidden; }}
.pbar i {{ display:block; height:100%; background:var(--c); }}
.badge {{ background:{PRIMARY_SOFT}; color:{PRIMARY_DARK}; font-size:12px; font-weight:700; border-radius:10px; padding:1px 9px; margin-left:8px; }}
div[class*="st-key-cv_"] {{ background:{CARD}; border:1px solid {BORDER}; border-radius:12px; padding:10px 12px; margin-bottom:8px; }}
div[class*="st-key-cv_"].sel {{ border-color:{PRIMARY}; }}
div[class*="st-key-cvsel_"] {{ background:{CARD}; border:2px solid {PRIMARY}; border-radius:12px; padding:10px 12px; margin-bottom:8px; }}

/* ---------- table ---------- */
.tw {{ overflow-x:auto; }}
table.t {{ width:100%; border-collapse:collapse; font-size:13.5px; }}
table.t th {{ background:{PRIMARY_SOFT}; color:{PRIMARY_DARK}; font-size:12px; font-weight:700; padding:11px 12px; text-align:left; white-space:nowrap; }}
table.t td {{ padding:9px 12px; border-bottom:1px solid #F0F1FA; color:{TEXT}; }}
table.t tr:nth-child(odd) td {{ background:#F8F9FE; }}
table.t .c {{ text-align:center; }}

/* ---------- skill matrix ---------- */
table.m {{ border-collapse:separate; border-spacing:3px; font-size:12px; }}
table.m th.sk {{ writing-mode:vertical-rl; transform:rotate(180deg); font-weight:600; color:{MUTED}; padding:4px 0; height:120px; white-space:nowrap; font-size:11.5px; }}
table.m td.nm {{ padding-right:12px; white-space:nowrap; font-weight:600; color:{TEXT}; text-align:right; }}
table.m td.cell {{ width:24px; height:24px; border-radius:6px; background:{TRACK}; }}

/* ---------- detail ---------- */
.dh {{ display:flex; align-items:center; gap:18px; }}
.dh .av {{ width:72px; height:72px; flex:0 0 72px; font-size:24px; }}
.dh h2 {{ margin:0; font-size:25px; font-weight:700; color:{TEXT}; padding:0; }}
.pill {{ display:inline-block; background:{PRIMARY_SOFT}; color:{PRIMARY_DARK}; font-size:12.5px; padding:3px 11px; border-radius:6px; margin:7px 7px 0 0; }}
.stats {{ margin-top:10px; font-size:13px; color:{MUTED}; }}
.stats b {{ color:{PRIMARY}; font-size:20px; margin-left:0; }}
.stats span {{ margin-right:22px; }}
.rings {{ margin-left:auto; display:flex; gap:20px; flex:0 0 auto; }}
.ring {{ width:108px; height:108px; border-radius:50%; flex:0 0 108px; position:relative;
  background: conic-gradient(var(--c) calc(var(--p)*1%), {TRACK} 0); }}
.ring::after {{ content:""; position:absolute; inset:12px; background:{CARD}; border-radius:50%; }}
.ring span {{ position:absolute; inset:0; display:flex; flex-direction:column; align-items:center; justify-content:center; z-index:2; }}
.ring b {{ font-size:22px; color:{TEXT}; line-height:1; }}
.ring small {{ font-size:11px; color:{MUTED}; }}
.dt {{ font-size:19px; font-weight:700; color:{TEXT}; margin:2px 0 2px 0; }}
.dm {{ font-size:13px; color:{MUTED}; margin-bottom:10px; }}
.sec {{ font-size:14px; font-weight:700; color:{PRIMARY_DARK}; margin:16px 0 5px 0; }}
.fld {{ font-size:14px; margin:3px 0; }} .fld b {{ color:{TEXT}; }}
.num {{ color:{PRIMARY_DARK}; font-weight:700; }}
.li {{ font-size:14px; margin:4px 0; }}
.metric {{ font-size:16px; font-weight:700; color:{PRIMARY_DARK}; margin:2px 0 6px 0; }}
.role {{ font-size:15px; font-weight:700; color:{TEXT}; margin-top:12px; }}
.subh {{ font-size:12.5px; font-style:italic; color:{MUTED}; margin-bottom:3px; }}
.okk {{ color:{SUCCESS}; font-weight:700; }} .bad {{ color:#9AA0C8; }}

/* ---------- empty state ---------- */
.empty {{ text-align:center; padding:70px 20px; }}
.empty .ic {{ width:76px; height:76px; border-radius:50%; background:{PRIMARY_SOFT}; margin:0 auto 16px auto;
  display:flex; align-items:center; justify-content:center; font-size:34px; }}
.empty h3 {{ font-size:20px; margin:0 0 6px 0; color:{TEXT}; }}
.empty p {{ color:{MUTED}; font-size:14px; max-width:520px; margin:0 auto; }}

/* ---------- streamlit widgets ---------- */
[data-testid="stFileUploader"] section {{ background:{CARD}; border:2px dashed #C5CAF0; border-radius:14px; padding:22px 16px; }}
[data-testid="stFileUploader"] section:hover {{ border-color:{PRIMARY}; background:{PRIMARY_SOFT}; }}
[data-testid="stFileUploader"] small {{ color:{MUTED}; }}
.stButton > button, .stDownloadButton > button {{ border-radius:10px; border:1px solid {BORDER}; background:{CARD}; color:{PRIMARY_DARK}; font-weight:600; }}
.stButton > button:hover, .stDownloadButton > button:hover {{ border-color:{PRIMARY}; color:{PRIMARY}; background:{PRIMARY_SOFT}; }}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{ background:{PRIMARY}; color:#fff; border-color:{PRIMARY}; padding:.55rem 1.6rem; }}
.stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {{ background:{PRIMARY_DARK}; color:#fff; }}
button[data-baseweb="tab"] {{ font-weight:600; color:{MUTED}; }}
button[data-baseweb="tab"][aria-selected="true"] {{ color:{PRIMARY}; }}
[data-baseweb="tab-highlight"] {{ background:{PRIMARY} !important; }}
div[data-testid="stTextInput"] input {{ border-radius:10px; background:{CARD}; border:1px solid {BORDER}; }}
div[class*="st-key-modebar"] [data-testid="stSegmentedControl"] button, div[class*="st-key-modebar"] label {{ border-radius:22px; }}
.foot {{ font-size:12px; color:{MUTED}; display:flex; justify-content:space-between; margin-top:14px; }}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

CSS2 = f"""
<style>
[data-testid="stMainBlockContainer"], .block-container {{ padding-top: 0 !important; }}
[data-testid="stAppViewContainer"] > .main {{ padding-top: 0 !important; }}
[data-testid="stVerticalBlock"] {{ gap: .7rem; }}
div[data-testid="stElementContainer"]:has(style) {{ display:none; }}
.cvc {{ margin-bottom:16px; }}
h1 a, h2 a, h3 a, [data-testid="stHeaderActionElements"] {{ display:none !important; }}

/* header like the desktop app */
.hdr {{ position:relative; overflow:hidden; padding:30px 34px; gap:20px; }}
.hdr::before {{ content:""; position:absolute; width:260px; height:260px; border-radius:50%;
  background:rgba(255,255,255,.06); left:44%; top:-170px; }}
.hdr::after {{ content:""; position:absolute; width:340px; height:340px; border-radius:50%;
  background:rgba(255,255,255,.08); right:-30px; top:30px; }}
.hdr > * {{ position:relative; z-index:1; }}
.hdr .logo {{ width:68px; height:68px; border-radius:12px; background:#fff; box-shadow:0 4px 14px rgba(0,0,0,.18); }}
.hdr .logo .ok {{ right:-7px; bottom:-7px; width:24px; height:24px; line-height:20px; font-size:14px; }}
.hdr .chip {{ border-radius:9px; background:rgba(255,255,255,.14); padding:10px 16px; }}
.strip {{ margin-bottom: 18px; }}

/* drop card = the uploader itself */
.st-key-dropcard [data-testid="stFileUploaderDropzone"] {{
  height:190px; padding:0; border:1.5px dotted #C5CAF0; border-radius:14px; background-color:{CARD};
  background-image:url("data:image/svg+xml,DROPSVG"); background-repeat:no-repeat; background-position:center; background-size:contain;
  position:relative; }}
.st-key-dropcard [data-testid="stFileUploaderDropzone"]:hover {{ border-color:{PRIMARY}; background-color:{PRIMARY_SOFT}; }}
.st-key-dropcard [data-testid="stFileUploaderDropzoneInstructions"] {{ display:none; }}
.st-key-dropcard [data-testid="stFileUploaderDropzone"] button {{ position:absolute; inset:0; width:100%; height:100%; opacity:0; cursor:pointer; }}

/* big indigo "+ Add CVs" button (also an uploader) */
.st-key-addbtn [data-testid="stFileUploaderDropzone"] {{ padding:0; border:none; background:transparent; }}
.st-key-addbtn [data-testid="stFileUploaderDropzoneInstructions"] {{ display:none; }}
.st-key-addbtn [data-testid="stFileUploaderDropzone"] button {{ width:100%; height:48px; background:{PRIMARY}; border:none;
  border-radius:10px; font-size:0; color:#fff; }}
.st-key-addbtn [data-testid="stFileUploaderDropzone"] button::after {{ content:"+  Add CVs"; font-size:15px; font-weight:700; color:#fff; }}
.st-key-addbtn [data-testid="stFileUploaderDropzone"] button > * {{ display:none; }}
.st-key-addbtn [data-testid="stFileUploaderDropzone"] button:hover {{ background:{PRIMARY_DARK}; }}
.st-key-addbtn [data-testid="stFileUploader"] > div:last-child {{ display:none; }}
.st-key-folderbtn [data-testid="stPopoverButton"] {{ width:100%; height:48px; background:{PRIMARY_SOFT}; color:{PRIMARY_DARK};
  border:none; border-radius:10px; font-weight:700; justify-content:center; }}
.st-key-folderbtn [data-testid="stPopoverButton"] [data-testid="stIconMaterial"] {{ display:none; }}

/* list card */
.st-key-listcard {{ background:{CARD}; border:1px solid {BORDER}; border-radius:12px; padding:14px 16px 10px 16px; min-height:520px; }}
.st-key-rerun button, .st-key-clearall button {{ border:none !important; background:transparent !important; box-shadow:none;
  padding:0; min-height:0; font-size:13px; font-weight:700; color:{PRIMARY}; }}
.st-key-clearall button {{ color:{MUTED}; }}
.st-key-clearall button:hover {{ color:{DANGER}; }}
.st-key-listcard .hint {{ text-align:center; color:{MUTED}; font-size:13.5px; padding:46px 8px; line-height:1.7; }}

/* pill tabs */
.st-key-modebar [data-testid="stButtonGroup"] > div[role="radiogroup"] {{ background:#E9ECF8; border-radius:12px; padding:4px; gap:2px; display:inline-flex; }}
.st-key-modebar button[data-variant="segmented_control"] {{ background:transparent; border:none; border-radius:9px; padding:9px 24px; box-shadow:none; }}
.st-key-modebar button[data-variant="segmented_control"] p {{ color:#33385f; font-weight:700; font-size:14px; }}
.st-key-modebar button[data-variant="segmented_control"][aria-checked="true"] {{ background:{PRIMARY}; }}
.st-key-modebar button[data-variant="segmented_control"][aria-checked="true"] p {{ color:#fff; }}

/* download button */
.st-key-dl_on [data-testid="stPopoverButton"], .st-key-dl_off [data-testid="stPopoverButton"] {{ height:46px; border:none; border-radius:10px; font-weight:700; justify-content:center; }}
.st-key-dl_on [data-testid="stPopoverButton"] {{ background:{PRIMARY}; color:#fff; }}
.st-key-dl_on [data-testid="stPopoverButton"] p {{ color:#fff; font-weight:700; }}
.st-key-dl_off [data-testid="stPopoverButton"] {{ background:#E9EBF5; color:#A3A9C9; }}
.st-key-dl_off [data-testid="stPopoverButton"] p {{ color:#A3A9C9; font-weight:700; }}
.st-key-dl_on [data-testid="stIconMaterial"], .st-key-dl_off [data-testid="stIconMaterial"] {{ display:none; }}

.foot2 {{ display:flex; justify-content:space-between; font-size:12.5px; color:{MUTED}; margin-top:6px; }}
.foot2 .r {{ color:#9AA0C8; font-size:12px; }}
</style>
"""
DROP_SVG = (
    "%3Csvg xmlns='http://www.w3.org/2000/svg' width='420' height='190' viewBox='0 0 420 190'%3E"
    "%3Ccircle cx='210' cy='46' r='28' fill='%23EEF0FF'/%3E"
    "%3Cpath d='M210 33 v26 M199 44 l11 -11 l11 11' stroke='%234F46E5' stroke-width='4' fill='none' stroke-linecap='round' stroke-linejoin='round'/%3E"
    "%3Cpath d='M197 66 h26' stroke='%234F46E5' stroke-width='4' stroke-linecap='round'/%3E"
    "%3Ctext x='210' y='108' text-anchor='middle' font-family='Segoe UI,Arial,sans-serif' font-size='19' font-weight='700' fill='%231E2140'%3EAdd your CVs%3C/text%3E"
    "%3Ctext x='210' y='132' text-anchor='middle' font-family='Segoe UI,Arial,sans-serif' font-size='13.5' fill='%236C7293'%3EClick to browse — or drop files here%3C/text%3E"
    "%3Ctext x='210' y='154' text-anchor='middle' font-family='Segoe UI,Arial,sans-serif' font-size='12' fill='%239AA0C8'%3EPDF · DOCX · DOC · RTF · ODT · TXT · PNG · JPG%3C/text%3E"
    "%3C/svg%3E"
)
st.markdown(CSS2.replace("DROPSVG", DROP_SVG), unsafe_allow_html=True)

CSS3 = f"""
<style>
/* tab strip like the desktop app: bold dark text, indigo active tab + underline.
   Works with both old (button) and new (div) Streamlit tab markup, and with any theme colour. */
[role="tablist"] {{ gap:8px; border-bottom:2px solid {BORDER}; }}
[data-baseweb="tab-border"] {{ background:{BORDER} !important; }}
[data-baseweb="tab-highlight"], .react-aria-SelectionIndicator {{ background:{PRIMARY} !important; background-color:{PRIMARY} !important; border-color:{PRIMARY} !important; height:3px !important; }}
[role="tab"] {{ background:transparent !important; padding:10px 18px; color:#33385F !important; cursor:pointer; }}
[role="tab"] p, [role="tab"] div, [role="tab"] span {{ color:inherit !important; font-weight:700 !important; font-size:15px !important; }}
[role="tab"][aria-selected="true"] {{ color:{PRIMARY} !important; }}
[role="tab"]:hover {{ color:{PRIMARY} !important; background:{PRIMARY_SOFT} !important; }}

/* target skills: sits on the right of the tab strip, label beside the box */
.st-key-combwrap {{ position:relative; }}
.st-key-targetbox {{ position:absolute; right:0; top:-2px; width:min(430px, 46%); z-index:20; }}
.st-key-targetbox [data-testid="stHorizontalBlock"] {{ align-items:center; gap:.6rem; }}
.tlab {{ font-weight:700; font-size:13px; color:{MUTED}; text-align:right; white-space:nowrap; }}
.st-key-targetbox input {{ height:40px; }}
div[data-testid="stTextInput"] [data-baseweb="input"], div[data-testid="stTextInput"] [data-baseweb="base-input"] {{ border-radius:10px; border:1px solid {BORDER} !important; background:{CARD}; }}
div[data-testid="stTextInput"] [data-baseweb="input"]:focus-within {{ border-color:{PRIMARY} !important; box-shadow:0 0 0 2px {PRIMARY_SOFT}; }}
@media (max-width: 900px) {{
  .st-key-targetbox {{ position:static; width:100%; margin-bottom:8px; }}
}}
</style>
"""
st.markdown(CSS3, unsafe_allow_html=True)

EMPTY_SVG = """<svg width="250" height="190" viewBox="0 0 250 190" xmlns="http://www.w3.org/2000/svg">
<ellipse cx="125" cy="170" rx="62" ry="8" fill="#E4E6F3"/>
<rect x="38" y="22" width="130" height="140" rx="8" fill="#E7E9FA"/>
<rect x="50" y="14" width="136" height="146" rx="9" fill="#fff" stroke="#DADDF2" stroke-width="1.5"/>
<circle cx="78" cy="48" r="16" fill="#EEF0FF"/><circle cx="78" cy="44" r="5.5" fill="#4F46E5"/>
<path d="M67 58 q11 -13 22 0 z" fill="#4F46E5"/>
<rect x="103" y="36" width="58" height="7" rx="3.5" fill="#4F46E5" opacity=".75"/>
<rect x="103" y="49" width="34" height="6" rx="3" fill="#DADDF2"/>
<rect x="62" y="80" width="112" height="6" rx="3" fill="#E2E5F3"/>
<rect x="62" y="95" width="112" height="6" rx="3" fill="#E2E5F3"/>
<rect x="62" y="110" width="112" height="6" rx="3" fill="#E2E5F3"/>
<rect x="62" y="125" width="70" height="6" rx="3" fill="#E2E5F3"/>
<rect x="62" y="141" width="34" height="9" rx="4" fill="#E7E9FA"/><rect x="100" y="141" width="44" height="9" rx="4" fill="#CFF3E6"/>
<circle cx="190" cy="128" r="25" fill="#4F46E5" stroke="#fff" stroke-width="4"/>
<path d="M190 116 v24 M178 128 h24" stroke="#fff" stroke-width="5" stroke-linecap="round"/></svg>"""


# ------------------------------------------------------------
# small helpers (same behaviour as the desktop app)
# ------------------------------------------------------------
def esc(x):
    return html.escape(str(x if x is not None else ""))


def shorten(text, limit):
    text = text or ""
    return text if len(text) <= limit else text[:limit - 1].rstrip() + "…"


def initials(text):
    words = re.findall(r"[A-Za-zÀ-ÿ]+", text or "")
    if not words:
        return "?"
    if len(words) == 1:
        return words[0][:2].upper()
    return (words[0][0] + words[-1][0]).upper()


def avatar_color(key):
    return AVATAR_COLORS[sum(ord(ch) for ch in (key or "?")) % len(AVATAR_COLORS)]


def level_color(p):
    return SUCCESS if p >= 75 else WARNING if p >= 50 else DANGER


def pct_text(v):
    t = f"{v + 1e-9:.1f}"
    return (t[:-2] if t.endswith(".0") else t) + "%"


def _mix(c1, c2, t):
    a = [int(c1.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    b = [int(c2.lstrip("#")[i:i + 2], 16) for i in (0, 2, 4)]
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def chips(items, bg, fg):
    if not items:
        return ""
    return '<div class="chips2">' + "".join(
        f'<span class="c2" style="--bg:{bg};--fg:{fg}">{esc(i)}</span>' for i in items) + "</div>"


def bar_rows(items, color=PRIMARY, label_w=172, fmt=None, color_fn=None, max_value=None):
    if not items:
        return '<div class="muted">Nothing to show yet.</div>'
    top = max_value or max(v for _l, v in items) or 1
    out = []
    for label, value in items:
        c = color_fn(value) if color_fn else color
        w = max(2, min(100, value / top * 100))
        out.append(
            f'<div class="bar" style="--lw:{label_w}px"><div class="lab" title="{esc(label)}">{esc(label)}</div>'
            f'<div class="track"><div class="fill" style="width:{w:.1f}%;--c:{c}"></div></div>'
            f'<div class="num">{esc(fmt(value) if fmt else value)}</div></div>')
    return "".join(out)


def card(title, subtitle, body):
    sub = f'<div class="sub">{esc(subtitle)}</div>' if subtitle else ""
    return f'<div class="card"><h3>{esc(title)}</h3>{sub}{body}</div>'


def empty_state(title, text, icon="➕"):
    art = EMPTY_SVG if icon == "➕" else f'<div class="ic">{icon}</div>'
    st.markdown(f'<div class="empty" style="padding-top:60px">{art}<h3 style="font-size:24px;margin-top:12px">{esc(title)}</h3>'
                f'<p>{esc(text)}</p></div>', unsafe_allow_html=True)


# ------------------------------------------------------------
# state
# ------------------------------------------------------------
S = st.session_state
S.setdefault("entries", [])
S.setdefault("next_id", 1)
S.setdefault("selected", None)
S.setdefault("uploader_key", 0)
S.setdefault("notice", None)


SUPPORTED = {".pdf", ".docx", ".doc", ".rtf", ".odt", ".txt", ".png", ".jpg", ".jpeg",
             ".pptx", ".xlsx", ".xls", ".csv"}


def run_analysis(entry):
    suffix = Path(entry["name"]).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(entry["data"])
        path = tmp.name
    try:
        analysis = core.analyze_resume(path)
        analysis["_id"] = entry["id"]
        analysis["filename"] = entry["name"]
        analysis["file"] = entry["name"]
        entry.update(status="done", analysis=analysis, error=None)
    except Exception as error:  # shown on the red card, like the desktop app
        entry.update(status="error", analysis=None, error=str(error))
    finally:
        try:
            os.unlink(path)
        except OSError:
            pass


def analyse_uploads(items):
    """items: list of (file name, bytes)."""
    known = {e["sig"] for e in S.entries}
    added = dup = 0
    for name, data in items:
        sig = (name, len(data))
        if sig in known:
            dup += 1
            continue
        known.add(sig)
        entry = {"id": S.next_id, "sig": sig, "name": name, "data": data,
                 "status": "queued", "analysis": None, "error": None}
        S.next_id += 1
        run_analysis(entry)
        S.entries.append(entry)
        added += 1
        if S.selected is None and entry["status"] == "done":
            S.selected = entry["id"]
    msg = f"Added {added} CV{'s' if added != 1 else ''} — analysis finished."
    if dup:
        msg += f"  ({dup} already in the list)"
    S.notice = ("success" if added else "warning",
                msg if added else "Those CVs are already in your list.")


def unzip_items(zf_file):
    import zipfile
    items = []
    with zipfile.ZipFile(zf_file) as z:
        for info in z.infolist():
            if info.is_dir() or Path(info.filename).name.startswith((".", "~$")):
                continue
            if Path(info.filename).suffix.lower() in SUPPORTED:
                items.append((Path(info.filename).name, z.read(info)))
    return items


def get_entry(eid):
    return next((e for e in S.entries if e["id"] == eid), None)


done = [e for e in S.entries if e["status"] == "done" and e["analysis"]]
required = core.parse_required_skills(S.get("target", "")) if S.get("target") else []


def ats(analysis):
    """ATS estimate for one CV (uses the target skills when typed)."""
    return core.calculate_ats(analysis, required)


# ------------------------------------------------------------
# header
# ------------------------------------------------------------
LOGO = ('<svg width="40" height="40" viewBox="0 0 40 40"><path d="M10 5h14l8 8v22H10z" fill="#EEF0FF" stroke="#4F46E5" '
        'stroke-width="2.4" stroke-linejoin="round"/><path d="M24 5v8h8" fill="none" stroke="#4F46E5" stroke-width="2.4"/>'
        '<path d="M15 20h12M15 25h12M15 30h8" stroke="#4F46E5" stroke-width="2.2" stroke-linecap="round"/></svg>')
st.markdown(
    f'<div class="hdr"><div class="logo">{LOGO}<span class="ok">✓</span></div>'
    f'<div><h1>{esc(core.APP_TITLE)}</h1>'
    f'<p>Add one or many CVs, analyze them together, and download a Word or Excel report.</p></div>'
    f'<div class="chips"><span class="chip"><b>{len(S.entries)}</b>CVs loaded</span>'
    f'<span class="chip"><b>{len(done)}</b>analyzed</span></div></div>'
    f'<div class="strip"></div>', unsafe_allow_html=True)

if S.notice:
    kind, text = S.notice
    {"success": st.success, "warning": st.warning}.get(kind, st.info)(text)
    S.notice = None

left, right = st.columns([1, 2.7], gap="large")

# ------------------------------------------------------------
# LEFT: add + list
# ------------------------------------------------------------
with left:
    TYPES = ["pdf", "docx", "doc", "rtf", "odt", "txt", "png", "jpg", "jpeg", "pptx", "xlsx", "xls", "csv"]
    with st.container(key="dropcard"):
        up1 = st.file_uploader("Add CVs", accept_multiple_files=True, label_visibility="collapsed",
                               type=TYPES, key=f"upa_{S.uploader_key}")
    b1c, b2c = st.columns(2)
    with b1c:
        with st.container(key="addbtn"):
            up2 = st.file_uploader("Add CVs button", accept_multiple_files=True, label_visibility="collapsed",
                                   type=TYPES, key=f"upb_{S.uploader_key}")
    with b2c:
        with st.container(key="folderbtn"):
            with st.popover("Add Folder", use_container_width=True):
                st.caption("Put your CV folder into a .zip file and upload it here.")
                zp = st.file_uploader("Folder (.zip)", type=["zip"], key=f"zip_{S.uploader_key}")

    items = [(f.name, f.getvalue()) for f in (up1 or []) + (up2 or [])]
    if zp is not None:
        try:
            items += unzip_items(zp)
        except Exception:
            S.notice = ("warning", "That zip file could not be opened.")
    if items:
        with st.spinner("Reading your CVs…"):
            analyse_uploads(items)
        S.uploader_key += 1
        st.rerun()

    with st.container(key="listcard"):
        c1, c2, c3 = st.columns([1.5, 0.6, 0.8])
        c1.markdown(f'<span style="font-weight:700;font-size:16px">Your CVs</span>'
                    f'<span class="badge">{len(S.entries)}</span>', unsafe_allow_html=True)
        if c2.button("Re-run", key="rerun", disabled=not S.entries):
            with st.spinner("Analyzing again…"):
                for e in S.entries:
                    run_analysis(e)
            st.rerun()
        if c3.button("Clear all", key="clearall", disabled=not S.entries):
            S.entries, S.selected = [], None
            st.rerun()

        if not S.entries:
            st.markdown('<div class="hint">No CVs yet.<br>Drop files on the box above<br>or click Add CVs to choose them.</div>',
                        unsafe_allow_html=True)

        for e in S.entries:
            sel = e["id"] == S.selected
            with st.container(key=f"{'cvsel' if sel else 'cv'}_{e['id']}"):
                if e["status"] == "done":
                    a = e["analysis"]
                    label = core.candidate_label(a)
                    pct = a["completeness"]
                    st.markdown(
                        f'<div class="cvc"><div class="av" style="--c:{avatar_color(label)}">{esc(initials(label))}</div>'
                        f'<div style="flex:1;min-width:0"><div class="t">{esc(shorten(label, 30))}</div>'
                        f'<div class="s">{pct_text(pct)} complete · ATS {ats(a)["score"]:.0f} · {len(a["all_skills"])} skills</div>'
                        f'<div class="pbar"><i style="width:{pct}%;--c:{level_color(pct)}"></i></div></div></div>',
                        unsafe_allow_html=True)
                else:
                    st.markdown(
                        f'<div class="cvc"><div class="av" style="--c:{DANGER}">!</div><div style="flex:1;min-width:0">'
                        f'<div class="t">{esc(shorten(e["name"], 30))}</div>'
                        f'<div class="s" style="color:{DANGER}">Could not be analyzed</div></div></div>',
                        unsafe_allow_html=True)
                b1, b2 = st.columns(2)
                if b1.button("Open", key=f"open_{e['id']}", use_container_width=True):
                    S.selected = e["id"]
                    S["mode"] = "Candidate Detail"
                    st.rerun()
                if b2.button("Remove", key=f"rm_{e['id']}", use_container_width=True):
                    S.entries = [x for x in S.entries if x["id"] != e["id"]]
                    if S.selected == e["id"]:
                        S.selected = next((x["id"] for x in S.entries if x["status"] == "done"), None)
                    st.rerun()

# ------------------------------------------------------------
# RIGHT: analysis
# ------------------------------------------------------------
with right:
    top_l, top_r = st.columns([3, 1])
    with top_l:
        with st.container(key="modebar"):
            mode = st.segmented_control("View", ["Combined Analysis", "Candidate Detail"],
                                        default="Combined Analysis", key="mode",
                                        label_visibility="collapsed") or "Combined Analysis"
    with top_r:
        export_slot = st.container(key="dl_on" if done else "dl_off")

    required = core.parse_required_skills(S.get("target", "")) if S.get("target") else []
    combined = core.build_combined_analysis([e["analysis"] for e in done], required) if done else None

    # ---------- download report ----------
    with export_slot:
        with st.popover("↓  Download Report", use_container_width=True):
            if not done:
                st.caption("Analyze at least one CV first, then download the report.")
            else:
                fmt = st.radio("File type", ["Word Document", "Excel Workbook", "Text File"], key="fmt")
                sel_entry = get_entry(S.selected)
                scope_opts = ["All CVs"] + (["Selected CV only"] if sel_entry and sel_entry["status"] == "done" else [])
                scope = st.radio("What to include", scope_opts, key="scope")
                key = {"Word Document": ("docx", ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                       "Excel Workbook": ("xlsx", ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
                       "Text File": ("txt", ".txt", "text/plain")}[fmt]
                analyses = [sel_entry["analysis"]] if scope == "Selected CV only" else [e["analysis"] for e in done]
                try:
                    comb = core.build_combined_analysis(analyses, required)
                    with tempfile.TemporaryDirectory() as tmpdir:
                        out = os.path.join(tmpdir, "report" + key[1])
                        core.EXPORTERS[key[0]](out, analyses, comb)
                        data = Path(out).read_bytes()
                    st.download_button("Download", data=data, type="primary", use_container_width=True,
                                       file_name=core.default_report_name(len(analyses)) + key[1], mime=key[2])
                except Exception as err:
                    st.error(f"Export error: {err}")

    # ---------- COMBINED ----------
    if mode == "Combined Analysis":
        if not done:
            if S.entries:
                empty_state("Nothing to compare yet",
                            "None of the CVs could be analyzed. Check the red cards on the left.", "!")
            else:
                empty_state("Add CVs to see them analyzed together",
                            "Drop several CVs on the left. You'll get a ranking, a skill heat-map, "
                            "shared and distinctive skills — and a Word or Excel report to download.")
        else:
            count = combined["count"]
            with st.container(key="combwrap"):
                with st.container(key="targetbox"):
                    tl, ti = st.columns([0.8, 2.4])
                    tl.markdown('<div class="tlab">Target skills</div>', unsafe_allow_html=True)
                    ti.text_input("Target skills", placeholder="e.g. Python, SQL, Power BI",
                                  key="target", label_visibility="collapsed")
                t_dash, t_cmp, t_mat = st.tabs(["Dashboard", "Compare Candidates", "Skill Matrix"])

            # ---- dashboard ----
            with t_dash:
                k = st.columns(5)
                avg_ats = sum(ats(e["analysis"])["score"] for e in done) / len(done)
                kp = [("CVs analyzed", str(count), PRIMARY, False),
                      ("Avg. complete", f"{combined['avg_completeness']:.1f}%", SUCCESS, False),
                      ("Avg. ATS", f"{avg_ats:.1f}", PRIMARY_DARK, False),
                      ("Unique skills", str(combined["total_unique_skills"]), ACCENT, False),
                      ("Top area", combined["top_area"], WARNING, len(combined["top_area"]) > 12)]
                for col, (cap, val, c, small) in zip(k, kp):
                    col.markdown(f'<div class="kpi" style="--c:{c}"><div class="cap">{esc(cap)}</div>'
                                 f'<div class="val{" s" if small else ""}">{esc(val)}</div></div>',
                                 unsafe_allow_html=True)
                st.write("")
                a, b = st.columns([3, 2])
                a.markdown(card("Most common skills", "The skills that appear in most CVs",
                                bar_rows([(s, c) for s, c, _h in combined["skill_freq"][:10]], PRIMARY, 150,
                                         lambda v: f"{int(v)}/{count}", max_value=count)), unsafe_allow_html=True)
                b.markdown(card("Career areas", "Where the candidates work",
                                bar_rows([(s, c) for s, c, _h in combined["career_freq"][:8]], ACCENT, 150,
                                         lambda v: f"{int(v)}/{count}", max_value=count)), unsafe_allow_html=True)
                st.write("")
                by_c = sorted(combined["rows"], key=lambda r: (-r["completeness"], r["label"].lower()))
                pc1, pc2 = st.columns(2)
                pc1.markdown(card("Profile completeness",
                                  "How many CV sections were found for each candidate",
                                  bar_rows([(r["label"], r["completeness"]) for r in by_c[:12]], label_w=130,
                                           fmt=lambda v: f"{v:.1f}%", color_fn=level_color, max_value=100)),
                             unsafe_allow_html=True)
                by_a = sorted(combined["rows"], key=lambda r: (-ats(r["analysis"])["score"], r["label"].lower()))
                pc2.markdown(card("ATS score",
                                  "How easily an applicant tracking system can read each CV (out of 100)",
                                  bar_rows([(r["label"], ats(r["analysis"])["score"]) for r in by_a[:12]], label_w=130,
                                           fmt=lambda v: f"{v:.0f}", color_fn=level_color, max_value=100)),
                             unsafe_allow_html=True)
                st.write("")
                c1, c2 = st.columns(2)
                common = chips(combined["common_skills"], "#D8F5EE", "#0B7A57") or \
                    '<div class="muted">Add at least two CVs to see shared skills.</div>'
                c1.markdown(card("Skills every CV shares", "Common ground across the whole group", common),
                            unsafe_allow_html=True)
                uniq = combined["unique_skills"]
                if not uniq:
                    ub = ('<div class="muted">Add at least two CVs to see what makes each candidate different.</div>'
                          if count < 2 else '<div class="muted">No skills are unique to a single candidate.</div>')
                else:
                    parts = []
                    for n, (label, skills) in enumerate(uniq.items()):
                        if n >= 6:
                            parts.append(f'<div class="muted">+ {len(uniq) - n} more candidates (see Compare)</div>')
                            break
                        col = avatar_color(label)
                        parts.append(f'<div style="margin-top:{0 if n == 0 else 12}px;margin-bottom:5px">'
                                     f'<span class="dot" style="--c:{col}"></span><b>{esc(shorten(label, 34))}</b></div>'
                                     + chips(skills[:14], _mix(col, "#FFFFFF", .86), _mix(col, "#000000", .3)))
                    ub = "".join(parts)
                c2.markdown(card("Distinctive skills", "Skills that only one candidate has", ub), unsafe_allow_html=True)

            # ---- compare ----
            with t_cmp:
                f1, f2, f3 = st.columns([2.2, 1.2, 1])
                query = f1.text_input("Filter", placeholder="Filter by name, skill, role or career area…",
                                      label_visibility="collapsed").lower().strip()
                has_match = bool(combined["required"])
                sort_opts = ["Rank", "Candidate", "Email", "Phone", "Completeness", "ATS", "Skills", "Career Area"] + \
                            (["Match"] if has_match else [])
                sort_by = f2.selectbox("Sort by", sort_opts, label_visibility="collapsed")
                desc = f3.toggle("Descending", value=sort_by in ("Completeness", "ATS", "Skills", "Match"))
                rows = []
                for r in combined["rows"]:
                    hay = " ".join([r["label"], r["email"], r["filename"], r["top_area"]]
                                   + r["skills"] + r["career_areas"] + r["roles"]).lower()
                    if not query or query in hay:
                        rows.append(r)
                keyf = {"Rank": lambda r: r["rank"], "Candidate": lambda r: r["label"].lower(),
                        "Email": lambda r: (r["email"] or "~").lower(), "Phone": lambda r: r["phone"] or "~",
                        "Completeness": lambda r: r["completeness"],
                        "ATS": lambda r: ats(r["analysis"])["score"], "Skills": lambda r: r["skill_count"],
                        "Career Area": lambda r: r["top_area"].lower(),
                        "Match": lambda r: r["match_pct"] if r["match_pct"] is not None else -1}[sort_by]
                rows.sort(key=keyf, reverse=desc)
                head = ["#", "Candidate", "Email", "Phone", "Completeness", "ATS", "Skills", "Career Area"] + \
                       (["Match"] if has_match else [])
                body = []
                for r in rows:
                    cells = [f'<td class="c">{r["rank"]}</td>',
                             f'<td><b>{esc(r["label"])}</b></td>', f'<td>{esc(r["email"] or "—")}</td>',
                             f'<td>{esc(r["phone"] or "—")}</td>',
                             f'<td class="c" style="color:{level_color(r["completeness"])};font-weight:700">'
                             f'{r["completeness"]:.1f}%</td>',
                             f'<td class="c" style="color:{level_color(ats(r["analysis"])["score"])};font-weight:700">'
                             f'{ats(r["analysis"])["score"]:.0f}</td>',
                             f'<td class="c">{r["skill_count"]}</td>', f'<td>{esc(r["top_area"])}</td>']
                    if has_match:
                        cells.append(f'<td class="c" style="color:{level_color(r["match_pct"])};font-weight:700">'
                                     f'{r["match_pct"]:.0f}%  ({len(r["matched"])}/{len(combined["required"])})</td>')
                    body.append("<tr>" + "".join(cells) + "</tr>")
                info = f"{len(rows)} of {len(combined['rows'])} CVs"
                if has_match and combined["rows"]:
                    best = combined["rows"][0]
                    info += f"  ·  Best: {shorten(best['label'], 18)} ({best['match_pct']:.0f}%)"
                st.markdown(
                    f'<div class="card" style="padding:6px 6px 12px 6px"><div class="tw"><table class="t"><thead><tr>'
                    + "".join(f"<th>{h}</th>" for h in head) + "</tr></thead><tbody>" + "".join(body)
                    + f'</tbody></table></div><div class="muted" style="padding:10px 12px 0 12px">{esc(info)}  ·  '
                      f'Use "Open" on a candidate card to see the full profile.</div></div>', unsafe_allow_html=True)

            # ---- matrix ----
            with t_mat:
                skills = [s for s, _c, _h in combined["skill_freq"][:30]]
                cat_of = {}
                for e in done:
                    for cat, lst in e["analysis"]["skills_by_category"].items():
                        for s in lst:
                            cat_of.setdefault(s, cat)
                head = "".join(f'<th class="sk">{esc(s)}</th>' for s in skills)
                body = []
                for r in combined["rows"]:
                    has = set(r["skills"])
                    tds = []
                    for s in skills:
                        if s in has:
                            col = CATEGORY_COLORS.get(cat_of.get(s, ""), PRIMARY)
                            tds.append(f'<td class="cell" title="{esc(r["label"])} — {esc(s)}" style="background:{col}"></td>')
                        else:
                            tds.append('<td class="cell"></td>')
                    body.append(f'<tr><td class="nm">{esc(shorten(r["label"], 26))}</td>{"".join(tds)}</tr>')
                legend = "".join(f'<span style="margin-right:14px"><span class="dot" style="--c:{c}"></span>'
                                 f'<span class="muted" style="font-size:12px">{esc(n)}</span></span>'
                                 for n, c in CATEGORY_COLORS.items())
                st.markdown(
                    f'<div class="card"><h3>Who has which skill</h3><div class="sub">The 30 most common skills across '
                    f'all CVs. Hover a cell for details.</div><div class="tw"><table class="m"><thead><tr><th></th>{head}'
                    f'</tr></thead><tbody>{"".join(body)}</tbody></table></div>'
                    f'<div style="margin-top:12px">{legend}</div></div>', unsafe_allow_html=True)

    # ---------- DETAIL ----------
    else:
        entry = get_entry(S.selected)
        if not S.entries:
            empty_state("No CVs yet", "Add a CV and its detailed profile will show up here.")
        elif entry is None:
            empty_state("Pick a CV from your list", "Click Open on any card on the left to see that candidate's full profile.")
        elif entry["status"] == "error":
            empty_state("This CV couldn't be analyzed", f"{entry['name']} — {entry['error']}", "!")
        else:
            d = entry["analysis"]
            label = core.candidate_label(d)
            pills = "".join(f'<span class="pill">{esc(t)}</span>' for t in
                            (d["email"] or "No email found", d["phone"] or "No phone found", d["filename"]))
            stats = "".join(f'<span><b>{v}</b> {c}</span>' for v, c in
                            ((len(d["all_skills"]), "skills"), (len(d["roles"]), "job roles"),
                             (len(d["career_areas"]), "career areas"), (len(d["languages"]), "languages")))
            pct = d["completeness"]
            ats_d = ats(d)
            st.markdown(
                f'<div class="card"><div class="dh"><div class="av" style="--c:{avatar_color(label)}">{esc(initials(label))}</div>'
                f'<div><h2>{esc(label)}</h2><div>{pills}</div><div class="stats">{stats}</div></div>'
                f'<div class="rings"><div class="ring" style="--p:{pct};--c:{level_color(pct)}"><span><b>{pct_text(pct)}</b><small>complete</small></span></div>'
                f'<div class="ring" style="--p:{ats_d["score"]};--c:{level_color(ats_d["score"])}"><span><b>{ats_d["score"]:.0f}</b><small>ATS score</small></span></div></div>'
                f'</div></div>', unsafe_allow_html=True)
            st.write("")

            def numbered(items):
                return "".join(f'<div class="li"><span class="num">{i}.</span> {esc(t)}</div>'
                               for i, t in enumerate(items, 1))

            t_ov, t_sk, t_ex, t_ed, t_ce = st.tabs(
                ["Overview", "Skills", "Experience", "Education", "Courses & Certifications"])

            with t_ov:
                h = ['<div class="dt">Resume Overview</div>'
                     '<div class="dm">A structured summary of the candidate profile and resume quality.</div>',
                     '<div class="sec">Candidate Information</div>']
                for lab, val in (("Candidate Name", d["name"]), ("Email", d["email"]),
                                 ("Phone", d["phone"]), ("Resume File", d["filename"])):
                    h.append(f'<div class="fld"><b>{lab}:</b> {esc(val or "Not detected")}</div>')
                h.append('<div class="sec">Career Areas Identified</div>')
                h.append(numbered(d["career_areas"]) if d["career_areas"] else
                         '<div class="muted">No specific career area confidently identified.</div>')
                h.append('<div class="sec">Job Roles / Positions Detected</div>')
                h.append(numbered(d["roles"]) if d["roles"] else
                         '<div class="muted">No specific job roles detected.</div>')
                if required:
                    matched = [r for r in required if core.skill_in_resume(d, r)]
                    missing = [r for r in required if r not in matched]
                    h.append('<div class="sec">Match With Target Skills</div>'
                             f'<div class="metric">{len(matched) / len(required) * 100:.0f}% match</div>'
                             f'<div class="fld"><b>Matched:</b> {esc(", ".join(matched) or "none")}</div>'
                             f'<div class="fld"><b>Missing:</b> {esc(", ".join(missing) or "none")}</div>')
                h.append('<div class="sec">Resume Completeness</div>'
                         f'<div class="metric">Completeness Score: {pct:.1f}%</div>'
                         '<div class="sec" style="margin-top:6px">Section Check</div>')
                for i, (sec, ok) in enumerate(d["section_status"].items(), 1):
                    mark = '<span class="okk">✓ Present</span>' if ok else '<span class="bad">Not detected</span>'
                    h.append(f'<div class="li"><span class="num">{i}.</span> <b>{esc(sec)}</b> — {mark}</div>')
                h.append('<div class="sec">ATS Score</div>'
                         f'<div class="metric">ATS Score: {ats_d["score"]:.1f} / 100</div>'
                         '<div class="dm" style="margin-bottom:4px">An estimate of how easily an Applicant Tracking System can read and rank this CV. '
                         'It is a guide, not the result of any real ATS.</div>')
                for i, (lab, got, top, note) in enumerate(ats_d["parts"], 1):
                    col = level_color(got / top * 100)
                    h.append(f'<div class="li"><span class="num">{i}.</span> <b>{esc(lab)}</b> — '
                             f'<span style="color:{col};font-weight:700">{got:g}/{top}</span> '
                             f'<span class="muted">· {esc(note)}</span></div>')
                if ats_d["tips"]:
                    h.append('<div class="sec" style="margin-top:10px">How to raise the ATS score</div>')
                    h.extend(f'<div class="li">• {esc(t)}</div>' for t in ats_d["tips"])
                parts = []
                if d["name"]:
                    parts.append(f"The resume belongs to {d['name']}.")
                if d["career_areas"]:
                    parts.append(f"The strongest career areas identified are {', '.join(d['career_areas'][:3])}.")
                if d["all_skills"]:
                    parts.append(f"The resume contains {len(d['all_skills'])} recognized skills.")
                if d["experience"]:
                    parts.append("Professional experience information is present in the resume.")
                if d["education"]:
                    parts.append("Educational qualifications are present.")
                if d["certifications"]:
                    parts.append("Courses or certifications were detected.")
                if d["languages"]:
                    parts.append("Language information is also present.")
                h.append('<div class="sec">Analysis Summary</div><div class="fld">'
                         + esc(" ".join(parts) or "The application was unable to identify enough structured information.")
                         + "</div>")
                st.markdown("".join(h), unsafe_allow_html=True)

            with t_sk:
                h = ['<div class="dt">Skills Identified From Resume</div>'
                     '<div class="dm">Skills are grouped by category for easier evaluation.</div>']
                if not d["skills_by_category"]:
                    h.append('<div class="muted">No skills were detected from the master skill dictionary.</div>')
                else:
                    for cat, lst in d["skills_by_category"].items():
                        col = CATEGORY_COLORS.get(cat, PRIMARY)
                        h.append(f'<div style="margin:16px 0 7px 0"><span class="dot" style="--c:{col}"></span>'
                                 f'<b>{esc(cat)}</b> <span class="muted">&nbsp;{len(lst)}</span></div>'
                                 + chips(lst, _mix(col, "#FFFFFF", .87), _mix(col, "#000000", .28)))
                    h.append(f'<div style="margin-top:20px;font-weight:700;color:{PRIMARY_DARK}">'
                             f'Total Unique Skills Detected: {len(d["all_skills"])}</div>')
                st.markdown("".join(h), unsafe_allow_html=True)

            with t_ex:
                h = ['<div class="dt">Professional Experience</div>'
                     '<div class="dm">Relevant experience is presented in a compact, structured format.</div>']
                if not d["experience"]:
                    h.append('<div class="muted">No dedicated Experience section was detected.</div>')
                else:
                    formatted = core.format_experience_short(d["experience"])
                    for i, blk in enumerate([b.strip() for b in formatted.split("\n\n") if b.strip()], 1):
                        lines = blk.splitlines()
                        h.append(f'<div class="role"><span class="num">{i}.</span> {esc(lines[0])}</div>')
                        if len(lines) > 1:
                            h.append(f'<div class="subh">{esc(lines[1])}</div>')
                        for j, ln in enumerate(lines[2:], 1):
                            h.append(f'<div class="li" style="margin-left:18px"><span class="num">{j}.</span> '
                                     f'{esc(ln.strip().lstrip("•").strip())}</div>')
                st.markdown("".join(h), unsafe_allow_html=True)

            with t_ed:
                h = ['<div class="dt">Education</div><div class="dm">Educational qualifications identified from the resume.</div>']
                lines = [x.strip() for x in (d["education"] or "").splitlines() if x.strip()]
                h.append(numbered(lines) if lines else '<div class="muted">No dedicated Education information was detected.</div>')
                st.markdown("".join(h), unsafe_allow_html=True)

            with t_ce:
                h = ['<div class="dt">Courses &amp; Certifications</div>'
                     '<div class="dm">Courses and certification information identified from the resume.</div>'
                     '<div class="sec">Courses / Certifications</div>']
                lines = [x.strip() for x in (d["certifications"] or "").splitlines() if x.strip()]
                h.append(numbered(lines) if lines else
                         '<div class="muted">No dedicated Courses or Certifications section was detected.</div>')
                h.append('<div class="sec">Languages</div>')
                h.append(numbered(d["languages"]) if d["languages"] else
                         '<div class="muted">No languages were confidently detected.</div>')
                st.markdown("".join(h), unsafe_allow_html=True)

status_text = "Ready — add CVs to begin." if not S.entries else f"{len(done)} of {len(S.entries)} CVs analyzed."
st.markdown(f'<div class="foot2"><span>{status_text}</span><span class="r">Drag &amp; drop files on the box  ·  '
            f'Download report any time  ·  Remove a CV with its Remove button</span></div>', unsafe_allow_html=True)
