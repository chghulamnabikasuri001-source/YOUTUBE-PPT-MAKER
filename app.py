"""
Lecture to PPT - website (Streamlit).  The heavy lifting lives in engine.py and pptx_builder.py.
"""
import os
import re
from datetime import datetime, timedelta, timezone

import streamlit as st
from google import genai

import engine as E
from pptx_builder import THEMES

st.set_page_config(page_title="Lecture to PPT", page_icon="📚", layout="centered")

BUCKET = "ppts"  # Supabase storage bucket name
PKT = timezone(timedelta(hours=5))
MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
VIS_ICON = {"nested": "⭕", "flow": "🔀", "cards": "🗂️", "steps": "🪜", "table": "📊", "tree": "🌳",
            "network": "🧠", "code": "💻", "compare": "⚖️", None: "📝"}

CSS = """
<style>
:root { --blue:#2F80ED; --sky:#56CCF2; --ink:#12324F; --soft:#EAF4FF; --line:#CFE3FA; }
html, body, .stApp { color-scheme: light; }
.stApp {
  background:
    radial-gradient(900px 420px at 8% -8%, #D6EAFF 0%, rgba(214,234,255,0) 60%),
    radial-gradient(800px 420px at 100% 0%, #DDF5FF 0%, rgba(221,245,255,0) 55%),
    #F7FBFF;
}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer { display:none !important; }
.block-container { max-width: 900px; padding-top: 1.4rem; padding-bottom: 4rem; }

/* text */
.stApp, .stApp p, .stApp li, .stApp label, .stApp h1, .stApp h2, .stApp h3, .stApp h4,
.stApp [data-testid="stMarkdownContainer"], .stApp [data-testid="stCaptionContainer"] { color: var(--ink); }
[data-testid="stWidgetLabel"] p { color:#1F4E79 !important; font-weight:600; font-size:.92rem; }

/* hero */
.hero { position:relative; overflow:hidden; border-radius:28px; padding:2.4rem 2.2rem 2.1rem;
  background: linear-gradient(135deg, #2F80ED 0%, #4AA8F5 55%, #56CCF2 100%);
  box-shadow: 0 20px 44px rgba(47,128,237,.28); margin-bottom:1.3rem; animation: rise .6s ease both; }
.hero .blob { position:absolute; border-radius:50%; background:rgba(255,255,255,.16); animation: floaty 7s ease-in-out infinite; }
.hero .b1 { width:210px; height:210px; right:-50px; top:-60px; }
.hero .b2 { width:120px; height:120px; right:140px; bottom:-45px; animation-delay:-3s; }
.hero .b3 { width:70px; height:70px; left:55%; top:22px; animation-delay:-5s; background:rgba(255,255,255,.22); }
.hero .badge { display:inline-block; background:rgba(255,255,255,.22); color:#fff !important; padding:.3rem .85rem;
  border-radius:999px; font-weight:700; font-size:.82rem; letter-spacing:.02em; }
.hero h1 { color:#fff !important; font-size:2.15rem; line-height:1.15; margin:.75rem 0 .55rem; font-weight:800; padding:0; }
.hero p.sub { color:rgba(255,255,255,.94) !important; font-size:1.02rem; max-width:560px; margin:0 0 1.1rem; }
.hero .chip { display:inline-block; padding:.42rem .95rem; margin:.2rem .4rem .2rem 0; border-radius:999px;
  background:rgba(255,255,255,.2); color:#fff !important; font-weight:600; font-size:.9rem; transition:all .2s; cursor:default; }
.hero .chip:hover { background:rgba(255,255,255,.34); transform:translateY(-2px); }
@keyframes floaty { 0%,100% { transform:translateY(0) } 50% { transform:translateY(-14px) } }
@keyframes rise { from { opacity:0; transform:translateY(12px) } to { opacity:1; transform:none } }

/* cards */
div[data-testid="stVerticalBlockBorderWrapper"] {
  background:#fff; border:1px solid #D6E8FB !important; border-radius:22px !important;
  box-shadow:0 12px 34px rgba(47,128,237,.09); }

/* inputs */
.stTextInput input, .stTextArea textarea, .stNumberInput input {
  background:#fff !important; color:var(--ink) !important; border:1.5px solid var(--line) !important;
  border-radius:12px !important; transition:all .15s; }
.stTextInput input:focus, .stTextArea textarea:focus { border-color:var(--blue) !important; box-shadow:0 0 0 3px rgba(47,128,237,.16) !important; }
.stTextInput input::placeholder, .stTextArea textarea::placeholder { color:#8AA6C4 !important; }
div[data-baseweb="select"] > div { background:#fff !important; border:1.5px solid var(--line) !important; border-radius:12px !important; color:var(--ink) !important; }
div[data-baseweb="select"] * { color:var(--ink) !important; }
div[data-baseweb="select"] svg { fill:var(--blue) !important; }
div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li, div[data-baseweb="popover"] > div { background:#fff !important; color:var(--ink) !important; }
div[data-baseweb="popover"] li:hover { background:var(--soft) !important; }
[data-testid="stCheckbox"] label span { color:var(--ink) !important; }

/* buttons */
.stButton > button, .stDownloadButton > button {
  border-radius:14px !important; font-weight:700 !important; padding:.7rem 1.3rem !important; transition:all .18s !important; }
.stButton > button[kind="primary"], button[data-testid="stBaseButton-primary"] {
  background:linear-gradient(135deg,#2F80ED,#56CCF2) !important; color:#fff !important; border:none !important;
  box-shadow:0 10px 24px rgba(47,128,237,.32) !important; font-size:1.05rem !important; }
.stButton > button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover { transform:translateY(-2px); box-shadow:0 14px 30px rgba(47,128,237,.42) !important; }
.stButton > button[kind="primary"] p, button[data-testid="stBaseButton-primary"] p { color:#fff !important; }
.stButton > button:not([kind="primary"]), .stDownloadButton > button {
  background:#fff !important; color:#1F6FD8 !important; border:1.5px solid var(--line) !important; }
.stButton > button:not([kind="primary"]) p, .stDownloadButton > button p { color:#1F6FD8 !important; }
.stButton > button:not([kind="primary"]):hover, .stDownloadButton > button:hover { background:var(--soft) !important; border-color:var(--blue) !important; transform:translateY(-1px); }

/* expander / status / alerts */
[data-testid="stExpander"] { background:#fff; border:1px solid #D6E8FB !important; border-radius:16px !important; }
[data-testid="stExpander"] summary p { color:#1F4E79 !important; font-weight:600; }
[data-testid="stStatus"] { background:#fff !important; border:1px solid #D6E8FB !important; border-radius:18px !important; }
[data-testid="stAlert"] { border-radius:14px !important; }
[data-testid="stProgress"] > div > div > div { background:linear-gradient(90deg,#2F80ED,#56CCF2) !important; }
[data-testid="stProgress"] > div > div { background:#DCEBFB !important; }
[data-testid="stTooltipContent"], div[data-baseweb="tooltip"] > div { background:#12324F !important; color:#fff !important; }

/* theme swatches & stats */
.sw { display:flex; align-items:center; gap:.4rem; margin:.1rem 0 .6rem; }
.sw i { width:22px; height:22px; border-radius:50%; display:inline-block; box-shadow:0 2px 6px rgba(0,0,0,.15); transition:transform .15s; }
.sw i:hover { transform:scale(1.25); }
.sw span { font-size:.82rem; color:#5B7A99; margin-left:.4rem; }
.stats { display:flex; gap:.8rem; margin:.6rem 0 1rem; }
.stat { flex:1; background:linear-gradient(180deg,#F2F8FF,#fff); border:1px solid #D6E8FB; border-radius:16px; padding:.9rem 1rem; text-align:center; transition:transform .18s; }
.stat:hover { transform:translateY(-3px); }
.stat b { display:block; font-size:1.7rem; color:#2F80ED; line-height:1.1; }
.stat span { font-size:.82rem; color:#5B7A99; }
.ok-title { font-size:1.25rem; font-weight:800; color:#12324F; margin:.2rem 0 0; }
</style>
"""

HERO = """
<div class="hero">
  <div class="blob b1"></div><div class="blob b2"></div><div class="blob b3"></div>
  <span class="badge">✨ AI presentation maker for teachers</span>
  <h1>Turn any lecture video into a beautiful presentation</h1>
  <p class="sub">Paste a YouTube link and get designed slides with diagrams, flow charts, tables,
  code and speaker notes, ready to teach.</p>
  <span class="chip">1 · Paste the link</span><span class="chip">2 · AI designs the slides</span><span class="chip">3 · Download your PPT</span>
</div>
"""


def get_secret(name, default=None):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


# ----------------------------------------------------------------------
# History (permanent with Supabase, otherwise this browser session only)
# ----------------------------------------------------------------------
@st.cache_resource
def get_store():
    url, key = get_secret("SUPABASE_URL"), get_secret("SUPABASE_KEY")
    if not (url and key):
        return None
    from supabase import create_client

    return create_client(url, key)


def save_history(title, data):
    slug = re.sub(r"[^A-Za-z0-9]+", "-", title).strip("-")[:40] or "lecture"
    name = f"{datetime.now(PKT).strftime('%Y%m%d-%H%M%S')}__{slug}.pptx"
    store = get_store()
    if store:
        try:
            store.storage.from_(BUCKET).upload(path=name, file=data, file_options={"content-type": MIME})
            return ""
        except Exception as e:
            return f"Could not save to history: {str(e)[:200]}"
    st.session_state.setdefault("history", []).insert(0, {"name": name, "data": data})
    return ""


def list_history():
    store = get_store()
    if store:
        try:
            rows = store.storage.from_(BUCKET).list(
                "", {"limit": 50, "sortBy": {"column": "created_at", "order": "desc"}})
            names = sorted((r["name"] for r in rows if r["name"].endswith(".pptx")), reverse=True)
            return [{"name": n, "data": None} for n in names[:30]]
        except Exception:
            return None
    return st.session_state.get("history", [])


def pretty(name):
    try:
        d = datetime.strptime(name[:15], "%Y%m%d-%H%M%S").strftime("%d %b %Y, %I:%M %p")
        return f"{d}  —  {name.split('__', 1)[1][:-5].replace('-', ' ')}"
    except Exception:
        return name


def show_history():
    with st.expander("📂 Previous presentations"):
        items = list_history()
        if items is None:
            st.warning("History could not be loaded. Please check the storage settings.")
            return
        if not get_store():
            st.caption("History is kept only for this browser session.")
        if not items:
            st.write("No presentations yet.")
        for it in items:
            st.write(pretty(it["name"]))
            data, key = it["data"], f"dl::{it['name']}"
            if data is None:
                data = st.session_state.get(key)
                if data is None:
                    if st.button("Prepare download", key="btn::" + it["name"]):
                        try:
                            st.session_state[key] = get_store().storage.from_(BUCKET).download(it["name"])
                            st.rerun()
                        except Exception as e:
                            st.error(f"File not found: {str(e)[:150]}")
                    continue
            st.download_button("⬇️ Download", data, file_name=it["name"].split("__", 1)[-1],
                               mime=MIME, key="dlb::" + it["name"])


# ----------------------------------------------------------------------
# Page
# ----------------------------------------------------------------------
def password_gate():
    pw = get_secret("APP_PASSWORD")
    if not pw or st.session_state.get("authed"):
        return True
    with st.container(border=True):
        st.markdown("**🔒 This app is private.** Enter the password to continue.")
        entered = st.text_input("Password", type="password")
    if entered:
        if entered == pw:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def swatches(theme_name):
    T = THEMES[theme_name]
    cols = [T["dark"]] + T["acc"] + [T["light"]]
    dots = "".join(f'<i style="background:#{c}"></i>' for c in cols)
    return f'<div class="sw">{dots}<span>{theme_name} palette</span></div>'


def show_result(res):
    st.markdown(f'<p class="ok-title">🎉 Your presentation is ready</p>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="stats"><div class="stat"><b>{res["slides"]}</b><span>content slides</span></div>'
        f'<div class="stat"><b>{res["visuals"]}</b><span>with diagrams</span></div>'
        f'<div class="stat"><b>{res["shots"]}</b><span>screenshots</span></div></div>',
        unsafe_allow_html=True)
    safe = re.sub(r"[^\w\- ]", "", res["title"]).strip()[:60] or "lecture"
    st.download_button("⬇️ Download PowerPoint", res["data"], file_name=f"{safe}.pptx", mime=MIME,
                       use_container_width=True, type="primary")
    for n in res["notes"]:
        st.info(n)
    with st.expander("👀 Slide outline"):
        for i, (head, vis, shot) in enumerate(res["outline"], 1):
            st.write(f"{VIS_ICON.get(vis, '🧩')} **{i}.** {head}" + ("  📸" if shot else ""))
    st.caption("Please review the slides before presenting. The speaker notes hold the full explanation.")


def main():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown(HERO, unsafe_allow_html=True)

    if not password_gate():
        st.stop()

    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        st.error("The app is not configured yet: GEMINI_API_KEY is missing in Secrets.")
        st.stop()

    with st.container(border=True):
        url = st.text_input("YouTube video link", placeholder="https://www.youtube.com/watch?v=...")
        with st.expander("Having trouble with a video? Paste the transcript instead"):
            manual = st.text_area("Transcript (optional)", height=140, key="manual_transcript")
            st.caption('On YouTube: open the video, expand the description, click "Show transcript", '
                       "select all the text and copy it here. Timestamps are supported.")
        c1, c2, c3 = st.columns(3)
        lang_label = c1.selectbox("Slide language", list(E.LANGS.keys()))
        detail_label = c2.selectbox("Detail level", list(E.DETAIL.keys()), index=1)
        theme_name = c3.selectbox("Theme", list(THEMES.keys()))
        st.markdown(swatches(theme_name), unsafe_allow_html=True)
        want_shots = st.checkbox(
            "Capture notes and code from the video (slower)", value=True,
            help="Adds screenshots only where handwritten notes, a board, text slides or code are shown. "
                 "Photos of the teacher are never added.")
        go = st.button("✨ Create presentation", type="primary", use_container_width=True)

    if go:
        if not url.strip():
            st.warning("Please paste a YouTube link first.")
        else:
            warnings, err = [], None
            with st.status("Creating your presentation…", expanded=True) as status:
                bar = st.progress(0)

                def say(m):
                    st.write(m)

                def prog(f, t):
                    bar.progress(max(0, min(int(f * 100), 100)), text=t)

                def warn(m):
                    warnings.append(m)
                    st.write("⚠️ " + m)

                try:
                    model = get_secret("GEMINI_MODEL", E.DEFAULT_MODEL)
                    client = genai.Client(api_key=api_key)
                    res = E.generate_presentation(
                        url, manual, lang_label, detail_label, theme_name, want_shots, client, model,
                        proxy=get_secret("PROXY_URL"), say=say, prog=prog, warn=warn)
                    hist_err = save_history(res["title"], res["data"])
                    if hist_err:
                        warnings.append(hist_err)
                    res["notes"] = warnings + res["notes"]
                    st.session_state["result"] = res
                    status.update(label="Your presentation is ready!", state="complete", expanded=False)
                    st.toast("Presentation ready! 🎉")
                except Exception as e:
                    err = e
                    status.update(label="Could not create the presentation", state="error", expanded=True)
            if err is not None:
                st.session_state.pop("result", None)
                if isinstance(err, E.UserError):
                    st.error(err.msg)
                else:
                    st.error(E.explain_error(err))
                st.caption("Tip: if YouTube is blocking the server, paste the transcript manually in the box above.")
                with st.expander("Technical details"):
                    st.code(f"{type(err).__name__}: {str(err)[:1500]}")

    res = st.session_state.get("result")
    if res:
        with st.container(border=True):
            show_result(res)

    show_history()


main()
