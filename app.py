"""
YouTube Lecture to PPT  (Gemini + Streamlit)
Features: language, detail level, colour theme, screenshots, history, quiz.
"""
import json
import os
import re
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

import streamlit as st
from google import genai
from google.genai import types
from youtube_transcript_api import YouTubeTranscriptApi

from pptx_builder import THEMES, build_pptx

st.set_page_config(page_title="Lecture to PPT", page_icon="📚", layout="centered")

# Hide the Streamlit toolbar (GitHub icon, menu) for a cleaner public site
st.markdown(
    """
    <style>
    [data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

DEFAULT_MODEL = "gemini-2.5-flash"                 # works on the free tier
FALLBACK_MODELS = ["gemini-2.5-flash-lite"]        # tried automatically if the first one fails
CHUNK_CHARS = 9000                                 # roughly 10 minutes of lecture per part
MAX_TRANSCRIPT_CHARS = 150000                      # cost / limit protection for very long videos
MAX_SHOTS = 20                                     # maximum screenshots per deck
MAX_VIDEO_SECONDS = 2 * 60 * 60                    # no screenshots for videos longer than 2 hours
BUCKET = "ppts"                                    # Supabase storage bucket name
PKT = timezone(timedelta(hours=5))
MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

LANGS = {
    "English": "English",
    "Urdu (اردو)": "Urdu (written in Urdu/Arabic script)",
    "Roman Urdu": "Roman Urdu (Urdu written in English letters)",
    "Same as the video": "the same language as the transcript",
}
DETAIL = {
    "Brief (fewer slides)": "3-4",
    "Standard": "5-7",
    "Detailed (more slides)": "8-10",
}


def get_secret(name, default=None):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


def explain_error(e) -> str:
    """Turn a technical error into a clear message for the user."""
    m = str(e)
    low = m.lower()
    if "api_key_invalid" in low or "api key not valid" in low or "permission_denied" in low or "unauthenticated" in low:
        return "The AI service rejected the API key. Check GEMINI_API_KEY in the app's Secrets."
    if "429" in m or "resource_exhausted" in low or "quota" in low:
        return ("The free AI usage limit has been reached. Please wait a minute and try again, "
                "or try a shorter video / lower detail level.")
    if "404" in m or "not_found" in low:
        return "The selected AI model is not available. Set GEMINI_MODEL in Secrets (for example gemini-2.5-flash-lite)."
    if "503" in m or "unavailable" in low or "overloaded" in low:
        return "The AI service is busy right now. Please try again in a minute."
    if "valid json" in low or "format" in low:
        return "The AI returned an unexpected format. Please try again."
    return "Something went wrong while generating the slides. Please try again."


# ----------------------------------------------------------------------
# YouTube: link, info, transcript
# ----------------------------------------------------------------------
def extract_video_id(url: str):
    url = url.strip()
    m = re.search(r"(?:v=|/videos/|embed/|youtu\.be/|/shorts/|/live/)([A-Za-z0-9_-]{11})", url)
    if m:
        return m.group(1)
    if re.fullmatch(r"[A-Za-z0-9_-]{11}", url):
        return url
    return None


def get_video_info(video_id: str):
    try:
        watch = f"https://www.youtube.com/watch?v={video_id}"
        api = "https://www.youtube.com/oembed?format=json&url=" + urllib.parse.quote(watch)
        with urllib.request.urlopen(api, timeout=10) as r:
            data = json.loads(r.read().decode("utf-8"))
        return data.get("title", ""), data.get("author_name", "")
    except Exception:
        return "", ""


def transcript_via_api(video_id: str):
    """Main method: youtube-transcript-api. Returns [(start, duration, text), ...]."""
    proxy = get_secret("PROXY_URL")
    if proxy:
        from youtube_transcript_api.proxies import GenericProxyConfig

        api = YouTubeTranscriptApi(proxy_config=GenericProxyConfig(http_url=proxy, https_url=proxy))
    else:
        api = YouTubeTranscriptApi()
    transcripts = api.list(video_id)
    try:
        t = transcripts.find_transcript(["ur", "hi", "en"])
    except Exception:
        t = next(iter(transcripts))
    out = []
    for s in t.fetch():
        txt = re.sub(r"\s+", " ", s.text.replace("\n", " ")).strip()
        if txt:
            out.append((float(s.start), float(s.duration), txt))
    return out


def transcript_via_ytdlp(video_id: str):
    """Backup method: read the caption track through yt-dlp."""
    from yt_dlp import YoutubeDL

    opts = {"quiet": True, "no_warnings": True, "skip_download": True, "socket_timeout": 20}
    proxy = get_secret("PROXY_URL")
    if proxy:
        opts["proxy"] = proxy
    with YoutubeDL(opts) as ydl:
        info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=False)
    for source in (info.get("subtitles") or {}, info.get("automatic_captions") or {}):
        for lang in ("ur", "hi", "en"):
            for key in [k for k in source if k == lang or k.startswith(lang + "-")]:
                for fmt in source[key]:
                    if fmt.get("ext") != "json3":
                        continue
                    with urllib.request.urlopen(fmt["url"], timeout=20) as r:
                        data = json.loads(r.read().decode("utf-8"))
                    out = []
                    for ev in data.get("events", []):
                        txt = "".join(sg.get("utf8", "") for sg in (ev.get("segs") or []))
                        txt = re.sub(r"\s+", " ", txt.replace("\n", " ")).strip()
                        if txt:
                            out.append((ev.get("tStartMs", 0) / 1000, ev.get("dDurationMs", 0) / 1000, txt))
                    if out:
                        return out
    raise RuntimeError("No caption track found")


def get_transcript(video_id: str):
    try:
        return transcript_via_api(video_id)
    except Exception as first_error:
        try:
            return transcript_via_ytdlp(video_id)
        except Exception:
            raise first_error


def transcript_error_message(e) -> str:
    name = type(e).__name__
    if name in ("TranscriptsDisabled", "NoTranscriptFound"):
        return ("This video has no captions, so a transcript is not available. "
                "Try another video, or paste the transcript manually.")
    if name in ("VideoUnavailable", "VideoUnplayable", "InvalidVideoId"):
        return "This video is unavailable or private."
    if name in ("IpBlocked", "RequestBlocked", "YouTubeRequestFailed", "AgeRestricted"):
        return ("YouTube is temporarily blocking requests from this server. "
                "Please try again later, or paste the transcript manually below.")
    return ("The transcript for this video could not be retrieved. YouTube may be blocking the server. "
            "Please try again later, or paste the transcript manually.")


def parse_manual_transcript(text: str):
    """Parses text copied from YouTube's 'Show transcript' panel.
    Returns (snippets, has_timestamps)."""
    ts = re.compile(r"^(?:(\d{1,2}):)?(\d{1,2}):(\d{2})$")
    snips, cur_time, has_ts, n = [], 0.0, False, 0
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = ts.match(line)
        if m:
            h, mi, sec = int(m.group(1) or 0), int(m.group(2)), int(m.group(3))
            cur_time, has_ts = h * 3600 + mi * 60 + sec, True
            continue
        if not has_ts:
            cur_time = n * 5.0  # no timestamps: rough placeholder times
        snips.append((float(cur_time), 5.0, line))
        n += 1
    return snips, has_ts


def split_chunks(snips, size=CHUNK_CHARS, max_chars=MAX_TRANSCRIPT_CHARS):
    chunks, cur, n, total, truncated = [], [], 0, 0, False

    def flush():
        if cur:
            chunks.append({
                "text": " ".join(x[2] for x in cur),
                "start": cur[0][0],
                "end": cur[-1][0] + cur[-1][1],
            })

    for sn in snips:
        if total + len(sn[2]) > max_chars:
            truncated = True
            break
        cur.append(sn)
        n += len(sn[2]) + 1
        total += len(sn[2]) + 1
        if n >= size:
            flush()
            cur, n = [], 0
    flush()
    return chunks, truncated


# ----------------------------------------------------------------------
# Gemini AI
# ----------------------------------------------------------------------
def ask_ai(client, model, system, user, max_tokens=4000):
    """Calls Gemini. Retries on rate limits and falls back to a second model if needed."""
    models = [model] + [m for m in FALLBACK_MODELS if m != model]
    last_err = None
    for mdl in models:
        for attempt in range(2):
            try:
                r = client.models.generate_content(
                    model=mdl,
                    contents=user,
                    config=types.GenerateContentConfig(
                        system_instruction=system,
                        max_output_tokens=max_tokens + 4000,
                        response_mime_type="application/json",
                    ),
                )
                return r.text or ""
            except Exception as e:
                last_err = e
                msg = str(e)
                if "404" in msg or "NOT_FOUND" in msg:
                    break  # model missing -> try the next model
                if any(k in msg for k in ("429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE")):
                    if attempt == 0:
                        time.sleep(20)  # free tier: wait a little, then retry once
                        continue
                    break  # still limited -> try the next model
                raise
    raise last_err


def parse_json(raw: str):
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()
    a, b = raw.find("{"), raw.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("No JSON found")
    return json.loads(raw[a : b + 1])


def ask_json(client, model, system, user, max_tokens=4000):
    for _ in range(2):
        raw = ask_ai(client, model, system, user, max_tokens)
        try:
            return parse_json(raw)
        except Exception:
            user += "\n\nYour previous answer was not valid JSON. Return only valid JSON."
    raise ValueError("The AI did not return a valid JSON format.")


def slides_from_chunk(client, model, chunk, idx, total, lang, per_chunk, done_headings):
    system = f"""You are an expert teacher's assistant. You turn a part of a lecture
transcript into detailed presentation slides for a classroom.

Rules:
- Cover EVERY concept, definition, example, formula, step and important point the teacher explains in this part.
- Do not invent facts that are not in the transcript. The transcript comes from auto-captions, so silently fix obvious spelling mistakes.
- Output language: {lang}.
- Make about {per_chunk} slides for this part.
- Each slide: a short heading, 3-6 bullets (each bullet a complete, useful line), and speaker "notes" (3-6 sentences with the full explanation for the teacher).
- Skip greetings, subscribe requests, and ads.
- Return ONLY valid JSON, no extra text, in this exact shape:
{{"slides": [{{"heading": "...", "bullets": ["...", "..."], "notes": "..."}}]}}"""
    user = f"""This is part {idx} of {total} of the lecture.
Headings already made in earlier parts (do not repeat them): {done_headings[-15:]}

TRANSCRIPT PART:
{chunk}"""
    return ask_json(client, model, system, user, max_tokens=6000).get("slides", [])


def summary_and_quiz(client, model, slides, lang, video_title):
    outline = "\n".join(
        f"- {s.get('heading', '')}: " + "; ".join(s.get("bullets", [])[:3]) for s in slides
    )
    system = f"""You help a teacher. From the lecture outline, produce a good presentation title,
a summary and a short quiz. Output language: {lang}.
Return ONLY valid JSON in this shape:
{{"title": "...", "summary": ["5-7 key takeaways"], "quiz": [{{"q": "...", "a": "..."}}]}}
Give 5 quiz questions with short answers."""
    return ask_json(client, model, system, f"Video title: {video_title}\n\nOUTLINE:\n{outline}", max_tokens=3000)


# ----------------------------------------------------------------------
# Screenshots
# ----------------------------------------------------------------------
def grab_frames(video_id, times):
    """Grabs one frame at each given time (seconds). Returns {position: jpg_bytes}.
    Uses yt-dlp to get the stream URL and ffmpeg to pull only the needed frames."""
    from yt_dlp import YoutubeDL

    opts = {
        "format": "best[height<=480][ext=mp4]/best[height<=480]/best",
        "quiet": True, "no_warnings": True, "noplaylist": True, "socket_timeout": 20,
    }
    proxy = get_secret("PROXY_URL")
    if proxy:
        opts["proxy"] = proxy
    with YoutubeDL(opts) as ydl:
        info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=False)
    stream = info.get("url")
    if not stream:
        return {}
    headers = info.get("http_headers", {}) or {}
    hdr = "".join(f"{k}: {v}\r\n" for k, v in headers.items())

    frames = {}
    with tempfile.TemporaryDirectory() as tmp:
        for i, t in enumerate(times):
            fp = os.path.join(tmp, f"f{i}.jpg")
            cmd = ["ffmpeg", "-y", "-loglevel", "error"]
            if hdr:
                cmd += ["-headers", hdr]
            cmd += ["-ss", str(int(t)), "-i", stream, "-frames:v", "1",
                    "-vf", "scale=720:-2", "-q:v", "4", fp]
            try:
                subprocess.run(cmd, capture_output=True, timeout=45)
            except Exception:
                continue
            if os.path.exists(fp) and os.path.getsize(fp) > 2000:
                with open(fp, "rb") as f:
                    frames[i] = f.read()
    return frames


def youtube_thumbnails(video_id):
    """Backup: YouTube's own 3 preview frames (about 25%, 50%, 75% into the video)."""
    out = []
    for name in ("hq1.jpg", "hq2.jpg", "hq3.jpg"):
        try:
            with urllib.request.urlopen(f"https://img.youtube.com/vi/{video_id}/{name}", timeout=10) as r:
                b = r.read()
            if len(b) > 2000:
                out.append(b)
        except Exception:
            pass
    return out


def attach_screenshots(video_id, slides):
    """Adds an 'image' key to slides. Returns a status message."""
    n = len(slides)
    if n == 0:
        return ""
    if n <= MAX_SHOTS:
        idxs = list(range(n))
    else:
        idxs = sorted({round(k * (n - 1) / (MAX_SHOTS - 1)) for k in range(MAX_SHOTS)})
    times = [slides[i].get("time", 0) for i in idxs]

    try:
        frames = grab_frames(video_id, times)
    except Exception:
        frames = {}
    if frames:
        for pos, data in frames.items():
            slides[idxs[pos]]["image"] = data
        return f"Video screenshots were added to {len(frames)} slides."

    thumbs = youtube_thumbnails(video_id)
    if thumbs:
        spots = [round(n * f) for f in (0.25, 0.5, 0.75)]
        for k, data in enumerate(thumbs):
            slides[min(spots[k], n - 1)]["image"] = data
        return ("Frames could not be taken from the video (YouTube blocked the server), "
                f"so {len(thumbs)} YouTube preview images were used instead.")
    return "Screenshots were not available, so the presentation was created without them."


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
            store.storage.from_(BUCKET).upload(
                path=name, file=data, file_options={"content-type": MIME}
            )
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
                "", {"limit": 50, "sortBy": {"column": "created_at", "order": "desc"}}
            )
            names = sorted((r["name"] for r in rows if r["name"].endswith(".pptx")), reverse=True)
            return [{"name": n, "data": None} for n in names[:30]]
        except Exception:
            return None
    return st.session_state.get("history", [])


def pretty(name):
    try:
        d = datetime.strptime(name[:15], "%Y%m%d-%H%M%S").strftime("%d %b %Y, %I:%M %p")
        title = name.split("__", 1)[1][:-5].replace("-", " ")
        return f"{d}  —  {title}"
    except Exception:
        return name


def show_history():
    with st.expander("📂 Previous presentations"):
        items = list_history()
        if items is None:
            st.warning("History could not be loaded. Please check the storage settings.")
            return
        if not get_store():
            st.caption("History is currently kept only for this browser session.")
        if not items:
            st.write("No presentations yet.")
        for it in items:
            st.write(pretty(it["name"]))
            data = it["data"]
            key = f"dl::{it['name']}"
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
# Website
# ----------------------------------------------------------------------
def password_gate():
    pw = get_secret("APP_PASSWORD")
    if not pw or st.session_state.get("authed"):
        return True
    entered = st.text_input("Password", type="password")
    if entered:
        if entered == pw:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def main():
    st.title("📚 Lecture to PPT")
    st.caption("Paste a YouTube lecture link and get a detailed, ready-to-use presentation.")

    if not password_gate():
        st.stop()

    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        st.error("The app is not configured yet: GEMINI_API_KEY is missing in Secrets.")
        st.stop()

    url = st.text_input("YouTube video link", placeholder="https://www.youtube.com/watch?v=...")
    with st.expander("Having trouble with a video? Paste the transcript instead"):
        manual = st.text_area("Transcript (optional)", height=150, key="manual_transcript")
        st.caption("On YouTube: open the video, expand the description, click \"Show transcript\", "
                   "select all the text and copy it here. Timestamps are supported.")
    c1, c2 = st.columns(2)
    lang_label = c1.selectbox("Slide language", list(LANGS.keys()))
    detail_label = c2.selectbox("Detail level", list(DETAIL.keys()), index=1)
    c3, c4 = st.columns(2)
    theme_name = c3.selectbox("Theme", list(THEMES.keys()))
    want_shots = c4.checkbox(
        "Add video screenshots to slides", value=False,
        help="Slower, and may not work on every server. If it fails, YouTube preview images are used.",
    )

    if st.button("Create presentation", type="primary") and url:
        vid = extract_video_id(url)
        if not vid:
            st.error("This does not look like a valid YouTube link. Please check it and try again.")
            st.stop()

        lang = LANGS[lang_label]
        rtl = lang_label.startswith("Urdu")
        per_chunk = DETAIL[detail_label]
        model = get_secret("GEMINI_MODEL", DEFAULT_MODEL)
        client = genai.Client(api_key=api_key)
        progress = st.progress(0, text="Fetching video details...")
        st.session_state.pop("shot_msg", None)

        try:
            video_title, channel = get_video_info(vid)

            progress.progress(5, text="Getting the transcript...")
            manual_text = (manual or "").strip()
            if len(manual_text) >= 200:
                snips, has_ts = parse_manual_transcript(manual_text)
                if not has_ts:
                    want_shots = False  # screenshots need real timestamps
            else:
                try:
                    snips = get_transcript(vid)
                except Exception as e:
                    progress.empty()
                    st.error(transcript_error_message(e))
                    with st.expander("Technical details"):
                        st.code(f"{type(e).__name__}: {str(e)[:1500]}")
                    st.stop()

            if sum(len(s[2]) for s in snips) < 200:
                progress.empty()
                st.error("The transcript is too short to create a presentation.")
                st.stop()

            chunks, truncated = split_chunks(snips)
            if truncated:
                st.warning("This video is very long, so only the first part is used.")
            video_end = chunks[-1]["end"] if chunks else 0

            all_slides, headings = [], []
            for i, ch in enumerate(chunks, 1):
                progress.progress(10 + int(65 * (i - 1) / len(chunks)),
                                  text=f"Creating slides... (part {i} of {len(chunks)})")
                part = slides_from_chunk(client, model, ch["text"], i, len(chunks), lang, per_chunk, headings)
                for j, sl in enumerate(part):
                    sl["time"] = ch["start"] + (j + 0.5) / len(part) * (ch["end"] - ch["start"])
                all_slides.extend(part)
                headings.extend(s.get("heading", "") for s in part)

            progress.progress(78, text="Creating summary and quiz...")
            extra = summary_and_quiz(client, model, all_slides, lang, video_title)

            shot_msg = ""
            if want_shots:
                if video_end > MAX_VIDEO_SECONDS:
                    shot_msg = "The video is longer than 2 hours, so screenshots were skipped."
                else:
                    progress.progress(82, text="Capturing video screenshots (this takes a while)...")
                    shot_msg = attach_screenshots(vid, all_slides)

            progress.progress(95, text="Building the PowerPoint file...")
            title = extra.get("title") or video_title or "Lecture"
            data = build_pptx(title, channel, f"https://www.youtube.com/watch?v={vid}",
                              all_slides, extra.get("summary", []), extra.get("quiz", []),
                              rtl, theme_name)

            hist_err = save_history(title, data)
            st.session_state["result"] = {"data": data, "title": title, "count": len(all_slides)}
            st.session_state["shot_msg"] = shot_msg
            if hist_err:
                st.warning(hist_err)
            progress.progress(100, text="Done!")
        except Exception as e:
            progress.empty()
            st.error(explain_error(e))
            with st.expander("Technical details"):
                st.code(f"{type(e).__name__}: {str(e)[:1500]}")

    res = st.session_state.get("result")
    if res:
        st.success(f"Your presentation is ready: {res['count']} content slides plus summary and quiz.")
        if st.session_state.get("shot_msg"):
            st.info(st.session_state["shot_msg"])
        safe = re.sub(r"[^\w\- ]", "", res["title"]).strip()[:60] or "lecture"
        st.download_button("⬇️ Download PPT", res["data"], file_name=f"{safe}.pptx", mime=MIME)
        st.info("Please review the slides before presenting. The speaker notes contain the full explanation.")

    show_history()
    st.divider()
    st.caption("The presentation credits the original channel. Please do not share someone else's content as your own.")


main()
