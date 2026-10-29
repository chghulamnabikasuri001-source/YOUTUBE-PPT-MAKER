"""
YouTube Video -> PPT Maker  (Gemini + Streamlit)
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

st.set_page_config(page_title="YouTube → PPT Maker", page_icon="📚", layout="centered")

DEFAULT_MODEL = "gemini-2.5-flash"   # free tier par chalta hai
CHUNK_CHARS = 9000                   # ~10 minute ka transcript ek hissa
MAX_TRANSCRIPT_CHARS = 150000        # bohat lambi video se bachne ke liye
MAX_SHOTS = 20                       # zyada se zyada kitne screenshots
MAX_VIDEO_SECONDS = 2 * 60 * 60      # 2 ghante se lambi video par screenshots nahi
BUCKET = "ppts"                      # Supabase storage bucket ka naam
PKT = timezone(timedelta(hours=5))

LANGS = {
    "English": "English",
    "Urdu (اردو script)": "Urdu (written in Urdu/Arabic script)",
    "Roman Urdu": "Roman Urdu (Urdu written in English letters)",
    "Video wali language": "the same language as the transcript",
}
DETAIL = {
    "Mukhtasar (kam slides)": "3-4",
    "Darmiyana": "5-7",
    "Tafseeli (zyada detail)": "8-10",
}


def get_secret(name, default=None):
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    return os.environ.get(name, default)


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


def get_transcript(video_id: str):
    """[(start_seconds, duration_seconds, text), ...] return karta hai."""
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
    last_err = None
    for attempt in range(4):
        try:
            r = client.models.generate_content(
                model=model,
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
            if any(k in msg for k in ("429", "RESOURCE_EXHAUSTED", "503", "UNAVAILABLE")):
                time.sleep(20 * (attempt + 1))  # free limit: thoda intezar
                continue
            raise
    raise last_err


def parse_json(raw: str):
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()
    a, b = raw.find("{"), raw.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("JSON nahi mila")
    return json.loads(raw[a : b + 1])


def ask_json(client, model, system, user, max_tokens=4000):
    for _ in range(2):
        raw = ask_ai(client, model, system, user, max_tokens)
        try:
            return parse_json(raw)
        except Exception:
            user += "\n\nPichla jawab valid JSON nahi tha. Sirf valid JSON return karo."
    raise ValueError("AI ne sahi format mein jawab nahi diya. Dobara try karo.")


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
    """Video ke di hue waqt (seconds) par frame nikalta hai. {position: jpg_bytes} return karta hai.
    yt-dlp se seedha stream ka link leta hai aur ffmpeg se sirf zaroori frame uthata hai."""
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
    """Backup: YouTube ke khud ke 3 frames (takreeban 25%, 50%, 75% par)."""
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


def attach_screenshots(video_id, slides, progress):
    """slides mein 'image' key lagata hai. Message return karta hai."""
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
        return f"{len(frames)} slides par video ke screenshots lag gaye."

    # backup
    thumbs = youtube_thumbnails(video_id)
    if thumbs:
        spots = [round(n * f) for f in (0.25, 0.5, 0.75)]
        for k, data in enumerate(thumbs):
            slides[min(spots[k], n - 1)]["image"] = data
        return ("Video se frames nahi nikal sake (server par YouTube ne rok diya hoga). "
                f"Is liye YouTube ki apni {len(thumbs)} tasveeren lagayi gayi hain.")
    return "Screenshots nahi mil sake, PPT bina screenshots ke bani hai."


# ----------------------------------------------------------------------
# History (Supabase ho to hamesha ke liye, warna sirf is session mein)
# ----------------------------------------------------------------------
MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"


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
            return f"History mein save nahi ho saki: {str(e)[:200]}"
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
    with st.expander("📂 Purani PPTs (History)"):
        items = list_history()
        if items is None:
            st.warning("History load nahi ho saki. Supabase ki settings check karo.")
            return
        if not get_store():
            st.caption("Abhi history sirf is browser session tak hai. Hamesha ke liye rakhne ke liye "
                       "README mein Supabase wala step karo.")
        if not items:
            st.write("Abhi koi PPT nahi bani.")
        for it in items:
            st.write(pretty(it["name"]))
            data = it["data"]
            key = f"dl::{it['name']}"
            if data is None:
                data = st.session_state.get(key)
                if data is None:
                    if st.button("Download ke liye tayyar karo", key="btn::" + it["name"]):
                        try:
                            st.session_state[key] = get_store().storage.from_(BUCKET).download(it["name"])
                            st.rerun()
                        except Exception as e:
                            st.error(f"File nahi mili: {str(e)[:150]}")
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
            st.error("Password ghalat hai.")
    return False


def main():
    st.title("📚 YouTube → PPT Maker")
    st.caption("YouTube lecture ka link paste karo, detailed PPT download karo.")

    if not password_gate():
        st.stop()

    api_key = get_secret("GEMINI_API_KEY")
    if not api_key:
        st.error("GEMINI_API_KEY set nahi hai. README mein 'Secrets' wala step dekho.")
        st.stop()

    url = st.text_input("YouTube video ka link", placeholder="https://www.youtube.com/watch?v=...")
    c1, c2 = st.columns(2)
    lang_label = c1.selectbox("Slides ki language", list(LANGS.keys()))
    detail_label = c2.selectbox("Detail level", list(DETAIL.keys()), index=1)
    c3, c4 = st.columns(2)
    theme_name = c3.selectbox("Theme / colour", list(THEMES.keys()))
    want_shots = c4.checkbox("Video ke screenshots slides mein lagao", value=False,
                             help="Slow hai aur server par kabhi kabhi nahi chalta. Na chale to YouTube ki 3 tasveeren lagti hain.")

    if st.button("PPT banao", type="primary") and url:
        vid = extract_video_id(url)
        if not vid:
            st.error("Ye YouTube link sahi nahi lag raha. Dobara check karo.")
            st.stop()

        lang = LANGS[lang_label]
        rtl = lang_label.startswith("Urdu")
        per_chunk = DETAIL[detail_label]
        model = get_secret("GEMINI_MODEL", DEFAULT_MODEL)
        client = genai.Client(api_key=api_key)
        progress = st.progress(0, text="Video ki info la rahe hain...")
        st.session_state.pop("shot_msg", None)

        try:
            video_title, channel = get_video_info(vid)

            progress.progress(5, text="Transcript nikal rahe hain...")
            try:
                snips = get_transcript(vid)
            except Exception as e:
                st.error("Is video ka transcript nahi mil saka. Ho sakta hai captions band hon, "
                         "ya YouTube ne server ko block kiya ho.")
                with st.expander("Technical detail"):
                    st.code(str(e)[:1500])
                st.stop()

            if sum(len(s[2]) for s in snips) < 200:
                st.error("Transcript bohat chhota hai, isse PPT nahi ban sakti.")
                st.stop()

            chunks, truncated = split_chunks(snips)
            if truncated:
                st.warning("Video bohat lambi hai, sirf shuru ka hissa use ho raha hai.")
            video_end = chunks[-1]["end"] if chunks else 0

            all_slides, headings = [], []
            for i, ch in enumerate(chunks, 1):
                progress.progress(10 + int(65 * (i - 1) / len(chunks)),
                                  text=f"AI slides bana raha hai... (hissa {i}/{len(chunks)})")
                part = slides_from_chunk(client, model, ch["text"], i, len(chunks), lang, per_chunk, headings)
                for j, sl in enumerate(part):
                    sl["time"] = ch["start"] + (j + 0.5) / len(part) * (ch["end"] - ch["start"])
                all_slides.extend(part)
                headings.extend(s.get("heading", "") for s in part)

            progress.progress(78, text="Summary aur quiz bana rahe hain...")
            extra = summary_and_quiz(client, model, all_slides, lang, video_title)

            shot_msg = ""
            if want_shots:
                if video_end > MAX_VIDEO_SECONDS:
                    shot_msg = "Video 2 ghante se lambi hai, isliye screenshots nahi lagaye."
                else:
                    progress.progress(82, text="Video ke screenshots nikal rahe hain (thoda waqt lagega)...")
                    shot_msg = attach_screenshots(vid, all_slides, progress)

            progress.progress(95, text="PPT file bana rahe hain...")
            title = extra.get("title") or video_title or "Lecture"
            data = build_pptx(title, channel, f"https://www.youtube.com/watch?v={vid}",
                              all_slides, extra.get("summary", []), extra.get("quiz", []),
                              rtl, theme_name)

            hist_err = save_history(title, data)
            st.session_state["result"] = {"data": data, "title": title, "count": len(all_slides)}
            st.session_state["shot_msg"] = shot_msg
            if hist_err:
                st.warning(hist_err)
            progress.progress(100, text="Ho gaya!")
        except Exception as e:
            st.error("Kuch masla aa gaya. Dobara try karo.")
            with st.expander("Technical detail"):
                st.code(str(e)[:1500])

    res = st.session_state.get("result")
    if res:
        st.success(f"PPT tayyar hai: {res['count']} content slides (+ summary aur quiz)")
        if st.session_state.get("shot_msg"):
            st.info(st.session_state["shot_msg"])
        safe = re.sub(r"[^\w\- ]", "", res["title"]).strip()[:60] or "lecture"
        st.download_button("⬇️ PPT download karo", res["data"], file_name=f"{safe}.pptx", mime=MIME)
        st.info("Class mein dikhane se pehle slides ek baar parh lein. Speaker notes mein poori detail hai.")

    show_history()
    st.divider()
    st.caption("Tip: PPT mein channel ka credit shamil hota hai. Kisi aur ka content apne naam se share na karein.")


main()
