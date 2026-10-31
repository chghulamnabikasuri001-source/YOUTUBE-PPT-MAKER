"""
Core logic (no Streamlit code here): YouTube transcript, Gemini slide generation,
parallel processing and smart screenshots.
"""
import json
import os
import re
import subprocess
import tempfile
import time
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

try:
    from google.genai import types
except Exception:  # lets this module be imported where the SDK is missing (tests)
    types = None

DEFAULT_MODEL = "gemini-3.8-flash"
FALLBACK_MODELS = ["gemini-3.6-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite"]
CHUNK_CHARS = 14000            # about 15 minutes of lecture per AI call
MAX_TRANSCRIPT_CHARS = 150000
SEGMENT_SECONDS = 600          # video mode: minutes of video per AI call
MAX_VIDEO_SECONDS = 2 * 60 * 60
MAX_SHOTS = 12                 # most screenshots we try to capture
PARALLEL = 3                   # AI calls running at the same time

LANGS = {
    "English": "English",
    "Urdu (اردو)": "Urdu (written in Urdu/Arabic script)",
    "Roman Urdu": "Roman Urdu (Urdu written in English letters)",
    "Same as the video": "the same language as the video",
}
DETAIL = {
    "Brief": "4-5",
    "Standard": "6-8",
    "Detailed": "9-12",
}


# ----------------------------------------------------------------------
# Friendly errors
# ----------------------------------------------------------------------
def explain_error(e) -> str:
    m = str(e)
    low = m.lower()
    if "api_key_invalid" in low or "api key not valid" in low or "permission_denied" in low or "unauthenticated" in low:
        return "The AI service rejected the API key. Check GEMINI_API_KEY in the app's Secrets."
    if "429" in m or "resource_exhausted" in low or "quota" in low:
        return ("The free AI usage limit has been reached. Please wait a minute and try again, "
                "or try a shorter video / lower detail level.")
    if "404" in m or "not_found" in low:
        return ("The selected AI model is not available. Remove GEMINI_MODEL from Secrets, "
                "or set it to a current model such as gemini-3.8-flash.")
    if "503" in m or "unavailable" in low or "overloaded" in low:
        return "The AI service is busy right now. Please try again in a minute."
    if "invalid_argument" in low and ("video" in low or "youtube" in low or "url" in low):
        return "The AI could not access this video. It must be a public YouTube video (not private or age-restricted)."
    if "valid json" in low or "format" in low:
        return "The AI returned an unexpected format. Please try again."
    return "Something went wrong while generating the slides. Please try again."


def transcript_error_message(e) -> str:
    name = type(e).__name__
    if name in ("TranscriptsDisabled", "NoTranscriptFound"):
        return "This video has no captions, so a transcript is not available."
    if name in ("VideoUnavailable", "VideoUnplayable", "InvalidVideoId"):
        return "This video is unavailable or private."
    return "The transcript could not be retrieved."


# ----------------------------------------------------------------------
# YouTube helpers
# ----------------------------------------------------------------------
def extract_video_id(url: str):
    url = (url or "").strip()
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


def clean_title(t: str) -> str:
    t = (t or "").strip()
    if len(t) > 70:
        for sep in (" | ", " - ", " – ", " : "):
            if sep in t:
                t = t.split(sep)[0]
                break
    return t[:80].strip() or "Lecture"


def transcript_via_api(video_id, proxy=None):
    from youtube_transcript_api import YouTubeTranscriptApi

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


def transcript_via_ytdlp(video_id, proxy=None):
    from yt_dlp import YoutubeDL

    opts = {"quiet": True, "no_warnings": True, "skip_download": True, "socket_timeout": 20}
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


def get_transcript(video_id, proxy=None):
    try:
        return transcript_via_api(video_id, proxy)
    except Exception as first_error:
        try:
            return transcript_via_ytdlp(video_id, proxy)
        except Exception:
            raise first_error


def parse_manual_transcript(text: str):
    """Parses text copied from YouTube's 'Show transcript'. Returns (snippets, has_timestamps)."""
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
            cur_time = n * 5.0
        snips.append((float(cur_time), 5.0, line))
        n += 1
    return snips, has_ts


def split_chunks(snips, size=CHUNK_CHARS, max_chars=MAX_TRANSCRIPT_CHARS):
    chunks, cur, n, total, truncated = [], [], 0, 0, False

    def flush():
        if cur:
            chunks.append({"text": " ".join(x[2] for x in cur), "start": cur[0][0],
                           "end": max(cur[-1][0] + cur[-1][1], cur[0][0] + 1)})

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


def get_video_duration(video_id, proxy=None):
    try:
        req = urllib.request.Request(
            f"https://www.youtube.com/watch?v={video_id}",
            headers={"User-Agent": "Mozilla/5.0", "Accept-Language": "en-US,en;q=0.9"},
        )
        with urllib.request.urlopen(req, timeout=10) as r:
            html = r.read().decode("utf-8", "ignore")
        m = re.search(r'"lengthSeconds":"(\d+)"', html)
        if m:
            return int(m.group(1))
    except Exception:
        pass
    try:
        from yt_dlp import YoutubeDL

        opts = {"quiet": True, "no_warnings": True, "skip_download": True, "socket_timeout": 15}
        if proxy:
            opts["proxy"] = proxy
        with YoutubeDL(opts) as ydl:
            info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=False)
        return int(info.get("duration") or 0) or None
    except Exception:
        return None


# ----------------------------------------------------------------------
# Gemini
# ----------------------------------------------------------------------
_UNAVAILABLE = {}  # model -> time until which it is skipped


def _config(system, max_tokens, think):
    kw = dict(system_instruction=system, max_output_tokens=max_tokens + 4000,
              response_mime_type="application/json")
    if think:
        try:
            kw["thinking_config"] = types.ThinkingConfig(thinking_level="low")  # faster answers
        except Exception:
            pass
    return types.GenerateContentConfig(**kw)


def ask_ai(client, model, system, contents, max_tokens=4000):
    """Calls Gemini. Falls back to other free models if one is missing or its quota is used up."""
    models = [model] + [m for m in FALLBACK_MODELS if m != model]
    last_err = None
    for round_no in range(2):
        retryable_seen = False
        for mdl in models:
            if _UNAVAILABLE.get(mdl, 0) > time.time():
                continue
            try:
                try:
                    r = client.models.generate_content(model=mdl, contents=contents,
                                                       config=_config(system, max_tokens, True))
                except Exception as e:
                    if "thinking" in str(e).lower():
                        r = client.models.generate_content(model=mdl, contents=contents,
                                                           config=_config(system, max_tokens, False))
                    else:
                        raise
                return r.text or ""
            except Exception as e:
                last_err = e
                msg = str(e)
                if "404" in msg or "NOT_FOUND" in msg:
                    _UNAVAILABLE[mdl] = time.time() + 3600
                elif "429" in msg or "RESOURCE_EXHAUSTED" in msg:
                    _UNAVAILABLE[mdl] = time.time() + 300
                    retryable_seen = True
                elif "503" in msg or "UNAVAILABLE" in msg or "overloaded" in msg.lower():
                    retryable_seen = True
                else:
                    raise
        if round_no == 0 and retryable_seen:
            time.sleep(20)
            _UNAVAILABLE.clear()
            continue
        break
    raise last_err or RuntimeError("No AI model is available.")


def parse_json(raw: str):
    raw = (raw or "").strip()
    raw = re.sub(r"^```(?:json)?", "", raw).strip()
    raw = re.sub(r"```$", "", raw).strip()
    a, b = raw.find("{"), raw.rfind("}")
    if a == -1 or b == -1:
        raise ValueError("No JSON found")
    return json.loads(raw[a: b + 1])


def ask_json(client, model, system, user, max_tokens=4000, media=None):
    for _ in range(2):
        contents = [media, user] if media is not None else user
        raw = ask_ai(client, model, system, contents, max_tokens)
        try:
            return parse_json(raw)
        except Exception:
            user += "\n\nYour previous answer was not valid JSON. Return only valid JSON."
    raise ValueError("The AI did not return a valid JSON format.")


# ----------------------------------------------------------------------
# Prompts
# ----------------------------------------------------------------------
SCHEMA_HELP = """
Each slide is a JSON object:
{
 "heading": "short, catchy title (max 8 words)",
 "bullets": ["0-4 SHORT lines (max 12 words each). Use **bold** for the key term, e.g. \\"**AI**: making machines smart\\""],
 "visual": null OR exactly one of:
   {"type":"nested","items":["Outer set","Middle set","Inner set"]}                 // sets inside sets, 2-4 items, outermost first
   {"type":"flow","items":["Step 1","Step 2","Step 3"],"direction":"vertical"}      // process, pipeline, cause->effect (2-6 items; direction vertical or horizontal)
   {"type":"cards","items":[{"icon":"🐕","title":"Short title","text":"One short line"}]}   // 2-4 options / categories / examples, one emoji icon each
   {"type":"steps","items":[{"title":"Layer 1","text":"Simple lines"}]}              // ordered stages (2-5)
   {"type":"table","headers":["","Option A","Option B"],"rows":[["Feature","x","y"]]} // comparison table, max 6 rows, max 4 columns
   {"type":"tree","root":{"label":"Topic","children":[{"label":"Part","children":[{"label":"Sub-part"}]}]}}  // classification / hierarchy / taxonomy, max 3 levels
   {"type":"network","layers":[4,6,6,2],"labels":["Input","Hidden layers","Output"]}  // ONLY for neural networks
   {"type":"code","language":"python","code":"line1\\nline2"}                          // code that is shown or explained (max 12 short lines)
   {"type":"compare","left":{"title":"A","points":["..."]},"right":{"title":"B","points":["..."]}}   // two-sided contrast
 "footer": "optional one-line key takeaway (use it with table / cards / steps / tree / compare)",
 "notes": "speaker notes: 3-6 sentences with the FULL explanation, examples and details the teacher gave",
 "position": 0.0 to 1.0   // where in this part of the lecture this topic is discussed
 "show_board": true or false  // true ONLY when, at that moment, handwritten/typed notes, a whiteboard/blackboard, a text slide or code is clearly on screen and worth capturing
}
"""


def slide_system(lang, per_chunk, video_mode, first):
    src = ("one segment of a lecture VIDEO (use what the teacher says AND what is shown on screen: text, formulas, diagrams, code)"
           if video_mode else "one part of a lecture TRANSCRIPT (auto-captions may contain spelling mistakes; fix them silently)")
    cover, tail = "", ""
    if first:
        cover = ('\n- Because this is the FIRST part, also return "tagline" (a short subtitle, max 6 words) and '
                 '"cover" (one visual object of type nested, flow, tree or network that symbolises the whole lecture topic).')
        tail = ', "tagline": "...", "cover": {...}'
    return f"""You are an expert teacher and presentation designer. You receive {src}.
Turn it into a visually rich, engaging classroom slide deck - like a polished explainer, NOT a transcript.

{SCHEMA_HELP}
Rules:
- Make about {per_chunk} slides. Cover every important concept, definition, example, formula and step; put the fine detail in "notes".
- Slide text must be short and punchy. Never copy sentences from the transcript.
- At least 70% of slides MUST have a visual. Pick the visual that fits the content best and vary them (never the same type 3 times in a row).
- With table / cards / steps / tree / compare use NO bullets and use "footer" instead. With nested / flow / network / code use 2-3 bullets.
- Use real, specific labels in visuals (not "Item 1"). Emoji icons must be simple and relevant.
- If a slide is about code or notes shown on screen, set "visual" to null (or "code") and set show_board=true.
- Do not invent facts. Skip greetings, subscribe requests and ads.
- Output language for all text: {lang}.
- Also return "takeaways": 2-3 short key takeaways for this part, and "quiz": 1-2 questions [{{"q":"...","a":"..."}}] for this part.{cover}
- Return ONLY valid JSON: {{"slides": [...], "takeaways": ["..."], "quiz": [{{"q":"...","a":"..."}}]{tail}}}"""


def _norm_part(d):
    if not isinstance(d, dict):
        d = {}
    slides = []
    for s in d.get("slides") or []:
        if not isinstance(s, dict):
            continue
        try:
            pos = float(s.get("position"))
        except Exception:
            pos = None
        slides.append(dict(
            heading=str(s.get("heading", "")).strip(),
            bullets=[str(b).strip() for b in (s.get("bullets") or []) if str(b).strip()][:5],
            visual=s.get("visual") if isinstance(s.get("visual"), dict) else None,
            footer=str(s.get("footer") or "").strip(),
            notes=str(s.get("notes") or "").strip(),
            position=min(max(pos, 0.0), 1.0) if pos is not None else None,
            show_board=bool(s.get("show_board")),
        ))
    quiz = [dict(q=str(x.get("q", "")).strip(), a=str(x.get("a", "")).strip())
            for x in (d.get("quiz") or []) if isinstance(x, dict) and x.get("q")]
    return dict(slides=slides,
                takeaways=[str(t).strip() for t in (d.get("takeaways") or []) if str(t).strip()],
                quiz=quiz, tagline=str(d.get("tagline") or "").strip(),
                cover=d.get("cover") if isinstance(d.get("cover"), dict) else None)


def slides_from_chunk(client, model, text, idx, total, lang, per_chunk):
    system = slide_system(lang, per_chunk, False, idx == 1)
    user = f"This is part {idx} of {total} of the lecture.\n\nTRANSCRIPT PART:\n{text}"
    return _norm_part(ask_json(client, model, system, user, max_tokens=9000))


def slides_from_video(client, model, video_id, start, end, idx, total, lang, per_chunk):
    system = slide_system(lang, per_chunk, True, idx == 1)
    user = f"This is segment {idx} of {total} of the lecture (minutes {start // 60} to {end // 60})."
    media = types.Part(
        file_data=types.FileData(file_uri=f"https://www.youtube.com/watch?v={video_id}"),
        video_metadata=types.VideoMetadata(start_offset=f"{int(start)}s", end_offset=f"{int(end)}s"),
    )
    return _norm_part(ask_json(client, model, system, user, max_tokens=9000, media=media))


def run_parallel(funcs, workers=PARALLEL, on_progress=None):
    """Runs funcs in threads. Returns (results, errors) in the original order."""
    results, errors = [None] * len(funcs), [None] * len(funcs)
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(f): i for i, f in enumerate(funcs)}
        for fut in as_completed(futs):
            i = futs[fut]
            try:
                results[i] = fut.result()
            except Exception as e:
                errors[i] = e
            done += 1
            if on_progress:
                on_progress(done, len(funcs))
    return results, errors


def _spread(items, n):
    if len(items) <= n:
        return items
    idx = sorted({round(k * (len(items) - 1) / (n - 1)) for k in range(n)})
    return [items[i] for i in idx]


def merge_parts(parts):
    """parts: list of (part_dict, start_seconds, end_seconds) in lecture order."""
    slides, takeaways, quiz, tagline, cover = [], [], [], "", None
    for part, start, end in parts:
        n = len(part["slides"]) or 1
        for j, s in enumerate(part["slides"]):
            if not s["heading"]:
                continue
            pos = s["position"] if s["position"] is not None else (j + 0.5) / n
            s = dict(s)
            s["time"] = start + pos * (end - start)
            if slides and slides[-1]["heading"].lower() == s["heading"].lower():
                continue
            slides.append(s)
        takeaways += part["takeaways"]
        quiz += part["quiz"]
        tagline = tagline or part["tagline"]
        cover = cover or part["cover"]
    seen, tk = set(), []
    for t in takeaways:
        if t.lower() not in seen:
            seen.add(t.lower())
            tk.append(t)
    return dict(slides=slides, takeaways=_spread(tk, 6), quiz=_spread(quiz, 5), tagline=tagline, cover=cover)


# ----------------------------------------------------------------------
# Screenshots: only notes / board / code, never the teacher
# ----------------------------------------------------------------------
def grab_frames(video_id, times, proxy=None, workers=4):
    """One frame per time (seconds). Returns {position: jpg_bytes}."""
    from yt_dlp import YoutubeDL

    opts = {"format": "best[height<=480][ext=mp4]/best[height<=480]/best",
            "quiet": True, "no_warnings": True, "noplaylist": True, "socket_timeout": 20}
    if proxy:
        opts["proxy"] = proxy
    with YoutubeDL(opts) as ydl:
        info = ydl.extract_info(f"https://www.youtube.com/watch?v={video_id}", download=False)
    stream = info.get("url")
    if not stream:
        return {}
    headers = info.get("http_headers", {}) or {}
    hdr = "".join(f"{k}: {v}\r\n" for k, v in headers.items())

    with tempfile.TemporaryDirectory() as tmp:
        def one(i, t):
            fp = os.path.join(tmp, f"f{i}.jpg")
            cmd = ["ffmpeg", "-y", "-loglevel", "error"]
            if hdr:
                cmd += ["-headers", hdr]
            cmd += ["-ss", str(int(t)), "-i", stream, "-frames:v", "1", "-vf", "scale=960:-2", "-q:v", "4", fp]
            try:
                subprocess.run(cmd, capture_output=True, timeout=45)
            except Exception:
                return None
            if os.path.exists(fp) and os.path.getsize(fp) > 2000:
                with open(fp, "rb") as f:
                    return f.read()
            return None

        frames = {}
        with ThreadPoolExecutor(max_workers=workers) as ex:
            futs = {ex.submit(one, i, t): i for i, t in enumerate(times)}
            for fut in as_completed(futs):
                data = fut.result()
                if data:
                    frames[futs[fut]] = data
    return frames


def keep_board_frames(client, model, frames):
    """Asks Gemini which frames show notes / board / slide text / code. Returns a set of frame keys."""
    keys = sorted(frames)
    if not keys:
        return set()
    contents = []
    for n, k in enumerate(keys):
        contents.append(types.Part.from_bytes(data=frames[k], mime_type="image/jpeg"))
        contents.append(f"Image {n}")
    contents.append("Which images should be kept?")
    system = ("You review frames from a lecture video. KEEP an image only if its main content is written or typed "
              "notes, a whiteboard/blackboard, a presentation slide with text or formulas, code, or a diagram. "
              "DROP images that mainly show a person (teacher's face or body), a room, a title/intro screen, or are blurry. "
              'Return ONLY JSON: {"keep": [indexes of the images to keep]}')
    try:
        data = parse_json(ask_ai(client, model, system, contents, max_tokens=500))
        return {keys[int(i)] for i in data.get("keep", []) if 0 <= int(i) < len(keys)}
    except Exception:
        return set()  # when unsure, add nothing


def attach_screenshots(client, model, video_id, slides, proxy=None):
    """Adds 'image' to slides that show notes/code. Returns (count, message)."""
    cand = [i for i, s in enumerate(slides) if s.get("show_board")]
    if not cand:
        return 0, "No notes or code were shown on screen, so no screenshots were added."
    cand = _spread(cand, MAX_SHOTS)
    try:
        frames = grab_frames(video_id, [slides[i]["time"] for i in cand], proxy)
    except Exception:
        frames = {}
    if not frames:
        return 0, "Screenshots were not available (YouTube blocked the server), so none were added."
    keep = keep_board_frames(client, model, frames)
    for pos in keep:
        slides[cand[pos]]["image"] = frames[pos]
    if not keep:
        return 0, "The captured frames showed no readable notes or code, so none were added."
    return len(keep), f"Added {len(keep)} screenshot(s) of notes/code from the video."


# ----------------------------------------------------------------------
# Whole pipeline (called by the website)
# ----------------------------------------------------------------------
class UserError(Exception):
    """An expected problem with a clear message for the user."""

    def __init__(self, msg, detail=""):
        super().__init__(msg)
        self.msg, self.detail = msg, detail


def _retry_failed(funcs, results, errors):
    for i, e in enumerate(errors):
        if e is not None:
            try:
                results[i] = funcs[i]()
                errors[i] = None
            except Exception as e2:
                errors[i] = e2


def generate_presentation(url, manual_text, lang_label, detail_label, theme_name, want_shots,
                          client, model, proxy=None, say=None, prog=None, warn=None):
    from pptx_builder import build_pptx, norm_visual

    say = say or (lambda m: None)
    prog = prog or (lambda f, t: None)
    warn = warn or (lambda m: None)

    vid = extract_video_id(url)
    if not vid:
        raise UserError("This does not look like a valid YouTube link. Please check it and try again.")
    lang, rtl, per = LANGS[lang_label], lang_label.startswith("Urdu"), DETAIL[detail_label]

    say("🔎 Reading the video details")
    prog(0.04, "Reading the video details...")
    video_title, channel = get_video_info(vid)

    notes, mode, snips = [], "transcript", None
    manual = (manual_text or "").strip()
    if len(manual) >= 200:
        snips, has_ts = parse_manual_transcript(manual)
        if not has_ts:
            want_shots = False
            notes.append("Screenshots need timestamps, so they were skipped for the pasted transcript.")
    else:
        say("📝 Getting the transcript")
        prog(0.08, "Getting the transcript...")
        try:
            snips = get_transcript(vid, proxy)
        except Exception as e:
            if type(e).__name__ in ("VideoUnavailable", "VideoUnplayable", "InvalidVideoId"):
                raise UserError(transcript_error_message(e), str(e))
            mode = "video"
            notes.append("The transcript could not be retrieved, so the AI watched the video directly. "
                         "This is slower for long videos.")

    parts, video_end = [], 0
    if mode == "transcript":
        if sum(len(x[2]) for x in snips) < 200:
            raise UserError("The transcript is too short to create a presentation.")
        chunks, truncated = split_chunks(snips)
        if truncated:
            warn("This video is very long, so only the first part is used.")
        video_end = chunks[-1]["end"]
        say(f"🎨 Designing slides ({len(chunks)} part{'s' if len(chunks) > 1 else ''} at the same time)")
        funcs = [(lambda ch=ch, i=i: slides_from_chunk(client, model, ch["text"], i, len(chunks), lang, per))
                 for i, ch in enumerate(chunks, 1)]
        results, errors = run_parallel(
            funcs, on_progress=lambda d, n: prog(0.1 + 0.6 * d / n, f"Designing slides... {d} of {n} parts done"))
        _retry_failed(funcs, results, errors)
        parts = [(results[i], chunks[i]["start"], chunks[i]["end"]) for i in range(len(chunks)) if results[i]]
        if not parts:
            raise next(e for e in errors if e)
        if len(parts) < len(chunks):
            warn(f"{len(chunks) - len(parts)} part(s) of the lecture could not be processed and were skipped.")
    else:
        say("🎬 Watching the video with AI")
        duration = get_video_duration(vid, proxy)
        limit = min(duration, MAX_VIDEO_SECONDS) if duration else MAX_VIDEO_SECONDS
        if duration and duration > MAX_VIDEO_SECONDS:
            warn("This video is longer than 2 hours, so only the first 2 hours are used.")
        starts = list(range(0, int(limit), SEGMENT_SECONDS))
        spans = [(s0, min(s0 + SEGMENT_SECONDS, int(limit))) for s0 in starts]
        if duration:
            funcs = [(lambda k=k, sp=sp: slides_from_video(client, model, vid, sp[0], sp[1], k, len(spans), lang, per))
                     for k, sp in enumerate(spans, 1)]
            results, errors = run_parallel(
                funcs, on_progress=lambda d, n: prog(0.1 + 0.6 * d / n, f"Watching the video... {d} of {n} parts done"))
            _retry_failed(funcs, results, errors)
            parts = [(results[i], spans[i][0], spans[i][1]) for i in range(len(spans)) if results[i]]
            if not parts:
                raise next(e for e in errors if e)
            if len(parts) < len(spans):
                warn(f"{len(spans) - len(parts)} part(s) of the video could not be processed and were skipped.")
        else:
            for k, (a, b) in enumerate(spans, 1):
                prog(min(0.1 + 0.05 * k, 0.7), f"Watching the video... part {k}")
                try:
                    part = slides_from_video(client, model, vid, a, b, k, len(spans), lang, per)
                except Exception:
                    if parts:
                        break  # probably past the end of the video
                    raise
                if not part["slides"]:
                    break
                parts.append((part, a, b))
        video_end = max((p[2] for p in parts), default=0)

    merged = merge_parts(parts)
    if not merged["slides"]:
        raise UserError("No slides could be created from this video.")

    shots = 0
    if want_shots:
        if video_end > MAX_VIDEO_SECONDS:
            notes.append("The video is longer than 2 hours, so screenshots were skipped.")
        else:
            say("📸 Looking for notes and code on screen")
            prog(0.75, "Capturing notes and code from the video...")
            shots, msg = attach_screenshots(client, model, vid, merged["slides"], proxy)
            notes.append(msg)

    say("📦 Building your PowerPoint")
    prog(0.92, "Building your PowerPoint...")
    title = clean_title(video_title)
    data = build_pptx(title, merged["tagline"], channel, f"https://www.youtube.com/watch?v={vid}",
                      merged["slides"], merged["takeaways"], merged["quiz"], rtl, theme_name, merged["cover"])
    outline = [(s["heading"], (norm_visual(s["visual"]) or {}).get("type") if s["visual"] else None,
                bool(s.get("image"))) for s in merged["slides"]]
    prog(1.0, "Done!")
    return dict(data=data, title=title, slides=len(merged["slides"]), visuals=sum(1 for o in outline if o[1]),
                shots=shots, outline=outline, notes=notes)
