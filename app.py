from flask import Flask, request, render_template, jsonify, send_file
import requests
import anthropic
import json
import os
import re
import io
import time
import threading
import uuid
from datetime import datetime, timezone
from collections import Counter
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.chart.data import ChartData
from pptx.enum.chart import XL_CHART_TYPE
from bs4 import BeautifulSoup

load_dotenv()
SEARCHAPI_KEY     = os.getenv("SEARCHAPI_KEY")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY")
APIFY_API_KEY     = os.getenv("APIFY_API_KEY")

app = Flask(__name__)

TOTAL_SLIDES = 26

# ═══════════════════════════════════════════════════════════════
# MIDNIGHT PANDA BRAND DESIGN TOKENS
# ═══════════════════════════════════════════════════════════════

BG_DARK        = RGBColor(0x0A, 0x0A, 0x0A)
BG_LIGHT       = RGBColor(0xF5, 0xF3, 0xEF)
GOLD           = RGBColor(0xB8, 0x94, 0x5A)
CARD_LIGHT     = RGBColor(0xEC, 0xE9, 0xE2)
CARD_DARK      = RGBColor(0x1A, 0x1A, 0x1A)
TEXT_DARK      = RGBColor(0x0A, 0x0A, 0x0A)
TEXT_LIGHT     = RGBColor(0xF5, 0xF3, 0xEF)
TEXT_GRAY      = RGBColor(0x66, 0x66, 0x66)
TEXT_GRAY_LT   = RGBColor(0xD0, 0xCF, 0xC9)
TEXT_FOOTER    = RGBColor(0x99, 0x99, 0x99)
PIE_TAN        = RGBColor(0xD9, 0xC4, 0xA0)
MOMENTUM_UP    = RGBColor(0x6F, 0xA8, 0x6B)
MOMENTUM_DOWN  = RGBColor(0xC1, 0x6E, 0x6E)

FONT_MONO      = "DM Mono"
FONT_MONO_MED  = "DM Mono Medium"
FONT_MONO_LT   = "DM Mono Light"
FONT_SERIF     = "Playfair Display Black"
FONT_STAT      = "Bebas Neue"


# ═══════════════════════════════════════════════════════════════
# STEP 1 — DISCOVER SOCIAL HANDLES
# ═══════════════════════════════════════════════════════════════

def clean_domain(url):
    url = url.strip().rstrip("/")
    if not url.startswith("http"):
        url = "https://" + url
    return url

def is_youtube_channel_id(handle):
    raw = handle.lstrip("@")
    return raw.startswith("UC") and len(raw) >= 20

def extract_social_from_html(html):
    soup = BeautifulSoup(html, "html.parser")
    handles = {}
    SKIP = {"p","explore","accounts","stories","reel","tv","share","sharer",
            "login","signup","intent","home","search"}
    for tag in soup.find_all("a", href=True):
        href = tag["href"].lower()

        if "instagram.com" in href and "instagram" not in handles:
            m = re.search(r'instagram\.com/([A-Za-z0-9_.]+)/?', href, re.I)
            if m and m.group(1).lower() not in SKIP:
                handles["instagram"] = m.group(1)

        if "facebook.com" in href and "facebook" not in handles:
            if "profile.php" not in href:
                m = re.search(r'facebook\.com/([A-Za-z0-9_.]+)/?', href, re.I)
                if m and m.group(1).lower() not in SKIP:
                    handles["facebook"] = m.group(1)

        if "linkedin.com/company" in href and "linkedin" not in handles:
            m = re.search(r'linkedin\.com/company/([A-Za-z0-9_.-]+)/?', href, re.I)
            if m:
                handles["linkedin"] = m.group(1)

        if "youtube.com" in href and "youtube" not in handles:
            m = re.search(r'youtube\.com/@([A-Za-z0-9_.]+)/?', href, re.I)
            if m:
                handles["youtube"] = "@" + m.group(1)
            else:
                m = re.search(r'youtube\.com/(?:channel|c|user)/([A-Za-z0-9_.-]+)/?', href, re.I)
                if m:
                    handles["youtube"] = m.group(1)
    return handles

def find_handles_from_website(website_url):
    base = clean_domain(website_url)
    pages = [base, base+"/contact", base+"/about",
             base+"/contact-us", base+"/about-us"]
    headers = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Chrome/120.0.0.0"}
    all_handles = {}
    for page in pages:
        try:
            r = requests.get(page, headers=headers, timeout=10)
            if r.status_code == 200:
                found = extract_social_from_html(r.text)
                all_handles.update(found)
                if len(all_handles) >= 4:
                    break
        except:
            continue
    return all_handles

def find_handles_from_searchapi(website_url, missing_platforms):
    domain = website_url.replace("https://","").replace("http://","").rstrip("/")
    brand  = re.sub(r'\.(com|in|io|co|net|org|app).*','', domain).lower()
    brand  = brand.replace("www.","")
    handles = {}
    SKIP = {"p","explore","accounts","stories","reel","tv","share",
            "sharer","login","signup","intent","home","search"}

    brand_words = re.split(r'[-_.]', brand)
    print(f"   Brand: '{brand}' | Words: {brand_words} | Domain: '{domain}'")

    queries = {
        "instagram": [
            f'{domain} instagram',
            f'"{brand}" instagram official',
            f'site:instagram.com "{brand}"',
        ],
        "facebook": [
            f'"{brand}" facebook page',
            f'{domain} facebook',
            f'site:facebook.com "{brand}"',
        ],
        "linkedin": [
            f'"{brand}" site:linkedin.com/company',
            f'{domain} linkedin',
            f'"{brand}" linkedin company followers',
        ],
        "youtube": [
            f'"{brand}" youtube channel',
            f'{domain} youtube',
        ],
    }

    for platform in missing_platforms:
        found = False
        for q in queries[platform]:
            if found:
                break
            try:
                print(f"   🔍 {platform}: {q}")
                r = requests.get(
                    "https://www.searchapi.io/api/v1/search",
                    params={"engine":"google","q":q,
                            "api_key":SEARCHAPI_KEY,"num":5}
                )
                if r.status_code != 200:
                    continue

                results = r.json().get("organic_results",[])
                for idx, result in enumerate(results):
                    link    = result.get("link","")
                    title   = result.get("title","").lower()
                    snippet = result.get("snippet","").lower()

                    def brand_match(text):
                        text_clean = text.replace("-","").replace("_","").replace(" ","").lower()
                        brand_clean = brand.replace("-","").replace("_","").replace(" ","")
                        if brand_clean in text_clean:
                            return True
                        for word in brand_words:
                            if len(word) > 3 and word in text_clean:
                                return True
                        return False

                    if platform == "instagram" and "instagram.com" in link:
                        m = re.search(r'instagram\.com/([A-Za-z0-9_.]+)/?', link, re.I)
                        if m:
                            handle = m.group(1)
                            if handle.lower() in SKIP:
                                continue
                            if brand_match(handle) or brand_match(title) or brand_match(snippet):
                                handles["instagram"] = handle
                                print(f"   ✅ Instagram: @{handle}")
                                found = True; break
                            elif idx == 0 and (domain.replace("www.","").split(".")[0] in handle.lower()
                                               or brand.split("-")[0] in handle.lower()):
                                handles["instagram"] = handle
                                print(f"   ✅ Instagram (domain match): @{handle}")
                                found = True; break
                            else:
                                print(f"   ⚠️ Instagram skipped: @{handle}")

                    elif platform == "facebook" and "facebook.com" in link:
                        if "profile.php" in link:
                            print(f"   ⚠️ Facebook skipped profile.php")
                            continue
                        m = re.search(r'facebook\.com/([A-Za-z0-9_.]+)/?', link, re.I)
                        if m:
                            handle = m.group(1)
                            if handle.lower() in SKIP:
                                continue
                            if brand_match(handle) or brand_match(title) or brand_match(snippet):
                                handles["facebook"] = handle
                                print(f"   ✅ Facebook: {handle}")
                                found = True; break
                            else:
                                print(f"   ⚠️ Facebook skipped wrong brand: {handle}")

                    elif platform == "linkedin" and "linkedin.com/company" in link:
                        m = re.search(r'linkedin\.com/company/([A-Za-z0-9_.-]+)/?', link, re.I)
                        if m:
                            handles["linkedin"] = m.group(1)
                            print(f"   ✅ LinkedIn: {m.group(1)}")
                            found = True; break

                    elif platform == "youtube" and "youtube.com" in link:
                        m = re.search(r'youtube\.com/@([A-Za-z0-9_.]+)/?', link, re.I)
                        if not m:
                            m = re.search(r'youtube\.com/(?:channel|c|user)/([A-Za-z0-9_.-]+)/?', link, re.I)
                        if m:
                            handle = m.group(1)
                            if brand_match(handle) or brand_match(title) or brand_match(snippet):
                                if is_youtube_channel_id(handle):
                                    handles["youtube"] = handle.lstrip("@")
                                else:
                                    handles["youtube"] = "@" + handle.lstrip("@")
                                print(f"   ✅ YouTube: {handles['youtube']}")
                                found = True; break
                            else:
                                print(f"   ⚠️ YouTube skipped: {handle}")

            except Exception as e:
                print(f"   ⚠️ SearchAPI failed for {platform}: {e}")

    return handles

def find_youtube_via_engine(brand):
    try:
        print(f"   🔍 YouTube engine search: {brand}")
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine":"youtube","q":brand,"api_key":SEARCHAPI_KEY}
        )
        if r.status_code != 200:
            return None
        brand_clean = brand.replace("-","").replace("_","").replace(" ","").lower()
        for video in r.json().get("videos", [])[:5]:
            ch = video.get("channel", {})
            ch_id    = ch.get("id","")
            ch_title = ch.get("title","").replace(" ","").replace("-","").lower()
            if ch_id and brand_clean in ch_title:
                print(f"   ✅ YouTube via engine: {ch.get('title')} ({ch_id})")
                return ch_id
    except Exception as e:
        print(f"   ⚠️ YouTube engine search failed: {e}")
    return None

def discover_all_handles(website_url):
    print(f"\n🔎 Finding all social handles for: {website_url}")
    handles = find_handles_from_website(website_url)
    print(f"   Found on website: {handles}")
    missing = [p for p in ["instagram","facebook","linkedin","youtube"]
               if p not in handles]
    if missing:
        print(f"   Searching for missing: {missing}")
        google_handles = find_handles_from_searchapi(website_url, missing)
        handles.update(google_handles)
        print(f"   Final handles: {handles}")

    if "youtube" not in handles:
        domain = website_url.replace("https://","").replace("http://","").rstrip("/")
        brand  = re.sub(r'\.(com|in|io|co|net|org|app).*','', domain).replace("www.","")
        yt_id = find_youtube_via_engine(brand)
        if yt_id:
            handles["youtube"] = yt_id
            print(f"   Final handles: {handles}")

    if not handles:
        raise ValueError(
            f"Could not find any social media accounts for {website_url}. "
            "Please check the website URL."
        )
    return handles


# ═══════════════════════════════════════════════════════════════
# STEP 2 — FETCH DATA FROM EACH PLATFORM
# ═══════════════════════════════════════════════════════════════

def _get_first(d, keys, default=None):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return default

def _to_timestamp(val):
    """
    Convert various timestamp formats to a unix timestamp (epoch seconds).

    IMPORTANT — timezone handling:
    Instagram / Apify / SearchAPI all report post dates in UTC. The trailing
    "Z" on strings like "2026-04-01T13:30:41.000Z" means "Zulu time", i.e. UTC.
    We explicitly attach UTC tzinfo for naive date strings instead of letting
    Python guess based on wherever the server happens to be running — otherwise
    the exact same post produces a different timestamp on a Mac in IST vs a
    Railway server, which silently throws off posting frequency, day-of-week
    calculations, and the pinned-post date-sort below.
    """
    if val is None:
        return None
    try:
        if isinstance(val, (int, float)):
            v = float(val)
            # Millisecond-epoch detection: HarvestAPI (and some other
            # APIs) send epoch MILLISECONDS. Read as seconds, those put
            # posts in the year ~58509 and crash every date calculation
            # (confirmed live: "ValueError: year 58509 is out of range").
            # Any value above 1e11 seconds (~year 5138) can't be a real
            # post date in seconds, so it must be milliseconds.
            if v > 1e11:
                v = v / 1000.0
            return v
        s = str(val)
        if s.isdigit():
            v = float(s)
            if v > 1e11:
                v = v / 1000.0
            return v

        # Formats that already carry explicit timezone/offset info
        for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"):
            try:
                dt = datetime.strptime(s[:32], fmt)
                return dt.timestamp()
            except:
                continue

        # Formats with NO timezone marker — Instagram data here is always UTC
        for fmt in ("%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S",
                    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                clean = s[:26].rstrip("Z")
                dt = datetime.strptime(clean, fmt)
                return dt.replace(tzinfo=timezone.utc).timestamp()
            except:
                continue
    except:
        pass
    return None

def _classify_post(post):
    t = str(_get_first(post, ["type","productType","product_type"], "")).lower()
    if "video" in t or "reel" in t or "clip" in t:
        return "reels"
    elif "carousel" in t or "sidecar" in t or "album" in t:
        return "carousels"
    # If from reel scraper, always classify as reel
    if post.get("_source") == "reel_scraper":
        return "reels"
    return "images"

def compute_day_distribution(parsed):
    """
    Works out which day of the week the brand posts most/least on,
    based on each post's real UTC date (not Instagram's pinned-post
    display order, and not local server time).

    Returns most_active_day and least_active_day. If a weekday never
    appears in the sample, it's reported as "never posted" rather than
    just being called the least-active day, since that's more accurate
    than implying it *did* happen, just rarely.
    """
    day_names = ["Monday","Tuesday","Wednesday","Thursday",
                 "Friday","Saturday","Sunday"]
    counts = Counter()
    for p in parsed:
        if p["ts"]:
            weekday = datetime.fromtimestamp(p["ts"], tz=timezone.utc).strftime("%A")
            counts[weekday] += 1

    if not counts:
        return {"most_active_day": "N/A", "least_active_day": "N/A"}

    most_active_day = max(counts, key=counts.get)
    zero_days = [d for d in day_names if d not in counts]
    if zero_days:
        least_active_day = f"{zero_days[0]} (never posted)"
    else:
        least_active_day = min(counts, key=counts.get)

    return {"most_active_day": most_active_day, "least_active_day": least_active_day}

def _score(p):
    return p["likes"] + p["comments"] + (p["shares"] or 0) + (p["reposts"] or 0)


def compute_hour_insights(parsed, min_posts_to_qualify=2, outlier_multiplier=3):
    """
    Most Frequent Hour: which hour of day (UTC) has the most posts — pure count,
    doesn't consider engagement at all.

    Best Performing Hour: which hour has the highest AVERAGE engagement. Only
    hours with `min_posts_to_qualify` or more posts are allowed to compete —
    this stops a single lucky post from falsely looking like a repeatable
    "best time to post."

    Outlier note: separately, checks every hour that has exactly ONE post
    (too few to qualify above) and flags it if that single post scored
    dramatically higher than the sample's overall average (>= outlier_multiplier
    times the average). This keeps the standout post visible to the client
    without mislabeling it as a proven pattern.
    """
    hour_posts = {}
    for p in parsed:
        if p["ts"] is not None:
            hour = datetime.fromtimestamp(p["ts"], tz=timezone.utc).hour
            hour_posts.setdefault(hour, []).append(p)

    if not hour_posts:
        return {"most_frequent_hour": "N/A", "best_performing_hour": "N/A", "outlier_note": None}

    def hour_avg(posts):
        return sum(_score(pp) for pp in posts) / len(posts)

    # ── Most Frequent Hour — pure count ──────────────────────────────
    most_frequent_hour = max(hour_posts, key=lambda h: len(hour_posts[h]))

    # ── Best Performing Hour — avg engagement, min posts to qualify ──
    qualifying = {h: posts for h, posts in hour_posts.items() if len(posts) >= min_posts_to_qualify}
    if qualifying:
        best_hour = max(qualifying, key=lambda h: hour_avg(qualifying[h]))
        best_performing_hour = f"{best_hour:02d}:00 UTC"
    else:
        best_performing_hour = "N/A (not enough data)"

    # ── Outlier note — single-post hours only ────────────────────────
    all_scores  = [_score(pp) for pp in parsed]
    overall_avg = sum(all_scores) / len(all_scores) if all_scores else 0
    outlier_note = None
    lonely_hours = {h: posts[0] for h, posts in hour_posts.items() if len(posts) == 1}
    if lonely_hours and overall_avg > 0:
        lonely_hour, lonely_post = max(lonely_hours.items(), key=lambda kv: _score(kv[1]))
        lonely_score = _score(lonely_post)
        if lonely_score >= overall_avg * outlier_multiplier:
            outlier_note = (f"Outlier: one post at {lonely_hour:02d}:00 UTC scored {lonely_score:,} "
                             f"vs. an average of {overall_avg:.0f} — worth reviewing, though it's "
                             f"just one post so far.")

    return {
        "most_frequent_hour":   f"{most_frequent_hour:02d}:00 UTC",
        "best_performing_hour": best_performing_hour,
        "outlier_note":         outlier_note,
    }


def compute_posting_consistency(parsed):
    """
    Looks at the real gap (in days) between every consecutive pair of posts
    and measures how much those gaps vary, using the coefficient of variation
    (std deviation ÷ average gap). A low value means posts land at fairly
    even intervals ("Consistent"); a high value means bursts of activity
    followed by long silences ("Irregular") — something a single flat
    "X per week" number hides completely.
    """
    timestamps = sorted(p["ts"] for p in parsed if p["ts"] is not None)
    if len(timestamps) < 3:
        return "N/A"

    gaps = [(timestamps[i+1] - timestamps[i]) / 86400 for i in range(len(timestamps)-1)]
    avg_gap = sum(gaps) / len(gaps)
    if avg_gap == 0:
        return "N/A"
    variance = sum((g - avg_gap) ** 2 for g in gaps) / len(gaps)
    coefficient_of_variation = (variance ** 0.5) / avg_gap

    return "Consistent" if coefficient_of_variation < 0.75 else "Irregular"


def compute_momentum(parsed):
    """
    Splits posts into two halves by real date — the most recent half vs.
    the older half — and compares average engagement between them. Returns
    a signed percentage plus a direction ("up"/"down") so the report can
    show a colored arrow: green ▲ if recent posts are outperforming older
    ones, red ▼ if engagement is slipping.
    """
    dated = sorted((p for p in parsed if p["ts"] is not None), key=lambda p: p["ts"], reverse=True)
    n = len(dated)
    if n < 10:
        return {"momentum_pct": "N/A", "momentum_direction": None}

    half = n // 2
    recent_half = dated[:half]
    older_half  = dated[half:half*2] if len(dated) >= half*2 else dated[half:]
    if not older_half:
        return {"momentum_pct": "N/A", "momentum_direction": None}

    recent_avg = sum(_score(p) for p in recent_half) / len(recent_half)
    older_avg  = sum(_score(p) for p in older_half) / len(older_half)
    if older_avg == 0:
        return {"momentum_pct": "N/A", "momentum_direction": None}

    pct_change = (recent_avg - older_avg) / older_avg * 100
    return {
        "momentum_pct":       f"{pct_change:+.0f}%",
        "momentum_direction": "up" if pct_change >= 0 else "down",
    }

def parse_num(val):
    """
    Module-level numeric parser for follower counts etc. — handles
    values like "233,374", "1.2M", "45K", "1000+".

    NOTE: create_ppt also defines its own local parse_num (identical
    logic); that local copy harmlessly shadows this one inside
    create_ppt. This module-level version exists because
    fetch_linkedin needs it too, and previously calling it there
    raised a silent NameError (the local one wasn't visible outside
    create_ppt), which wiped out all LinkedIn post-level metrics.
    """
    try:
        v = str(val).replace(",","").upper().replace("+","")
        if "M" in v: return float(v.replace("M","")) * 1_000_000
        if "K" in v: return float(v.replace("K","")) * 1_000
        return float(v)
    except:
        return 0


def _smart_round(value, max_decimals=2):
    """
    Rounds to a whole number normally (e.g. 4.2 -> "4"), but if that
    would hide a real, non-zero value behind a flat "0" (e.g. 0.375
    likes/post), keeps adding decimal places until the number is
    actually visible — up to max_decimals. A genuinely-zero value still
    correctly shows "0". This fixes averages that looked misleadingly
    flat on low-volume post groups (e.g. 8 videos with only 3 total
    likes between them) while every other stat on the same slide showed
    a proper decimal.
    """
    if value == 0:
        return "0"
    rounded = round(value)
    if rounded != 0:
        return str(rounded)
    for d in range(1, max_decimals + 1):
        r = round(value, d)
        if r != 0:
            return f"{r:.{d}f}"
    return f"{value:.{max_decimals}f}"


def _compute_group_metrics(posts_group, followers, is_reels=False):
    n = len(posts_group)
    if n == 0:
        return {"n": 0}

    total_likes    = sum(p["likes"] for p in posts_group)
    total_comments = sum(p["comments"] for p in posts_group)
    shares_known   = [p["shares"] for p in posts_group if p["shares"] is not None]
    reposts_known  = [p["reposts"] for p in posts_group if p["reposts"] is not None]
    total_shares   = sum(shares_known) if shares_known else 0
    total_reposts  = sum(reposts_known) if reposts_known else 0

    avg_engagement  = (total_likes + total_comments + total_shares + total_reposts) / n
    er_per_follower = (total_likes + total_comments) / n / followers * 100 if followers else 0

    scored = sorted(posts_group, key=_score, reverse=True)
    top    = scored[0]
    worst  = scored[-1]

    result = {
        "n":               n,
        "avg_engagement":  round(avg_engagement, 1),
        "er_per_follower": f"{er_per_follower:.2f}%",
        "avg_likes":       _smart_round(total_likes / n),
        "avg_comments":    _smart_round(total_comments / n),
        "avg_shares":      _smart_round(total_shares / n) if shares_known else "N/A",
        "avg_reposts":     _smart_round(total_reposts / n) if reposts_known else "N/A",
        "top_score":       _score(top),
        "top_url":         top["url"],
        "worst_score":     _score(worst),
        "worst_url":       worst["url"],
    }

    if is_reels:
        views_known = [p["views"] for p in posts_group if p["views"] is not None]
        total_views = sum(views_known) if views_known else 0
        result["avg_views"] = _smart_round(total_views / n) if views_known else "N/A"
        if views_known:
            max_v = max(posts_group, key=lambda p: p["views"] or 0)
            result["max_views"]     = max_v["views"]
            result["max_views_url"] = max_v["url"]
        else:
            result["max_views"]     = "N/A"
            result["max_views_url"] = "N/A"

        # ── Rates are per VIEW (not per follower) ──────────────────
        result["like_rate"]    = f"{(total_likes/total_views*100):.2f}%"    if total_views else "N/A"
        result["comment_rate"] = f"{(total_comments/total_views*100):.2f}%" if total_views else "N/A"
        result["share_rate"]   = f"{(total_shares/total_views*100):.2f}%"   if (total_views and shares_known) else "N/A"
        result["repost_rate"]  = f"{(total_reposts/total_views*100):.2f}%"  if (total_views and reposts_known) else "N/A"
        result["er_by_view"]   = f"{((total_likes+total_comments+total_shares+total_reposts)/total_views*100):.2f}%" if total_views else "N/A"


    return result


def fetch_apify_async(username, actor_id, input_json):
    """Generic async Apify runner: start run, poll until done, fetch dataset."""
    try:
        print(f"   🔄 Apify async: starting {actor_id} for @{username}...")
        start_r = requests.post(
            f"https://api.apify.com/v2/acts/{actor_id}/runs",
            params={"token": APIFY_API_KEY},
            json=input_json,
            timeout=30
        )
        if start_r.status_code not in (200, 201):
            print(f"   ⚠️ Apify async start failed: {start_r.status_code}")
            return []

        run_id = start_r.json().get("data", {}).get("id")
        if not run_id:
            print("   ⚠️ Apify async: no run ID returned")
            return []

        print(f"   ⏳ Apify run started: {run_id} — polling...")

        last_status_r = None
        for attempt in range(24):
            time.sleep(5)
            last_status_r = requests.get(
                f"https://api.apify.com/v2/acts/{actor_id}/runs/{run_id}",
                params={"token": APIFY_API_KEY},
                timeout=15
            )
            status = last_status_r.json().get("data", {}).get("status", "")
            print(f"   ⏳ Apify status: {status} (attempt {attempt+1})")
            if status == "SUCCEEDED":
                break
            elif status in ("FAILED", "ABORTED", "TIMED-OUT"):
                print(f"   ⚠️ Apify run ended with: {status}")
                return []

        if last_status_r is None:
            return []

        dataset_id = last_status_r.json().get("data", {}).get("defaultDatasetId")
        if not dataset_id:
            print("   ⚠️ Apify async: no dataset ID")
            return []

        items_r = requests.get(
            f"https://api.apify.com/v2/datasets/{dataset_id}/items",
            params={"token": APIFY_API_KEY, "limit": 35},
            timeout=30
        )
        posts = items_r.json()
        if isinstance(posts, list) and len(posts) > 0:
            print(f"   ✅ Apify async returned {len(posts)} items")
            return posts
        else:
            print("   ⚠️ Apify async dataset was empty")
            return []

    except Exception as e:
        print(f"   ⚠️ Apify async error: {e}")
        return []


def fetch_post_metrics_via_analytics_actor(post_urls):
    """
    Enriches posts with shares / reposts / saves using the Apify actor
    'patient_discovery/instagram-reel-analytics-by-url'.

    Why this specific actor: after testing 6 different tools (SearchAPI,
    two official Apify scrapers, ScraperAPI, SocialCrawl, anonymous
    GraphQL), this was the ONLY one that returned real values. Verified
    against a live TalkingLands post on 2026-07-15:
        metrics.repost_count = 17   (matched browser exactly)
        metrics.share_count  = 750
        metrics.save_count   = 410
    Instagram only exposes these numbers to logged-in viewers; this actor
    handles that on its side, so no Instagram credentials are needed here.

    Takes a list of post URLs (supports bulk input), returns a dict of
    {url: {"shares": int|None, "reposts": int|None, "saves": int|None}}.
    Cost: one Apify result per post (~$2.30-2.70 per 1,000 → roughly
    ₹6-7 extra per 30-post report).
    """
    if not APIFY_API_KEY:
        print("   ⚠️ No APIFY_API_KEY set — skipping shares/reposts enrichment")
        return {}

    urls = [u for u in post_urls if u and u != "N/A"]
    if not urls:
        return {}

    results = {}
    try:
        print(f"   🔄 Fetching shares/reposts/saves for {len(urls)} posts via analytics actor...")
        actor_input = {"urls": urls}
        r = requests.post(
            "https://api.apify.com/v2/acts/patient_discovery~instagram-reel-analytics-by-url/run-sync-get-dataset-items",
            params={"token": APIFY_API_KEY, "timeout": 180},
            json=actor_input,
            timeout=200
        )
        items = []
        if r.status_code in (200, 201):
            items = r.json()
            if not isinstance(items, list):
                items = []
        if not items:
            print(f"   ⚠️ Analytics actor sync empty/failed (status {r.status_code}), trying async...")
            items = fetch_apify_async("bulk", "patient_discovery~instagram-reel-analytics-by-url", actor_input)

        for item in items:
            metrics = item.get("metrics", {}) or {}
            # Match result back to its post URL: the actor echoes the input
            # URL and also provides the shortcode in `code`.
            item_url = _get_first(item, ["inputUrl", "url"]) or ""
            code     = item.get("code", "")
            key = None
            for u in urls:
                if (item_url and item_url.rstrip("/") == u.rstrip("/")) or (code and f"/{code}" in u):
                    key = u
                    break
            if key is None:
                continue
            results[key] = {
                "shares":  metrics.get("share_count"),
                "reposts": metrics.get("repost_count"),
                "saves":   metrics.get("save_count"),
            }

        found = sum(1 for v in results.values() if v.get("reposts") is not None or v.get("shares") is not None)
        print(f"   ✅ Shares/reposts found for {found}/{len(urls)} posts")
    except Exception as e:
        print(f"   ⚠️ Analytics actor error: {e}")

    return results


def fetch_instagram_posts_via_apify(username):
    """
    Two-scraper approach:
    1. apify/instagram-post-scraper  → images + carousels (likes, comments, views)
    2. apify/instagram-reel-scraper  → reels (likes, comments, views, SHARES)

    NOTE on the 35-post fetch limit (up from 30):
    Instagram displays pinned posts first in profile grid order, regardless of
    when they were actually posted — a post pinned 3 months ago can occupy a
    "top 30" slot and silently push out a genuinely recent post. We fetch a
    small buffer of extra posts here (35 instead of 30) so that, after the
    real-date sort applied later in fetch_instagram(), there's always enough
    genuinely recent posts to fill a true top-30 even when a few pinned posts
    turn out to be old. We intentionally do NOT slice down to 30 here — the
    final trim happens after sorting by real post date, not fetch order.
    """
    if not APIFY_API_KEY:
        print("   ⚠️ No APIFY_API_KEY set")
        return []

    all_posts = []

    # ── 1. Post scraper (images + carousels + some reels) ────────
    try:
        print(f"   🔄 Apify posts: fetching for @{username}...")
        post_input = {
            "username":      [username],
            "resultsLimit":  35,
            "addParentData": False,
        }
        r = requests.post(
            "https://api.apify.com/v2/acts/apify~instagram-post-scraper/run-sync-get-dataset-items",
            params={"token": APIFY_API_KEY, "timeout": 120},
            json=post_input,
            timeout=150
        )
        if r.status_code in (200, 201):
            posts = r.json()
            if isinstance(posts, list) and len(posts) > 0:
                print(f"   ✅ Post scraper returned {len(posts)} posts")
                all_posts.extend(posts)
            else:
                print("   ⚠️ Post scraper sync empty, trying async...")
                posts = fetch_apify_async(username, "apify~instagram-post-scraper", post_input)
                all_posts.extend(posts)
        else:
            print(f"   ⚠️ Post scraper status {r.status_code}, trying async...")
            posts = fetch_apify_async(username, "apify~instagram-post-scraper", post_input)
            all_posts.extend(posts)
    except Exception as e:
        print(f"   ⚠️ Post scraper error: {e}")

    # ── 2. Reel scraper (reels with shares) ──────────────────────
    try:
        print(f"   🔄 Apify reels: fetching reels with shares for @{username}...")
        reel_input = {
            "username":     [username],
            "resultsLimit": 35,
        }
        r2 = requests.post(
            "https://api.apify.com/v2/acts/apify~instagram-reel-scraper/run-sync-get-dataset-items",
            params={"token": APIFY_API_KEY, "timeout": 120},
            json=reel_input,
            timeout=150
        )
        if r2.status_code in (200, 201):
            reels = r2.json()
            if isinstance(reels, list) and len(reels) > 0:
                print(f"   ✅ Reel scraper returned {len(reels)} reels")
                for reel in reels:
                    reel["_source"] = "reel_scraper"
                existing_urls = {p.get("url","") for p in all_posts}
                new_reels = [r for r in reels if r.get("url","") not in existing_urls]
                all_posts.extend(new_reels)
                print(f"   ✅ Added {len(new_reels)} new reels (deduped)")
            else:
                print("   ⚠️ Reel scraper sync empty, trying async...")
                reels = fetch_apify_async(username, "apify~instagram-reel-scraper", reel_input)
                for reel in reels:
                    reel["_source"] = "reel_scraper"
                existing_urls = {p.get("url","") for p in all_posts}
                new_reels = [r for r in reels if r.get("url","") not in existing_urls]
                all_posts.extend(new_reels)
        else:
            print(f"   ⚠️ Reel scraper status {r2.status_code}, trying async...")
            reels = fetch_apify_async(username, "apify~instagram-reel-scraper", reel_input)
            for reel in reels:
                reel["_source"] = "reel_scraper"
            existing_urls = {p.get("url","") for p in all_posts}
            new_reels = [r for r in reels if r.get("url","") not in existing_urls]
            all_posts.extend(new_reels)
    except Exception as e:
        print(f"   ⚠️ Reel scraper error: {e}")

    print(f"   ✅ Total combined posts fetched (pre date-sort): {len(all_posts)}")
    return all_posts


def fetch_facebook_posts_via_apify(page_handle):
    """
    Fetches recent Facebook posts using apify/facebook-posts-scraper (the
    official Apify actor for Facebook Pages).

    HONEST CAVEAT: unlike the Instagram actors in this file, which went
    through many rounds of real-world testing to confirm exact field
    names, this Facebook actor has NOT yet been verified against real
    output. The field-name guesses below are based on common Apify
    Facebook-scraper conventions, with a wide fallback list per field
    (same defensive pattern used for Instagram). Use the /debug_facebook
    route to run one real test and confirm/adjust the field names before
    trusting this at scale — same process that eventually nailed Instagram.
    """
    if not APIFY_API_KEY:
        print("   ⚠️ No APIFY_API_KEY set")
        return []

    page_url = page_handle if page_handle.startswith("http") else f"https://www.facebook.com/{page_handle}"
    fb_input = {
        "startUrls": [{"url": page_url}],
        "resultsLimit": 35,  # same buffer logic as Instagram — covers pinned-post displacement
    }
    try:
        print(f"   🔄 Apify Facebook posts: fetching for {page_handle}...")
        r = requests.post(
            "https://api.apify.com/v2/acts/apify~facebook-posts-scraper/run-sync-get-dataset-items",
            params={"token": APIFY_API_KEY, "timeout": 120},
            json=fb_input,
            timeout=150
        )
        if r.status_code in (200, 201):
            posts = r.json()
            if isinstance(posts, list) and len(posts) > 0:
                print(f"   ✅ Facebook post scraper returned {len(posts)} posts")
                return posts
        print(f"   ⚠️ Facebook post scraper sync empty/failed (status {r.status_code}), trying async...")
        return fetch_apify_async(page_handle, "apify~facebook-posts-scraper", fb_input)
    except Exception as e:
        print(f"   ⚠️ Facebook post scraper error: {e}")
        return []


def _classify_facebook_post(post):
    """
    Facebook doesn't have a Carousel format like Instagram, so posts are
    split into Photos / Videos / Links instead.

    CONFIRMED against real apify/facebook-posts-scraper output (2026-07-16):
    there is no top-level "type" field on a post. The real signal lives
    inside the "media" array — each media item has a "__typename" field
    ("Photo" for images; expected "Video" for actual video/reel posts,
    unconfirmed against a real video post yet). This replaced an earlier
    guess that checked for a top-level type field, which doesn't exist
    and was silently classifying every post as "photos".
    """
    media_items = post.get("media") or []
    typenames = [str(m.get("__typename","")).lower() for m in media_items if isinstance(m, dict)]
    if any("video" in t for t in typenames):
        return "videos"
    if any("photo" in t or "image" in t for t in typenames):
        return "photos"
    # No media at all, or an unrecognized shape — check the caption as a
    # last resort (posts often say "reel" or "video" somewhere in their
    # own text, e.g. "In this reel, we break down..."), then fall back
    # to "links" for pure text/link posts.
    text = str(post.get("text","")).lower()
    if "reel" in text or "video" in text:
        return "videos"
    if not media_items:
        return "links"
    return "photos"


def fetch_instagram(username):
    print(f"📸 Fetching Instagram: @{username}")
    data = {}
    try:
        # ── Profile data from SearchAPI ───────────────────────────────
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine":"instagram_profile","username":username,
                    "api_key":SEARCHAPI_KEY}
        )
        if r.status_code != 200:
            return data

        ig = r.json()
        p  = ig.get("profile", {})

        followers   = p.get("followers", "N/A")
        following   = p.get("following", "N/A")
        total_posts = p.get("posts", "N/A")
        bio         = p.get("biography", "")
        full_name   = p.get("full_name", "")
        is_verified = p.get("is_verified", False)
        is_business = p.get("is_business", False)
        avatar_url  = _get_first(p, ["avatar_hd", "avatar"])

        try:
            followers_n = int(str(followers).replace(",",""))
        except:
            followers_n = 0

        # ── Posts via Apify (post + reel scrapers), fallback to SearchAPI ──
        apify_posts = fetch_instagram_posts_via_apify(username)
        if apify_posts:
            posts_raw = apify_posts
            source    = "apify"
        else:
            posts_raw = ig.get("posts", [])
            source    = "searchapi"
            print(f"   ↩️  Using SearchAPI posts ({len(posts_raw)} posts)")

        print(f"   🔍 Source: {source} | Posts fetched: {len(posts_raw)}")

        # ── Parse every post ──────────────────────────────────────────
        parsed = []

        for post in posts_raw:
            if source == "apify":
                reel_src = post.get("_source") == "reel_scraper"
                likes    = post.get("likesCount", 0) or 0
                comments = post.get("commentsCount", 0) or 0
                if reel_src:
                    # Reel scraper has reliable sharesCount
                    shares  = _get_first(post, ["sharesCount","videoSharesCount","shares"])
                    reposts = _get_first(post, ["repostsCount","reposts"])
                    views   = _get_first(post, ["videoViewCount","videoPlayCount","playsCount","views"])
                else:
                    shares  = _get_first(post, ["sharesCount","share_count","shares"])
                    reposts = _get_first(post, ["repostsCount","repost_count","reposts"])
                    views   = _get_first(post, ["videoViewCount","viewsCount","videoPlayCount","views"])
                url = _get_first(post, ["url","permalink","link"]) or "N/A"
                ts  = _to_timestamp(_get_first(post, ["timestamp","iso_date"]))
            else:
                # SearchAPI confirmed field names
                likes    = post.get("likes", 0) or 0
                comments = post.get("comments", 0) or 0
                shares   = _get_first(post, ["share_count","shares","reshare_count"])
                reposts  = _get_first(post, ["repost_count","reposts","reshares"])
                views    = _get_first(post, ["views","view_count","play_count"])
                url      = _get_first(post, ["permalink","link","url","post_url"]) or "N/A"
                ts       = _to_timestamp(_get_first(post, ["iso_date","timestamp","taken_at","created_at"]))

            ptype = _classify_post(post)

            parsed.append({
                "type":    ptype,
                "likes":   likes,
                "comments":comments,
                "shares":  (int(shares)  if shares  is not None else None),
                "reposts": (int(reposts) if reposts is not None else None),
                "views":   (int(views)   if views   is not None else None),
                "url":     url,
                "ts":      ts,
            })

        # ── Sort by REAL post date, ignoring Instagram's pinned-post
        #    display order, and keep only the true most-recent 30.
        #
        #    Why: Instagram shows pinned posts first in the grid no matter
        #    how old they are. If we trusted fetch/display order, an old
        #    pinned post could occupy a "top 30" slot and quietly displace
        #    a genuinely recent post — skewing posting frequency, averages,
        #    and which post gets crowned Top/Worst performer. A pinned post
        #    that IS genuinely recent is unaffected by this — it earns its
        #    place in the top 30 purely by having a recent real date, same
        #    as any other post. Pin status itself is never checked. ───────
        posts_with_dates    = [pp for pp in parsed if pp["ts"] is not None]
        posts_without_dates = [pp for pp in parsed if pp["ts"] is None]
        posts_with_dates.sort(key=lambda pp: pp["ts"], reverse=True)
        parsed = (posts_with_dates + posts_without_dates)[:30]

        # ── Enrich with shares / reposts / saves via the analytics actor ──
        #    Runs on the final 30 posts only (after the date-sort/trim), so
        #    we never pay for posts that won't appear in the report. Values
        #    already present from the base scrapers are kept; the actor
        #    fills the gaps.
        posts_needing = [pp for pp in parsed
                         if pp["reposts"] is None or pp["shares"] is None]
        if posts_needing and APIFY_API_KEY:
            metric_results = fetch_post_metrics_via_analytics_actor(
                [pp["url"] for pp in posts_needing]
            )
            for pp in parsed:
                found = metric_results.get(pp["url"])
                if not found:
                    continue
                if pp["reposts"] is None and found.get("reposts") is not None:
                    pp["reposts"] = int(found["reposts"])
                if pp["shares"] is None and found.get("shares") is not None:
                    pp["shares"] = int(found["shares"])

        # ── Content-type counts — computed AFTER the date-sort/trim above,
        #    so the Content Strategy pie chart reflects the true final 30
        #    posts, not the larger raw fetch batch. ───────────────────────
        img_c = sum(1 for pp in parsed if pp["type"] == "images")
        car_c = sum(1 for pp in parsed if pp["type"] == "carousels")
        vid_c = sum(1 for pp in parsed if pp["type"] == "reels")

        post_count      = len(parsed)
        images_group    = [pp for pp in parsed if pp["type"] == "images"]
        carousels_group = [pp for pp in parsed if pp["type"] == "carousels"]
        reels_group     = [pp for pp in parsed if pp["type"] == "reels"]

        # ── Overall engagement ────────────────────────────────────────
        total_l          = sum(pp["likes"] for pp in parsed)
        total_c          = sum(pp["comments"] for pp in parsed)
        engagement_total = total_l + total_c
        eng_rate         = f"{(engagement_total/post_count/followers_n*100):.2f}%" if (post_count and followers_n) else "N/A"

        reel_l = sum(pp["likes"]    for pp in reels_group)
        reel_c = sum(pp["comments"] for pp in reels_group)
        eng_rate_reels = (f"{(reel_l+reel_c)/len(reels_group)/followers_n*100:.2f}%"
                          if (reels_group and followers_n) else "N/A")

        # ── Posting frequency ─────────────────────────────────────────
        timestamps = [pp["ts"] for pp in parsed if pp["ts"]]
        if len(timestamps) >= 2:
            span_days         = max((max(timestamps) - min(timestamps)) / 86400, 1)
            posting_frequency = f"{(post_count / span_days * 7):.1f} / week"
        else:
            posting_frequency = "N/A"

        # ── Day-of-week posting distribution (UTC) ──────────────────────
        day_dist = compute_day_distribution(parsed)

        # ── Best/most-frequent posting hour, consistency, momentum (UTC) ──
        hour_insights = compute_hour_insights(parsed)
        consistency   = compute_posting_consistency(parsed)
        momentum      = compute_momentum(parsed)

        # ── Per-group metrics ─────────────────────────────────────────
        images_metrics    = _compute_group_metrics(images_group,    followers_n)
        carousels_metrics = _compute_group_metrics(carousels_group, followers_n)
        reels_metrics     = _compute_group_metrics(reels_group,     followers_n, is_reels=True)

        data = {
            "username":    username,
            "full_name":   full_name,
            "followers":   str(followers),
            "following":   str(following),
            "posts":       str(total_posts),
            "bio":         bio,
            "is_verified": "Yes" if is_verified else "No",
            "is_business": "Yes" if is_business else "No",

            "sample_size":           post_count,
            "engagement_rate":       eng_rate,
            "engagement_rate_reels": eng_rate_reels,
            "engagement_total":      engagement_total,
            "posting_frequency":     posting_frequency,

            "most_active_day":  day_dist["most_active_day"],
            "least_active_day": day_dist["least_active_day"],

            "avatar_url": avatar_url,

            "most_frequent_hour":   hour_insights["most_frequent_hour"],
            "best_performing_hour": hour_insights["best_performing_hour"],
            "outlier_note":         hour_insights["outlier_note"],
            "posting_consistency":  consistency,
            "momentum_pct":         momentum["momentum_pct"],
            "momentum_direction":   momentum["momentum_direction"],

            "img_count":  img_c,
            "car_count":  car_c,
            "vid_count":  vid_c,
            "post_count": post_count,

            "images":    images_metrics,
            "carousels": carousels_metrics,
            "reels":     reels_metrics,
        }
        print(f"   ✅ {followers} followers | {post_count} posts "
              f"({img_c} img / {car_c} carousel / {vid_c} reels) | ER: {eng_rate} | "
              f"Most active: {day_dist['most_active_day']} | Least active: {day_dist['least_active_day']}")
    except Exception as e:
        import traceback
        print(f"   ⚠️ Instagram failed: {e}")
        print(traceback.format_exc())
    return data


def fetch_facebook(username):
    print(f"👥 Fetching Facebook: {username}")
    data = {}
    try:
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine":"facebook_business_page","username":username,
                    "api_key":SEARCHAPI_KEY}
        )
        if r.status_code != 200:
            return data

        p = r.json().get("page", {})
        followers_raw = p.get("followers",{}).get("count","N/A")
        try:
            followers_n = int(str(followers_raw).replace(",",""))
        except:
            followers_n = 0

        data = {
            "name":        p.get("name", username),
            "followers":   str(followers_raw),
            "following":   str(p.get("following",{}).get("count","N/A")),
            "category":    ", ".join(p.get("about",{}).get("category_formatted",[]) or []) or "N/A",
            "about":       p.get("about",{}).get("description","") or p.get("about",{}).get("general_info",""),
            "address":     p.get("address","") or "",
            "phone":       p.get("phone","") or "",
            "email":       p.get("email","") or "",
            "website":     p.get("website","") or "",
            "rating":      str(p.get("ratings",{}).get("value","N/A")),
            "is_verified": "Yes" if p.get("is_verified") else "No",
        }
        print(f"   ✅ {data['followers']} followers")

        # ── Posts via Apify — same pattern as Instagram ─────────────────
        posts_raw = fetch_facebook_posts_via_apify(username)
        print(f"   🔍 Facebook posts fetched: {len(posts_raw)}")

        parsed = []
        for post in posts_raw:
            likes    = _get_first(post, ["likes", "likesCount", "reactionsCount", "reactions"], 0) or 0
            comments = _get_first(post, ["comments", "commentsCount"], 0) or 0
            shares   = _get_first(post, ["shares", "sharesCount"])
            views    = _get_first(post, ["viewsCount", "views", "videoViewCount"])
            url      = _get_first(post, ["url", "postUrl", "facebookUrl", "link"]) or "N/A"
            ts       = _to_timestamp(_get_first(post, ["time", "timestamp", "date", "publishedAt"]))
            ptype    = _classify_facebook_post(post)
            try: likes = int(likes)
            except: likes = 0
            try: comments = int(comments)
            except: comments = 0
            parsed.append({
                "type":     ptype,
                "likes":    likes,
                "comments": comments,
                "shares":   (int(shares) if shares is not None else None),
                "reposts":  None,  # Facebook has no repost concept — always None, reused
                                    # formulas treat this as 0 automatically, same as Instagram's N/A handling
                "views":    (int(views) if views is not None else None),
                "url":      url,
                "ts":       ts,
            })

        # ── Sort by REAL post date, keep the true most-recent 30 —
        #    identical logic to the Instagram pinned-post fix. ──────────
        posts_with_dates    = [pp for pp in parsed if pp["ts"] is not None]
        posts_without_dates = [pp for pp in parsed if pp["ts"] is None]
        posts_with_dates.sort(key=lambda pp: pp["ts"], reverse=True)
        parsed = (posts_with_dates + posts_without_dates)[:30]

        photos_group = [pp for pp in parsed if pp["type"] == "photos"]
        videos_group = [pp for pp in parsed if pp["type"] == "videos"]
        links_group  = [pp for pp in parsed if pp["type"] == "links"]
        post_count   = len(parsed)

        total_l = sum(pp["likes"] for pp in parsed)
        total_c = sum(pp["comments"] for pp in parsed)
        engagement_total = total_l + total_c
        eng_rate = f"{(engagement_total/post_count/followers_n*100):.2f}%" if (post_count and followers_n) else "N/A"

        vid_l = sum(pp["likes"] for pp in videos_group)
        vid_c = sum(pp["comments"] for pp in videos_group)
        eng_rate_videos = (f"{(vid_l+vid_c)/len(videos_group)/followers_n*100:.2f}%"
                            if (videos_group and followers_n) else "N/A")

        timestamps = [pp["ts"] for pp in parsed if pp["ts"]]
        if len(timestamps) >= 2:
            span_days = max((max(timestamps) - min(timestamps)) / 86400, 1)
            posting_frequency = f"{(post_count / span_days * 7):.1f} / week"
        else:
            posting_frequency = "N/A"

        # ── Reused directly from the Instagram pipeline — same formulas,
        #    just fed Facebook's parsed posts instead. ───────────────────
        day_dist      = compute_day_distribution(parsed)
        hour_insights = compute_hour_insights(parsed)
        consistency   = compute_posting_consistency(parsed)
        momentum      = compute_momentum(parsed)

        photos_metrics = _compute_group_metrics(photos_group, followers_n)
        videos_metrics = _compute_group_metrics(videos_group, followers_n, is_reels=True)
        links_metrics  = _compute_group_metrics(links_group,  followers_n)

        data.update({
            "sample_size":            post_count,
            "engagement_rate":        eng_rate,
            "engagement_rate_videos": eng_rate_videos,
            "engagement_total":       engagement_total,
            "posting_frequency":      posting_frequency,

            "most_active_day":  day_dist["most_active_day"],
            "least_active_day": day_dist["least_active_day"],

            "most_frequent_hour":   hour_insights["most_frequent_hour"],
            "best_performing_hour": hour_insights["best_performing_hour"],
            "outlier_note":         hour_insights["outlier_note"],
            "posting_consistency":  consistency,
            "momentum_pct":         momentum["momentum_pct"],
            "momentum_direction":   momentum["momentum_direction"],

            "photo_count": len(photos_group),
            "video_count": len(videos_group),
            "link_count":  len(links_group),

            "photos": photos_metrics,
            "videos": videos_metrics,
            "links":  links_metrics,
        })
        print(f"   ✅ {post_count} posts ({len(photos_group)} photos / {len(videos_group)} videos / "
              f"{len(links_group)} links) | ER: {eng_rate}")

    except Exception as e:
        import traceback
        print(f"   ⚠️ Facebook failed: {e}")
        print(traceback.format_exc())
    return data


def fetch_youtube(channel_id):
    print(f"▶️  Fetching YouTube: {channel_id}")
    data = {}
    raw = channel_id.strip().lstrip("@")
    if raw.startswith("UC") and len(raw) >= 20:
        param = raw
    else:
        param = "@" + raw
    print(f"   Using param: {param}")
    try:
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine":"youtube_channel","channel_id":param,
                    "api_key":SEARCHAPI_KEY}
        )
        if r.status_code == 200:
            result  = r.json()
            channel = result.get("channel", {})
            about   = result.get("about", {})
            data = {
                "handle":      channel.get("handle", channel_id),
                "title":       channel.get("title",""),
                "subscribers": str(channel.get("subscribers","N/A")),
                "videos":      str(channel.get("videos","N/A")),
                "views":       str(about.get("views","N/A")),
                "description": channel.get("description","") or about.get("description",""),
                "joined":      about.get("joined_date","N/A"),
                "is_verified": "Yes" if channel.get("is_verified") else "No",
            }
            print(f"   ✅ {data['subscribers']} subscribers | {data['videos']} videos")
    except Exception as e:
        print(f"   ⚠️ YouTube failed: {e}")
    return data


def fetch_linkedin_posts_via_apify(company_handle):
    """
    Fetches recent LinkedIn company posts. Primary actor:
    brilliant_gum/linkedin-company-post-scraper — chosen because its
    documentation explicitly supports sorting by "recent" (chronological,
    newest first) and up to 500 posts per run.

    Fallback: data-slayer/linkedin-company-posts-scraper — already
    verified working against real data (field names confirmed), but it
    accepts ONLY a URL: it ignores any limit parameter (docs: "One
    field. Paste your company URLs and run."), which is why it
    returned just 10 posts. Kept as a safety net in case the primary
    actor's input guesses need adjusting.

    HONEST CAVEAT: brilliant_gum's exact input field names are
    best-effort guesses from its documentation (sort "recent"/"top",
    a post limit, pagination via start page) — not yet verified against
    a real run. Use /debug_linkedin to confirm. If the primary returns
    nothing, the proven fallback keeps reports working meanwhile.
    """
    if not APIFY_API_KEY:
        print("   ⚠️ No APIFY_API_KEY set")
        return []

    company_url = company_handle if company_handle.startswith("http") else f"https://www.linkedin.com/company/{company_handle}"

    # ── Primary: harvestapi/linkedin-company-posts ──────────────────
    # Input format CONFIRMED verbatim from Apify's official API example
    # for this actor: {"targetUrls": [...], "maxPosts": N}.
    # HarvestAPI: 5.8K users, 5.0 rating, $2/1k posts. Scrapes newest-
    # first by design ("from now up to postedLimitDate").
    # (Previous primary, brilliant_gum, was a 15-user unrated actor that
    # rejected every API run — likely needed manual rental activation.)
    primary_input = {
        "targetUrls": [company_url],
        "maxPosts":   35,
    }
    try:
        print(f"   🔄 Apify LinkedIn posts (harvestapi, newest-first): {company_handle}...")
        r = requests.post(
            "https://api.apify.com/v2/acts/harvestapi~linkedin-company-posts/run-sync-get-dataset-items",
            params={"token": APIFY_API_KEY, "timeout": 120},
            json=primary_input,
            timeout=150
        )
        if r.status_code in (200, 201):
            posts = r.json()
            if isinstance(posts, list) and len(posts) > 0:
                print(f"   ✅ Primary LinkedIn scraper returned {len(posts)} posts")
                return posts
        print(f"   ⚠️ Primary LinkedIn scraper empty/failed (status {r.status_code}), falling back...")
    except Exception as e:
        print(f"   ⚠️ Primary LinkedIn scraper error: {e}, falling back...")

    # ── Fallback: data-slayer (proven field names, URL-only input) ──
    fallback_input = {"startUrls": [{"url": company_url}]}
    try:
        print(f"   🔄 Apify LinkedIn posts (fallback): {company_handle}...")
        r = requests.post(
            "https://api.apify.com/v2/acts/data-slayer~linkedin-company-posts-scraper/run-sync-get-dataset-items",
            params={"token": APIFY_API_KEY, "timeout": 120},
            json=fallback_input,
            timeout=150
        )
        if r.status_code in (200, 201):
            posts = r.json()
            if isinstance(posts, list) and len(posts) > 0:
                print(f"   ✅ Fallback LinkedIn scraper returned {len(posts)} posts")
                return posts
        print(f"   ⚠️ Fallback LinkedIn scraper sync empty/failed (status {r.status_code}), trying async...")
        return fetch_apify_async(company_handle, "data-slayer~linkedin-company-posts-scraper", fallback_input)
    except Exception as e:
        print(f"   ⚠️ Fallback LinkedIn scraper error: {e}")
        return []


def _classify_linkedin_post(post):
    """
    7 categories (Reposted Content deliberately excluded — a reposted
    video is just classified as a video, based on its actual attached
    format): Videos / Multi-Image / Single-Image / Documents / Polls /
    Text-Only / Other (Articles, Newsletters, Celebrations combined).

    Handles BOTH actors' confirmed output shapes:
    - brilliant_gum: "images" (array of URLs), "video" (object with
      stream_url), "reshared" / "reshared_from"
    - data-slayer:   "attachments" (array of typed objects),
      "is_repost" / "shared_post"
    """
    # ── HarvestAPI shape (CONFIRMED from official docs:
    #    postImages / postVideo / article / newsletterUrl / repost) ──
    #    For pure reposts with no media of their own, classify by the
    #    reposted content's media instead.
    src = post
    repost_obj = post.get("repost")
    if isinstance(repost_obj, dict) and repost_obj:
        has_own_media = (post.get("postImages") or post.get("postVideo") or post.get("article"))
        if not has_own_media:
            src = repost_obj

    post_video = src.get("postVideo")
    if isinstance(post_video, dict) and (post_video.get("videoUrl") or post_video.get("thumbnailUrl")):
        return "videos"
    post_images = src.get("postImages")
    if isinstance(post_images, list) and len(post_images) > 0:
        return "multi_image" if len(post_images) >= 2 else "single_image"
    document_obj = src.get("document")
    if isinstance(document_obj, dict) and document_obj:
        return "documents"
    if src.get("newsletterUrl") or (isinstance(src.get("article"), dict) and src.get("article")):
        return "other"          # Articles & Newsletters

    # ── brilliant_gum shape (images array / video object) ───────────
    video_obj = post.get("video")
    if isinstance(video_obj, dict) and video_obj.get("stream_url"):
        return "videos"
    images = post.get("images")
    if isinstance(images, list):
        if len(images) >= 2:
            return "multi_image"
        if len(images) == 1:
            return "single_image"

    # ── data-slayer shape (attachments array) ───────────────────────
    attachments = post.get("attachments") or []
    if not attachments and post.get("is_repost") and isinstance(post.get("shared_post"), dict):
        attachments = post["shared_post"].get("attachments") or []

    types = [str(a.get("type","")).lower() for a in attachments if isinstance(a, dict)]

    if any("video" in t for t in types):
        return "videos"
    if any("document" in t or "pdf" in t for t in types):
        return "documents"
    image_count = sum(1 for t in types if "image" in t or "photo" in t)
    if image_count >= 2:
        return "multi_image"
    if image_count == 1:
        return "single_image"
    if post.get("poll") or _get_first(post, ["poll_options","pollOptions"]):
        return "polls"

    text = str(post.get("text","") or post.get("commentary","") or post.get("content","")).lower()
    if not attachments and not text:
        return "text_only"
    if "linkedin.com/pulse" in text or _get_first(post, ["article_url","articleUrl"]):
        return "other"          # Article
    if any(k in text for k in ["newsletter", "subscribe to our"]):
        return "other"          # Newsletter
    if not attachments:
        return "text_only"
    return "other"               # Celebrations / anything unrecognized


def fetch_linkedin(company_handle, brand_name=None):
    print(f"💼 Fetching LinkedIn: {company_handle}")
    company_name = company_handle.replace("-"," ").replace("_"," ")
    data = {
        "company":   company_handle,
        "followers": "N/A",
        "employees": "N/A",
        "snippet":   "",
        "url":       f"https://www.linkedin.com/company/{company_handle}",
    }

    def extract_followers(text):
        for pattern in [
            r'([\d,]+\.?\d*[KMB]?\+?)\s*followers',
            r'followers[:\s]+([\d,]+\.?\d*[KMB]?\+?)',
            r'([\d,]+)\s*follow',
        ]:
            m = re.search(pattern, text, re.I)
            if m:
                return m.group(1).strip()
        return None

    def extract_employees(text):
        for pattern in [
            r'([\d,\-]+\.?\d*[KMB]?\+?)\s*employees',
            r'employees[:\s]+([\d,\-]+)',
            r'(\d+[-–]\d+)\s*employees',
        ]:
            m = re.search(pattern, text, re.I)
            if m:
                return m.group(1).strip()
        return None

    queries = [f'{company_name} linkedin followers']
    if brand_name and brand_name.lower() != company_name.lower():
        queries.append(f'{brand_name} linkedin followers')
        queries.append(f'"{brand_name}" linkedin company followers')
    queries.append(f'site:linkedin.com/company/{company_handle}')
    queries.append(f'{company_name} site:linkedin.com followers')

    for q in queries:
        try:
            print(f"   🔍 LinkedIn query: {q}")
            r = requests.get(
                "https://www.searchapi.io/api/v1/search",
                params={"engine":"google","q":q,
                        "api_key":SEARCHAPI_KEY,"num":5}
            )
            if r.status_code != 200:
                continue

            result_json = r.json()

            for key in ["knowledge_graph","answer_box"]:
                section = result_json.get(key, {})
                if section:
                    s = json.dumps(section)
                    f = extract_followers(s)
                    e = extract_employees(s)
                    if f and data["followers"] == "N/A":
                        data["followers"] = f
                        print(f"   ✅ {key} followers: {f}")
                    if e and data["employees"] == "N/A":
                        data["employees"] = e

            for res in result_json.get("organic_results", []):
                link     = res.get("link","")
                snippet  = res.get("snippet","")
                title    = res.get("title","")
                all_text = f"{title} {snippet}"
                sitelinks = res.get("sitelinks", [])
                if isinstance(sitelinks, list):
                    for sl in sitelinks:
                        all_text += f" {sl.get('title','')} {sl.get('snippet','')}"
                elif isinstance(sitelinks, dict):
                    all_text += json.dumps(sitelinks)

                f = extract_followers(all_text)
                e = extract_employees(all_text)
                if f and data["followers"] == "N/A":
                    data["followers"] = f
                    print(f"   ✅ Followers: {f}")
                if e and data["employees"] == "N/A":
                    data["employees"] = e
                if "linkedin.com/company" in link:
                    if not data.get("snippet"):
                        data["snippet"] = snippet
                    if title:
                        data["company"] = title.replace("| LinkedIn","").replace("LinkedIn","").strip()

            if data["followers"] != "N/A":
                print(f"   ✅ LinkedIn final: {data['followers']} followers")
                break  # stop trying more search queries, but still fall through
                       # to the post-level fetch below — a plain 'return' here
                       # would have skipped it entirely

        except Exception as e:
            print(f"   ⚠️ LinkedIn query failed: {e}")

    # ── Posts via Apify — reuses the exact same metric machinery as
    #    Facebook (day distribution, hour insights, consistency,
    #    momentum, per-category group metrics). ─────────────────────
    try:
        followers_n = parse_num(data.get("followers", 0))
        posts_raw = fetch_linkedin_posts_via_apify(company_handle)
        print(f"   🔍 LinkedIn posts fetched: {len(posts_raw)}")

        parsed = []
        for post in posts_raw:
            # Covers all three actors seen so far:
            # - harvestapi:    possibly nested "engagement" {likes, comments,
            #                  shares} and "postedAt" {timestamp/date}; text
            #                  in "content"
            # - brilliant_gum: num_likes / num_comments / num_reposts / posted
            # - data-slayer:   likes / comments / shares / created_at
            engagement = post.get("engagement") if isinstance(post.get("engagement"), dict) else {}
            likes    = _get_first(engagement, ["likes", "reactions"], None)
            if likes is None:
                likes = _get_first(post, ["num_likes", "likes", "reactions", "likesCount", "reactionsCount"], 0) or 0
            comments = _get_first(engagement, ["comments"], None)
            if comments is None:
                comments = _get_first(post, ["num_comments", "comments", "commentsCount"], 0) or 0
            shares   = _get_first(engagement, ["shares", "reposts"], None)
            if shares is None:
                shares = _get_first(post, ["num_reposts", "shares", "reposts", "sharesCount", "repostsCount"])

            url = _get_first(post, ["linkedinUrl", "url", "postUrl", "link"]) or "N/A"

            # Timestamp: try each candidate until one actually PARSES —
            # a plain "first non-empty" pick would grab relative strings
            # like "2 days ago" and silently lose the date. Nested
            # postedAt objects (harvestapi) are checked first.
            ts = None
            posted_at = post.get("postedAt")
            if isinstance(posted_at, dict):
                for k in ["timestamp", "date", "postedDate"]:
                    candidate = posted_at.get(k)
                    if candidate not in (None, ""):
                        ts = _to_timestamp(candidate)
                        if ts is not None:
                            break
            if ts is None:
                for ts_field in ["posted", "postedAt", "created_at", "timestamp", "date", "publishedAt", "postedDate", "time"]:
                    candidate = post.get(ts_field)
                    if isinstance(candidate, dict):
                        continue
                    if candidate not in (None, ""):
                        ts = _to_timestamp(candidate)
                        if ts is not None:
                            break

            ptype    = _classify_linkedin_post(post)
            try: likes = int(likes)
            except: likes = 0
            try: comments = int(comments)
            except: comments = 0
            parsed.append({
                "type": ptype, "likes": likes, "comments": comments,
                "shares": (int(shares) if shares is not None else None),
                "reposts": None, "views": None, "url": url, "ts": ts,
            })

        posts_with_dates    = [pp for pp in parsed if pp["ts"] is not None]
        posts_without_dates = [pp for pp in parsed if pp["ts"] is None]
        posts_with_dates.sort(key=lambda pp: pp["ts"], reverse=True)
        parsed = (posts_with_dates + posts_without_dates)[:30]

        post_count = len(parsed)
        li_categories = {}
        for cat in ["videos", "multi_image", "single_image", "documents", "polls", "text_only", "other"]:
            group = [pp for pp in parsed if pp["type"] == cat]
            li_categories[cat] = _compute_group_metrics(group, followers_n)
            li_categories[cat]["n"] = len(group)

        total_l = sum(pp["likes"] for pp in parsed)
        total_c = sum(pp["comments"] for pp in parsed)
        engagement_total = total_l + total_c
        eng_rate = f"{(engagement_total/post_count/followers_n*100):.2f}%" if (post_count and followers_n) else "N/A"

        timestamps = [pp["ts"] for pp in parsed if pp["ts"]]
        if len(timestamps) >= 2:
            span_days = max((max(timestamps) - min(timestamps)) / 86400, 1)
            posting_frequency = f"{(post_count / span_days * 7):.1f} / week"
        else:
            posting_frequency = "N/A"

        day_dist      = compute_day_distribution(parsed)
        hour_insights = compute_hour_insights(parsed)
        consistency   = compute_posting_consistency(parsed)
        momentum      = compute_momentum(parsed)

        data.update({
            "sample_size":       post_count,
            "engagement_rate":   eng_rate,
            "engagement_total":  engagement_total,
            "posting_frequency": posting_frequency,
            "most_active_day":  day_dist["most_active_day"],
            "least_active_day": day_dist["least_active_day"],
            "most_frequent_hour":   hour_insights["most_frequent_hour"],
            "best_performing_hour": hour_insights["best_performing_hour"],
            "outlier_note":         hour_insights["outlier_note"],
            "posting_consistency":  consistency,
            "momentum_pct":         momentum["momentum_pct"],
            "momentum_direction":   momentum["momentum_direction"],
            "categories": li_categories,
        })
        print(f"   ✅ {post_count} posts across 7 categories | ER: {eng_rate}")
    except Exception as e:
        import traceback
        print(f"   ⚠️ LinkedIn post-level fetch failed: {e}")
        print(traceback.format_exc())

    return data


# ═══════════════════════════════════════════════════════════════
# STEP 3 — CLAUDE ANALYSIS
# ═══════════════════════════════════════════════════════════════

def analyse_with_claude(website_url, handles, ig, fb, yt, li):
    print("\n🤖 Sending all data to Claude...")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    prompt = f"""
You are a social media analyst. Analyse this brand's social media presence.

Website: {website_url}

INSTAGRAM DATA (includes overall stats plus per-format breakdown for Images, Carousels, Reels):
{json.dumps(ig, indent=2)}

FACEBOOK DATA:
{json.dumps(fb, indent=2)}

YOUTUBE DATA:
{json.dumps(yt, indent=2)}

LINKEDIN DATA:
{json.dumps(li, indent=2)}

For the "instagram_analysis" field, write EXACTLY 3 sentences (strict maximum ~420
characters total — this must fit in a fixed-size slide card, so be concise) using these
industry benchmark bands for Engagement Rate:
- Under 1% = Low engagement
- 1% to 3.5% = Average / healthy
- 3.5% to 6% = Good
- 6%+ = Excellent
Reference the account's engagement_rate against this scale, compare Reels engagement vs
overall engagement, and comment on posting_frequency vs the general best-practice of
3-5 posts per week. Be specific and use the real numbers, but stay concise — no filler
sentences, no repeated points.

For the "facebook_analysis" field, write EXACTLY 3 sentences (strict maximum ~420
characters total) using the SAME benchmark bands as above, but applied to the FACEBOOK
DATA's engagement_rate and engagement_rate_videos fields instead. Compare video
engagement vs overall engagement, and comment on Facebook's posting_frequency vs the
3-5 posts/week best practice. If Facebook's sample_size is 0 or very low, say so plainly
instead of inventing benchmarks from missing data.

For the "linkedin_analysis" field, write EXACTLY 3 sentences (strict maximum ~420
characters total) using the SAME benchmark bands as above, applied to the LINKEDIN
DATA's engagement_rate field. LinkedIn's "categories" object breaks posts into videos,
multi_image, single_image, documents, polls, text_only, and other — mention whichever
category has the most posts or the highest engagement as the account's strongest
content format. Comment on posting_frequency vs LinkedIn's B2B best practice of
2-3 posts/week. If sample_size is 0 or very low, say so plainly instead of inventing
benchmarks from missing data.

Return ONLY a JSON object:
{{
  "brand_name": "brand name",
  "niche": "3-5 words describing the brand",
  "overall_summary": "2-3 sentence overview of their social media presence",
  "instagram": {{
    "strength": "biggest strength",
    "recommendation": "one actionable tip"
  }},
  "instagram_analysis": "EXACTLY 3 sentences, ~420 characters max, as instructed above",
  "facebook": {{
    "followers": "exact from data",
    "category": "exact from data",
    "rating": "exact from data",
    "about": "1 sentence summary",
    "strength": "biggest strength",
    "recommendation": "one actionable tip"
  }},
  "facebook_analysis": "EXACTLY 3 sentences, ~420 characters max, as instructed above",
  "youtube": {{
    "subscribers": "exact from data",
    "videos": "exact from data",
    "views": "exact from data",
    "joined": "exact from data",
    "description_summary": "1 sentence",
    "strength": "biggest strength",
    "recommendation": "one actionable tip"
  }},
  "linkedin": {{
    "followers": "exact from data",
    "employees": "exact from data",
    "summary": "1 sentence",
    "strength": "biggest strength",
    "recommendation": "one actionable tip"
  }},
  "linkedin_analysis": "EXACTLY 3 sentences, ~420 characters max, as instructed above",
  "cross_platform": {{
    "strongest_platform": "which platform performs best and why",
    "weakest_platform": "which needs most work and why",
    "overall_recommendation": "top priority action for the brand",
    "content_consistency": "are they consistent across platforms?",
    "growth_opportunity": "biggest growth opportunity"
  }}
}}
Return ONLY JSON. No markdown. No explanation.
Use EXACT numbers from data — never write N/A if data exists.
"""
    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2200,
        messages=[{"role":"user","content":prompt}]
    )
    raw = msg.content[0].text.strip()
    try:
        return json.loads(raw)
    except:
        return json.loads(raw.replace("```json","").replace("```","").strip())


# ═══════════════════════════════════════════════════════════════
# STEP 4 — MIDNIGHT PANDA BRANDED PPT HELPERS
# ═══════════════════════════════════════════════════════════════

def bg(slide, prs, color):
    s = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    s.fill.solid(); s.fill.fore_color.rgb = color; s.line.fill.background()

def tb(slide, text, x, y, w, h, size=14, bold=False, color=TEXT_DARK,
       align=PP_ALIGN.LEFT, font=FONT_MONO, italic=False, line_spacing=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf  = box.text_frame; tf.word_wrap = True
    p   = tf.paragraphs[0]
    p.text = str(text); p.alignment = align
    p.font.size = Pt(size); p.font.bold = bold
    p.font.color.rgb = color; p.font.italic = italic
    p.font.name = font
    if line_spacing:
        p.line_spacing = line_spacing
    return box

def brand_mark(slide, dark_bg=False, top_right=False):
    outer = TEXT_LIGHT if dark_bg else TEXT_DARK
    inner = TEXT_DARK if dark_bg else TEXT_LIGHT
    if top_right:
        dot_xs = [10.05, 10.29]; text_x = 10.58
    else:
        dot_xs = [0.55, 0.79];   text_x = 1.08
    for ox in dot_xs:
        c1 = slide.shapes.add_shape(9, Inches(ox), Inches(0.45), Inches(0.16), Inches(0.16))
        c1.fill.solid(); c1.fill.fore_color.rgb = outer; c1.line.fill.background()
        c2 = slide.shapes.add_shape(9, Inches(ox+0.05), Inches(0.50), Inches(0.05), Inches(0.05))
        c2.fill.solid(); c2.fill.fore_color.rgb = inner; c2.line.fill.background()
    tb(slide, "MIDNIGHT PANDA", text_x, 0.42, 2.20, 0.32, size=9.5, color=outer, font=FONT_MONO)

def footer_bar(slide, page_num, dark_bg=False):
    line_color = RGBColor(0x2A,0x2A,0x2A) if dark_bg else RGBColor(0xD8,0xD5,0xCC)
    ln = slide.shapes.add_shape(1, Inches(0), Inches(7.46), Inches(13.33), Pt(0.75))
    ln.fill.solid(); ln.fill.fore_color.rgb = line_color; ln.line.fill.background()
    tb(slide, "MIDNIGHTPANDA.AI · SOCIAL INTELLIGENCE", 0.55, 7.08, 5.00, 0.30,
       size=8.5, color=TEXT_FOOTER, font=FONT_MONO)
    tb(slide, f"{page_num:02d} / {TOTAL_SLIDES}", 11.93, 7.08, 0.85, 0.30,
       size=8.5, color=TEXT_FOOTER, font=FONT_MONO, align=PP_ALIGN.RIGHT)

def kicker_header(slide, kicker, title, subtitle, dark_bg=False):
    title_color = TEXT_LIGHT if dark_bg else TEXT_DARK
    sub_color   = TEXT_GRAY_LT if dark_bg else TEXT_GRAY
    tb(slide, kicker.upper(), 0.55, 0.42, 6.00, 0.30, size=11, color=GOLD, font=FONT_MONO_MED)
    tb(slide, title, 0.55, 0.72, 11.50, 0.65, size=26, color=title_color, font=FONT_SERIF)
    if subtitle:
        tb(slide, subtitle, 0.55, 1.34, 11.50, 0.35, size=13, color=sub_color, font=FONT_MONO_LT)

def stat_block(slide, x, y, w, value, label, size=54, dark_bg=False):
    label_color = TEXT_GRAY_LT if dark_bg else TEXT_GRAY
    vh = 0.95 if size >= 50 else (0.75 if size >= 40 else 0.55)
    tb(slide, value, x, y, w, vh, size=size, color=GOLD, font=FONT_STAT)
    tb(slide, label.upper(), x, y+vh+0.05, w, 0.30, size=9.5, color=label_color, font=FONT_MONO)

def pill(slide, x, y, text):
    w, h = 1.90, 0.50
    r = slide.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = CARD_DARK; r.line.fill.background()
    try: r.adjustments[0] = 0.5
    except: pass
    tb(slide, text, x, y+0.09, w, 0.32, size=11, color=TEXT_LIGHT, font=FONT_MONO, align=PP_ALIGN.CENTER)

def card(slide, x, y, w, h, label, body, dark_bg=False):
    fill_color = CARD_DARK if dark_bg else CARD_LIGHT
    body_color = TEXT_LIGHT if dark_bg else TEXT_DARK
    r = slide.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill_color; r.line.fill.background()
    try: r.adjustments[0] = 0.06
    except: pass
    tb(slide, label.upper(), x+0.28, y+0.18, w-0.56, 0.30, size=10.5, color=GOLD, font=FONT_MONO_MED)
    tb(slide, body, x+0.28, y+0.50, w-0.56, h-0.6, size=12.5, color=body_color,
       font=FONT_MONO_LT, line_spacing=1.15)

def card_url(slide, x, y, w, h, label, url, dark_bg=False):
    """
    Same visual style as card(), but the body is a real clickable
    hyperlink to the actual Instagram post — opens the post when
    clicked in PowerPoint / Keynote / Google Slides, instead of just
    showing plain, unclickable text.
    """
    fill_color = CARD_DARK if dark_bg else CARD_LIGHT
    body_color = TEXT_LIGHT if dark_bg else TEXT_DARK
    r = slide.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill_color; r.line.fill.background()
    try: r.adjustments[0] = 0.06
    except: pass
    tb(slide, label.upper(), x+0.28, y+0.18, w-0.56, 0.30, size=10.5, color=GOLD, font=FONT_MONO_MED)
    box = slide.shapes.add_textbox(Inches(x+0.28), Inches(y+0.50), Inches(w-0.56), Inches(h-0.6))
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.line_spacing = 1.15
    run = p.add_run()
    if url and url != "N/A":
        run.text = shorten_url(url, 50)
        run.hyperlink.address = url
        run.font.color.rgb = GOLD
    else:
        run.text = "N/A"
        run.font.color.rgb = body_color
    run.font.size = Pt(12.5)
    run.font.name = FONT_MONO_LT
    return box

def link_list_card(slide, x, y, w, h, label, items, dark_bg=False):
    """
    A compact card showing several "label: link" lines stacked vertically
    (e.g. Top reel / Worst reel / Max views) rather than squeezed onto
    one long concatenated line — that's what was overflowing past the
    card edge and colliding with the footer before. Each url becomes a
    real clickable hyperlink.

    items: list of (prefix_text, url) tuples, one per line.
    """
    fill_color = CARD_DARK if dark_bg else CARD_LIGHT
    body_color = TEXT_LIGHT if dark_bg else TEXT_DARK
    r = slide.shapes.add_shape(5, Inches(x), Inches(y), Inches(w), Inches(h))
    r.fill.solid(); r.fill.fore_color.rgb = fill_color; r.line.fill.background()
    try: r.adjustments[0] = 0.06
    except: pass
    tb(slide, label.upper(), x+0.28, y+0.14, w-0.56, 0.22, size=10, color=GOLD, font=FONT_MONO_MED)

    n = max(len(items), 1)
    line_h = (h - 0.38) / n
    for i, (prefix, url) in enumerate(items):
        box = slide.shapes.add_textbox(Inches(x+0.28), Inches(y+0.36+i*line_h), Inches(w-0.56), Inches(line_h))
        tf = box.text_frame; tf.word_wrap = False
        p = tf.paragraphs[0]
        run1 = p.add_run(); run1.text = prefix
        run1.font.size = Pt(11.5); run1.font.name = FONT_MONO_LT; run1.font.color.rgb = body_color
        run2 = p.add_run()
        if url and url != "N/A":
            run2.text = shorten_url(url, 44)
            run2.hyperlink.address = url
            run2.font.color.rgb = GOLD
        else:
            run2.text = "N/A"
            run2.font.color.rgb = body_color
        run2.font.size = Pt(11.5); run2.font.name = FONT_MONO

def linkedin_performance_slide(prs, blank, slide_num, category_label, category_data, sample_size):
    """
    One shared layout for all 7 LinkedIn content-type performance
    slides (Videos, Multi-Image, Single-Image, Documents, Polls,
    Text-Only, Other) — same visual structure as Facebook's Photos
    Performance slide, just parameterized per category so the 7 nearly
    identical slides don't need 7 copies of the same ~15 lines.
    """
    s, dark = start_slide(prs, blank, slide_num)
    n = category_data.get("n", 0)
    kicker_header(s, f"LinkedIn — {category_label}", f"{category_label} Performance",
                  f"Based on {n} {category_label.lower()} post{'s' if n != 1 else ''} from the last {sample_size}",
                  dark_bg=dark)
    stat_block(s, 0.55, 1.75, 3.86, category_data.get("avg_engagement","N/A"), "Avg Engagement / Post", size=44, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, category_data.get("er_per_follower","N/A"), "Engagement Rate / Follower", size=44, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, category_data.get("avg_likes","N/A"), "Avg Likes / Post", size=32, dark_bg=dark)
    stat_block(s, 0.55, 3.10, 3.86, category_data.get("avg_comments","N/A"), "Avg Comments / Post", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.10, 3.86, category_data.get("avg_shares","N/A"), "Avg Reposts / Post", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.10, 3.86, n, "Posts in Category", size=32, dark_bg=dark)

    if n > 0:
        card_url(s, 0.55, 4.75, 5.95, 1.75,
                 f"Top-Performing {category_label} (score {category_data.get('top_score','N/A')})",
                 category_data.get("top_url","N/A"), dark_bg=dark)
        card_url(s, 6.65, 4.75, 5.80, 1.75,
                 f"Lowest-Performing {category_label} (score {category_data.get('worst_score','N/A')})",
                 category_data.get("worst_url","N/A"), dark_bg=dark)
    else:
        card(s, 0.55, 4.75, 12.23, 1.75, "No Posts in This Category",
             f"No {category_label.lower()} posts appeared in the sample of LinkedIn posts analyzed for this report.",
             dark_bg=dark)

    footer_bar(s, slide_num, dark_bg=dark)


def branded_table(slide, headers, rows, x, y, w, h, dark_bg=False):
    cols = len(headers)
    tbl  = slide.shapes.add_table(len(rows)+1, cols, Inches(x), Inches(y), Inches(w), Inches(h)).table
    try: tbl.first_row = False; tbl.horz_banding = False
    except: pass
    cw = Inches(w/cols)
    for i in range(cols): tbl.columns[i].width = cw

    header_fill = GOLD if dark_bg else BG_DARK
    header_text = BG_DARK if dark_bg else TEXT_LIGHT
    body_fill   = CARD_DARK if dark_bg else BG_LIGHT
    body_text   = TEXT_LIGHT if dark_bg else TEXT_DARK

    for ci, hdr in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(13); p.font.name = FONT_MONO_MED
        p.font.color.rgb = header_text
        p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
        cell.margin_left = Inches(0.12)
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.cell(ri+1, ci)
            cell.text = str(val)
            cell.fill.solid(); cell.fill.fore_color.rgb = body_fill
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(13); p.font.name = FONT_MONO
            p.font.color.rgb = body_text
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.CENTER
            cell.margin_left = Inches(0.12)
    return tbl

def start_slide(prs, blank, slide_num):
    is_dark = (slide_num % 2 == 1)
    s = prs.slides.add_slide(blank)
    bg(s, prs, BG_DARK if is_dark else BG_LIGHT)
    brand_mark(s, dark_bg=is_dark, top_right=(slide_num != 1))
    return s, is_dark

def shorten_url(url, maxlen=42):
    if not url or url == "N/A":
        return "N/A"
    return url if len(url) <= maxlen else url[:maxlen-1] + "…"

def truncate_to_sentence(text, max_chars=480):
    """
    Hard safety net for card text: guarantees the string never exceeds
    max_chars, regardless of how long Claude's generated text turns out
    to be. Cuts at the last full sentence that fits (so it never ends
    mid-word), falling back to a plain word-boundary cut only if no
    sentence break exists anywhere within the limit. This is the actual
    guarantee that text can never overflow a fixed-size slide card — the
    prompt-level length instruction is just a soft, best-effort request
    on top of this.
    """
    if not text or len(text) <= max_chars:
        return text
    window = text[:max_chars]
    best_cut = -1
    for sep in [". ", "! ", "? "]:
        cut = window.rfind(sep)
        if cut > best_cut:
            best_cut = cut
    if best_cut > 0:
        return window[:best_cut+1]
    # No sentence break found anywhere in the window — cut at the last
    # full word instead, so we at least never chop a word in half.
    cut = window.rfind(" ")
    return (window[:cut] if cut > 0 else window) + "…"


# ═══════════════════════════════════════════════════════════════
# STEP 5 — BUILD 17-SLIDE MIDNIGHT PANDA BRANDED PPT
# ═══════════════════════════════════════════════════════════════

def create_ppt(analysis, handles, ig_raw, fb_raw, yt_raw, li_raw, website_url):
    prs   = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    ig = analysis.get("instagram", {})
    fb = analysis.get("facebook", {})
    yt = analysis.get("youtube", {})
    li = analysis.get("linkedin", {})
    cp = analysis.get("cross_platform", {})
    brand = analysis.get("brand_name", website_url)
    niche = analysis.get("niche", "")

    ig_images    = ig_raw.get("images", {})
    ig_carousels = ig_raw.get("carousels", {})
    ig_reels     = ig_raw.get("reels", {})

    def parse_num(val):
        try:
            v = str(val).replace(",","").upper().replace("+","")
            if "M" in v: return float(v.replace("M","")) * 1_000_000
            if "K" in v: return float(v.replace("K","")) * 1_000
            return float(v)
        except: return 0

    # ── SLIDE 1: COVER (dark) ─────────────────────────────────
    s, dark = start_slide(prs, blank, 1)
    tb(s, "SOCIAL INTELLIGENCE REPORT", 0.55, 1.55, 6.00, 0.30, size=11, color=GOLD, font=FONT_MONO_MED)
    tb(s, brand, 0.55, 1.95, 11.50, 1.50, size=58, color=TEXT_LIGHT, font=FONT_SERIF)
    tb(s, niche, 0.55, 2.95, 11.50, 0.45, size=15, color=TEXT_GRAY_LT, font=FONT_MONO_LT)
    tb(s, website_url.replace("https://","").replace("http://",""), 0.55, 3.42, 11.50, 0.35,
       size=11.5, color=GOLD, font=FONT_MONO)
    platform_labels = []
    if handles.get("instagram"): platform_labels.append("Instagram")
    if handles.get("facebook"):  platform_labels.append("Facebook")
    if handles.get("youtube"):   platform_labels.append("YouTube")
    if handles.get("linkedin"):  platform_labels.append("LinkedIn")
    for i, label in enumerate(platform_labels):
        pill(s, 0.55 + i*2.12, 4.05, label)
    stat_block(s, 0.55, 5.05, 2.83, ig_raw.get("followers","N/A"), "Instagram Followers", size=44, dark_bg=True)
    stat_block(s, 3.68, 5.05, 2.83, fb_raw.get("followers","N/A"), "Facebook Followers",  size=44, dark_bg=True)
    stat_block(s, 6.81, 5.05, 2.83, yt_raw.get("subscribers","N/A"), "YouTube Subscribers", size=44, dark_bg=True)
    stat_block(s, 9.94, 5.05, 2.83, li_raw.get("followers","N/A"), "LinkedIn Followers",   size=44, dark_bg=True)
    tb(s, f"Powered by SearchAPI.io + Apify + Claude AI  ·  {website_url.replace('https://','').replace('http://','')}",
       0.55, 6.85, 11.50, 0.30, size=10, color=TEXT_FOOTER, font=FONT_MONO)
    ln = s.shapes.add_shape(1, Inches(0), Inches(7.46), Inches(13.33), Pt(0.75))
    ln.fill.solid(); ln.fill.fore_color.rgb = RGBColor(0x2A,0x2A,0x2A); ln.line.fill.background()

    # ── SLIDE 2: Overview (light) ─────────────────────────────
    s, dark = start_slide(prs, blank, 2)
    kicker_header(s, "Report Overview", "Social Media Overview", f"{brand} — all platforms at a glance", dark_bg=dark)
    branded_table(s, ["Platform","Handle","Followers","Key Metric","Status"], [
        ["Instagram", f"@{handles.get('instagram','N/A')}", ig_raw.get("followers","N/A"), f"ER: {ig_raw.get('engagement_rate','N/A')}", "Active" if handles.get("instagram") else "Not Found"],
        ["Facebook",  handles.get("facebook","N/A"),  fb_raw.get("followers","N/A"), f"Rating: {fb.get('rating','N/A')}",       "Active" if handles.get("facebook") else "Not Found"],
        ["YouTube",   handles.get("youtube","N/A"),   yt_raw.get("subscribers","N/A"), f"Videos: {yt_raw.get('videos','N/A')}", "Active" if handles.get("youtube") else "Not Found"],
        ["LinkedIn",  handles.get("linkedin","N/A"),  li_raw.get("followers","N/A"), f"Employees: {li_raw.get('employees','N/A')}", "Active" if handles.get("linkedin") else "Not Found"],
    ], 0.55, 1.65, 12.23, 2.35, dark_bg=dark)
    card(s, 0.55, 4.35, 12.23, 2.15, "Overall Analysis", analysis.get("overall_summary",""), dark_bg=dark)
    footer_bar(s, 2, dark_bg=dark)

    # ── SLIDE 3: Instagram Overview (dark) ────────────────────
    s, dark = start_slide(prs, blank, 3)
    kicker_header(s, "Instagram", "Instagram Overview",
                  f"@{handles.get('instagram','N/A')}  ·  Based on last {ig_raw.get('sample_size','N/A')} posts", dark_bg=dark)

    # ── Core metrics (unchanged) ──────────────────────────────────────
    stat_block(s, 0.55, 1.75, 3.86, ig_raw.get("followers","N/A"), "Followers", size=54, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, ig_raw.get("posts","N/A"), "Total Posts (all time)", size=54, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, ig_raw.get("posting_frequency","N/A"), "Posting Frequency", size=44, dark_bg=dark)
    stat_block(s, 0.55, 3.15, 3.86, ig_raw.get("engagement_rate","N/A"), "Engagement Rate / Follower", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.15, 3.86, ig_raw.get("engagement_rate_reels","N/A"), "Engagement Rate (Reels Only)", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.15, 3.86, ig_raw.get("engagement_total","N/A"), "Engagement", size=32, dark_bg=dark)

    # ── Thin divider — separates the core metrics above from the
    #    posting-behavior insights below, for cleaner visual grouping ──
    divider_color = RGBColor(0x2A,0x2A,0x2A) if dark else RGBColor(0xD8,0xD5,0xCC)
    ln3 = s.shapes.add_shape(1, Inches(0.55), Inches(4.12), Inches(12.23), Pt(0.75))
    ln3.fill.solid(); ln3.fill.fore_color.rgb = divider_color; ln3.line.fill.background()

    # ── Row 3: hour insights, consistency, momentum — compact text
    #    lines rather than giant gold numbers, so these extra stats
    #    don't visually overwhelm an already busy slide. ────────────
    sub_color = TEXT_GRAY_LT if dark else TEXT_GRAY
    tb(s, f"Most Frequent Hour: {ig_raw.get('most_frequent_hour','N/A')}   ·   Best Performing Hour: {ig_raw.get('best_performing_hour','N/A')}",
       0.55, 4.24, 12.23, 0.24, size=12, color=sub_color, font=FONT_MONO_LT)
    tb(s, f"Posting Consistency: {ig_raw.get('posting_consistency','N/A')}",
       0.55, 4.50, 5.95, 0.24, size=12, color=sub_color, font=FONT_MONO_LT)

    momentum_direction = ig_raw.get("momentum_direction")
    momentum_pct       = ig_raw.get("momentum_pct", "N/A")
    if momentum_direction == "up":
        momentum_color, momentum_text = MOMENTUM_UP,   f"Momentum: \u25B2 {momentum_pct}"
    elif momentum_direction == "down":
        momentum_color, momentum_text = MOMENTUM_DOWN, f"Momentum: \u25BC {momentum_pct}"
    else:
        momentum_color, momentum_text = sub_color,      f"Momentum: {momentum_pct}"
    tb(s, momentum_text, 6.60, 4.50, 5.65, 0.24, size=12, bold=True, color=momentum_color, font=FONT_MONO_MED)

    # ── Instagram Analysis — full slide width now that the photo is
    #    gone, with a hard character cap so long Claude output can
    #    never overflow the card, regardless of what gets generated. ──
    analysis_text = truncate_to_sentence(analysis.get("instagram_analysis",""), max_chars=480)
    card(s, 0.55, 4.86, 12.23, 1.55, "Instagram Analysis", analysis_text, dark_bg=dark)

    tb(s, f"Most Active Day: {ig_raw.get('most_active_day','N/A')}   ·   Least Active Day: {ig_raw.get('least_active_day','N/A')}   ·   (all dates in UTC)",
       0.55, 6.50, 12.23, 0.22, size=11, color=sub_color, font=FONT_MONO_LT)

    outlier_note = ig_raw.get("outlier_note")
    if outlier_note:
        tb(s, outlier_note, 0.55, 6.74, 12.23, 0.22, size=11, color=GOLD, font=FONT_MONO_LT)

    footer_bar(s, 3, dark_bg=dark)

    # ── SLIDE 4: Instagram Content Strategy (light) ───────────
    s, dark = start_slide(prs, blank, 4)
    kicker_header(s, "Instagram", "Content Strategy",
                  f"Content mix from the last {ig_raw.get('sample_size',30)} posts", dark_bg=dark)
    img_c = ig_raw.get("img_count",0); car_c = ig_raw.get("car_count",0); vid_c = ig_raw.get("vid_count",0)
    total = max(img_c+car_c+vid_c, 1)
    cd = ChartData()
    cd.categories = ["Images","Carousels","Videos / Reels"]
    cd.add_series("Posts",(img_c,car_c,vid_c))
    chart = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.85),Inches(1.85),Inches(5.40),Inches(4.50),cd).chart
    chart.has_legend = True
    for i, col in enumerate([GOLD, CARD_DARK, PIE_TAN]):
        pt = chart.plots[0].series[0].points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    for lbl, cnt, y in [("Videos / Reels",vid_c,2.10),("Carousels",car_c,3.25),("Images",img_c,4.40)]:
        pct = int(round(cnt/total*100))
        stat_block(s, 6.85, y, 5.40, f"{pct}%", f"{lbl} — {cnt} posts", size=32, dark_bg=dark)
    card(s, 6.85, 5.55, 5.60, 1.00, "Instagram Strength", ig.get("strength",""), dark_bg=dark)
    footer_bar(s, 4, dark_bg=dark)

    # ── SLIDE 5: Images Performance (dark) ────────────────────
    s, dark = start_slide(prs, blank, 5)
    kicker_header(s, "Instagram — Images", "Images Performance",
                  f"Based on {ig_images.get('n',0)} image posts from the last {ig_raw.get('sample_size',30)}", dark_bg=dark)
    stat_block(s, 0.55, 1.75, 3.86, ig_images.get("avg_engagement","N/A"), "Avg Engagement / Post", size=44, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, ig_images.get("er_per_follower","N/A"), "Engagement Rate / Follower", size=44, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, ig_images.get("avg_likes","N/A"), "Avg Likes / Post", size=32, dark_bg=dark)
    stat_block(s, 0.55, 3.10, 3.86, ig_images.get("avg_comments","N/A"), "Avg Comments / Post", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.10, 3.86, ig_images.get("avg_shares","N/A"), "Avg Shares / Post", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.10, 3.86, ig_images.get("avg_reposts","N/A"), "Avg Reposts / Post", size=32, dark_bg=dark)
    card_url(s, 0.55, 4.75, 5.95, 1.75,
             f"Top-Performing Image (score {ig_images.get('top_score','N/A')})",
             ig_images.get("top_url","N/A"), dark_bg=dark)
    card_url(s, 6.65, 4.75, 5.80, 1.75,
             f"Lowest-Performing Image (score {ig_images.get('worst_score','N/A')})",
             ig_images.get("worst_url","N/A"), dark_bg=dark)
    footer_bar(s, 5, dark_bg=dark)

    # ── SLIDE 6: Carousels Performance (light) ────────────────
    s, dark = start_slide(prs, blank, 6)
    kicker_header(s, "Instagram — Carousels", "Carousels Performance",
                  f"Based on {ig_carousels.get('n',0)} carousel posts from the last {ig_raw.get('sample_size',30)}", dark_bg=dark)
    stat_block(s, 0.55, 1.75, 3.86, ig_carousels.get("avg_engagement","N/A"), "Avg Engagement / Post", size=44, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, ig_carousels.get("er_per_follower","N/A"), "Engagement Rate / Follower", size=44, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, ig_carousels.get("avg_likes","N/A"), "Avg Likes / Post", size=32, dark_bg=dark)
    stat_block(s, 0.55, 3.10, 3.86, ig_carousels.get("avg_comments","N/A"), "Avg Comments / Post", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.10, 3.86, ig_carousels.get("avg_shares","N/A"), "Avg Shares / Post", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.10, 3.86, ig_carousels.get("avg_reposts","N/A"), "Avg Reposts / Post", size=32, dark_bg=dark)
    card_url(s, 0.55, 4.75, 5.95, 1.75,
             f"Top-Performing Carousel (score {ig_carousels.get('top_score','N/A')})",
             ig_carousels.get("top_url","N/A"), dark_bg=dark)
    card_url(s, 6.65, 4.75, 5.80, 1.75,
             f"Lowest-Performing Carousel (score {ig_carousels.get('worst_score','N/A')})",
             ig_carousels.get("worst_url","N/A"), dark_bg=dark)
    footer_bar(s, 6, dark_bg=dark)

    # ── SLIDE 7: Reels Performance (dark) ─────────────────────
    s, dark = start_slide(prs, blank, 7)
    kicker_header(s, "Instagram — Reels", "Reels Performance",
                  f"Based on {ig_reels.get('n',0)} reels from the last {ig_raw.get('sample_size',30)}", dark_bg=dark)
    stat_block(s, 0.55, 1.75, 3.86, ig_reels.get("avg_engagement","N/A"), "Avg Engagement / Reel", size=40, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, ig_reels.get("er_per_follower","N/A"), "Engagement Rate / Follower", size=40, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, ig_reels.get("avg_views","N/A"), "Avg Views / Reel", size=40, dark_bg=dark)
    stat_block(s, 0.55, 2.95, 3.86, ig_reels.get("avg_likes","N/A"), "Avg Likes / Reel", size=28, dark_bg=dark)
    stat_block(s, 4.73, 2.95, 3.86, ig_reels.get("avg_comments","N/A"), "Avg Comments / Reel", size=28, dark_bg=dark)
    stat_block(s, 8.92, 2.95, 3.86, ig_reels.get("avg_shares","N/A"), "Avg Shares / Reel", size=28, dark_bg=dark)
    stat_block(s, 0.55, 4.05, 3.86, ig_reels.get("avg_reposts","N/A"), "Avg Reposts / Reel", size=28, dark_bg=dark)
    stat_block(s, 4.73, 4.05, 3.86, ig_reels.get("like_rate","N/A"), "Like Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 8.92, 4.05, 3.86, ig_reels.get("comment_rate","N/A"), "Comment Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 0.55, 5.15, 3.86, ig_reels.get("share_rate","N/A"), "Share Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 4.73, 5.15, 3.86, ig_reels.get("repost_rate","N/A"), "Repost Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 8.92, 5.15, 3.86, ig_reels.get("er_by_view","N/A"), "Engagement Rate by View", size=28, dark_bg=dark)
    link_list_card(s, 0.55, 6.10, 12.23, 0.92, "Top / Worst / Max Views", [
        (f"Top reel (score {ig_reels.get('top_score','N/A')}):  ",       ig_reels.get("top_url","N/A")),
        (f"Worst reel (score {ig_reels.get('worst_score','N/A')}):  ",   ig_reels.get("worst_url","N/A")),
        (f"Max views ({ig_reels.get('max_views','N/A')}):  ",            ig_reels.get("max_views_url","N/A")),
    ], dark_bg=dark)
    footer_bar(s, 7, dark_bg=dark)

    # ── SLIDE 8: Facebook Overview (light) ──────────────────────
    s, dark = start_slide(prs, blank, 8)
    kicker_header(s, "Facebook", "Facebook Overview",
                  f"{handles.get('facebook','N/A')}  ·  Based on last {fb_raw.get('sample_size','N/A')} posts", dark_bg=dark)

    stat_block(s, 0.55, 1.75, 3.86, fb_raw.get("followers","N/A"), "Page Followers", size=54, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, fb_raw.get("sample_size","N/A"), "Posts Analyzed", size=54, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, fb_raw.get("posting_frequency","N/A"), "Posting Frequency", size=44, dark_bg=dark)
    stat_block(s, 0.55, 3.15, 3.86, fb_raw.get("engagement_rate","N/A"), "Engagement Rate / Follower", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.15, 3.86, fb_raw.get("engagement_rate_videos","N/A"), "Engagement Rate (Videos Only)", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.15, 3.86, fb_raw.get("engagement_total","N/A"), "Engagement", size=32, dark_bg=dark)

    fb_divider_color = RGBColor(0x2A,0x2A,0x2A) if dark else RGBColor(0xD8,0xD5,0xCC)
    ln8 = s.shapes.add_shape(1, Inches(0.55), Inches(4.12), Inches(12.23), Pt(0.75))
    ln8.fill.solid(); ln8.fill.fore_color.rgb = fb_divider_color; ln8.line.fill.background()

    fb_sub_color = TEXT_GRAY_LT if dark else TEXT_GRAY
    tb(s, f"Most Frequent Hour: {fb_raw.get('most_frequent_hour','N/A')}   ·   Best Performing Hour: {fb_raw.get('best_performing_hour','N/A')}",
       0.55, 4.24, 12.23, 0.24, size=12, color=fb_sub_color, font=FONT_MONO_LT)
    tb(s, f"Posting Consistency: {fb_raw.get('posting_consistency','N/A')}",
       0.55, 4.50, 5.95, 0.24, size=12, color=fb_sub_color, font=FONT_MONO_LT)

    fb_momentum_direction = fb_raw.get("momentum_direction")
    fb_momentum_pct       = fb_raw.get("momentum_pct", "N/A")
    if fb_momentum_direction == "up":
        fb_momentum_color, fb_momentum_text = MOMENTUM_UP,   f"Momentum: \u25B2 {fb_momentum_pct}"
    elif fb_momentum_direction == "down":
        fb_momentum_color, fb_momentum_text = MOMENTUM_DOWN, f"Momentum: \u25BC {fb_momentum_pct}"
    else:
        fb_momentum_color, fb_momentum_text = fb_sub_color,   f"Momentum: {fb_momentum_pct}"
    tb(s, fb_momentum_text, 6.60, 4.50, 5.65, 0.24, size=12, bold=True, color=fb_momentum_color, font=FONT_MONO_MED)

    fb_analysis_text = truncate_to_sentence(analysis.get("facebook_analysis",""), max_chars=480)
    card(s, 0.55, 4.86, 12.23, 1.55, "Facebook Analysis", fb_analysis_text, dark_bg=dark)

    tb(s, f"Most Active Day: {fb_raw.get('most_active_day','N/A')}   ·   Least Active Day: {fb_raw.get('least_active_day','N/A')}   ·   (all dates in UTC)",
       0.55, 6.50, 12.23, 0.22, size=11, color=fb_sub_color, font=FONT_MONO_LT)

    fb_outlier_note = fb_raw.get("outlier_note")
    if fb_outlier_note:
        tb(s, fb_outlier_note, 0.55, 6.74, 12.23, 0.22, size=11, color=GOLD, font=FONT_MONO_LT)

    footer_bar(s, 8, dark_bg=dark)

    # ── SLIDE 9: Facebook Content Strategy (dark) ───────────────
    s, dark = start_slide(prs, blank, 9)
    kicker_header(s, "Facebook", "Content Strategy",
                  f"Content mix from the last {fb_raw.get('sample_size',30)} posts", dark_bg=dark)
    fb_photo_c = fb_raw.get("photo_count",0); fb_video_c = fb_raw.get("video_count",0); fb_link_c = fb_raw.get("link_count",0)
    fb_total = max(fb_photo_c+fb_video_c+fb_link_c, 1)
    fb_cd = ChartData()
    fb_cd.categories = ["Photos","Videos","Links"]
    fb_cd.add_series("Posts",(fb_photo_c,fb_video_c,fb_link_c))
    fb_chart = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.85),Inches(1.85),Inches(5.40),Inches(4.50),fb_cd).chart
    fb_chart.has_legend = True
    for i, col in enumerate([GOLD, CARD_DARK, PIE_TAN]):
        pt = fb_chart.plots[0].series[0].points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    for lbl, cnt, y in [("Videos",fb_video_c,2.10),("Photos",fb_photo_c,3.25),("Links",fb_link_c,4.40)]:
        pct = int(round(cnt/fb_total*100))
        stat_block(s, 6.85, y, 5.40, f"{pct}%", f"{lbl} — {cnt} posts", size=32, dark_bg=dark)
    card(s, 6.85, 5.55, 5.60, 1.00, "Facebook Strength", fb.get("strength",""), dark_bg=dark)
    footer_bar(s, 9, dark_bg=dark)

    # ── SLIDE 10: Facebook Photos Performance (light) ───────────
    fb_photos = fb_raw.get("photos", {})
    s, dark = start_slide(prs, blank, 10)
    kicker_header(s, "Facebook — Photos", "Photos Performance",
                  f"Based on {fb_photos.get('n',0)} photo posts from the last {fb_raw.get('sample_size',30)}", dark_bg=dark)
    stat_block(s, 0.55, 1.75, 3.86, fb_photos.get("avg_engagement","N/A"), "Avg Engagement / Post", size=44, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, fb_photos.get("er_per_follower","N/A"), "Engagement Rate / Follower", size=44, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, fb_photos.get("avg_likes","N/A"), "Avg Likes / Post", size=32, dark_bg=dark)
    stat_block(s, 0.55, 3.10, 3.86, fb_photos.get("avg_comments","N/A"), "Avg Comments / Post", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.10, 3.86, fb_photos.get("avg_shares","N/A"), "Avg Shares / Post", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.10, 3.86, fb_raw.get("engagement_rate","N/A"), "Overall Engagement Rate", size=32, dark_bg=dark)
    card_url(s, 0.55, 4.75, 5.95, 1.75,
             f"Top-Performing Photo (score {fb_photos.get('top_score','N/A')})",
             fb_photos.get("top_url","N/A"), dark_bg=dark)
    card_url(s, 6.65, 4.75, 5.80, 1.75,
             f"Lowest-Performing Photo (score {fb_photos.get('worst_score','N/A')})",
             fb_photos.get("worst_url","N/A"), dark_bg=dark)
    footer_bar(s, 10, dark_bg=dark)

    # ── SLIDE 11: Facebook Videos Performance (dark) ────────────
    fb_videos = fb_raw.get("videos", {})
    s, dark = start_slide(prs, blank, 11)
    kicker_header(s, "Facebook — Videos", "Videos Performance",
                  f"Based on {fb_videos.get('n',0)} videos from the last {fb_raw.get('sample_size',30)}", dark_bg=dark)
    stat_block(s, 0.55, 1.75, 3.86, fb_videos.get("avg_engagement","N/A"), "Avg Engagement / Video", size=40, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, fb_videos.get("er_per_follower","N/A"), "Engagement Rate / Follower", size=40, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, fb_videos.get("avg_views","N/A"), "Avg Views / Video", size=40, dark_bg=dark)
    stat_block(s, 0.55, 2.95, 3.86, fb_videos.get("avg_likes","N/A"), "Avg Likes / Video", size=28, dark_bg=dark)
    stat_block(s, 4.73, 2.95, 3.86, fb_videos.get("avg_comments","N/A"), "Avg Comments / Video", size=28, dark_bg=dark)
    stat_block(s, 8.92, 2.95, 3.86, fb_videos.get("avg_shares","N/A"), "Avg Shares / Video", size=28, dark_bg=dark)
    stat_block(s, 0.55, 4.05, 3.86, fb_videos.get("like_rate","N/A"), "Like Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 4.73, 4.05, 3.86, fb_videos.get("comment_rate","N/A"), "Comment Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 8.92, 4.05, 3.86, fb_videos.get("share_rate","N/A"), "Share Rate (per view)", size=28, dark_bg=dark)
    stat_block(s, 0.55, 5.15, 3.86, fb_videos.get("er_by_view","N/A"), "Engagement Rate by View", size=28, dark_bg=dark)
    link_list_card(s, 0.55, 6.10, 12.23, 0.92, "Top / Worst / Max Views", [
        (f"Top video (score {fb_videos.get('top_score','N/A')}):  ",     fb_videos.get("top_url","N/A")),
        (f"Worst video (score {fb_videos.get('worst_score','N/A')}):  ", fb_videos.get("worst_url","N/A")),
        (f"Max views ({fb_videos.get('max_views','N/A')}):  ",           fb_videos.get("max_views_url","N/A")),
    ], dark_bg=dark)
    footer_bar(s, 11, dark_bg=dark)

    # ── SLIDE 12: YouTube Channel (light) ───────────────────────
    s, dark = start_slide(prs, blank, 12)
    kicker_header(s, "YouTube", "YouTube Channel", handles.get('youtube','N/A'), dark_bg=dark)
    tb(s, str(yt.get("description_summary","") or yt_raw.get("description",""))[:220],
       0.55, 1.62, 12.23, 0.55, size=12.5, color=(TEXT_GRAY_LT if dark else TEXT_GRAY), font=FONT_MONO_LT)
    stat_block(s, 0.55, 2.50, 3.86, yt_raw.get("subscribers","N/A"), "Subscribers", size=54, dark_bg=dark)
    stat_block(s, 4.73, 2.50, 3.86, yt_raw.get("videos","N/A"), "Total Videos", size=54, dark_bg=dark)
    stat_block(s, 8.92, 2.50, 3.86, yt_raw.get("views","N/A"), "Total Views", size=54, dark_bg=dark)
    stat_block(s, 0.55, 3.85, 3.86, yt_raw.get("is_verified","No"), "Verified", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.85, 3.86, yt_raw.get("joined","N/A"), "Joined Date", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.85, 3.86, "YouTube", "Platform", size=32, dark_bg=dark)
    card(s, 0.55, 5.35, 12.23, 1.15, "Reach Signal", yt.get("strength",""), dark_bg=dark)
    footer_bar(s, 12, dark_bg=dark)

    # ── SLIDE 13: YouTube Insights (dark) ───────────────────────
    s, dark = start_slide(prs, blank, 13)
    kicker_header(s, "YouTube", "Channel Insights", "Channel performance and strategic analysis", dark_bg=dark)
    branded_table(s, ["Metric","Value"], [
        ["Subscribers",  yt_raw.get("subscribers","N/A")],
        ["Total Videos", yt_raw.get("videos","N/A")],
        ["Total Views",  yt_raw.get("views","N/A")],
        ["Channel Age",  yt_raw.get("joined","N/A")],
    ], 0.55, 1.65, 5.60, 2.75, dark_bg=dark)
    try:
        vpv = int(str(yt_raw.get("views","0")).replace(",","")) // max(int(str(yt_raw.get("videos","1")).replace(",","")),1)
        vpv_str = f"{vpv:,}"
    except:
        vpv_str = "N/A"
    stat_block(s, 0.55, 4.75, 5.60, vpv_str, "Views Per Video (Est.)", size=54, dark_bg=dark)
    card(s, 6.75, 1.65, 5.70, 1.75, "Strength", yt.get("strength","N/A"), dark_bg=dark)
    card(s, 6.75, 3.60, 5.70, 2.35, "Recommendation", yt.get("recommendation","N/A"), dark_bg=dark)
    footer_bar(s, 13, dark_bg=dark)

    # ── SLIDE 14: LinkedIn Overview (light) ─────────────────────
    li_categories = li_raw.get("categories", {})
    s, dark = start_slide(prs, blank, 14)
    kicker_header(s, "LinkedIn", "LinkedIn Overview",
                  f"{handles.get('linkedin','N/A')}  ·  Based on last {li_raw.get('sample_size','N/A')} posts", dark_bg=dark)

    stat_block(s, 0.55, 1.75, 3.86, li_raw.get("followers","N/A"), "Followers", size=54, dark_bg=dark)
    stat_block(s, 4.73, 1.75, 3.86, li_raw.get("sample_size","N/A"), "Posts Analyzed", size=54, dark_bg=dark)
    stat_block(s, 8.92, 1.75, 3.86, li_raw.get("posting_frequency","N/A"), "Posting Frequency", size=44, dark_bg=dark)
    stat_block(s, 0.55, 3.15, 3.86, li_raw.get("engagement_rate","N/A"), "Engagement Rate / Follower", size=32, dark_bg=dark)
    stat_block(s, 4.73, 3.15, 3.86, li_raw.get("engagement_total","N/A"), "Engagement", size=32, dark_bg=dark)
    stat_block(s, 8.92, 3.15, 3.86, li_raw.get("employees","N/A"), "Employees", size=32, dark_bg=dark)

    li_divider_color = RGBColor(0x2A,0x2A,0x2A) if dark else RGBColor(0xD8,0xD5,0xCC)
    ln14 = s.shapes.add_shape(1, Inches(0.55), Inches(4.12), Inches(12.23), Pt(0.75))
    ln14.fill.solid(); ln14.fill.fore_color.rgb = li_divider_color; ln14.line.fill.background()

    li_sub_color = TEXT_GRAY_LT if dark else TEXT_GRAY
    tb(s, f"Most Frequent Hour: {li_raw.get('most_frequent_hour','N/A')}   ·   Best Performing Hour: {li_raw.get('best_performing_hour','N/A')}",
       0.55, 4.24, 12.23, 0.24, size=12, color=li_sub_color, font=FONT_MONO_LT)
    tb(s, f"Posting Consistency: {li_raw.get('posting_consistency','N/A')}",
       0.55, 4.50, 5.95, 0.24, size=12, color=li_sub_color, font=FONT_MONO_LT)

    li_momentum_direction = li_raw.get("momentum_direction")
    li_momentum_pct       = li_raw.get("momentum_pct", "N/A")
    if li_momentum_direction == "up":
        li_momentum_color, li_momentum_text = MOMENTUM_UP,   f"Momentum: \u25B2 {li_momentum_pct}"
    elif li_momentum_direction == "down":
        li_momentum_color, li_momentum_text = MOMENTUM_DOWN, f"Momentum: \u25BC {li_momentum_pct}"
    else:
        li_momentum_color, li_momentum_text = li_sub_color,   f"Momentum: {li_momentum_pct}"
    tb(s, li_momentum_text, 6.60, 4.50, 5.65, 0.24, size=12, bold=True, color=li_momentum_color, font=FONT_MONO_MED)

    li_analysis_text = truncate_to_sentence(analysis.get("linkedin_analysis",""), max_chars=480)
    card(s, 0.55, 4.86, 12.23, 1.55, "LinkedIn Analysis", li_analysis_text, dark_bg=dark)

    tb(s, f"Most Active Day: {li_raw.get('most_active_day','N/A')}   ·   Least Active Day: {li_raw.get('least_active_day','N/A')}   ·   (all dates in UTC)",
       0.55, 6.50, 12.23, 0.22, size=11, color=li_sub_color, font=FONT_MONO_LT)

    li_outlier_note = li_raw.get("outlier_note")
    if li_outlier_note:
        tb(s, li_outlier_note, 0.55, 6.74, 12.23, 0.22, size=11, color=GOLD, font=FONT_MONO_LT)

    footer_bar(s, 14, dark_bg=dark)

    # ── SLIDE 15: LinkedIn Content Strategy — 7-category pie ────
    s, dark = start_slide(prs, blank, 15)
    kicker_header(s, "LinkedIn", "Content Strategy",
                  f"Content mix from the last {li_raw.get('sample_size',30)} posts", dark_bg=dark)

    li_cat_order = [
        ("videos",       "Videos"),
        ("multi_image",  "Multi-Image"),
        ("single_image", "Single-Image"),
        ("documents",    "Documents"),
        ("polls",        "Polls"),
        ("text_only",    "Text-Only"),
        ("other",        "Other"),
    ]
    li_counts = [li_categories.get(key, {}).get("n", 0) for key, _ in li_cat_order]
    li_total  = max(sum(li_counts), 1)

    li_cd = ChartData()
    li_cd.categories = [label for _, label in li_cat_order]
    li_cd.add_series("Posts", tuple(li_counts))
    li_chart = s.shapes.add_chart(XL_CHART_TYPE.PIE, Inches(0.55),Inches(1.85),Inches(5.90),Inches(4.60),li_cd).chart
    li_chart.has_legend = True
    li_chart.legend.font.size = Pt(9)
    LINKEDIN_PIE_COLORS = [
        GOLD, RGBColor(0x8B,0x6F,0x47), CARD_DARK if not dark else RGBColor(0x3A,0x3A,0x3A),
        PIE_TAN, RGBColor(0x6B,0x59,0x40), RGBColor(0xA8,0xA2,0x96), RGBColor(0x4A,0x38,0x26),
    ]
    for i, col in enumerate(LINKEDIN_PIE_COLORS):
        pt = li_chart.plots[0].series[0].points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col

    # Compact 7-row list, right of the pie — a full stat_block per
    # category (like Instagram's 3-slice version) wouldn't fit 7 rows
    # cleanly, so this uses smaller text rows instead.
    li_list_y = 1.95
    for key, label in li_cat_order:
        cnt = li_categories.get(key, {}).get("n", 0)
        pct = int(round(cnt / li_total * 100))
        tb(s, f"{label}", 6.70, li_list_y, 3.20, 0.30, size=13, bold=True,
           color=(TEXT_LIGHT if dark else TEXT_DARK), font=FONT_MONO_MED)
        tb(s, f"{pct}%  —  {cnt} post{'s' if cnt != 1 else ''}", 10.00, li_list_y, 2.78, 0.30,
           size=13, color=GOLD, font=FONT_MONO_MED)
        li_list_y += 0.44

    footer_bar(s, 15, dark_bg=dark)

    # ── SLIDES 16-22: LinkedIn Performance — one per content type ──
    li_sample_size = li_raw.get("sample_size", 30)
    for i, (key, label) in enumerate(li_cat_order):
        linkedin_performance_slide(
            prs, blank, 16 + i, label,
            li_categories.get(key, {"n": 0}),
            li_sample_size
        )

    # ── SLIDE 23: Cross-Platform Comparison (dark) ──────────────
    s, dark = start_slide(prs, blank, 23)
    kicker_header(s, "Cross-Platform", "Follower Comparison", "Audience size across all platforms", dark_bg=dark)
    ig_f = parse_num(ig_raw.get("followers",0)); fb_f = parse_num(fb_raw.get("followers",0))
    yt_f = parse_num(yt_raw.get("subscribers",0)); li_f = parse_num(li_raw.get("followers",0))
    cd2 = ChartData()
    cd2.categories = ["Instagram","Facebook","YouTube","LinkedIn"]
    cd2.add_series("Followers",(ig_f,fb_f,yt_f,li_f))
    chart2 = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                                 Inches(0.55),Inches(1.75),Inches(12.23),Inches(3.65),cd2).chart
    chart2.has_legend = False
    pt0 = chart2.plots[0].series[0]
    for i in range(4):
        pt0.points[i].format.fill.solid(); pt0.points[i].format.fill.fore_color.rgb = GOLD
    chart2.category_axis.tick_labels.font.size = Pt(11)
    chart2.category_axis.tick_labels.font.color.rgb = TEXT_GRAY_LT
    chart2.value_axis.tick_labels.font.size = Pt(10)
    chart2.value_axis.tick_labels.font.color.rgb = TEXT_GRAY_LT
    for i,(lbl,val) in enumerate([
        ("Instagram", ig_raw.get("followers","N/A")),
        ("Facebook",  fb_raw.get("followers","N/A")),
        ("YouTube",   yt_raw.get("subscribers","N/A")),
        ("LinkedIn",  li_raw.get("followers","N/A")),
    ]):
        stat_block(s, 0.55 + i*3.13, 5.65, 2.83, val, lbl, size=32, dark_bg=dark)
    footer_bar(s, 23, dark_bg=dark)

    # ── SLIDE 24: Engagement Benchmarks (light) ─────────────────
    s, dark = start_slide(prs, blank, 24)
    kicker_header(s, "Benchmarks", "Engagement Benchmarks",
                  "Performance metrics vs. industry standards", dark_bg=dark)
    def er_status(val):
        try:
            v = float(str(val).replace("%",""))
            return "Strong" if v>=3 else ("Average" if v>=1 else "Low")
        except: return "—"
    branded_table(s, ["Platform","Key Metric","Value","Benchmark","Status"], [
        ["Instagram","Engagement Rate",   ig_raw.get("engagement_rate","N/A"),       ">3% is good",        er_status(ig_raw.get("engagement_rate","0"))],
        ["Instagram","Reels Engagement",  ig_raw.get("engagement_rate_reels","N/A"), ">3% is good",        er_status(ig_raw.get("engagement_rate_reels","0"))],
        ["Instagram","Posting Frequency", ig_raw.get("posting_frequency","N/A"),     "3–5 / week",         "—"],
        ["Facebook", "Engagement Rate",   fb_raw.get("engagement_rate","N/A"),       ">3% is good",        er_status(fb_raw.get("engagement_rate","0"))],
        ["Facebook", "Posting Frequency", fb_raw.get("posting_frequency","N/A"),     "3–5 / week",         "—"],
        ["YouTube",  "Subscribers",       yt_raw.get("subscribers","N/A"),           "Varies by niche",    "—"],
        ["YouTube",  "Views / Video",     vpv_str,                                   ">1,000 is good",     "—"],
        ["LinkedIn", "Engagement Rate",   li_raw.get("engagement_rate","N/A"),       ">2% is good (B2B)",  er_status(li_raw.get("engagement_rate","0"))],
        ["LinkedIn", "Posting Frequency", li_raw.get("posting_frequency","N/A"),     "2–3 / week",         "—"],
    ], 0.55, 1.65, 12.23, 4.90, dark_bg=dark)
    footer_bar(s, 24, dark_bg=dark)

    # ── SLIDE 25: Strengths & Gaps (dark) ───────────────────────
    s, dark = start_slide(prs, blank, 25)
    kicker_header(s, "Strengths & Gaps", "Key Strengths & Gaps",
                  "What's working, and what needs attention", dark_bg=dark)
    card(s, 0.55, 1.65, 5.96, 1.35, "Strongest Platform",  cp.get("strongest_platform","N/A"),  dark_bg=dark)
    card(s, 6.83, 1.65, 5.96, 1.35, "Needs Most Work",     cp.get("weakest_platform","N/A"),    dark_bg=dark)
    card(s, 0.55, 3.25, 5.96, 1.35, "Content Consistency", cp.get("content_consistency","N/A"), dark_bg=dark)
    card(s, 6.83, 3.25, 5.96, 1.35, "Growth Opportunity",  cp.get("growth_opportunity","N/A"),  dark_bg=dark)
    card(s, 0.55, 4.95, 12.23, 1.55, "Overall Summary",    analysis.get("overall_summary","N/A"), dark_bg=dark)
    footer_bar(s, 25, dark_bg=dark)

    # ── SLIDE 26: Recommendations (light) ───────────────────────
    s, dark = start_slide(prs, blank, 26)
    tb(s, "ACTION PLAN", 0.55, 0.42, 6.00, 0.30, size=11, color=GOLD, font=FONT_MONO_MED)
    tb(s, "Strategic Recommendations", 0.55, 0.72, 11.50, 0.65,
       size=26, color=(TEXT_LIGHT if dark else TEXT_DARK), font=FONT_SERIF)
    tb(s, "Prioritised next steps by platform", 0.55, 1.34, 11.50, 0.35,
       size=13, color=(TEXT_GRAY_LT if dark else TEXT_GRAY), font=FONT_MONO_LT)
    card(s, 0.55, 1.70, 5.96, 1.30, "Instagram", ig.get("recommendation","N/A"), dark_bg=dark)
    card(s, 6.83, 1.70, 5.96, 1.30, "YouTube",   yt.get("recommendation","N/A"), dark_bg=dark)
    card(s, 0.55, 3.30, 5.96, 1.30, "Facebook",  fb.get("recommendation","N/A"), dark_bg=dark)
    card(s, 6.83, 3.30, 5.96, 1.30, "LinkedIn",  li.get("recommendation","N/A"), dark_bg=dark)
    card(s, 0.55, 4.95, 12.23, 1.15, "Top Priority", cp.get("overall_recommendation","N/A"), dark_bg=dark)
    tb(s, f"Report for {brand} ({website_url.replace('https://','').replace('http://','')})  ·  SearchAPI.io + Apify + Claude AI",
       0.55, 6.35, 12.23, 0.30, size=10, color=TEXT_FOOTER, font=FONT_MONO)
    footer_bar(s, 26, dark_bg=dark)

    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    safe_name = re.sub(r'[^A-Za-z0-9_]', '_', analysis.get("brand_name", website_url).replace(" ","_"))
    path = os.path.join(output_dir, f"{safe_name}_social_report.pptx")
    prs.save(path)
    print(f"✅ PPT saved: {path}")
    return path, safe_name


# ═══════════════════════════════════════════════════════════════
# ASYNC JOB SYSTEM
# ═══════════════════════════════════════════════════════════════
#
# WHY THIS EXISTS: the old /generate route did all the work — handle
# discovery, 4 platform fetches, Claude analysis, PPT building — inside
# one single HTTP request. As the pipeline grew (especially the new
# shares/reposts enrichment step), that request started taking long
# enough that the browser or Railway would give up waiting and show a
# generic "upstream error", EVEN THOUGH THE SERVER KEPT WORKING AND
# FINISHED THE REPORT ANYWAY IN THE BACKGROUND. Proof of this: the logs
# showed "PPT saved" and a 200 response for a request the user had
# already seen fail.
#
# THE FIX: /generate now only does one thing — kick off the real work
# on a background thread and return a job_id immediately (in well under
# a second, every time). The page then polls /status/<job_id> every
# couple of seconds until the job is done. No single request is ever
# open long enough to time out, no matter how slow a report gets.
#
# IMPORTANT DEPLOYMENT NOTE: job progress is stored in memory (JOBS
# dict below), not a database. This is intentional — simple and
# sufficient for a low-traffic internal tool — but it ONLY works
# correctly if the app runs as a SINGLE process/worker. If your
# Procfile or Railway start command runs gunicorn with more than one
# worker (e.g. "--workers 2"), a job started on worker A may not be
# visible when a later /status request lands on worker B, and polling
# would incorrectly show "not found". Check your Procfile — it should
# specify a single worker, e.g.:
#     web: gunicorn app:app --workers 1 --timeout 300
# A single worker is completely fine for one agency's internal usage.

JOBS = {}
JOBS_LOCK = threading.Lock()
JOB_MAX_AGE_SECONDS = 2 * 60 * 60  # auto-forget jobs older than 2 hours

def _cleanup_old_jobs():
    cutoff = time.time() - JOB_MAX_AGE_SECONDS
    stale = [jid for jid, j in JOBS.items() if j.get("created_at", 0) < cutoff]
    for jid in stale:
        del JOBS[jid]

def _job_update(job_id, **fields):
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(fields)

def _job_add_step(job_id, label):
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id]["steps"].append(label)

def run_generate_job(job_id, website_url):
    """Runs the full report pipeline on a background thread. Never touches
    the HTTP request/response cycle directly — only writes progress into
    the JOBS dict, which /status/<job_id> reads."""
    try:
        _job_update(job_id, status="running")

        _job_add_step(job_id, "Scanning website for social handles")
        handles = discover_all_handles(website_url)

        ig_raw = {}
        if handles.get("instagram"):
            _job_add_step(job_id, "Fetching Instagram profile data")
            ig_raw = fetch_instagram(handles.get("instagram"))

        fb_raw = {}
        if handles.get("facebook"):
            _job_add_step(job_id, "Fetching Facebook page data")
            fb_raw = fetch_facebook(handles.get("facebook"))

        yt_raw = {}
        if handles.get("youtube"):
            _job_add_step(job_id, "Fetching YouTube channel data")
            yt_raw = fetch_youtube(handles.get("youtube"))

        domain_for_brand = website_url.replace("https://","").replace("http://","").rstrip("/")
        brand_for_search = re.sub(r'\.(com|in|io|co|net|org|app).*','', domain_for_brand).replace("www.","")
        li_raw = {}
        if handles.get("linkedin"):
            _job_add_step(job_id, "Fetching LinkedIn data")
            li_raw = fetch_linkedin(handles.get("linkedin"), brand_for_search)

        _job_add_step(job_id, "Analysing with Claude AI")
        analysis = analyse_with_claude(website_url, handles, ig_raw, fb_raw, yt_raw, li_raw)

        _job_add_step(job_id, f"Building {TOTAL_SLIDES}-slide PowerPoint")
        ppt_path, safe_name = create_ppt(analysis, handles, ig_raw, fb_raw, yt_raw, li_raw, website_url)

        _job_update(job_id, status="done", result={
            "success":      True,
            "brand":        analysis.get("brand_name",""),
            "handles":      handles,
            "analysis":     analysis,
            "download_url": f"/download/{safe_name}"
        })

    except ValueError as e:
        _job_update(job_id, status="error", error=str(e), error_code=404)
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        _job_update(job_id, status="error", error=str(e), error_code=500)


# ═══════════════════════════════════════════════════════════════
# FLASK ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    body        = request.get_json()
    website_url = (body or {}).get("website_url","").strip()
    if not website_url:
        return jsonify({"error":"Please enter a website URL"}), 400

    with JOBS_LOCK:
        _cleanup_old_jobs()
        job_id = str(uuid.uuid4())
        JOBS[job_id] = {
            "status": "starting",
            "steps": [],
            "result": None,
            "error": None,
            "created_at": time.time(),
        }

    thread = threading.Thread(target=run_generate_job, args=(job_id, website_url), daemon=True)
    thread.start()

    # Returns almost instantly — this request never does the slow work,
    # so it can never be the thing that times out.
    return jsonify({"job_id": job_id}), 202

@app.route("/status/<job_id>")
def job_status(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Job not found (it may have expired, or the server restarted)"}), 404
        snapshot = dict(job)  # copy fields out before releasing the lock

    if snapshot["status"] == "error":
        return jsonify({
            "status": "error",
            "steps":  snapshot["steps"],
            "error":  snapshot["error"],
        }), snapshot.get("error_code", 500)

    if snapshot["status"] == "done":
        return jsonify({
            "status": "done",
            "steps":  snapshot["steps"],
            **snapshot["result"],
        })

    return jsonify({"status": snapshot["status"], "steps": snapshot["steps"]})

@app.route("/download/<safe_name>")
def download(safe_name):
    path = f"output/{safe_name}_social_report.pptx"
    if not os.path.exists(path):
        return jsonify({"error":"Report not found"}), 404
    return send_file(path, as_attachment=True,
                     download_name=f"{safe_name}_social_report.pptx")

@app.route("/debug_repost")
def debug_repost():
    """Quick test of the analytics actor for one post URL — returns the
    shares/reposts/saves it found, for verifying the enrichment source."""
    post_url = request.args.get("url", "")
    if not post_url:
        return jsonify({"error": "Pass a post URL like ?url=https://www.instagram.com/p/XXXX/"})
    results = fetch_post_metrics_via_analytics_actor([post_url])
    return jsonify({
        "requested_url": post_url,
        "metrics_found": results.get(post_url) or results,
    })

@app.route("/debug_linkedin/<company_handle>")
def debug_linkedin(company_handle):
    """Run this once against a real LinkedIn company page to check the
    raw Apify output — field names in fetch_linkedin_posts_via_apify /
    _classify_linkedin_post are best-effort guesses and may need
    adjusting once you see real data, same process used for Facebook."""
    try:
        posts = fetch_linkedin_posts_via_apify(company_handle)
        return jsonify({
            "total_posts_returned": len(posts),
            "first_post_full":      posts[0] if posts else "no posts found",
            "second_post_full":     posts[1] if len(posts) > 1 else "n/a",
        })
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/debug_facebook/<username>")
def debug_facebook(username):
    """Run this once against a real Facebook Page to check the raw Apify
    output — the field names in fetch_facebook_posts_via_apify /
    _classify_facebook_post are best-effort guesses and may need
    adjusting once you see real data, same process used to verify
    Instagram's fields."""
    try:
        posts = fetch_facebook_posts_via_apify(username)
        return jsonify({
            "total_posts_returned": len(posts),
            "first_post_full":      posts[0] if posts else "no posts found",
            "second_post_full":     posts[1] if len(posts) > 1 else "n/a",
        })
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route("/debug_instagram/<username>")
def debug_instagram(username):
    try:
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine": "instagram_profile", "username": username,
                    "api_key": SEARCHAPI_KEY}
        )
        data  = r.json()
        posts = data.get("posts", [])
        return jsonify({
            "status_code":          r.status_code,
            "total_posts_returned": len(posts),
            "search_metadata":      data.get("search_metadata", {}),
            "profile_keys":         list(data.get("profile", {}).keys()),
            "first_post_full":      posts[0] if posts else "no posts found",
            "second_post_full":     posts[1] if len(posts) > 1 else "n/a",
        })
    except Exception as e:
        return jsonify({"error": str(e)})

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    # threaded=True is required here: without it, Flask's built-in server
    # handles one request at a time, which would block /status polling
    # while a report is being generated in the background.
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)