from flask import Flask, request, render_template, jsonify, send_file
import requests
import anthropic
import json
import os
import re
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

app = Flask(__name__)

NAVY   = RGBColor(0x0A, 0x19, 0x3C)
NAVY2  = RGBColor(0x0E, 0x1E, 0x3A)
GOLD   = RGBColor(0xD4, 0xAF, 0x37)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT  = RGBColor(0xC8, 0xD2, 0xE6)
GREEN  = RGBColor(0x2E, 0xCC, 0x71)
PURPLE = RGBColor(0x8E, 0x44, 0xAD)
TEAL   = RGBColor(0x16, 0xA0, 0x85)
ORANGE = RGBColor(0xE6, 0x7E, 0x22)
DARK   = RGBColor(0x05, 0x0C, 0x1A)
BLUE   = RGBColor(0x1A, 0x5A, 0xFF)
RED    = RGBColor(0xFF, 0x3B, 0x30)


# ═══════════════════════════════════════════════════════════════
# STEP 1 — DISCOVER SOCIAL HANDLES
# ═══════════════════════════════════════════════════════════════

def clean_domain(url):
    url = url.strip().rstrip("/")
    if not url.startswith("http"):
        url = "https://" + url
    return url

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

        # ── FIX: Skip profile.php Facebook URLs ──────────────
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
                m = re.search(r'youtube\.com/(?:channel|c|user)/([A-Za-z0-9_.]+)/?', href, re.I)
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

    # Extract all meaningful words from brand for flexible matching
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
                        """Check if ANY brand word appears in text."""
                        text_clean = text.replace("-","").replace("_","").replace(" ","").lower()
                        brand_clean = brand.replace("-","").replace("_","").replace(" ","")
                        # Full brand match
                        if brand_clean in text_clean:
                            return True
                        # Any individual word match (for brands like "talla jewellers")
                        for word in brand_words:
                            if len(word) > 3 and word in text_clean:
                                return True
                        return False

                    # ── Instagram ─────────────────────────────
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
                            # Accept first result if it's top result and domain matches
                            elif idx == 0 and (domain.replace("www.","").split(".")[0] in handle.lower()
                                               or brand.split("-")[0] in handle.lower()):
                                handles["instagram"] = handle
                                print(f"   ✅ Instagram (domain match): @{handle}")
                                found = True; break
                            else:
                                print(f"   ⚠️ Instagram skipped: @{handle}")

                    # ── Facebook ──────────────────────────────
                    elif platform == "facebook" and "facebook.com" in link:
                        # ── FIX: Skip profile.php URLs ────────
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

                    # ── LinkedIn ──────────────────────────────
                    elif platform == "linkedin" and "linkedin.com/company" in link:
                        m = re.search(r'linkedin\.com/company/([A-Za-z0-9_.-]+)/?', link, re.I)
                        if m:
                            handles["linkedin"] = m.group(1)
                            print(f"   ✅ LinkedIn: {m.group(1)}")
                            found = True; break

                    # ── YouTube ───────────────────────────────
                    elif platform == "youtube" and "youtube.com" in link:
                        m = re.search(r'youtube\.com/@([A-Za-z0-9_.]+)/?', link, re.I)
                        if not m:
                            m = re.search(r'youtube\.com/(?:channel|c|user)/([A-Za-z0-9_.]+)/?', link, re.I)
                        if m:
                            handle = m.group(1)
                            if brand_match(handle) or brand_match(title) or brand_match(snippet):
                                handles["youtube"] = "@"+handle
                                print(f"   ✅ YouTube: @{handle}")
                                found = True; break
                            else:
                                print(f"   ⚠️ YouTube skipped: {handle}")

            except Exception as e:
                print(f"   ⚠️ SearchAPI failed for {platform}: {e}")

    return handles

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

    # Don't require all platforms — just need at least one
    if not handles:
        raise ValueError(
            f"Could not find any social media accounts for {website_url}. "
            "Please check the website URL."
        )
    return handles


# ═══════════════════════════════════════════════════════════════
# STEP 2 — FETCH DATA FROM EACH PLATFORM
# ═══════════════════════════════════════════════════════════════

def fetch_instagram(username):
    print(f"📸 Fetching Instagram: @{username}")
    data = {}
    try:
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine":"instagram_profile","username":username,
                    "api_key":SEARCHAPI_KEY}
        )
        if r.status_code == 200:
            ig    = r.json()
            p     = ig.get("profile", {})
            posts = ig.get("posts", [])
            followers    = p.get("followers", "N/A")
            following    = p.get("following", "N/A")
            total_posts  = p.get("posts", "N/A")
            bio          = p.get("biography", "")
            full_name    = p.get("full_name", "")
            is_verified  = p.get("is_verified", False)
            is_business  = p.get("is_business", False)
            avg_likes    = p.get("avg_likes", None)
            avg_comments = p.get("avg_comments", None)
            eng_rate     = p.get("engagement_rate", None)
            vid_c = car_c = img_c = 0
            total_l = total_c = post_count = 0
            for post in posts[:12]:
                pt = post.get("type","").lower()
                if "video" in pt or "reel" in pt: vid_c += 1
                elif "carousel" in pt or "sidecar" in pt: car_c += 1
                else: img_c += 1
                total_l    += post.get("likes",0) or 0
                total_c    += post.get("comments",0) or 0
                post_count += 1
            if post_count > 0:
                if not avg_likes:    avg_likes = round(total_l / post_count)
                if not avg_comments: avg_comments = round(total_c / post_count)
                if not eng_rate:
                    try:
                        fn = int(str(followers).replace(",",""))
                        er = ((total_l + total_c) / post_count) / fn * 100
                        eng_rate = f"{er:.2f}%"
                    except: eng_rate = "N/A"
            if eng_rate and eng_rate != "N/A":
                try:
                    if "%" not in str(eng_rate):
                        eng_rate = f"{float(eng_rate):.2f}%"
                except: pass
            data = {
                "username":     username,
                "full_name":    full_name,
                "followers":    str(followers),
                "following":    str(following),
                "posts":        str(total_posts),
                "bio":          bio,
                "avg_likes":    str(avg_likes or "N/A"),
                "avg_comments": str(avg_comments or "N/A"),
                "eng_rate":     str(eng_rate or "N/A"),
                "is_verified":  "Yes" if is_verified else "No",
                "is_business":  "Yes" if is_business else "No",
                "img_count":    img_c,
                "car_count":    car_c,
                "vid_count":    vid_c,
                "post_count":   post_count,
            }
            print(f"   ✅ {followers} followers | ER: {eng_rate}")
    except Exception as e:
        print(f"   ⚠️ Instagram failed: {e}")
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
        if r.status_code == 200:
            p = r.json().get("page", {})
            data = {
                "name":        p.get("name", username),
                "followers":   str(p.get("followers",{}).get("count","N/A")),
                "following":   str(p.get("following",{}).get("count","N/A")),
                "category":    ", ".join(p.get("about",{}).get("category_formatted",[]) or []),
                "about":       p.get("about",{}).get("description","") or p.get("about",{}).get("general_info",""),
                "address":     p.get("address","") or "",
                "phone":       p.get("phone","") or "",
                "email":       p.get("email","") or "",
                "website":     p.get("website","") or "",
                "rating":      str(p.get("ratings",{}).get("value","N/A")),
                "is_verified": "Yes" if p.get("is_verified") else "No",
            }
            print(f"   ✅ {data['followers']} followers")
    except Exception as e:
        print(f"   ⚠️ Facebook failed: {e}")
    return data

def fetch_youtube(channel_id):
    print(f"▶️  Fetching YouTube: {channel_id}")
    data = {}
    try:
        r = requests.get(
            "https://www.searchapi.io/api/v1/search",
            params={"engine":"youtube_channel","channel_id":channel_id,
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

def fetch_linkedin(company_handle):
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

    queries = [
        f'{company_name} linkedin followers',
        f'"{company_name}" linkedin company followers',
        f'site:linkedin.com/company/{company_handle}',
        f'{company_name} site:linkedin.com followers',
    ]

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
                link    = res.get("link","")
                snippet = res.get("snippet","")
                title   = res.get("title","")
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
                return data

        except Exception as e:
            print(f"   ⚠️ LinkedIn query failed: {e}")

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

INSTAGRAM DATA:
{json.dumps(ig, indent=2)}

FACEBOOK DATA:
{json.dumps(fb, indent=2)}

YOUTUBE DATA:
{json.dumps(yt, indent=2)}

LINKEDIN DATA:
{json.dumps(li, indent=2)}

Return ONLY a JSON object:
{{
  "brand_name": "brand name",
  "niche": "3-5 words describing the brand",
  "overall_summary": "2-3 sentence overview of their social media presence",
  "instagram": {{
    "followers": "exact from data",
    "engagement_rate": "exact from data",
    "avg_likes": "exact from data",
    "avg_comments": "exact from data",
    "posts": "exact from data",
    "content_types": ["type1", "type2", "type3"],
    "bio_summary": "1 sentence",
    "strength": "biggest strength",
    "recommendation": "one actionable tip"
  }},
  "facebook": {{
    "followers": "exact from data",
    "category": "exact from data",
    "rating": "exact from data",
    "about": "1 sentence summary",
    "strength": "biggest strength",
    "recommendation": "one actionable tip"
  }},
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
        max_tokens=2000,
        messages=[{"role":"user","content":prompt}]
    )
    raw = msg.content[0].text.strip()
    try:
        return json.loads(raw)
    except:
        return json.loads(raw.replace("```json","").replace("```","").strip())


# ═══════════════════════════════════════════════════════════════
# STEP 4 — PPT HELPERS
# ═══════════════════════════════════════════════════════════════

def bg(slide, prs, color=None):
    c = color or NAVY
    s = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    s.fill.solid(); s.fill.fore_color.rgb = c; s.line.fill.background()

def tb(slide, text, x, y, w, h, size=14, bold=False,
       color=WHITE, align=PP_ALIGN.LEFT, italic=False):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf  = box.text_frame; tf.word_wrap = True
    p   = tf.paragraphs[0]
    p.text = str(text); p.alignment = align
    p.font.size = Pt(size); p.font.bold = bold
    p.font.color.rgb = color; p.font.italic = italic
    return box

def rect(slide, x, y, w, h, fill, line=None):
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fill
    if line: s.line.color.rgb = line; s.line.width = Pt(1)
    else: s.line.fill.background()

def stat_card(slide, x, y, w, h, label, value, accent):
    rect(slide, x, y, w, h, NAVY2, accent)
    s = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(0.06))
    s.fill.solid(); s.fill.fore_color.rgb = accent; s.line.fill.background()
    tb(slide, str(value), x+0.1, y+0.15, w-0.2, h*0.52,
       size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, label, x+0.08, y+h*0.6, w-0.16, h*0.35,
       size=10, color=LIGHT, align=PP_ALIGN.CENTER)

def slide_title(slide, title, subtitle=None):
    tb(slide, title, 0.5, 0.15, 12.3, 0.65, size=28, bold=True, color=GOLD)
    if subtitle:
        tb(slide, subtitle, 0.5, 0.75, 12.3, 0.35, size=12, color=LIGHT, italic=True)

def add_table(slide, headers, rows, x, y, w, h):
    cols = len(headers)
    tbl  = slide.shapes.add_table(
        len(rows)+1, cols, Inches(x), Inches(y), Inches(w), Inches(h)
    ).table
    cw = Inches(w/cols)
    for i in range(cols): tbl.columns[i].width = cw
    for ci, hdr in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.text = hdr
        cell.fill.solid(); cell.fill.fore_color.rgb = GOLD
        p = cell.text_frame.paragraphs[0]
        p.font.bold = True; p.font.size = Pt(12)
        p.font.color.rgb = DARK; p.alignment = PP_ALIGN.CENTER
    for ri, row in enumerate(rows):
        fc = NAVY2 if ri%2==0 else DARK
        for ci, val in enumerate(row):
            cell = tbl.cell(ri+1, ci)
            cell.text = str(val)
            cell.fill.solid(); cell.fill.fore_color.rgb = fc
            p = cell.text_frame.paragraphs[0]
            p.font.size = Pt(11); p.font.color.rgb = WHITE
            p.alignment = PP_ALIGN.CENTER

def insight_card(slide, x, y, w, h, title, body, accent):
    rect(slide, x, y, w, h, NAVY2, accent)
    tb(slide, title, x+0.15, y+0.1, w-0.3, 0.35, size=12, bold=True, color=accent)
    tb(slide, body,  x+0.15, y+0.45, w-0.3, h-0.55, size=11, color=WHITE)


# ═══════════════════════════════════════════════════════════════
# STEP 5 — BUILD 14-SLIDE PPT
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

    s = prs.slides.add_slide(blank)
    bg(s, prs, DARK)
    circ = s.shapes.add_shape(9, Inches(9.8), Inches(-2), Inches(6), Inches(6))
    circ.fill.solid(); circ.fill.fore_color.rgb = RGBColor(0x12,0x24,0x4A)
    circ.line.fill.background()
    tb(s, "SOCIAL MEDIA INTELLIGENCE REPORT", 0.7, 1.2, 10, 0.5, size=12, bold=True, color=GOLD, italic=True)
    tb(s, brand, 0.7, 1.75, 10, 1.1, size=42, bold=True, color=WHITE)
    tb(s, niche, 0.7, 2.85, 9, 0.45, size=16, color=GOLD, italic=True)
    tb(s, f"🌐  {website_url}", 0.7, 3.4, 9, 0.4, size=13, color=LIGHT)
    platforms_found = []
    if handles.get("instagram"): platforms_found.append(("📸 Instagram", GREEN))
    if handles.get("facebook"):  platforms_found.append(("👥 Facebook",  BLUE))
    if handles.get("youtube"):   platforms_found.append(("▶️  YouTube",   RED))
    if handles.get("linkedin"):  platforms_found.append(("💼 LinkedIn",  TEAL))
    for i, (label, col) in enumerate(platforms_found):
        rect(s, 0.7+i*3.15, 4.05, 2.9, 0.5, NAVY2, col)
        tb(s, label, 0.85+i*3.15, 4.12, 2.6, 0.36, size=13, bold=True, color=WHITE)
    for i, (lbl, val, col) in enumerate([
        ("INSTAGRAM FOLLOWERS", ig.get("followers","N/A"), GREEN),
        ("FACEBOOK FOLLOWERS",  fb.get("followers","N/A"), BLUE),
        ("YOUTUBE SUBSCRIBERS", yt.get("subscribers","N/A"), RED),
    ]):
        stat_card(s, 0.7+i*4.1, 4.8, 3.8, 1.5, lbl, val, col)
    tb(s, f"Powered by SearchAPI.io + Claude AI  •  {website_url}",
       0.7, 7.1, 12, 0.3, size=10, color=LIGHT, italic=True, align=PP_ALIGN.CENTER)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "Social Media Overview", f"{brand}  —  All platforms at a glance")
    add_table(s, ["Platform","Handle","Followers/Subscribers","Key Metric","Status"], [
        ["📸 Instagram", f"@{handles.get('instagram','N/A')}", ig.get("followers","N/A"), f"ER: {ig.get('engagement_rate','N/A')}", "✅ Active" if handles.get("instagram") else "❌ Not Found"],
        ["👥 Facebook",  handles.get("facebook","N/A"),        fb.get("followers","N/A"), f"Rating: {fb.get('rating','N/A')}",      "✅ Active" if handles.get("facebook") else "❌ Not Found"],
        ["▶️ YouTube",   handles.get("youtube","N/A"),          yt.get("subscribers","N/A"), f"Videos: {yt.get('videos','N/A')}",   "✅ Active" if handles.get("youtube") else "❌ Not Found"],
        ["💼 LinkedIn",  handles.get("linkedin","N/A"),         li.get("followers","N/A"), f"Employees: {li.get('employees','N/A')}", "✅ Active" if handles.get("linkedin") else "❌ Not Found"],
    ], 0.5, 1.3, 12.3, 3.5)
    rect(s, 0.5, 5.1, 12.3, 1.8, NAVY2, GOLD)
    tb(s, "OVERALL ANALYSIS", 0.75, 5.2, 5, 0.35, size=11, bold=True, color=GOLD)
    tb(s, analysis.get("overall_summary",""), 0.75, 5.6, 11.8, 1.1, size=13, color=WHITE)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "📸  Instagram Profile", f"@{handles.get('instagram','N/A')}  —  {ig.get('bio_summary','')}")
    for (lbl,val,col),(x,y) in zip([
        ("FOLLOWERS",         ig.get("followers","N/A"),      GREEN),
        ("FOLLOWING",         ig_raw.get("following","N/A"),  TEAL),
        ("TOTAL POSTS",       ig.get("posts","N/A"),          GOLD),
        ("ENGAGEMENT RATE",   ig.get("engagement_rate","N/A"),GREEN),
        ("AVG LIKES/POST",    ig.get("avg_likes","N/A"),      ORANGE),
        ("AVG COMMENTS/POST", ig.get("avg_comments","N/A"),   PURPLE),
    ], [(0.5,1.3),(4.7,1.3),(8.9,1.3),(0.5,4.1),(4.7,4.1),(8.9,4.1)]):
        stat_card(s, x, y, 3.8, 1.9, lbl, val, col)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "📸  Instagram Content Strategy", f"Content mix from recent {ig_raw.get('post_count',12)} posts")
    img_c = ig_raw.get("img_count",0)
    car_c = ig_raw.get("car_count",0)
    vid_c = ig_raw.get("vid_count",0)
    total = max(img_c+car_c+vid_c, 1)
    cd = ChartData()
    cd.categories = ["Images","Carousels","Videos/Reels"]
    cd.add_series("Posts",(img_c,car_c,vid_c))
    chart = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5),Inches(1.3),Inches(7.2),Inches(5.0),cd).chart
    chart.has_legend = False
    for i,col in enumerate([TEAL,GOLD,PURPLE]):
        pt = chart.plots[0].series[0].points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    chart.category_axis.tick_labels.font.size = Pt(12)
    chart.category_axis.tick_labels.font.color.rgb = WHITE
    chart.value_axis.tick_labels.font.size = Pt(11)
    chart.value_axis.tick_labels.font.color.rgb = WHITE
    for i,(lbl,cnt,col) in enumerate([("IMAGES",img_c,TEAL),("CAROUSELS",car_c,GOLD),("VIDEOS/REELS",vid_c,PURPLE)]):
        stat_card(s, 8.1, 1.3+i*1.9, 4.7, 1.65, lbl, f"{cnt} posts ({int(cnt/total*100)}%)", col)
    insight_card(s, 0.5, 6.55, 12.3, 0.75, "💡 Instagram Strength", ig.get("strength",""), GREEN)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "👥  Facebook Page", f"{handles.get('facebook','N/A')}  —  {fb.get('category','')}")
    for (lbl,val,col),(x,y) in zip([
        ("PAGE FOLLOWERS", fb.get("followers","N/A"),      BLUE),
        ("FOLLOWING",      fb_raw.get("following","N/A"),  TEAL),
        ("PAGE RATING",    fb.get("rating","N/A"),         GOLD),
        ("VERIFIED",       fb.get("is_verified","No"),     GREEN),
        ("CATEGORY",       str(fb.get("category","N/A"))[:20], PURPLE),
        ("PLATFORM",       "Facebook",                     BLUE),
    ], [(0.5,1.3),(4.7,1.3),(8.9,1.3),(0.5,4.1),(4.7,4.1),(8.9,4.1)]):
        stat_card(s, x, y, 3.8, 1.9, lbl, val, col)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "👥  Facebook Insights", "Page details and strategic analysis")
    rect(s, 0.5, 1.3, 12.3, 2.2, NAVY2, BLUE)
    tb(s, "PAGE ABOUT", 0.75, 1.4, 4, 0.35, size=11, bold=True, color=BLUE)
    tb(s, str(fb.get("about","N/A")), 0.75, 1.8, 11.8, 1.5, size=13, color=WHITE)
    add_table(s, ["Detail","Value"], [
        ["Address", str(fb_raw.get("address","N/A") or "N/A")],
        ["Phone",   str(fb_raw.get("phone","N/A")   or "N/A")],
        ["Email",   str(fb_raw.get("email","N/A")   or "N/A")],
        ["Website", str(fb_raw.get("website","N/A") or "N/A")],
    ], 0.5, 3.7, 6.0, 2.5)
    insight_card(s, 6.8, 3.7,  5.9, 1.1, "💡 Strength",       fb.get("strength","N/A"), BLUE)
    insight_card(s, 6.8, 5.05, 5.9, 1.1, "🎯 Recommendation", fb.get("recommendation","N/A"), GOLD)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "▶️   YouTube Channel", f"{handles.get('youtube','N/A')}  —  {yt.get('description_summary','')}")
    for (lbl,val,col),(x,y) in zip([
        ("SUBSCRIBERS",  yt.get("subscribers","N/A"), RED),
        ("TOTAL VIDEOS", yt.get("videos","N/A"),      ORANGE),
        ("TOTAL VIEWS",  yt.get("views","N/A"),        GOLD),
        ("VERIFIED",     yt.get("is_verified","No"),   GREEN),
        ("JOINED DATE",  yt.get("joined","N/A"),       TEAL),
        ("PLATFORM",     "YouTube",                    RED),
    ], [(0.5,1.3),(4.7,1.3),(8.9,1.3),(0.5,4.1),(4.7,4.1),(8.9,4.1)]):
        stat_card(s, x, y, 3.8, 1.9, lbl, val, col)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "▶️   YouTube Channel Insights", "Channel performance and strategic analysis")
    rect(s, 0.5, 1.3, 12.3, 2.0, NAVY2, RED)
    tb(s, "CHANNEL DESCRIPTION", 0.75, 1.4, 6, 0.35, size=11, bold=True, color=RED)
    tb(s, str(yt.get("description_summary","") or yt_raw.get("description",""))[:300], 0.75, 1.8, 11.8, 1.3, size=13, color=WHITE)
    insight_card(s, 0.5,  3.55, 6.0, 1.4, "💡 Strength",       yt.get("strength","N/A"), RED)
    insight_card(s, 6.85, 3.55, 5.9, 1.4, "🎯 Recommendation", yt.get("recommendation","N/A"), GOLD)
    add_table(s, ["Metric","Value"], [
        ["Subscribers",  yt.get("subscribers","N/A")],
        ["Total Videos", yt.get("videos","N/A")],
        ["Total Views",  yt.get("views","N/A")],
        ["Channel Age",  yt.get("joined","N/A")],
    ], 0.5, 5.1, 6.0, 2.1)
    rect(s, 6.85, 5.1, 5.9, 2.1, NAVY2, ORANGE)
    tb(s, "VIEWS PER VIDEO (EST.)", 7.1, 5.2, 5.4, 0.35, size=11, bold=True, color=ORANGE)
    try:
        vpv = int(str(yt.get("views","0")).replace(",","")) // max(int(str(yt.get("videos","1")).replace(",","")),1)
        tb(s, f"{vpv:,}", 7.1, 5.65, 5.4, 1.0, size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    except:
        tb(s, "N/A", 7.1, 5.65, 5.4, 1.0, size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "💼  LinkedIn Company Page", f"linkedin.com/company/{handles.get('linkedin','N/A')}")
    for (lbl,val,col),(x,y) in zip([
        ("FOLLOWERS",   li.get("followers","N/A"), TEAL),
        ("EMPLOYEES",   li.get("employees","N/A"), BLUE),
        ("PLATFORM",    "LinkedIn",                TEAL),
        ("DATA SOURCE", "Google Search",           LIGHT),
    ], [(0.5,1.3),(4.7,1.3),(8.9,1.3),(0.5,4.1)]):
        stat_card(s, x, y, 3.8, 1.9, lbl, val, col)
    rect(s, 0.5, 4.0, 12.3, 2.1, NAVY2, TEAL)
    tb(s, "COMPANY SUMMARY", 0.75, 4.1, 5, 0.35, size=11, bold=True, color=TEAL)
    tb(s, li.get("summary","N/A"), 0.75, 4.5, 11.8, 1.4, size=13, color=WHITE)
    insight_card(s, 0.5,  6.3, 6.0, 0.9, "💡 Strength",       li.get("strength","N/A"), TEAL)
    insight_card(s, 6.85, 6.3, 5.9, 0.9, "🎯 Recommendation", li.get("recommendation","N/A"), GOLD)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "💼  LinkedIn Strategic Analysis", "Professional presence and B2B opportunities")
    add_table(s, ["Metric","Value","Insight"], [
        ["Followers",     li.get("followers","N/A"), "Professional audience size"],
        ["Employees",     li.get("employees","N/A"), "Company scale indicator"],
        ["Content Focus", "B2B & Professional",      "Key content strategy"],
        ["Best Post Type","Articles & Updates",       "Highest LinkedIn engagement"],
        ["Posting Freq.", "2-3x per week",            "LinkedIn best practice"],
    ], 0.5, 1.3, 12.3, 4.0)
    insight_card(s, 0.5,  5.55, 6.0, 1.65, "💡 Strength",       li.get("strength","N/A"), TEAL)
    insight_card(s, 6.85, 5.55, 5.9, 1.65, "🎯 Recommendation", li.get("recommendation","N/A"), GOLD)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "📊  Cross-Platform Follower Comparison", "Audience size across all platforms")
    def parse_num(val):
        try:
            v = str(val).replace(",","").upper().replace("+","")
            if "M" in v: return float(v.replace("M","")) * 1_000_000
            if "K" in v: return float(v.replace("K","")) * 1_000
            return float(v)
        except: return 0
    ig_f = parse_num(ig.get("followers",0))
    fb_f = parse_num(fb.get("followers",0))
    yt_f = parse_num(yt.get("subscribers",0))
    li_f = parse_num(li.get("followers",0))
    cd2 = ChartData()
    cd2.categories = ["Instagram","Facebook","YouTube","LinkedIn"]
    cd2.add_series("Followers/Subscribers",(ig_f,fb_f,yt_f,li_f))
    chart2 = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.5),Inches(1.3),Inches(8.0),Inches(5.5),cd2).chart
    chart2.has_legend = False
    for i,col in enumerate([GREEN,BLUE,RED,TEAL]):
        pt = chart2.plots[0].series[0].points[i]
        pt.format.fill.solid(); pt.format.fill.fore_color.rgb = col
    chart2.category_axis.tick_labels.font.size = Pt(12)
    chart2.category_axis.tick_labels.font.color.rgb = WHITE
    chart2.value_axis.tick_labels.font.size = Pt(10)
    chart2.value_axis.tick_labels.font.color.rgb = WHITE
    for i,(lbl,val,col) in enumerate([
        ("INSTAGRAM", ig.get("followers","N/A"),   GREEN),
        ("FACEBOOK",  fb.get("followers","N/A"),   BLUE),
        ("YOUTUBE",   yt.get("subscribers","N/A"), RED),
        ("LINKEDIN",  li.get("followers","N/A"),   TEAL),
    ]):
        stat_card(s, 8.85, 1.3+i*1.55, 4.0, 1.3, lbl, val, col)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "📈  Engagement Benchmarks", "Performance metrics vs industry standards")
    def er_status(val):
        try:
            v = float(str(val).replace("%",""))
            return "🟢 Strong" if v>=3 else ("🟡 Average" if v>=1 else "🔴 Low")
        except: return "📊 Data"
    add_table(s, ["Platform","Key Metric","Value","Benchmark","Status"], [
        ["📸 Instagram","Engagement Rate",   ig.get("engagement_rate","N/A"),">3% is good",        er_status(ig.get("engagement_rate","0"))],
        ["📸 Instagram","Avg Likes/Post",    ig.get("avg_likes","N/A"),      "Varies by niche",    "📊 Data"],
        ["📸 Instagram","Avg Comments/Post", ig.get("avg_comments","N/A"),   ">20 is strong",      "📊 Data"],
        ["👥 Facebook", "Page Followers",    fb.get("followers","N/A"),      "Varies by industry", "📊 Data"],
        ["👥 Facebook", "Page Rating",       fb.get("rating","N/A"),         ">4.0 is good",       "📊 Data"],
        ["▶️ YouTube",  "Subscribers",       yt.get("subscribers","N/A"),    "Varies by niche",    "📊 Data"],
        ["▶️ YouTube",  "Views per Video",   "See slide 8",                  ">1000 is good",      "📊 Data"],
        ["💼 LinkedIn", "Followers",         li.get("followers","N/A"),      "Varies by industry", "📊 Data"],
    ], 0.5, 1.3, 12.3, 5.8)

    s = prs.slides.add_slide(blank)
    bg(s, prs)
    slide_title(s, "💪  Key Strengths & Gaps", "What's working and what needs attention")
    tb(s, "STRONGEST PLATFORM", 0.5, 1.25, 6.2, 0.35, size=11, bold=True, color=GREEN)
    rect(s, 0.5, 1.6, 6.2, 1.3, NAVY2, GREEN)
    tb(s, cp.get("strongest_platform","N/A"), 0.7, 1.7, 5.8, 1.1, size=13, color=WHITE)
    tb(s, "NEEDS MOST WORK", 6.85, 1.25, 5.9, 0.35, size=11, bold=True, color=RED)
    rect(s, 6.85, 1.6, 5.9, 1.3, NAVY2, RED)
    tb(s, cp.get("weakest_platform","N/A"), 7.05, 1.7, 5.5, 1.1, size=13, color=WHITE)
    insight_card(s, 0.5,  3.15, 6.2, 1.4, "🔄 Content Consistency", cp.get("content_consistency","N/A"), TEAL)
    insight_card(s, 6.85, 3.15, 5.9, 1.4, "🚀 Growth Opportunity",  cp.get("growth_opportunity","N/A"),  GOLD)
    rect(s, 0.5, 4.75, 12.3, 1.9, NAVY2, PURPLE)
    tb(s, "OVERALL SUMMARY", 0.75, 4.85, 5, 0.35, size=11, bold=True, color=PURPLE)
    tb(s, analysis.get("overall_summary","N/A"), 0.75, 5.25, 11.8, 1.25, size=13, color=WHITE)

    s = prs.slides.add_slide(blank)
    bg(s, prs, DARK)
    tb(s, "STRATEGIC RECOMMENDATIONS", 0.6, 0.25, 12, 0.5, size=13, bold=True, color=GOLD, italic=True)
    tb(s, "Action Plan", 0.6, 0.7, 12, 0.65, size=32, bold=True, color=WHITE)
    for (title,body,col,x,y) in [
        ("📸  Instagram", ig.get("recommendation","N/A"), GREEN, 0.5, 1.55),
        ("👥  Facebook",  fb.get("recommendation","N/A"), BLUE,  0.5, 3.1),
        ("▶️   YouTube",   yt.get("recommendation","N/A"), RED,   6.9, 1.55),
        ("💼  LinkedIn",  li.get("recommendation","N/A"), TEAL,  6.9, 3.1),
    ]:
        rect(s, x, y, 6.0, 1.3, NAVY2, col)
        tb(s, title, x+0.2, y+0.1,  5.6, 0.38, size=13, bold=True, color=col)
        tb(s, body,  x+0.2, y+0.52, 5.6, 0.68, size=12, color=WHITE)
    rect(s, 0.5, 4.75, 12.3, 1.5, NAVY2, GOLD)
    tb(s, "🎯  TOP PRIORITY", 0.75, 4.85, 5, 0.35, size=12, bold=True, color=GOLD)
    tb(s, cp.get("overall_recommendation","N/A"), 0.75, 5.25, 11.8, 0.85, size=14, color=WHITE)
    tb(s, f"Report for {brand} ({website_url})  •  SearchAPI.io + Claude AI",
       0.5, 7.1, 12.3, 0.3, size=10, color=LIGHT, align=PP_ALIGN.CENTER, italic=True)

    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)
    safe_name = re.sub(r'[^A-Za-z0-9_]', '_', analysis.get("brand_name", website_url).replace(" ","_"))
    path = os.path.join(output_dir, f"{safe_name}_social_report.pptx")
    prs.save(path)
    print(f"✅ PPT saved: {path}")
    return path, safe_name


# ═══════════════════════════════════════════════════════════════
# FLASK ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    try:
        body        = request.get_json()
        website_url = body.get("website_url","").strip()
        if not website_url:
            return jsonify({"error":"Please enter a website URL"}), 400
        handles  = discover_all_handles(website_url)
        ig_raw   = fetch_instagram(handles.get("instagram","")) if handles.get("instagram") else {}
        fb_raw   = fetch_facebook(handles.get("facebook",""))   if handles.get("facebook")  else {}
        yt_raw   = fetch_youtube(handles.get("youtube",""))     if handles.get("youtube")   else {}
        li_raw   = fetch_linkedin(handles.get("linkedin",""))   if handles.get("linkedin")  else {}
        analysis = analyse_with_claude(website_url, handles, ig_raw, fb_raw, yt_raw, li_raw)
        ppt_path, safe_name = create_ppt(analysis, handles, ig_raw, fb_raw, yt_raw, li_raw, website_url)
        return jsonify({
            "success":      True,
            "brand":        analysis.get("brand_name",""),
            "handles":      handles,
            "analysis":     analysis,
            "download_url": f"/download/{safe_name}"
        })
    except ValueError as e:
        return jsonify({"error": str(e)}), 404
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return jsonify({"error": str(e)}), 500

@app.route("/download/<safe_name>")
def download(safe_name):
    path = f"output/{safe_name}_social_report.pptx"
    if not os.path.exists(path):
        return jsonify({"error":"Report not found"}), 404
    return send_file(path, as_attachment=True,
                     download_name=f"{safe_name}_social_report.pptx")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=False)