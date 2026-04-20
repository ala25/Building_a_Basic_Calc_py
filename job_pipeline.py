"""
Multi-Agent Job Search Pipeline for Andres Altamirano
Agents: 1-Job Scraper → 2-Link Validator → 3-Resume Scorer → 4-Hiring Manager Researcher
Output: job_pipeline_results.xlsx
"""

import os
import re
import json
import time
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from urllib.parse import unquote
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
import anthropic
from dotenv import load_dotenv

# Walk up from script dir to find .env
_check = Path(__file__).parent
for _ in range(6):
    candidate = _check / ".env"
    if candidate.exists():
        load_dotenv(candidate, override=True)
        # Also manually parse in case of non-standard formatting
        for line in candidate.read_text().splitlines():
            line = line.strip()
            if "=" in line and not line.startswith("#"):
                k, _, v = line.partition("=")
                k, v = k.strip(), v.strip().strip('"').strip("'")
                if k and v and k not in os.environ:
                    os.environ[k] = v
        break
    _check = _check.parent

# ─────────────────────────────────────────────
# Data Model
# ─────────────────────────────────────────────

@dataclass
class JobRecord:
    job_title: str = ""
    company: str = ""
    location_type: str = "Unknown"   # Remote | NYC Hybrid | US-Based | Unknown
    job_link: str = ""
    match_score: float = -1.0        # 0–100; -1 = unscored
    date_posted: str = ""
    hiring_manager: str = ""
    link_valid: bool = False
    fallback_used: str = ""
    raw_description: str = ""
    error_log: list = field(default_factory=list)


# ─────────────────────────────────────────────
# Resume Content (parsed from resume.docx)
# ─────────────────────────────────────────────

RESUME_SUMMARY = """
Candidate: Andres Altamirano
Location: New York, NY
Education: B.S. Computer Science, New Jersey Institute of Technology (NJIT); Minor: Business Administration

EXPERIENCE:

1. News Corp, New York — Associate Product Manager, AI Incubation Team (May 2023 – Present)
   - Owned delivery of a multi-agent automation platform: phased milestones, user stories/acceptance criteria, execution; improved ticket completion and PR merges by 20%
   - Established intake and prioritization for requests and defects
   - Partnered with Legal Operations, Sourcing, and Cybersecurity to run compliance and risk assessments for AI vendor integrations
   - Translated constraints into technical requirements and rollout gates
   - Ran sprint planning, backlog refinement, and demos; delivery visibility across scope, dependencies, timelines

2. The Walt Disney Company, New York — Technical Program Manager (July 2018 – May 2023)
   - Managed global software delivery: planning, execution, QA coordination, release tracking
   - Wrote user stories/acceptance criteria; led defect triage
   - Coordinated offshore and on-site engineering teams; delivered automated video-overlay system
   - Partnered with HR and Legal to launch a centralized documentation hub

3. Viacom, New York — Technical Program Manager (April 2017 – July 2018)
   - Led rollout of developer tooling improvements; reduced recurring tickets by 20%, increased engineering velocity
   - Used A/B testing results to guide prioritization; increased user sign-ins by 10%, video views by 3%

4. Viacom, New York — Software Engineer (December 2015 – April 2017)
   - Led migration of key site components to React.js
   - Launched MTV's first tvOS app for the VMAs; 5% increase in digital video views

SKILLS:
- Program Execution (SDLC)
- Agile Delivery (Scrum)
- Intake/Triage & Prioritization
- Cross-Functional Leadership
- Rollout & Support Operations
- AI vendor integrations
- Multi-agent automation platforms
- User stories / acceptance criteria
- Sprint planning / backlog refinement
- Compliance and risk assessments
- React.js
- Product operations
- Technical requirements writing

INDUSTRIES: Media & Entertainment, AI/ML Tools, Enterprise SaaS, News/Publishing
"""

SCORING_SYSTEM_PROMPT = """You are an expert technical recruiter. Given a candidate resume and a job description, score the match 0-100.

Scoring weights:
- Job title alignment (25%): how closely target title matches candidate's experience
- Required skills overlap (35%): % of required/preferred skills candidate has
- Industry/domain fit (20%): media, tech, SaaS, AI relevance
- Tools/platform overlap (20%): specific tools, ATS, AI platforms

Return ONLY valid JSON, no extra text:
{
  "score": <integer 0-100>,
  "title_match": "<High|Medium|Low>",
  "skills_matched": ["skill1", "skill2"],
  "skills_missing": ["skill3"],
  "reasoning": "<2-3 sentence summary>"
}"""


# ─────────────────────────────────────────────
# Utilities
# ─────────────────────────────────────────────

JOB_TITLES = [
    "AI Product Manager",
    "Technical Product Manager",
    "Platform AI Implementation Specialist",
    "Solutions Engineer AI SaaS",
    "AI Tools Program Manager",
    "Product Operations Manager",
    "Automation Specialist AI",
]

SEARCH_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}


def safe_ddgs_search(query: str, max_results: int = 5, timelimit: str = None, retries: int = 3) -> list:
    """Search DuckDuckGo HTML endpoint and return list of {title, href, body} dicts."""
    for attempt in range(retries):
        try:
            params = {"q": query, "kl": "us-en"}
            if timelimit == "m":
                params["df"] = "m"
            resp = requests.get(
                "https://html.duckduckgo.com/html/",
                params=params,
                headers=SEARCH_HEADERS,
                timeout=20,
            )
            if resp.status_code == 202:
                raise requests.exceptions.Timeout("rate limited (202)")
            if resp.status_code != 200:
                raise Exception(f"HTTP {resp.status_code}")

            soup = BeautifulSoup(resp.text, "lxml")
            results = []
            seen_urls = set()

            for div in soup.find_all("div", class_=lambda c: c and "result" in c and "results_links" in c):
                title_tag = div.find("a", class_="result__a")
                snippet_tag = div.find("a", class_="result__snippet")
                if not title_tag:
                    continue
                raw_href = title_tag.get("href", "")
                # Decode DDG redirect: //duckduckgo.com/l/?uddg=<encoded_url>
                if "uddg=" in raw_href:
                    uddg_idx = raw_href.find("uddg=") + 5
                    real_url = unquote(raw_href[uddg_idx:].split("&")[0])
                elif raw_href.startswith("http"):
                    real_url = raw_href
                else:
                    continue

                # Skip DDG ad redirects
                if "duckduckgo.com/y.js" in real_url or "bing.com" in real_url:
                    continue
                if real_url in seen_urls:
                    continue
                seen_urls.add(real_url)

                results.append({
                    "title": title_tag.get_text(strip=True),
                    "href": real_url,
                    "body": snippet_tag.get_text(strip=True) if snippet_tag else "",
                })
                if len(results) >= max_results:
                    break

            return results
        except requests.exceptions.Timeout:
            if attempt < retries - 1:
                wait = 3 if attempt == 0 else 5
                time.sleep(wait)
            else:
                print(f"    [Search] Timeout exhausted: {query[:60]}")
                return []
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(3)
            else:
                print(f"    [Search] Error on '{query[:50]}': {e}")
                return []
    return []


def classify_location(text: str) -> str:
    t = text.lower()
    remote_signals = ["remote", "work from home", "wfh", "fully remote", "100% remote", "anywhere"]
    nyc_signals = ["new york", "nyc", "manhattan", "brooklyn", "hybrid"]
    us_signals = ["united states", "us-based", "usa", "anywhere in the us", "nationwide", "all states"]
    if any(s in t for s in remote_signals):
        return "Remote"
    if any(s in t for s in nyc_signals):
        return "NYC Hybrid"
    if any(s in t for s in us_signals):
        return "US-Based"
    return "Unknown"


def extract_date_posted(text: str) -> str:
    patterns = [
        r"(\d+)\s*days?\s*ago",
        r"(\d+)\s*hours?\s*ago",
        r"(\d+)\s*weeks?\s*ago",
        r"(\d+)\s*months?\s*ago",
        r"posted\s+(\w+\s+\d+)",
        r"(today|yesterday|just posted)",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.I)
        if m:
            return m.group(0)
    return ""


def clean_url(url: str) -> str:
    return url.split("?")[0].rstrip("/").lower()


def extract_domain(url: str) -> str:
    m = re.search(r"https?://(?:www\.)?([^/]+)", url, re.I)
    return m.group(1).lower() if m else ""


def parse_title_and_company(result_title: str, result_body: str, url: str) -> tuple:
    title = result_title
    company = ""

    # Strip common suffixes
    for suffix in [" - LinkedIn", " | LinkedIn", " - Indeed", " | Indeed",
                   " - Greenhouse", " | Greenhouse", " - Lever", " | Lever",
                   " - Glassdoor", " | Glassdoor", " - ZipRecruiter"]:
        title = title.replace(suffix, "")

    # Try extracting company from "Title at Company" or "Title - Company"
    at_match = re.search(r" at (.+?)(?:\s*[-|·]|$)", title, re.I)
    dash_match = re.search(r" [-–] (.+?)(?:\s*[-|·]|$)", title)
    if at_match:
        company = at_match.group(1).strip()
        title = title[:at_match.start()].strip()
    elif dash_match:
        company = dash_match.group(1).strip()
        title = title[:dash_match.start()].strip()

    # Fallback: try to get company from body snippet
    if not company:
        body_company = re.search(r"(?:at|@)\s+([A-Z][A-Za-z0-9\s&,\.]{2,30})", result_body)
        if body_company:
            company = body_company.group(1).strip()

    # Fallback: extract from URL domain
    if not company:
        domain = extract_domain(url)
        if domain and "linkedin" not in domain and "indeed" not in domain:
            company = domain.split(".")[0].title()

    title = title.strip(" -–|·").strip()
    company = company.strip(" -–|·").strip()
    return title, company


# ─────────────────────────────────────────────
# Agent 1 — Job Scraper
# ─────────────────────────────────────────────

class Agent1_JobScraper:

    SEARCH_TEMPLATES = [
        'site:linkedin.com/jobs/view "{title}" Remote',
        'site:job-boards.greenhouse.io "{title}" Remote',
        'site:jobs.lever.co "{title}" Remote',
        '"{title}" job remote apply 2025 -glassdoor.com -ziprecruiter.com -indeed.com',
        '"{title}" "New York" job apply 2025 -glassdoor.com -ziprecruiter.com',
    ]

    def run(self) -> list:
        # If pre-collected jobs JSON exists, load from it instead of searching
        json_path = Path(__file__).parent / "jobs_raw.json"
        if json_path.exists():
            print(f"[AGENT 1] Loading pre-collected jobs from {json_path.name}...")
            return self._load_from_json(json_path)

        print("[AGENT 1] Job Scraper — starting search across LinkedIn, Indeed, Greenhouse, Lever...")
        all_records = []
        remote_count = 0

        for title in JOB_TITLES:
            print(f"  Searching: {title}")
            for template in self.SEARCH_TEMPLATES:
                query = template.replace("{title}", title)
                results = safe_ddgs_search(query, max_results=5, timelimit="m")
                time.sleep(3)

                for r in results:
                    job = self._parse_result(r, title)
                    if job:
                        all_records.append(job)
                        if job.location_type == "Remote":
                            remote_count += 1

            if remote_count >= 10:
                print(f"  [AGENT 1] 10+ remote jobs found — stopping early search")
                break

        deduped = self._deduplicate(all_records)
        print(f"[AGENT 1] Found {len(deduped)} unique jobs after deduplication (from {len(all_records)} raw results)")
        return deduped

    def _load_from_json(self, json_path: Path) -> list:
        with open(json_path) as f:
            raw = json.load(f)
        records = []
        for item in raw:
            job = JobRecord(
                job_title=item.get("job_title", ""),
                company=item.get("company", ""),
                location_type=item.get("location_type", "Unknown"),
                job_link=item.get("job_link", ""),
                date_posted=item.get("date_posted", ""),
            )
            if job.job_link and job.job_title:
                records.append(job)
        deduped = self._deduplicate(records)
        print(f"[AGENT 1] Loaded {len(deduped)} unique jobs from JSON (from {len(raw)} raw entries)")
        return deduped

    def _parse_result(self, r: dict, search_title: str) -> JobRecord:
        try:
            url = r.get("href", "")
            result_title = r.get("title", "")
            body = r.get("body", "")

            if not url or not result_title:
                return None

            # Skip aggregator search/listing pages (not individual job pages)
            skip_patterns = [
                r"linkedin\.com/jobs/search",
                r"indeed\.com/jobs\b",
                r"indeed\.com/\?",
                r"indeed\.com/?$",
                r"glassdoor\.com/Job/jobs",
                r"ziprecruiter\.com/jobs",
            ]
            for pat in skip_patterns:
                if re.search(pat, url, re.I):
                    return None

            title, company = parse_title_and_company(result_title, body, url)
            if not title:
                title = search_title

            location_text = result_title + " " + body
            location_type = classify_location(location_text)
            date_posted = extract_date_posted(body)

            return JobRecord(
                job_title=title,
                company=company,
                location_type=location_type,
                job_link=url,
                date_posted=date_posted,
            )
        except Exception as e:
            return None

    def _deduplicate(self, records: list) -> list:
        seen_urls = set()
        seen_pairs = set()
        result = []
        for r in records:
            url_key = clean_url(r.job_link)
            pair_key = (r.company.lower().strip(), r.job_title.lower().strip())
            if url_key and url_key not in seen_urls and pair_key not in seen_pairs:
                seen_urls.add(url_key)
                seen_pairs.add(pair_key)
                result.append(r)

        # Sort: Remote first, then NYC Hybrid, then US-Based
        priority = {"Remote": 0, "NYC Hybrid": 1, "US-Based": 2, "Unknown": 3}
        result.sort(key=lambda j: priority.get(j.location_type, 3))
        return result


# ─────────────────────────────────────────────
# Agent 2 — Link Validator
# ─────────────────────────────────────────────

GENERIC_URL_PATTERNS = [
    r"linkedin\.com/jobs/?$",
    r"linkedin\.com/jobs/search",
    r"indeed\.com/?$",
    r"indeed\.com/jobs/?$",
    r"greenhouse\.io/?$",
    r"lever\.co/?$",
    r"/404",
    r"/not-found",
    r"/job-not-found",
    r"jobs/search",
    r"/careers/?$",
    r"error",
]

STALE_PAGE_SIGNALS = [
    "this job is no longer available",
    "job has expired",
    "position has been filled",
    "job not found",
    "page not found",
    "posting has expired",
    "no longer accepting",
    "job listing not found",
]


class Agent2_LinkValidator:

    def run(self, jobs: list) -> list:
        print(f"[AGENT 2] Link Validator — checking {len(jobs)} URLs...")
        for i, job in enumerate(jobs):
            try:
                valid, reason, final_url = self._validate_url(job.job_link)
                if valid:
                    job.link_valid = True
                    if final_url and final_url != job.job_link:
                        job.job_link = final_url
                else:
                    job.link_valid = False
                    job.fallback_used = "unverified"
            except Exception as e:
                job.error_log.append(f"Agent2 error: {e}")
                job.link_valid = False
                job.job_link = "Unverified"
            time.sleep(0.1)

        valid_count = sum(1 for j in jobs if j.link_valid)
        print(f"[AGENT 2] {valid_count}/{len(jobs)} links valid")
        return jobs

    def _validate_url(self, url: str) -> tuple:
        try:
            resp = requests.get(url, timeout=12, allow_redirects=True,
                                headers=SEARCH_HEADERS)
            if resp.status_code != 200:
                return False, f"HTTP {resp.status_code}", None
            if self._is_generic_page(url, resp.url, resp.text):
                return False, "generic_or_stale_page", resp.url
            return True, "", resp.url
        except requests.exceptions.Timeout:
            return False, "timeout", None
        except requests.exceptions.ConnectionError:
            return False, "connection_error", None
        except Exception as e:
            return False, str(e)[:80], None

    def _is_generic_page(self, original_url: str, final_url: str, html: str) -> bool:
        for pat in GENERIC_URL_PATTERNS:
            if re.search(pat, final_url, re.I):
                return True
        lowered = html.lower()
        return any(s in lowered for s in STALE_PAGE_SIGNALS)

    def _fallback_careers_page(self, job: JobRecord) -> tuple:
        domain = extract_domain(job.job_link)
        if not domain:
            return False, ""
        # Strip ATS platform domains
        for ats in ["greenhouse.io", "lever.co", "linkedin.com", "indeed.com"]:
            if ats in domain:
                return False, ""
        candidates = [
            f"https://{domain}/careers",
            f"https://{domain}/jobs",
            f"https://careers.{domain}",
        ]
        for url in candidates:
            try:
                resp = requests.get(url, timeout=10, headers=SEARCH_HEADERS)
                if resp.status_code == 200:
                    soup = BeautifulSoup(resp.text, "lxml")
                    links = soup.find_all("a", href=True)
                    title_lower = job.job_title.lower()
                    for link in links:
                        text = (link.get_text() or "").lower()
                        href = link["href"]
                        if title_lower[:10] in text and href.startswith("http"):
                            return True, href
                time.sleep(0.3)
            except Exception:
                continue
        return False, ""

    def _fallback_web_search(self, job: JobRecord) -> tuple:
        query = f'"{job.company}" "{job.job_title}" apply job'
        results = safe_ddgs_search(query, max_results=3)
        company_lower = job.company.lower().replace(" ", "")
        for r in results:
            href = r.get("href", "")
            href_domain = extract_domain(href)
            if company_lower[:6] in href_domain.replace(".", "") and href:
                return True, href
        return False, ""


# ─────────────────────────────────────────────
# Agent 3 — Resume Match Scorer
# ─────────────────────────────────────────────

class Agent3_ResumeScorer:

    JD_SELECTORS = [
        "div.job-description",
        "div.jobDescriptionContent",
        "div#job-content",
        "div.content",
        "div.job__description",
        "section.description",
        "div[class*='description']",
        "div[class*='job-detail']",
        "div[class*='posting']",
        "article",
        "main",
    ]

    def __init__(self):
        api_key = os.environ.get("ANTHROPIC_API_KEY")
        if not api_key:
            print("[AGENT 3] WARNING: ANTHROPIC_API_KEY not set — scoring will be skipped")
            self.client = None
        else:
            self.client = anthropic.Anthropic(api_key=api_key)

    def run(self, jobs: list) -> list:
        valid_jobs = [j for j in jobs if j.link_valid]
        print(f"[AGENT 3] Resume Scorer — scoring {len(valid_jobs)} valid jobs against resume...")

        for i, job in enumerate(valid_jobs):
            print(f"  [{i+1}/{len(valid_jobs)}] Scoring: {job.job_title} @ {job.company}")
            try:
                jd_text = self._fetch_jd(job.job_link)
                if jd_text:
                    job.raw_description = jd_text
                    score, reasoning = self._score(job)
                    job.match_score = score
                    if reasoning:
                        job.error_log.append(f"Scoring reasoning: {reasoning}")
                else:
                    job.error_log.append("JD extraction failed — partial score from title/company only")
                    score, reasoning = self._score_partial(job)
                    job.match_score = score
            except Exception as e:
                job.error_log.append(f"Agent3 error: {e}")
            time.sleep(0.5)

        # Include invalid-link jobs in output with score -1
        all_jobs = jobs  # already contains all
        high_match = sum(1 for j in all_jobs if j.match_score >= 75.0)
        print(f"[AGENT 3] Scoring complete. {high_match} jobs at 75%+ match")
        return all_jobs

    def _fetch_jd(self, url: str) -> str:
        try:
            resp = requests.get(url, timeout=15, headers=SEARCH_HEADERS)
            if resp.status_code != 200:
                return ""
            soup = BeautifulSoup(resp.text, "lxml")
            for tag in soup(["script", "style", "nav", "header", "footer", "aside", "form"]):
                tag.decompose()
            for sel in self.JD_SELECTORS:
                el = soup.select_one(sel)
                if el:
                    text = el.get_text(" ", strip=True)
                    if len(text) > 200:
                        return text[:6000]
            # Fallback: largest text block
            text = soup.get_text(" ", strip=True)
            return text[:6000] if len(text) > 200 else ""
        except Exception as e:
            return ""

    def _score(self, job: JobRecord) -> tuple:
        if not self.client:
            return self._heuristic_score(job), "API not available"
        prompt = (
            f"CANDIDATE RESUME SUMMARY:\n{RESUME_SUMMARY}\n\n"
            f"JOB TITLE: {job.job_title}\n"
            f"COMPANY: {job.company}\n\n"
            f"JOB DESCRIPTION:\n{job.raw_description}\n\n"
            "Score the match. Return JSON only."
        )
        for attempt in range(4):
            try:
                msg = self.client.messages.create(
                    model="claude-haiku-4-5",
                    max_tokens=512,
                    system=SCORING_SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": prompt}],
                )
                return self._parse_score(msg.content[0].text)
            except anthropic.InternalServerError:
                wait = 15 * (attempt + 1)
                print(f"    [API] 500 error — retrying in {wait}s (attempt {attempt+1}/4)...")
                time.sleep(wait)
            except anthropic.RateLimitError:
                print(f"    [API] Rate limited — sleeping 30s...")
                time.sleep(30)
            except anthropic.APIError as e:
                job.error_log.append(f"Claude API error: {e}")
                return self._heuristic_score(job), f"API error: {e}"
        job.error_log.append("Claude API: exhausted retries on 500 errors")
        return self._heuristic_score(job), "API 500 exhausted"

    def _score_partial(self, job: JobRecord) -> tuple:
        if not self.client:
            return self._heuristic_score(job), "API not available"
        prompt = (
            f"CANDIDATE RESUME SUMMARY:\n{RESUME_SUMMARY}\n\n"
            f"JOB TITLE: {job.job_title}\n"
            f"COMPANY: {job.company}\n\n"
            "No job description available. Score based on title and company only. Return JSON only."
        )
        for attempt in range(4):
            try:
                msg = self.client.messages.create(
                    model="claude-haiku-4-5",
                    max_tokens=256,
                    system=SCORING_SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": prompt}],
                )
                return self._parse_score(msg.content[0].text)
            except anthropic.InternalServerError:
                time.sleep(15 * (attempt + 1))
            except Exception as e:
                return self._heuristic_score(job), str(e)
        return self._heuristic_score(job), "API 500 exhausted"

    def _parse_score(self, text: str) -> tuple:
        try:
            # Strip markdown code fences if present
            cleaned = re.sub(r"```(?:json)?", "", text).strip().strip("`").strip()
            data = json.loads(cleaned)
            score = float(data.get("score", -1))
            reasoning = data.get("reasoning", "")
            return score, reasoning
        except (json.JSONDecodeError, ValueError):
            m = re.search(r'"score"\s*:\s*(\d+)', text)
            if m:
                return float(m.group(1)), ""
            return -1.0, "parse_error"

    def _heuristic_score(self, job: JobRecord) -> float:
        title_lower = job.job_title.lower()
        keywords = ["ai", "product manager", "technical", "program manager",
                    "automation", "platform", "solutions engineer", "operations"]
        matches = sum(1 for kw in keywords if kw in title_lower)
        return min(50.0 + matches * 8, 75.0)


# ─────────────────────────────────────────────
# Agent 4 — Hiring Manager Researcher
# ─────────────────────────────────────────────

KNOWN_DOMAINS = {
    "news corp": "newscorp.com",
    "disney": "disney.com",
    "viacom": "viacom.com",
    "google": "google.com",
    "microsoft": "microsoft.com",
    "meta": "meta.com",
    "amazon": "amazon.com",
    "apple": "apple.com",
    "salesforce": "salesforce.com",
    "openai": "openai.com",
    "anthropic": "anthropic.com",
    "stripe": "stripe.com",
    "spotify": "spotify.com",
    "airbnb": "airbnb.com",
    "netflix": "netflix.com",
    "linkedin": "linkedin.com",
    "slack": "slack.com",
    "zoom": "zoom.us",
    "hubspot": "hubspot.com",
    "atlassian": "atlassian.com",
    "datadog": "datadoghq.com",
    "snowflake": "snowflake.com",
}

EMAIL_FORMATS = [
    "{first}.{last}@{domain}",
    "{first}{last}@{domain}",
    "{fi}{last}@{domain}",
    "{first}@{domain}",
]


class Agent4_HiringManagerResearcher:

    def run(self, jobs: list) -> list:
        high_match = [j for j in jobs if j.match_score >= 75.0]
        print(f"[AGENT 4] Hiring Manager Researcher — researching {len(high_match)} high-match jobs...")

        for i, job in enumerate(high_match):
            print(f"  [{i+1}/{len(high_match)}] Researching: {job.job_title} @ {job.company}")
            try:
                linkedin_url, post_snippet = self._search_linkedin(job)
                if linkedin_url:
                    name = self._extract_name_from_linkedin_url(linkedin_url)
                    hm = f"{name} — {linkedin_url}"
                    if post_snippet:
                        snippet_clean = post_snippet[:100].replace("\n", " ")
                        hm += f" | Recent: '{snippet_clean}'"
                    job.hiring_manager = hm
                else:
                    email_str = self._guess_email_formats(job.company)
                    if email_str:
                        job.hiring_manager = f"[Email guesses] {email_str}"
            except Exception as e:
                job.error_log.append(f"Agent4 error: {e}")
            time.sleep(2.0)

        hm_found = sum(1 for j in jobs if j.hiring_manager)
        print(f"[AGENT 4] Found hiring manager info for {hm_found} roles")
        return jobs

    def _search_linkedin(self, job: JobRecord) -> tuple:
        queries = [
            f'site:linkedin.com/in "{job.company}" "head of product" OR "director of product" OR "VP product"',
            f'site:linkedin.com/in "{job.company}" "product manager" lead',
        ]
        for query in queries:
            results = safe_ddgs_search(query, max_results=3, retries=1)
            time.sleep(0.5)
            for r in results:
                href = r.get("href", "")
                if "linkedin.com/in/" in href:
                    return href, r.get("body", "")
        return "", ""

    def _extract_name_from_linkedin_url(self, url: str) -> str:
        try:
            slug = url.split("/in/")[-1].split("?")[0].rstrip("/")
            parts = re.split(r"[-_]", slug)
            name_parts = [p.capitalize() for p in parts if not p.isdigit() and len(p) > 1]
            return " ".join(name_parts[:3])
        except Exception:
            return "Unknown"

    def _guess_email_formats(self, company: str) -> str:
        domain = self._company_to_domain(company)
        if not domain:
            return ""
        samples = [
            f.format(first="firstname", last="lastname", fi="f", domain=domain)
            for f in EMAIL_FORMATS[:3]
        ]
        return " | ".join(samples)

    def _company_to_domain(self, company: str) -> str:
        company_lower = company.lower().strip()
        for key, domain in KNOWN_DOMAINS.items():
            if key in company_lower:
                return domain
        # Attempt slug
        slug = re.sub(r"[^a-z0-9]", "", company_lower)
        if slug:
            return f"{slug}.com"
        return ""


# ─────────────────────────────────────────────
# Output Writer
# ─────────────────────────────────────────────

class OutputWriter:

    HEADERS = ["Job Title", "Company", "Location/Type", "Job Link",
               "Match Score", "Posted", "Hiring Manager"]

    GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    HEADER_FILL = PatternFill(start_color="2E5090", end_color="2E5090", fill_type="solid")
    HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
    LINK_FONT = Font(color="0563C1", underline="single")

    COL_WIDTHS = [35, 25, 14, 60, 14, 16, 60]

    def write(self, jobs: list, path: str):
        wb = Workbook()
        ws = wb.active
        ws.title = "Job Pipeline Results"

        # Header row
        ws.append(self.HEADERS)
        for col_idx, cell in enumerate(ws[1], 1):
            cell.fill = self.HEADER_FILL
            cell.font = self.HEADER_FONT
        ws.freeze_panes = "A2"

        # Sort: by match score descending (unscored = -1 → bottom)
        sorted_jobs = sorted(jobs, key=lambda j: j.match_score if j.match_score >= 0 else -999, reverse=True)

        for job in sorted_jobs:
            score_display = f"{job.match_score:.0f}%" if job.match_score >= 0 else "N/A"
            row_data = [
                job.job_title,
                job.company,
                job.location_type,
                job.job_link,
                score_display,
                job.date_posted,
                job.hiring_manager,
            ]
            ws.append(row_data)
            row_num = ws.max_row

            # Green highlight for 75%+
            if job.match_score >= 75.0:
                for cell in ws[row_num]:
                    cell.fill = self.GREEN_FILL

            # Hyperlink on job link cell
            link_cell = ws.cell(row=row_num, column=4)
            if job.job_link and job.job_link.startswith("http"):
                link_cell.hyperlink = job.job_link
                link_cell.font = self.LINK_FONT

        # Column widths
        for col_idx, width in enumerate(self.COL_WIDTHS, 1):
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width

        # Row height for header
        ws.row_dimensions[1].height = 18

        wb.save(path)
        print(f"[OUTPUT] Spreadsheet saved: {path}")


# ─────────────────────────────────────────────
# Main Orchestration
# ─────────────────────────────────────────────

def main():
    output_path = str(Path(__file__).parent / "job_pipeline_results.xlsx")

    print("=" * 65)
    print("  JOB SEARCH PIPELINE — Andres Altamirano")
    print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 65)

    # ── Agent 1 ──────────────────────────────
    agent1 = Agent1_JobScraper()
    raw_jobs = agent1.run()
    print()

    if not raw_jobs:
        print("No jobs found by Agent 1. Exiting.")
        sys.exit(1)

    # ── Agent 2 ──────────────────────────────
    agent2 = Agent2_LinkValidator()
    validated_jobs = agent2.run(raw_jobs)
    valid_count = sum(1 for j in validated_jobs if j.link_valid)
    print()

    # ── Agent 3 ──────────────────────────────
    scores_cache = Path(__file__).parent / "scores_cache.json"
    if scores_cache.exists():
        print("[AGENT 3] Loading scores from cache...")
        cache = json.loads(scores_cache.read_text())
        for job in validated_jobs:
            key = f"{job.company}|{job.job_title}"
            if key in cache:
                job.match_score = cache[key]
        high_match = sum(1 for j in validated_jobs if j.match_score >= 75.0)
        print(f"[AGENT 3] Loaded from cache. {high_match} jobs at 75%+ match")
        scored_jobs = validated_jobs
    else:
        agent3 = Agent3_ResumeScorer()
        scored_jobs = agent3.run(validated_jobs)
        # Save cache
        cache = {f"{j.company}|{j.job_title}": j.match_score for j in scored_jobs if j.match_score >= 0}
        scores_cache.write_text(json.dumps(cache, indent=2))
    high_match = sum(1 for j in scored_jobs if j.match_score >= 75.0)
    print()

    # ── Agent 4 ──────────────────────────────
    agent4 = Agent4_HiringManagerResearcher()
    final_jobs = agent4.run(scored_jobs)
    hm_found = sum(1 for j in final_jobs if j.hiring_manager)
    print()

    # ── Output ───────────────────────────────
    writer = OutputWriter()
    writer.write(final_jobs, output_path)
    print()

    # ── Summary ──────────────────────────────
    print("=" * 65)
    print("  PIPELINE SUMMARY")
    print("=" * 65)
    print(f"  Total jobs found:          {len(raw_jobs)}")
    print(f"  Valid links:               {valid_count}")
    print(f"  75%+ match jobs:           {high_match}")
    print(f"  Hiring managers found:     {hm_found}")
    print(f"  Output file:               {output_path}")
    print("=" * 65)


if __name__ == "__main__":
    main()
