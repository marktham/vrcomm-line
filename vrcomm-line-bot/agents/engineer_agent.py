"""
agents/engineer_agent.py — VRCOMM Engineer / Pre-Sales Agent

Handles:
  1. technical_qa     — Technical questions about VRCOMM products
  2. tor_analysis     — TOR/SOW compliance: match requirements → products
  3. spec_comparison  — Compare products against each other or a requirement

Data sources (priority order):
  1. specs/ folder    — .txt/.md spec sheets per brand (add files here)
  2. product URLs     — fetch from VRCOMM_ProductList.xlsx URLs
  3. Claude knowledge — fallback for well-known products

Two-Step architecture (same as product_agent — prevents hallucination):
  Step 1: Claude selects relevant brands from VRCOMM list
  Step 2: Claude answers using only those brands + spec content
"""
import os, re, logging, time
import requests
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

_BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PRODUCT_LIST = os.path.join(_BASE_DIR, "product", "VRCOMM_ProductList.xlsx")
_SPECS_DIR    = os.path.join(_BASE_DIR, "specs")

# ── Product list cache ────────────────────────────────────────────────────────

_product_cache      = None
_product_cache_time = 0.0
_CACHE_TTL          = 300


def _load_product_list() -> list:
    """Load VRCOMM_ProductList.xlsx → list of {brand, url}. Cached 5 min."""
    global _product_cache, _product_cache_time

    path = _PRODUCT_LIST
    if not os.path.isfile(path):
        logger.error("ProductList not found: %s", path)
        return []

    mtime = os.path.getmtime(path)
    now   = time.time()
    if (_product_cache is not None
            and mtime == _product_cache_time
            and (now - _product_cache_time) < _CACHE_TTL):
        return _product_cache

    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        products = []
        for row in ws.iter_rows(values_only=True):
            brand = str(row[0]).strip() if row[0] else ""
            url   = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            if brand.lower() in ("brand", "product", "name", "") or not brand:
                continue
            products.append({"brand": brand, "url": url})
        _product_cache      = products
        _product_cache_time = mtime
        logger.info("[engineer_agent] ProductList: %d brands", len(products))
        return products
    except Exception as e:
        logger.error("[engineer_agent] ProductList load error: %s", e)
        return []


# ── Spec file loader ──────────────────────────────────────────────────────────

_spec_cache: dict = {}   # brand_lower → {"content": str, "mtime": float}


def _load_spec_file(brand: str) -> str:
    """
    Load spec file for a brand from specs/ folder.
    Tries: exact brand name, brand with underscores, partial match.
    Returns content string or empty if not found.
    """
    if not os.path.isdir(_SPECS_DIR):
        return ""

    brand_key = brand.lower().replace(" ", "_").replace("-", "_")

    # Try all .txt and .md files in specs/
    try:
        for fname in os.listdir(_SPECS_DIR):
            if not fname.endswith((".txt", ".md")):
                continue
            fname_key = fname.lower().replace("-", "_").rsplit(".", 1)[0]
            if fname_key == brand_key or brand_key in fname_key or fname_key in brand_key:
                fpath = os.path.join(_SPECS_DIR, fname)
                mtime = os.path.getmtime(fpath)
                cached = _spec_cache.get(brand_key)
                if cached and cached["mtime"] == mtime:
                    return cached["content"]
                with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
                    content = f.read()[:4000]
                _spec_cache[brand_key] = {"content": content, "mtime": mtime}
                logger.info("[engineer_agent] Loaded spec file: %s", fname)
                return content
    except Exception as e:
        logger.warning("[engineer_agent] Spec file load error: %s", e)
    return ""


# ── URL fetcher (cached 1 hr) ─────────────────────────────────────────────────

_url_cache: dict = {}
_URL_CACHE_TTL   = 3600


def _fetch_url(url: str) -> str:
    """Fetch URL → clean plain text (HTML stripped). Cached 1 hour."""
    url = url.strip()
    if not url:
        return ""

    now    = time.time()
    cached = _url_cache.get(url)
    if cached and (now - cached["fetched_at"]) < _URL_CACHE_TTL:
        return cached["content"]

    try:
        resp = requests.get(url, headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9,th;q=0.8",
        }, timeout=8, allow_redirects=True)
        resp.raise_for_status()

        html = resp.text
        html = re.sub(r'<(script|style|nav|footer|header)[^>]*>.*?</\1>',
                      ' ', html, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<[^>]+>', ' ', html)
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n{3,}', '\n\n', text).strip()[:2500]

        _url_cache[url] = {"content": text, "fetched_at": now}
        logger.info("[engineer_agent] Fetched %d chars: %s", len(text), url[:60])
        return text
    except Exception as e:
        logger.warning("[engineer_agent] URL fetch failed [%s]: %s", url[:60], e)
        return ""


# ── Mode detector ─────────────────────────────────────────────────────────────

_TOR_KEYWORDS = [
    "tor", "sow", "scope of work", "terms of reference",
    "tdr", "spec compliance", "ตาราง", "comply", "requirement",
    "ข้อกำหนด", "คุณสมบัติ", "เทียบ spec", "compliance",
    "tender", "bidding", "จัดซื้อ", "คุณลักษณะ", "รายการ",
]


def _detect_mode(message: str) -> str:
    """
    Detect whether this is a TOR analysis or general technical Q&A.
    Returns: "tor_analysis" or "technical_qa"
    """
    msg_lower = message.lower()
    if any(kw in msg_lower for kw in _TOR_KEYWORDS):
        return "tor_analysis"
    # Long messages with numbered items are likely TOR/requirements
    if len(message) > 400 and re.search(r'^\s*\d+[\.\)]\s+', message, re.MULTILINE):
        return "tor_analysis"
    return "technical_qa"


# ── STEP 1: Brand selector ────────────────────────────────────────────────────

_SELECT_PROMPT = """You are a technical product selector for VRCOMM, a cybersecurity company in Thailand.

Customer/engineer query:
\"{message}\"

VRCOMM's complete product list (these are the ONLY brands we sell):
{brand_list}

Select which brands from the list are technically relevant to this query.
Consider: product category, use case, technical requirements, complementary products.
Return ONLY exact brand names from the list, one per line.
If none are relevant, return: NONE
Do NOT add brands outside the list. No explanations."""


def _select_relevant_brands(message: str, product_list: list) -> list:
    """Step 1: Claude picks relevant brands from our list only."""
    brand_list_text = "\n".join(
        "%d. %s" % (i + 1, p["brand"]) for i, p in enumerate(product_list)
    )
    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            messages=[{"role": "user", "content": _SELECT_PROMPT.format(
                message=message[:800],
                brand_list=brand_list_text,
            )}],
        )
        raw = resp.content[0].text.strip()
        logger.info("[engineer_agent] Step1 selection: %s", raw[:120])

        if raw.strip().upper() == "NONE" or not raw:
            return []

        selected_names = [
            re.sub(r'^[\d\.\-\)\s]+', '', line).strip()
            for line in raw.splitlines()
            if line.strip() and line.strip().upper() != "NONE"
        ]
        matched = []
        for name in selected_names:
            for p in product_list:
                if (p["brand"].lower() == name.lower()
                        or name.lower() in p["brand"].lower()
                        or p["brand"].lower() in name.lower()):
                    if p not in matched:
                        matched.append(p)
        logger.info("[engineer_agent] Brands selected: %s",
                    [m["brand"] for m in matched])
        return matched
    except Exception as e:
        logger.error("[engineer_agent] Step1 error: %s", e)
        return []


# ── STEP 2: System prompts ────────────────────────────────────────────────────

_TECH_QA_SYSTEM = """You are VRCOMM's Senior Network and Cybersecurity Engineer.
VRCOMM is a solutions provider in Thailand.

The following VRCOMM products are relevant to this query:

{product_specs_section}

Answer the technical question using ONLY the VRCOMM products listed above.
Do NOT mention any brand not listed here (e.g. Fortinet, Cisco, Palo Alto, etc.).

Guidelines:
- Provide accurate technical specifications and comparisons
- Recommend specific models/configurations when possible
- For sizing questions: consider users, throughput, concurrent sessions
- NEVER quote prices — say: "ทีม Sales จะจัดทำใบเสนอราคาให้ครับ"
- Reply in the SAME LANGUAGE as the user (Thai → Thai, English → English)
- Plain text only — no markdown, no bullets, no headers
- Max 5 short paragraphs"""

_TOR_ANALYSIS_SYSTEM = """You are VRCOMM's Senior Pre-Sales Engineer specialising in TOR/SOW compliance.
VRCOMM is a Network and Cybersecurity solutions provider in Thailand.

The following VRCOMM products are available for this proposal:

{product_specs_section}

Analyse the TOR/SOW requirements and produce a compliance table.

Format each requirement as:
  [No.] Requirement | Recommended Product | Compliance | Notes

Compliance status:
  FULL    — product fully meets this requirement
  PARTIAL — product meets with configuration or optional add-on
  ALT     — alternative approach using VRCOMM product

After the table, add a brief summary of the proposed solution.

Rules:
- ONLY use products listed above — never suggest outside brands
- NEVER quote prices
- Reply in the SAME LANGUAGE as the document (Thai → Thai, English → English)
- Plain text only — no markdown"""

_NO_MATCH_SYSTEM = """You are VRCOMM's Senior Network and Cybersecurity Engineer.
VRCOMM is a solutions provider in Thailand.

VRCOMM's complete product portfolio:
{full_list}

The customer's query does not match any specific product in our catalog,
but you should still provide helpful technical guidance using ONLY the products above.

Rules:
- Suggest the most technically suitable products from the list for the requirement
- Explain why each suggested product fits
- NEVER mention brands outside this list
- NEVER quote prices
- Reply in the SAME LANGUAGE as the user
- Plain text only — max 5 paragraphs"""


def _build_product_specs_section(selected: list) -> str:
    """Build spec content for each selected brand."""
    sections = []
    for p in selected:
        # Priority: local spec file > URL fetch > just name+URL
        spec_content = _load_spec_file(p["brand"])
        if not spec_content:
            spec_content = _fetch_url(p["url"])

        section = "=== %s ===\nWebsite: %s" % (p["brand"], p["url"])
        if spec_content:
            section += "\n\n%s" % spec_content[:2000]
        sections.append(section)
    return "\n\n".join(sections) if sections else "(No spec content available)"


def _build_system(mode: str, selected: list, product_list: list) -> str:
    if not selected:
        full_list = "\n".join(
            "%d. %s" % (i + 1, p["brand"]) for i, p in enumerate(product_list)
        )
        return _NO_MATCH_SYSTEM.format(full_list=full_list)

    product_specs_section = _build_product_specs_section(selected)

    if mode == "tor_analysis":
        return _TOR_ANALYSIS_SYSTEM.format(
            product_specs_section=product_specs_section
        )
    else:
        return _TECH_QA_SYSTEM.format(
            product_specs_section=product_specs_section
        )


# ── Main handler ──────────────────────────────────────────────────────────────

def handle(message: str, user_name: str, user_id: str,
           source: str = "line", history: list = None,
           intent: str = "technical", **kwargs) -> str:
    """
    Two-step engineer handler:
      Step 0 — Detect mode (TOR analysis vs technical Q&A)
      Step 1 — SELECT: Claude picks relevant brands from our list
      Step 2 — ANSWER: Claude answers with spec context
    """
    if history is None:
        history = []

    product_list = _load_product_list()
    if not product_list:
        from agents.general_agent import handle as general_handle
        return general_handle(message=message, user_name=user_name,
                              user_id=user_id, source=source,
                              history=history, intent=intent)

    # Step 0: detect mode
    mode = _detect_mode(message)
    logger.info("[engineer_agent] mode=%s for: %s", mode, message[:60])

    # Step 1: select relevant brands
    selected = _select_relevant_brands(message, product_list)

    # Step 2: build system + answer
    system = _build_system(mode, selected, product_list)

    content = message if history else (
        "Engineer/Staff: %s%s\n\n%s" % (
            user_name,
            " (via email)" if source == "email" else "",
            message,
        )
    )
    messages = history + [{"role": "user", "content": content}]

    # TOR analysis may need more tokens for the compliance table
    max_tokens = 1500 if mode == "tor_analysis" else 768

    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=max_tokens,
            system=system,
            messages=messages,
        )
        reply = response.content[0].text.strip()
        logger.info("[engineer_agent] replied (mode=%s, brands=%s): %s",
                    mode, [s["brand"] for s in selected], reply[:80])
        return reply

    except Exception as e:
        logger.error("[engineer_agent] Claude API error: %s", e)
        return ("ขออภัยครับ ระบบขัดข้องชั่วคราว "
                "กรุณาติดต่อทีม VRCOMM ที่ iddhi.t@vrcomm.net ครับ")
