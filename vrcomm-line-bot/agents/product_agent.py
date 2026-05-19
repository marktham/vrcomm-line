"""
agents/product_agent.py — VRCOMM Product Information Agent

Data source: vrcomm-line-bot/product/VRCOMM_ProductList.xlsx
  - Column A: Brand/Product name
  - Column B: Website URL (vendor site or VRCOMM product page)

Logic:
  1. Load product list from Excel (cached)
  2. Check if any brand in the list matches the user's query
  3. If matched  → fetch that product's URL for real content → Claude answers
  4. If no match → Claude suggests comparable products or combined solutions from the list
"""
import os, re, logging, time
import requests
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ── Paths ─────────────────────────────────────────────────────────────────────

_BASE_DIR      = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_PRODUCT_LIST  = os.path.join(_BASE_DIR, "product", "VRCOMM_ProductList.xlsx")

# ── Product list cache ────────────────────────────────────────────────────────

_product_cache      = None   # list of {"brand": str, "url": str}
_product_cache_time = 0.0
_CACHE_TTL          = 300    # 5 minutes


def _load_product_list() -> list:
    """
    Load VRCOMM_ProductList.xlsx and return list of {brand, url}.
    Cached for 5 minutes.
    """
    global _product_cache, _product_cache_time

    path = _PRODUCT_LIST
    if not os.path.isfile(path):
        logger.error("ProductList not found at: %s", path)
        return []

    mtime = os.path.getmtime(path)
    now   = time.time()

    if (_product_cache is not None
            and mtime == _product_cache_time
            and (now - _product_cache_time) < _CACHE_TTL):
        return _product_cache

    try:
        import openpyxl
        wb   = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws   = wb.active
        rows = list(ws.iter_rows(values_only=True))

        products = []
        for row in rows:
            brand = str(row[0]).strip() if row[0] else ""
            url   = str(row[1]).strip() if len(row) > 1 and row[1] else ""
            # Skip header row if exists
            if brand.lower() in ("brand", "product", "name", ""):
                continue
            if brand:
                products.append({"brand": brand, "url": url})

        _product_cache      = products
        _product_cache_time = mtime
        logger.info("ProductList loaded: %d products", len(products))
        return products

    except Exception as e:
        logger.error("ProductList load error: %s", e)
        return []


# ── Product matching ──────────────────────────────────────────────────────────

def _match_products(message: str, product_list: list) -> list:
    """
    Return products from the list whose brand name appears in the message.
    Case-insensitive. Returns list of matched {brand, url} dicts.
    """
    msg_lower = message.lower()
    matched   = []
    for p in product_list:
        brand_lower = p["brand"].lower()
        # Match if brand name appears as a word/phrase in the message
        if re.search(r'\b' + re.escape(brand_lower) + r'\b', msg_lower):
            matched.append(p)
    return matched


# ── URL content fetcher (cached) ──────────────────────────────────────────────

_url_cache: dict = {}   # url → {"content": str, "fetched_at": float}
_URL_CACHE_TTL = 3600   # 1 hour


def _fetch_url(url: str) -> str:
    """
    Fetch a URL and return cleaned plain text (HTML stripped).
    Cached for 1 hour. Returns empty string on failure.
    """
    url = url.strip()
    if not url:
        return ""

    now = time.time()
    cached = _url_cache.get(url)
    if cached and (now - cached["fetched_at"]) < _URL_CACHE_TTL:
        logger.info("URL cache hit: %s", url[:60])
        return cached["content"]

    try:
        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/120.0.0.0 Safari/537.36"
            ),
            "Accept-Language": "en-US,en;q=0.9,th;q=0.8",
        }
        resp = requests.get(url, headers=headers, timeout=8, allow_redirects=True)
        resp.raise_for_status()

        html = resp.text

        # Remove scripts, styles, nav, footer
        html = re.sub(r'<(script|style|nav|footer|header)[^>]*>.*?</\1>',
                      ' ', html, flags=re.DOTALL | re.IGNORECASE)
        # Strip remaining HTML tags
        text = re.sub(r'<[^>]+>', ' ', html)
        # Clean whitespace
        text = re.sub(r'[ \t]+', ' ', text)
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = text.strip()

        # Cap at 3000 chars to keep token usage reasonable
        content = text[:3000]
        _url_cache[url] = {"content": content, "fetched_at": now}
        logger.info("Fetched URL (%d chars): %s", len(content), url[:60])
        return content

    except Exception as e:
        logger.warning("URL fetch failed [%s]: %s", url[:60], e)
        return ""


# ── System prompt builder ─────────────────────────────────────────────────────

_BASE_SYSTEM = """You are VRCOMM's Product Specialist.
VRCOMM is a Network and Cybersecurity solutions provider in Thailand.

Below is the COMPLETE and EXHAUSTIVE list of brands VRCOMM sells.
This list is final — VRCOMM does NOT sell any brand outside this list.

=== VRCOMM PRODUCT LIST (ALL brands we carry) ===
{product_list_text}
=================================================

ABSOLUTE RULES — violation is not acceptable:
1. NEVER mention, recommend, or reference ANY brand that does not appear in the list above.
   This includes — but is not limited to — Fortinet, Cisco, Palo Alto, Check Point,
   Juniper, SonicWall, Barracuda, Trend Micro, CrowdStrike, or any other brand
   not explicitly listed above.
2. When a customer asks about a category (e.g. firewall, endpoint, DLP, switch):
   → Recommend ONLY brands from the VRCOMM Product List that fit that category.
   → Do NOT use your general knowledge to suggest brands outside the list.
3. When a customer asks about a brand NOT in the list (e.g. "มี Fortinet ไหม"):
   → Say clearly that VRCOMM does not carry that brand.
   → Suggest the closest alternative(s) FROM THE LIST ONLY.
4. NEVER quote specific prices — if asked, say:
   "สำหรับราคา ทางทีม Sales ของ VRCOMM จะจัดทำใบเสนอราคาให้ครับ"
5. Reply in the SAME LANGUAGE as the customer (Thai → Thai, English → English)
6. Be concise and technical — max 4-5 short paragraphs
7. Plain text only — no markdown, no bullet symbols, no headers

{product_content_section}"""


def _build_system(product_list: list, fetched_contents: dict) -> str:
    # Format product list for context
    list_lines = ["%d. %s — %s" % (i+1, p["brand"], p["url"])
                  for i, p in enumerate(product_list)]
    product_list_text = "\n".join(list_lines) if list_lines else "(ไม่พบข้อมูล)"

    # Format fetched web content for matched products
    if fetched_contents:
        sections = []
        for brand, content in fetched_contents.items():
            if content:
                sections.append(
                    "=== %s (fetched from website) ===\n%s" % (brand, content)
                )
        product_content_section = (
            "\n=== PRODUCT DETAIL (from vendor websites) ===\n"
            + "\n\n".join(sections)
            if sections else ""
        )
    else:
        product_content_section = ""

    return _BASE_SYSTEM.format(
        product_list_text=product_list_text,
        product_content_section=product_content_section,
    )


# ── Main handler ──────────────────────────────────────────────────────────────

def handle(message: str, user_name: str, user_id: str,
           source: str = "line", history: list = None,
           intent: str = "product_info", **kwargs) -> str:
    """
    Handle product information requests.

    Steps:
    1. Load VRCOMM product list
    2. Match brands mentioned in the message
    3. Fetch URL content for matched products
    4. Build system prompt with list + content
    5. Claude responds (in-list: detailed answer; out-of-list: suggest alternatives)
    """
    if history is None:
        history = []

    # 1. Load product list
    product_list = _load_product_list()
    if not product_list:
        logger.warning("[product_agent] product list empty — falling back to general")
        from agents.general_agent import handle as general_handle
        return general_handle(message=message, user_name=user_name,
                              user_id=user_id, source=source,
                              history=history, intent=intent)

    # 2. Match products mentioned in the message
    matched = _match_products(message, product_list)
    logger.info("[product_agent] matched %d product(s): %s",
                len(matched), [m["brand"] for m in matched])

    # 3. Fetch URL content for matched products (up to 2 to stay fast)
    fetched = {}
    for p in matched[:2]:
        content = _fetch_url(p["url"])
        if content:
            fetched[p["brand"]] = content

    # 4. Build system + messages
    system  = _build_system(product_list, fetched)

    if history:
        content = message
    else:
        src_note = " (via email)" if source == "email" else ""
        content  = "Customer: %s%s\n\n%s" % (user_name, src_note, message)

    messages = history + [{"role": "user", "content": content}]

    # 5. Claude responds
    try:
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=768,
            system=system,
            messages=messages,
        )
        reply = response.content[0].text.strip()
        logger.info("[product_agent] replied to %s (matched=%s): %s",
                    user_name, [m["brand"] for m in matched], reply[:80])
        return reply

    except Exception as e:
        logger.error("[product_agent] Claude API error: %s", e)
        return ("ขออภัยครับ ระบบขัดข้องชั่วคราว "
                "กรุณาติดต่อทีม VRCOMM ที่ iddhi.t@vrcomm.net ครับ")
