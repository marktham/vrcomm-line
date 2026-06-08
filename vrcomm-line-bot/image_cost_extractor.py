"""
image_cost_extractor.py — Claude Vision cost-table extractor

Turns an image of a cost/price table (sent by PM via email)
into structured data that the Quotation Agent can use directly.

Two tasks in one call:
  1. Extract line items  → list of {brand, product, qty, unit_cost_thb}
  2. Extract customer    → customer name if visible in the image or email text

Supported image types: JPEG, PNG, GIF, WEBP (whatever Graph API returns)
"""
import os, re, json, base64, logging
from anthropic import Anthropic

logger = logging.getLogger(__name__)
client = Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ── Vision prompt ─────────────────────────────────────────────────────────────

_VISION_PROMPT = """This image was sent by a Product Manager at VRCOMM (a cybersecurity reseller in Thailand).
It may contain a cost/price table listing products for a customer quotation.

Extract ALL information from the table and return a single JSON object:

{
  "is_cost_table": true or false,
  "customer_name": "company name if visible anywhere in the image" or null,
  "items": [
    {
      "brand":         "vendor/brand name (e.g. Sangfor, Sectigo, LogPoint)",
      "product":       "product or model name",
      "qty":           <positive integer, default 1 if not shown>,
      "unit_cost_thb": <number in THB, null if currency unclear or not shown>
    }
  ]
}

Rules:
- Set is_cost_table=false if the image has no price/cost table at all
- Include EVERY row — do not skip any
- If a cost value is in USD or other currency, still capture it as unit_cost_thb with a note in the product field like "(USD 1200)"
- qty defaults to 1 if not shown
- Return ONLY raw JSON. No markdown, no explanation.
"""

_BODY_CUSTOMER_PROMPT = """Extract the customer company name from this email text.
The email is from a Product Manager sending a cost sheet for a quotation.
Look for phrases like: "ลูกค้า:", "Customer:", "Quote for", "สำหรับ", "Attn:", company names.

Return ONLY the company name as a plain string.
If no customer name is found, return: null
"""


# ── Image → items ─────────────────────────────────────────────────────────────

def extract_cost_from_image(image_bytes: bytes,
                             media_type: str = "image/jpeg") -> dict:
    """
    Use Claude Vision to extract a cost table from image bytes.

    Returns:
        {
          "is_cost_table": bool,
          "customer_name": str or None,
          "items": [{brand, product, qty, unit_cost_thb}, ...]
        }
    """
    # Normalise media_type
    mt = media_type.lower()
    if "png"  in mt: mt = "image/png"
    elif "gif" in mt: mt = "image/gif"
    elif "webp" in mt: mt = "image/webp"
    else:             mt = "image/jpeg"

    image_b64 = base64.standard_b64encode(image_bytes).decode()

    try:
        resp = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type":       "base64",
                            "media_type": mt,
                            "data":       image_b64,
                        },
                    },
                    {"type": "text", "text": _VISION_PROMPT},
                ],
            }],
        )
        raw = resp.content[0].text.strip()
        raw = re.sub(r'^```(?:json)?\s*', '', raw, flags=re.MULTILINE)
        raw = re.sub(r'\s*```$',          '', raw, flags=re.MULTILINE)
        result = json.loads(raw)
        logger.info(
            "[image_extractor] is_cost_table=%s, items=%d, customer=%s",
            result.get("is_cost_table"),
            len(result.get("items", [])),
            result.get("customer_name"),
        )
        return result

    except json.JSONDecodeError as e:
        logger.error("[image_extractor] JSON parse error: %s | raw: %s", e, raw[:200])
        return {"is_cost_table": False, "customer_name": None, "items": []}
    except Exception as e:
        logger.error("[image_extractor] Vision API error: %s", e)
        return {"is_cost_table": False, "customer_name": None, "items": []}


# ── Email body → customer name ─────────────────────────────────────────────────

def extract_customer_from_email(body_text: str, subject: str = "") -> str | None:
    """
    Use Haiku to find the customer company name in the email body/subject.
    Fast and cheap — called alongside the vision extraction.
    """
    combined = ("Subject: %s\n\n%s" % (subject, body_text))[:2000]
    try:
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=80,
            system=_BODY_CUSTOMER_PROMPT,
            messages=[{"role": "user", "content": combined}],
        )
        result = resp.content[0].text.strip()
        if result.lower() in ("null", "none", "", "n/a"):
            return None
        logger.info("[image_extractor] customer from email body: %s", result)
        return result
    except Exception as e:
        logger.error("[image_extractor] customer extraction error: %s", e)
        return None


# ── Combined entry point ───────────────────────────────────────────────────────

def process_email_cost_sheet(
    attachments: list,
    email_body: str = "",
    email_subject: str = "",
) -> dict:
    """
    Process all image attachments from an email and combine results.

    Args:
        attachments: list of {filename, content_type, content_bytes}
        email_body:  plain-text email body for customer name extraction
        email_subject: email subject line

    Returns:
        {
          "is_cost_sheet": bool,      — True if at least one image had a table
          "customer_name": str|None,  — from image or email body
          "items": [...],             — merged from all images
          "image_count": int,
        }
    """
    all_items    = []
    customer     = None
    found_table  = False
    image_count  = 0

    for att in attachments:
        ct = att.get("content_type", "")
        if not any(img in ct.lower() for img in ("image/", "jpeg", "jpg", "png", "gif", "webp")):
            continue

        image_count += 1
        img_bytes = att.get("content_bytes", b"")
        if not img_bytes:
            continue

        result = extract_cost_from_image(img_bytes, media_type=ct)

        if result.get("is_cost_table"):
            found_table = True
            all_items.extend(result.get("items", []))
            if not customer and result.get("customer_name"):
                customer = result["customer_name"]

    # If customer not found in images, try email body
    if not customer and email_body:
        customer = extract_customer_from_email(email_body, email_subject)

    logger.info(
        "[image_extractor] process_email_cost_sheet: images=%d, table=%s, items=%d, customer=%s",
        image_count, found_table, len(all_items), customer,
    )
    return {
        "is_cost_sheet": found_table,
        "customer_name": customer,
        "items":         all_items,
        "image_count":   image_count,
    }
