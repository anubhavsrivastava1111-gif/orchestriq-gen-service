"""
OrchestrIQ Document Intelligence Engine v4 — Input Sanitizer
Strips media URLs, markdown images, emojis, and control chars from
content before it reaches AI extraction or document engines.
Prevents the 'fal.ai URL in Excel' class of pollution bugs.
"""
import re

_MEDIA_URL = re.compile(
    r'https?://\S*?\.(?:png|jpe?g|gif|webp|mp4|mov|webm|avi|svg)(?:\?\S*)?',
    re.IGNORECASE)
_FAL_URL = re.compile(r'https?://(?:v\d+b?\.)?fal\.(?:media|ai)/\S+', re.IGNORECASE)
_MD_IMAGE = re.compile(r'!\[[^\]]*\]\([^)]*\)')
_IMG_LINE = re.compile(r'^.*(?:🖼|Image URL|Video URL|Generated Image|Generated Video).*$',
                       re.IGNORECASE | re.MULTILINE)
_EMOJI = re.compile(
    "["
    "\U0001F300-\U0001FAFF"
    "\U00002700-\U000027BF"
    "\U0001F000-\U0001F02F"
    "\U00002600-\U000026FF"
    "\U0000FE00-\U0000FE0F"
    "\U0001F1E6-\U0001F1FF"
    "]+", flags=re.UNICODE)
_CTRL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')


def sanitize_text(text: str, max_len: int = 12000) -> str:
    """Clean a text blob for use as AI context / document data."""
    if not text:
        return ""
    t = str(text)
    t = _FAL_URL.sub('', t)
    t = _MEDIA_URL.sub('', t)
    t = _MD_IMAGE.sub('', t)
    t = _IMG_LINE.sub('', t)
    t = _EMOJI.sub('', t)
    t = _CTRL.sub('', t)
    # collapse 3+ blank lines
    t = re.sub(r'\n{3,}', '\n\n', t)
    t = t.strip()
    return t[:max_len]


def is_media_only(text: str) -> bool:
    """True if the content is essentially just media URLs / image markup
    (i.e., useless as business data)."""
    if not text:
        return True
    cleaned = sanitize_text(text, max_len=100000)
    # After stripping media, is there any real substance left?
    words = re.findall(r'[A-Za-z]{3,}', cleaned)
    return len(words) < 12


def sanitize_request(objective: str, company_context: str, available_data: str):
    """Sanitize all inbound fields. If available_data is media-only, drop it
    so engines rely on objective + fallback modeling instead of garbage."""
    obj = sanitize_text(objective or "", 2000) or "Executive Business Review"
    ctx = sanitize_text(company_context or "", 3000)
    data = "" if is_media_only(available_data or "") else sanitize_text(available_data or "", 12000)
    return obj, ctx, data
