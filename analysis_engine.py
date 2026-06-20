import re
import difflib
import datetime

# ---------------------------------------------------------------------
# Product-type classification (consumable/recurring vs. one-time/durable)
# ---------------------------------------------------------------------
CONSUMABLE_KEYWORDS = [
    'subscription', 'recurring revenue', 'renewal', 'consumable',
    'replacement cycle', 'recurring maintenance', 'wear item',
    'usage-based', 'membership fee', 'service contract', 'replenishment',
    'recurring basis', 'auto-renew',
]
DURABLE_KEYWORDS = [
    'one-time purchase', 'durable good', 'useful life of',
    'capital equipment', 'long product life', 'infrequent purchase',
    'buy once', 'capital expenditure by customers', 'long replacement cycle',
]


def classify_product_type(business_text):
    text_lower = business_text.lower()
    consumable_score = sum(text_lower.count(k) for k in CONSUMABLE_KEYWORDS)
    durable_score = sum(text_lower.count(k) for k in DURABLE_KEYWORDS)

    if consumable_score == 0 and durable_score == 0:
        label = 'Undetermined (no clear language found)'
    elif consumable_score > durable_score:
        label = 'Consumable / recurring-leaning'
    elif durable_score > consumable_score:
        label = 'Durable / one-time-leaning'
    else:
        label = 'Mixed signal'

    return {
        'classification': label,
        'consumable_score': consumable_score,
        'durable_score': durable_score,
    }


# ---------------------------------------------------------------------
# End-market detection
# ---------------------------------------------------------------------
COMMON_END_MARKETS = [
    'automotive', 'aerospace', 'defense', 'healthcare', 'pharmaceutical',
    'consumer electronics', 'energy', 'oil and gas', 'retail',
    'financial services', 'telecommunications', 'industrial', 'construction',
    'agriculture', 'government', 'education', 'hospitality', 'semiconductor',
    'heavy machinery', 'logistics', 'transportation', 'manufacturing',
    'biotechnology', 'real estate', 'media', 'entertainment',
]


def extract_end_markets(business_text):
    text_lower = business_text.lower()
    return [m for m in COMMON_END_MARKETS if m in text_lower]


# ---------------------------------------------------------------------
# Year-over-year Risk Factors diffing
# ---------------------------------------------------------------------
def _split_sentences(text, min_len=40):
    sentences = re.split(r'(?<=[.!?])\s+', text)
    return [s.strip() for s in sentences if len(s.strip()) >= min_len]


def diff_risk_factors(old_text, new_text, min_len=40):
    """
    Returns sentences present in new_text's Risk Factors section that
    weren't (close to) anything in old_text's -- i.e. likely new/emerging
    risk language the company added this filing cycle.
    """
    if not old_text or not new_text:
        return []

    old_sents = _split_sentences(old_text, min_len)
    new_sents = _split_sentences(new_text, min_len)

    new_only = []
    for ns in new_sents:
        # Treat as "new" if there's no close match in the prior year's text
        matches = difflib.get_close_matches(ns, old_sents, n=1, cutoff=0.6)
        if not matches:
            new_only.append(ns)
    return new_only


# ---------------------------------------------------------------------
# Market share computation
# ---------------------------------------------------------------------
def compute_market_share(target_ticker, peer_tickers, timeline,
                          load_ticker_map_fn, resolve_cik_fn,
                          fetch_company_facts_fn, extract_annual_series_fn,
                          revenue_tags):
    """
    Dependency-injected so this module doesn't have to import
    sec_model_builder directly (keeps it testable standalone). Returns
    (revenue_by_ticker, share_by_ticker, names_by_ticker).
    """
    ticker_map = load_ticker_map_fn()

    all_tickers = [target_ticker] + [t for t in peer_tickers if t != target_ticker]
    revenue_by_ticker = {}
    names_by_ticker = {}

    for tk in all_tickers:
        cik, name = resolve_cik_fn(tk, ticker_map)
        if not cik:
            continue
        try:
            facts = fetch_company_facts_fn(cik)
        except Exception:
            continue
        series = extract_annual_series_fn(facts, revenue_tags, timeline)
        scaled = {yr: round(v / 1_000_000, 2) for yr, v in series.items()}
        if scaled:
            revenue_by_ticker[tk] = scaled
            names_by_ticker[tk] = name

    share_by_ticker = {tk: {} for tk in revenue_by_ticker}
    for yr in timeline:
        total = sum(revenue_by_ticker[tk].get(yr, 0) or 0 for tk in revenue_by_ticker)
        if total <= 0:
            continue
        for tk in revenue_by_ticker:
            val = revenue_by_ticker[tk].get(yr)
            if val is not None:
                share_by_ticker[tk][yr] = round(val / total, 4)

    return revenue_by_ticker, share_by_ticker, names_by_ticker
