# Analysis engine module. Kinda complicated

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

# Function that classifies a company's product type based on the text of its Business section. It counts the occurrences of keywords associated with consumable/recurring revenue and durable/one-time revenue, and then classifies the product type based on which set of keywords is more prevalent. If neither set of keywords is found, it returns "Undetermined". If both sets are found with equal frequency, it returns "Mixed signal". The function returns a dictionary containing the classification label and the raw scores for each category.
# Will need to be made more robust in later models, using more complex lexical values to assure whether durable or consumable
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
# Needs to be made more robust in future, larger list of end-markets, etc
COMMON_END_MARKETS = [
    'automotive', 'aerospace', 'defense', 'healthcare', 'pharmaceutical',
    'consumer electronics', 'energy', 'oil and gas', 'retail',
    'financial services', 'telecommunications', 'industrial', 'construction',
    'agriculture', 'government', 'education', 'hospitality', 'semiconductor',
    'heavy machinery', 'logistics', 'transportation', 'manufacturing',
    'biotechnology', 'real estate', 'media', 'entertainment',
]

# Function that extracts end markets from the Business section text by checking for the presence of common end market keywords. It converts the text to lowercase and returns a list of end markets that are mentioned in the text. The function relies on a predefined list of common end markets, which can be expanded in future iterations for greater robustness.
def extract_end_markets(business_text):
    text_lower = business_text.lower()
    return [m for m in COMMON_END_MARKETS if m in text_lower]


# ---------------------------------------------------------------------
# Year-over-year Risk Factors diffing
# ---------------------------------------------------------------------
# Function that splits a block of text into sentences, filtering out sentences that are shorter than a specified minimum length. It uses regular expressions to split the text on punctuation followed by whitespace, and then returns a list of sentences that meet the length requirement.
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
# This bit is really cool
# Function that computes market share for a target company against a list of peer companies over a specified timeline. It takes in the target ticker, a list of peer tickers, a timeline of years, and several dependency-injected functions for loading the ticker map, resolving CIKs, fetching company facts, and extracting annual revenue series. The function retrieves revenue data for the target and peer companies, scales it to millions, and then calculates the market share for each company in each year by dividing its revenue by the total revenue of all companies in the peer group for that year. It returns three dictionaries: one mapping tickers to their revenue by year, one mapping tickers to their market share by year, and one mapping tickers to their company names.
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
