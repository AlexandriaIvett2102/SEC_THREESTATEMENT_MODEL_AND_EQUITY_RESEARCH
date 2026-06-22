import re
from bs4 import BeautifulSoup

# Function that fetches the SEC submissions for a given CIK using the provided sec_get function. It constructs the URL for the submissions endpoint, makes the request, and returns the JSON response. If the request fails, it raises a RuntimeError with the HTTP status code.
def fetch_submissions(cik, sec_get):
    """One fetch of SEC's submissions endpoint, reused by everything below
    that needs either the filing history or the company's profile (SIC
    code etc.) -- avoids hitting the same endpoint twice per ticker."""
    url = f"https://data.sec.gov/submissions/CIK{cik}.json"
    resp = sec_get(url)
    if resp.status_code != 200:
        raise RuntimeError(f"submissions request failed: HTTP {resp.status_code}")
    return resp.json()

# Function that extracts the company profile information from the SEC submissions data. It retrieves the company name, SIC code, and SIC description from the JSON data and returns them in a dictionary.
def get_company_profile(submissions_data):
    return {
        'name': submissions_data.get('name'),
        'sic': submissions_data.get('sic'),
        'sicDescription': submissions_data.get('sicDescription'),
    }

# Function that iterates over the 10-K filings in the SEC submissions data. It retrieves the recent filings from the JSON data and yields the form type, accession number, primary document name, and filing date for each 10-K or 10-K/A filing.
def _iter_10k_filings(submissions_data):
    recent = submissions_data.get('filings', {}).get('recent', {})
    forms = recent.get('form', [])
    accessions = recent.get('accessionNumber', [])
    docs = recent.get('primaryDocument', [])
    dates = recent.get('filingDate', [])
    for form, accn, doc, fdate in zip(forms, accessions, docs, dates):
        if form in ('10-K', '10-K/A'):
            yield form, accn, doc, fdate

# Function that retrieves the N-th most recent 10-K filing from the SEC submissions data. It takes the submissions data and an optional parameter n (default is 0 for the most recent filing). It returns a tuple containing the document URL, filing date, and accession number of the N-th 10-K filing. If there are fewer than n filings, it returns None.
def get_nth_10k_filing(submissions_data, n=0):
    """n=0 -> most recent 10-K, n=1 -> the one before that, etc. Needed for
    year-over-year Risk Factors diffing."""
    cik = str(int(submissions_data.get('cik')))
    filings = list(_iter_10k_filings(submissions_data))
    if n >= len(filings):
        return None
    form, accn, doc, fdate = filings[n]
    accn_nodash = accn.replace('-', '')
    doc_url = f"https://www.sec.gov/Archives/edgar/data/{cik}/{accn_nodash}/{doc}"
    return doc_url, fdate, accn

# Function that retrieves the most recent 10-K filing for a given CIK using the provided sec_get function. It fetches the submissions data, gets the most recent 10-K filing, and returns a tuple containing the document URL, filing date, and accession number. If no 10-K filing is found, it raises a RuntimeError.
def get_latest_10k_filing(cik, sec_get):
    """Convenience wrapper kept for backwards compatibility / standalone use."""
    data = fetch_submissions(cik, sec_get)
    result = get_nth_10k_filing(data, n=0)
    if result is None:
        raise RuntimeError("No 10-K found in recent filings for this CIK.")
    return result


# ---------------------------------------------------------------------
# SIC-code peer lookup (the structured backstop for competitor discovery)
# ---------------------------------------------------------------------
# Function that looks up other companies sharing the same SIC code via the SEC's browse-edgar tool. It takes the SIC code, a function to make SEC GET requests, an optional CIK to exclude from the results, and a limit on the number of results to return. It constructs the URL for the browse-edgar endpoint, makes the request, and parses the XML response to extract the names and CIKs of peer companies. It returns a list of dictionaries containing the name and CIK of each peer company.
def sic_peer_lookup(sic_code, sec_get, exclude_cik=None, limit=20):
    """
    Pulls other companies sharing the same 4-digit SIC code via SEC's
    browse-edgar tool (Atom feed output -- structured XML, not scraped
    HTML). This is SEC's own public company search, so it's the same kind
    of request as everything else in this script, just a different
    endpoint. Returns [{name, cik}].
    """
    url = (f"https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany"
           f"&SIC={sic_code}&type=10-K&dateb=&owner=include&count={limit}&output=atom")
    resp = sec_get(url)
    if resp.status_code != 200:
        return []  # non-fatal: text-mined competitors still stand on their own

# Parses HTML and XML with BeautifulSoup, which is more robust than lxml.etree for this
    soup = BeautifulSoup(resp.text, 'xml')
    exclude_int = int(exclude_cik) if exclude_cik else None
    peers = []
    for entry in soup.find_all('entry'):
        title_tag = entry.find('title')
        if not title_tag:
            continue
        title = title_tag.get_text(strip=True)
        cik_match = re.search(r'(\d{10})', entry.get_text())
        name = re.sub(r'\(.*$', '', title).strip()
        if not cik_match:
            continue
        if exclude_int is not None and int(cik_match.group(1)) == exclude_int:
            continue
        peers.append({'name': name, 'cik': cik_match.group(1)})
    return peers

# Function that normalizes the text of a filing by converting HTML to clean, single-spaced plain text. It uses BeautifulSoup to parse the HTML, extracts the text, replaces certain Unicode characters with their ASCII equivalents, and collapses multiple whitespace characters into a single space. The resulting text is stripped of leading and trailing whitespace.
def normalize_filing_text(html):
    """HTML -> clean, single-spaced plain text. Flattening to single spaces
    (rather than trying to preserve paragraph breaks) sidesteps a real edge
    case: some filings have hard line-wraps embedded inside a single
    paragraph's text, which would otherwise get misread as paragraph
    boundaries."""
    soup = BeautifulSoup(html, 'html.parser')
    text = soup.get_text(separator=' ')
    text = text.replace('\xa0', ' ').replace('\u2019', "'").replace('\u2018', "'")
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


# Heading patterns. Flexible on punctuation/spacing since real filings vary
# ("Item 1. Business", "ITEM 1 - BUSINESS", "Item\xa01.Business", etc.)
# normalize_filing_text already collapses \xa0 -> ' ' before these run.
ITEM1_BUSINESS = r'item\s*1\.?\s*business\b'
ITEM1A_RISK = r'item\s*1a\.?\s*risk\s*factors\b'
ITEM1B_STAFF = r'item\s*1b\.?\s*unresolved\s*staff\s*comments\b'
ITEM2_PROPS = r'item\s*2\.?\s*properties\b'


def _heading_positions(text, pattern):
    return [m.start() for m in re.finditer(pattern, text, re.IGNORECASE)]


def extract_section(text, start_pattern, end_patterns, min_gap=400):
    """
    Finds every candidate occurrence of start_pattern, pairs each with the
    NEXT occurrence of any end_pattern after it, and keeps the pair with the
    largest gap -- this is what correctly skips the Table of Contents (where
    consecutive item headings sit only a few characters apart) and lands on
    the real section body (which runs for thousands of characters).
    """
    starts = _heading_positions(text, start_pattern)
    if not starts:
        return None

    end_positions = sorted(
        pos for p in end_patterns for pos in _heading_positions(text, p)
    )
    if not end_positions:
        return None

    best = None
    best_gap = -1
    for s in starts:
        nxt = next((e for e in end_positions if e > s), None)
        if nxt is None:
            continue
        gap = nxt - s
        if gap > best_gap:
            best_gap = gap
            best = (s, nxt)

    if best is None or best_gap < min_gap:
        return None

    s, e = best
    return text[s:e].strip()


def extract_business_section(text):
    return extract_section(text, ITEM1_BUSINESS, [ITEM1A_RISK, ITEM1B_STAFF, ITEM2_PROPS])


def extract_risk_factors_section(text):
    return extract_section(text, ITEM1A_RISK, [ITEM1B_STAFF, ITEM2_PROPS])


# ---------------------------------------------------------------------
# Competitor name extraction
# ---------------------------------------------------------------------
CORP_SUFFIX = r'(?:Inc|Incorporated|Corp|Corporation|LLC|Ltd|Limited|Co|Company|plc|PLC|Group|Holdings|N\.V\.|S\.A\.|AG)\.?'

# High-confidence: a capitalized phrase WITH a corporate suffix, anywhere in
# the section. Reliable on its own regardless of nearby wording.
COMPANY_NAME_RE = re.compile(
    r'\b([A-Z][A-Za-z0-9&\'\-]*(?:\s+[A-Z][A-Za-z0-9&\'\-]*){0,4}\s+' + CORP_SUFFIX + r')\b'
)

# Broader: ANY capitalized phrase, used only inside a window around an
# explicit competition-related cue -- this is what catches the much more
# common real-world case of competitors named without a suffix at all
# ("we compete with Apple, Samsung, and Dell"). Deliberately does NOT
# include the corporate-suffix alternative here -- COMPANY_NAME_RE already
# covers that case, and mixing the two let phrases bleed across sentence
# boundaries (a suffix's optional trailing period got consumed as part of
# the suffix match, so chaining could continue right past it into the next
# sentence's capitalized word).
CAPITALIZED_PHRASE_RE = re.compile(
    r'\b([A-Z][A-Za-z0-9&\'\-]*(?:\s+[A-Z][A-Za-z0-9&\'\-]*){0,3})\b'
)
COMPETITION_CUE_RE = re.compile(
    r'\b(?:compet(?:e|es|ing|ition|itor|itors|itive)|principal\s+rivals?|key\s+rivals?)\b',
    re.IGNORECASE
)
CUE_WINDOW = 350  # chars scanned after each cue word

_STOPWORDS = {
    'we', 'our', 'the', 'this', 'that', 'these', 'those', 'it', 'its', 'they',
    'their', 'in', 'on', 'at', 'as', 'if', 'is', 'are', 'was', 'were', 'be',
    'been', 'being', 'item', 'part', 'note', 'table', 'contents', 'company',
    'companies', 'inc', 'corp', 'llc', 'ltd', 'and', 'or', 'but', 'for',
    'with', 'from', 'to', 'of', 'a', 'an', 'such', 'other', 'others', 'no',
    'none', 'not', 'also', 'may', 'will', 'would', 'us', 'united', 'states',
    'north', 'america', 'europe', 'asia', 'including', 'among', 'these',
    'competition', 'competitive', 'competitors', 'competitor', 'competes',
    'compete', 'competing', 'oem', 'oems', 'gaap', 'sec', 'fda', 'usd',
    'ipo', 'ceo', 'cfo', 'ip',
}

# Generic words that technically match the pattern but aren't competitor
# names -- the company's own name (passed in) gets filtered out separately.
_GENERIC_FALSE_POSITIVES = {'the company', 'our company'}

# Helper that pulls capitalized phrases out of text, filtering out ones that are too short or start with a common stopword. This is deliberately broad
# since it's just a candidate generator, not a verifier; match_against_ticker_map does the real confirming by cross-referencing against actual SEC company titles. The competition cue window is already a strong signal, so this
def _phrase_candidates(text):
    out = []
    for m in CAPITALIZED_PHRASE_RE.finditer(text):
        phrase = m.group(1).strip()
        words = phrase.split()
        if words[0].lower().rstrip('.,') in _STOPWORDS or len(phrase) < 3:
            continue
        if len(words) == 1 and words[0].lower().rstrip('.,') in _STOPWORDS:
            continue
        out.append(phrase)
    return out

# Self explanatory from this point
def extract_competitor_candidates(business_text, own_company_name=None):
    """
    Pulls candidate company names out of the Business section text, two ways:
    (1) anywhere a corporate-suffixed name appears (high confidence on its
    own), and (2) any capitalized phrase found within a window around an
    explicit competition-related cue word (catches the far more common case
    of competitors named without a suffix). This is deliberately broad --
    it's a candidate generator, not a verifier; match_against_ticker_map
    does the real confirming.
    """
    seen = set()
    candidates = []

    def _add(phrase):
        key = phrase.lower().rstrip('.')
        if key in _GENERIC_FALSE_POSITIVES:
            return
        if own_company_name and own_company_name.lower().split()[0] in key:
            return
        if key not in seen:
            seen.add(key)
            candidates.append(phrase)

    for m in COMPANY_NAME_RE.finditer(business_text):
        _add(m.group(1).strip())

    for m in COMPETITION_CUE_RE.finditer(business_text):
        start = max(0, m.start() - 50)
        end = min(len(business_text), m.end() + CUE_WINDOW)
        for phrase in _phrase_candidates(business_text[start:end]):
            _add(phrase)

    return candidates


def match_against_ticker_map(candidates, ticker_map):
    """
    Cross-references each candidate name against SEC's official company
    titles. This is the real filter -- it confirms the candidate is an
    actual public company and resolves it to a ticker/CIK, discarding
    capitalized-noise false positives and private competitors alike (private
    companies are real competitors but won't have 10-K financials to compare
    against, so they're tracked separately rather than dropped silently).
    Word-boundary matching (not raw substring) is what avoids a short
    candidate like "We" spuriously matching inside an unrelated longer word
    like "Newell" -- there's no word boundary in the middle of "Newell" for
    \\bwe\\b to land on, so this is safe at any candidate length.
    """
    title_index = {item['title'].lower(): item for item in ticker_map.values()}

    confirmed = []
    unmatched = []
    for name in candidates:
        key = name.lower().rstrip('.')
        hit = title_index.get(key)
        if not hit:
            pattern = re.compile(r'\b' + re.escape(key) + r'\b')
            for title_lower, item in title_index.items():
                if pattern.search(title_lower):
                    hit = item
                    break
        if hit:
            confirmed.append({
                'mentioned_as': name,
                'matched_title': hit['title'],
                'ticker': hit['ticker'],
                'cik': str(hit['cik_str']).zfill(10),
            })
        else:
            unmatched.append(name)
    return confirmed, unmatched
