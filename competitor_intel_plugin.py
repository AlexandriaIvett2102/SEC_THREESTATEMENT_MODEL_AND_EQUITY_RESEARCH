import os
import datetime
from openpyxl import load_workbook
import ipywidgets as widgets
from IPython.display import display, clear_output, FileLink

from sec_model_builder import (
    sec_get, load_ticker_map, resolve_cik, fetch_company_facts,
    extract_annual_series, CONCEPT_TAXONOMIES, build_excel_model,
    pull_company_data, open_file_in_default_app,
)
from filing_extractor import (
    fetch_submissions, get_company_profile, get_nth_10k_filing,
    normalize_filing_text, extract_business_section, extract_risk_factors_section,
    extract_competitor_candidates, match_against_ticker_map, sic_peer_lookup,
)
from analysis_engine import (
    classify_product_type, extract_end_markets, diff_risk_factors, compute_market_share,
)
from excel_plugins import (
    build_competitive_landscape_sheet, build_industry_factors_sheet, build_scorecard_sheet,
)

MAX_PEERS = 6


def run_full_research_workbook(ticker, lookback=5, log=print):
    ticker = ticker.strip().upper()
    ticker_map = load_ticker_map()
    cik, company_name = resolve_cik(ticker, ticker_map)
    if not cik:
        raise ValueError(f"Ticker '{ticker}' not found in SEC's registry.")
    log(f"🎯 {company_name} ({ticker}) | CIK {cik}")

    # --- filings + profile (one submissions fetch, reused) ---
    submissions = fetch_submissions(cik, sec_get)
    profile = get_company_profile(submissions)

    latest = get_nth_10k_filing(submissions, n=0)
    if latest is None:
        raise RuntimeError("No 10-K on file for this company.")
    latest_url, latest_date, _ = latest
    prior = get_nth_10k_filing(submissions, n=1)

    log("📄 Pulling latest 10-K text...")
    resp = sec_get(latest_url)
    if resp.status_code != 200:
        raise RuntimeError(f"Could not fetch 10-K document: HTTP {resp.status_code}")
    latest_text = normalize_filing_text(resp.text)
    business = extract_business_section(latest_text) or ""
    risk = extract_risk_factors_section(latest_text) or ""
    if not business:
        log("   ⚠️ Could not isolate the Business section — competitor/product-type reads may be empty.")
    if not risk:
        log("   ⚠️ Could not isolate the Risk Factors section.")

    risk_prior = ""
    if prior is not None:
        prior_url, prior_date, _ = prior
        resp_prior = sec_get(prior_url)
        if resp_prior.status_code == 200:
            prior_text = normalize_filing_text(resp_prior.text)
            risk_prior = extract_risk_factors_section(prior_text) or ""
    else:
        log("   ℹ️ No prior-year 10-K found — skipping risk-factor diff.")

    # --- competitor discovery: text-mined + SIC backstop ---
    log("🔍 Mining competitor names from filing text...")
    candidates = extract_competitor_candidates(business, own_company_name=company_name)
    confirmed, unmatched = match_against_ticker_map(candidates, ticker_map)
    text_mined_tickers = [c['ticker'] for c in confirmed]
    log(f"   Found {len(confirmed)} named public competitor(s), {len(unmatched)} likely private.")

    sic_peer_tickers = []
    if profile.get('sic'):
        log(f"🔍 Cross-checking SIC {profile['sic']} ({profile.get('sicDescription')}) peers...")
        sic_entries = sic_peer_lookup(profile['sic'], sec_get, exclude_cik=cik)
        cik_to_ticker = {str(int(item['cik_str'])): item['ticker'] for item in ticker_map.values()}
        for entry in sic_entries:
            tk = cik_to_ticker.get(str(int(entry['cik'])))
            if tk and tk != ticker:
                sic_peer_tickers.append(tk)

    all_peer_tickers = list(dict.fromkeys(text_mined_tickers + sic_peer_tickers))[:MAX_PEERS]
    log(f"   Peer group ({len(all_peer_tickers)}): {', '.join(all_peer_tickers) if all_peer_tickers else '(none found)'}")

    # --- market share ---
    log("📊 Computing relative market share across peer group...")
    current_year = datetime.datetime.now().year
    timeline = [str(y) for y in range(current_year - lookback, current_year)]
    revenue_by_ticker, share_by_ticker, names_by_ticker = compute_market_share(
        ticker, all_peer_tickers, timeline,
        load_ticker_map_fn=load_ticker_map,
        resolve_cik_fn=resolve_cik,
        fetch_company_facts_fn=fetch_company_facts,
        extract_annual_series_fn=extract_annual_series,
        revenue_tags=CONCEPT_TAXONOMIES["Total Revenue"],
    )
    names_by_ticker[ticker] = company_name

    # --- industry factor analysis ---
    log("🧭 Classifying product type and end markets...")
    product_type_result = classify_product_type(business)
    end_markets = extract_end_markets(business)

    log("🆕 Diffing Risk Factors vs. prior year's filing...")
    new_risk_sentences = diff_risk_factors(risk_prior, risk) if risk_prior else []
    log(f"   {len(new_risk_sentences)} new risk statement(s) flagged.")

    # --- scorecard inputs ---
    target_shares = share_by_ticker.get(ticker, {})
    years_with_share = [yr for yr in timeline if yr in target_shares]
    share_trend = (target_shares[years_with_share[-1]] - target_shares[years_with_share[0]]
                   if len(years_with_share) >= 2 else 0.0)

    target_revenue = revenue_by_ticker.get(ticker, {})
    years_with_rev = [yr for yr in timeline if yr in target_revenue]
    if len(years_with_rev) >= 2:
        prev_rev = target_revenue[years_with_rev[-2]]
        last_rev = target_revenue[years_with_rev[-1]]
        revenue_growth = (last_rev - prev_rev) / prev_rev if prev_rev else 0.0
    else:
        revenue_growth = 0.0

    # --- build the workbook: model first, then attach plugin sheets ---
    log("💎 Building three-statement model...")
    _, _, model_timeline, model_data, _ = pull_company_data(ticker, lookback)
    out_path = os.path.abspath(f"{ticker}_research_workbook.xlsx")
    build_excel_model(ticker, company_name, cik, model_data, model_timeline, out_path)

    log("💎 Attaching Competitive Landscape, Industry Factors, and Signal Summary tabs...")
    wb = load_workbook(out_path)
    build_competitive_landscape_sheet(
        wb, ticker, names_by_ticker, timeline, revenue_by_ticker, share_by_ticker,
        unmatched_private_competitors=unmatched, sic_description=profile.get('sicDescription'),
    )
    build_industry_factors_sheet(
        wb, ticker, company_name, latest_date, product_type_result, end_markets, new_risk_sentences,
    )
    build_scorecard_sheet(
        wb, ticker, company_name, share_trend, revenue_growth,
        len(new_risk_sentences), product_type_result,
    )
    wb.save(out_path)
    log(f"\n💎 Full research workbook saved: {out_path}")
    return out_path


# =====================================================================
# UI
# =====================================================================
ticker_box = widgets.Text(value='ILMN', placeholder='Ticker...', description='Ticker:', layout=widgets.Layout(width='30%'))
years_box = widgets.IntText(value=5, description='Years:', layout=widgets.Layout(width='20%'))
action_btn = widgets.Button(description='Build Full Research Workbook', button_style='success', icon='search', layout=widgets.Layout(width='35%'))
display_output = widgets.Output()


def on_click(change_event):
    with display_output:
        clear_output()
        ticker = ticker_box.value.strip().upper()
        lookback = years_box.value
        if not ticker:
            print("⚠️ Please provide a valid stock ticker symbol.")
            return
        try:
            out_path = run_full_research_workbook(ticker, lookback)
        except Exception as e:
            print(f"❌ {e}")
            return
        display(FileLink(out_path))
        open_file_in_default_app(out_path)


action_btn.on_click(on_click)
ui_panel = widgets.HBox([ticker_box, years_box, action_btn])
print("🕸️  Competitor & Industry Intelligence Plugins")
display(ui_panel, display_output)
