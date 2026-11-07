"""Scrape every book on books.toscrape.com: title, price, rating, stock count,
category, UPC and link.

Saves the results to CSV and/or Excel (.xlsx), and compares prices with the
previous run to write a price-changes report. Pages are fetched several at a
time, book pages are cached on disk so re-runs skip them, and pages that fail
are retried, then logged and skipped, so one bad page never loses the whole run.
"""

import argparse
import csv
import hashlib
import json
import logging
import re
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://books.toscrape.com/catalogue/"
PAGE_URL = BASE_URL + "page-{}.html"
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
FIELDS = ["title", "price_gbp", "rating", "stock", "category", "upc", "url"]
# The site files some books under these placeholder names; we group them as one.
UNCATEGORIZED = {"Add a comment", "Default"}
CHANGE_FIELDS = ["title", "upc", "old_price_gbp", "new_price_gbp", "change_gbp", "url"]

log = logging.getLogger("scraper")


def make_session(retries, workers):
    """One shared HTTP session that retries connection errors and 429/5xx with backoff.
    The connection pool is sized so every worker thread can reuse a connection."""
    session = requests.Session()
    retry = Retry(total=retries, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504])
    session.mount("https://", HTTPAdapter(max_retries=retry, pool_connections=workers, pool_maxsize=workers))
    session.headers["User-Agent"] = "books-scraper-example/3.0"
    return session


class Fetcher:
    """Downloads pages with a polite delay, optionally through an on-disk HTML cache."""

    def __init__(self, session, delay, cache_dir=None, use_cache=True):
        self.session = session
        self.delay = delay
        self.cache_dir = Path(cache_dir) if cache_dir else None
        self.use_cache = use_cache
        self.cache_hits = 0
        self._lock = threading.Lock()
        if self.cache_dir:
            self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _cache_path(self, url):
        return self.cache_dir / (hashlib.sha1(url.encode()).hexdigest() + ".html")

    def get(self, url, cache=False):
        """Return the page as BeautifulSoup. With cache=True, a cached copy is used if
        present and a fresh download is saved for next time."""
        path = self._cache_path(url) if cache and self.cache_dir else None
        if path and self.use_cache and path.exists():
            with self._lock:
                self.cache_hits += 1
            return BeautifulSoup(path.read_text(encoding="utf-8"), "html.parser")
        time.sleep(self.delay)  # be polite: each worker pauses before every request
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        resp.encoding = "utf-8"
        if path:
            tmp = path.with_suffix(".tmp")
            tmp.write_text(resp.text, encoding="utf-8")
            tmp.replace(path)  # atomic, so an interrupted run never leaves a half-written page
        return BeautifulSoup(resp.text, "html.parser")


def parse_books(soup, page_url):
    """Return one dict per book on a catalogue page. Skips (and logs) any book it can't parse."""
    books = []
    for card in soup.select("article.product_pod"):
        try:
            link = card.select_one("h3 a")
            rating_word = next(c for c in card.select_one("p.star-rating")["class"] if c in RATINGS)
            books.append({
                "title": link["title"],
                "price_gbp": round(float(re.sub(r"[^\d.]", "", card.select_one("p.price_color").text)), 2),
                "rating": RATINGS[rating_word],
                "stock": None, "category": None, "upc": None,  # filled from the detail page
                "url": urljoin(page_url, link["href"]),
            })
        except (AttributeError, KeyError, TypeError, ValueError, StopIteration) as e:
            log.warning("Skipped a book on %s: %r", page_url, e)
    return books


def parse_detail(soup):
    """Stock count, category and UPC from a book's own page."""
    table = {tr.th.text.strip(): tr.td.text.strip() for tr in soup.select("table.table tr")}
    stock = re.search(r"(\d+) available", table.get("Availability", ""))
    crumbs = soup.select("ul.breadcrumb li a")
    category = crumbs[2].text.strip() if len(crumbs) > 2 else None
    return {
        "stock": int(stock.group(1)) if stock else 0,
        "category": "Uncategorized" if category in UNCATEGORIZED else category,
        "upc": table.get("UPC"),
    }


def total_pages(soup):
    """Read 'Page 1 of 50' from the pager. A single-page catalogue has no pager."""
    pager = soup.select_one("li.current")
    match = re.search(r"of\s+(\d+)", pager.text) if pager else None
    return int(match.group(1)) if match else 1


def run_pool(func, items, workers, desc):
    """Run func on every item with a thread pool and a progress bar.
    Returns the items whose call raised, after one slower sequential retry each."""
    def attempt(item):
        try:
            func(item)
            return None
        except (requests.RequestException, AttributeError, OSError) as e:
            return item, e

    with ThreadPoolExecutor(max_workers=workers) as pool:
        results = list(tqdm(pool.map(attempt, items), total=len(items), desc=desc, unit="page"))
    failed = [r for r in results if r]

    # Second chance for pages that still failed after the HTTP-level retries.
    still_failed = []
    for item, _ in failed:
        time.sleep(2)
        retry = attempt(item)
        if retry:
            still_failed.append(retry)
    return still_failed


def scrape(fetcher, max_pages=None, workers=8):
    """Scrape catalogue pages 1..N, then every book's own page.
    Returns (books, failed_page_numbers, failed_book_count)."""
    # Page 1 tells us how many pages exist. Catalogue pages are never cached,
    # so prices are always live for price-change tracking.
    first = fetcher.get(PAGE_URL.format(1))
    last = total_pages(first)
    if max_pages:
        last = min(last, max_pages)

    pages = {1: parse_books(first, PAGE_URL.format(1))}

    def load_page(n):
        url = PAGE_URL.format(n)
        pages[n] = parse_books(fetcher.get(url), url)

    failed = run_pool(load_page, list(range(2, last + 1)), workers, "Catalogue")
    for page, e in failed:
        log.error("Page %d failed, skipping: %s", page, e)
    books = [b for n in sorted(pages) for b in pages[n]]  # keep site order
    log.info("Catalogue: %d books from %d/%d pages", len(books), len(pages), last)

    # Each book's own page has the exact stock count, category and UPC.
    def load_detail(book):
        book.update(parse_detail(fetcher.get(book["url"], cache=True)))

    failed_books = run_pool(load_detail, books, workers, "Book pages")
    for book, e in failed_books:
        log.error("Detail page failed, leaving stock/category/UPC blank: %s (%s)", book["url"], e)
    log.info("Book pages: %d done, %d from cache, %d failed",
             len(books) - len(failed_books), fetcher.cache_hits, len(failed_books))

    return books, [p for p, _ in failed], len(failed_books)


def save_csv(books, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDS)
        writer.writeheader()
        writer.writerows({**b, "price_gbp": f"{b['price_gbp']:.2f}"} for b in books)


HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="1F6E43")
GBP = "£#,##0.00"


def style_sheet(ws, max_width=70):
    """Bold coloured frozen header, filters, and columns sized to fit their contents."""
    for cell in ws[1]:
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in ws.columns:
        longest = max(len(f"{c.value:,.2f}" if isinstance(c.value, float) else str(c.value or "")) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(longest + 3, max_width)


def save_xlsx(books, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Books"
    ws.append(["Title", "Price (£)", "Rating (1-5)", "Stock", "Category", "UPC", "Link"])
    for b in books:
        ws.append([b["title"], b["price_gbp"], b["rating"], b["stock"], b["category"], b["upc"], b["url"]])
        row = ws.max_row
        ws.cell(row, 2).number_format = GBP
        link = ws.cell(row, 7)
        link.hyperlink = b["url"]
        link.style = "Hyperlink"
    style_sheet(ws)

    # Summary: book count and average price per category.
    summary = wb.create_sheet("Summary")
    summary.append(["Category", "Books", "Average price (£)"])
    by_cat = {}
    for b in books:
        by_cat.setdefault(b["category"] or "(unknown)", []).append(b["price_gbp"])
    # Biggest categories first; ties in alphabetical order.
    for cat in sorted(by_cat, key=lambda c: (-len(by_cat[c]), c)):
        prices = by_cat[cat]
        summary.append([cat, len(prices), round(sum(prices) / len(prices), 2)])
    if books:
        summary.append(["All books", len(books), round(sum(b["price_gbp"] for b in books) / len(books), 2)])
        for cell in summary[summary.max_row]:
            cell.font = Font(bold=True)
    for row in summary.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = GBP
    style_sheet(summary)
    summary.auto_filter.ref = None  # the totals row would get sorted in with the categories
    wb.save(path)


def compare_prices(books, history_path, report_path):
    """Compare prices with the previous run's snapshot, write the changes report,
    then update the snapshot. Books are matched by UPC (or link if UPC is missing)."""
    history_path = Path(history_path)
    previous = json.loads(history_path.read_text(encoding="utf-8")) if history_path.exists() else None

    current = {b["upc"] or b["url"]: {"title": b["title"], "price_gbp": b["price_gbp"], "url": b["url"]}
               for b in books}

    if previous is None:
        log.info("No previous run found (%s); price changes will be reported from the next run.", history_path)
    else:
        changes = [
            {"title": now["title"], "upc": key if key != now["url"] else "",
             "old_price_gbp": f"{previous[key]['price_gbp']:.2f}", "new_price_gbp": f"{now['price_gbp']:.2f}",
             "change_gbp": f"{now['price_gbp'] - previous[key]['price_gbp']:+.2f}", "url": now["url"]}
            for key, now in current.items()
            if key in previous and round(previous[key]["price_gbp"], 2) != round(now["price_gbp"], 2)
        ]
        with open(report_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=CHANGE_FIELDS)
            writer.writeheader()
            writer.writerows(changes)
        new = sum(k not in previous for k in current)
        log.info("Price changes since last run: %d (report: %s). New books: %d.", len(changes), report_path, new)

    # Merge, so a partial run (--pages) doesn't forget books it didn't visit.
    history_path.write_text(json.dumps({**(previous or {}), **current}, indent=1), encoding="utf-8")


class TqdmHandler(logging.StreamHandler):
    """Log to the screen without breaking the progress bar."""

    def emit(self, record):
        tqdm.write(self.format(record), file=sys.stderr)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-o", "--output", default="books",
                    help="output file name; the extension is added for you (default: books)")
    ap.add_argument("-f", "--format", choices=["csv", "xlsx", "both"], default="both",
                    help="file type(s) to write (default: both)")
    ap.add_argument("-p", "--pages", type=int, help="scrape only the first N pages (default: all)")
    ap.add_argument("--delay", type=float, default=0.5,
                    help="seconds each worker waits before a request (default: 0.5)")
    ap.add_argument("-w", "--workers", type=int, default=8,
                    help="pages downloaded at the same time (default: 8)")
    ap.add_argument("--cache-dir", default=".cache",
                    help="folder for cached book pages (default: .cache)")
    ap.add_argument("--no-cache", action="store_true",
                    help="ignore cached book pages and download everything again")
    ap.add_argument("--retries", type=int, default=3, help="retries per request before skipping it (default: 3)")
    ap.add_argument("--history", default="last_prices.json",
                    help="prices from the previous run, used to spot changes (default: last_prices.json)")
    ap.add_argument("--changes", default="price_changes.csv",
                    help="where to write the price-changes report (default: price_changes.csv)")
    ap.add_argument("--log-file", default="scrape.log", help="where to write the log (default: scrape.log)")
    args = ap.parse_args()

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)-7s %(message)s",
        handlers=[TqdmHandler(), logging.FileHandler(args.log_file, encoding="utf-8")],
    )

    start = time.monotonic()
    workers = max(1, args.workers)
    fetcher = Fetcher(make_session(args.retries, workers), args.delay, args.cache_dir, use_cache=not args.no_cache)
    if args.no_cache:
        log.info("--no-cache: downloading every page fresh")
    try:
        books, failed, failed_books = scrape(fetcher, args.pages, workers)
    except requests.RequestException as e:
        log.error("Could not load the first page, nothing to scrape: %s", e)
        return 1

    base = Path(args.output)
    base = base.with_suffix("") if base.suffix in (".csv", ".xlsx") else base
    if args.format in ("csv", "both"):
        save_csv(books, base.with_suffix(".csv"))
        log.info("Saved %s", base.with_suffix(".csv"))
    if args.format in ("xlsx", "both"):
        save_xlsx(books, base.with_suffix(".xlsx"))
        log.info("Saved %s", base.with_suffix(".xlsx"))

    compare_prices(books, args.history, args.changes)

    log.info("Done: %d books in %.1f s.", len(books), time.monotonic() - start)
    if failed:
        log.warning("%d page(s) failed and were skipped: %s", len(failed), failed)
    if failed_books:
        log.warning("%d detail page(s) failed; those books have blank stock/category/UPC.", failed_books)
    if failed or failed_books:
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
