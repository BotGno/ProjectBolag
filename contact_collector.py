#!/usr/bin/env python3
"""Opens a Chrome window and writes company name, phone and email from the
allabolag.se/foretag/ pages you visit into companies.xlsx. Browse normally; close the window to stop.

Setup (once):
    pip install playwright openpyxl
    playwright install chromium

Run:
    python3 contact_collector.py            # record what you browse
    python3 contact_collector.py --auto     # pick an industry, the script reads every company in it
    python3 contact_collector.py --auto --click   # same, but visibly clicks each company (slower)
    python3 contact_collector.py --lookup   # also check company websites for missing emails
"""
import argparse
import html
import random
import re
import time
from copy import copy
from pathlib import Path
from urllib.parse import parse_qsl, quote, unquote, urlencode, urljoin, urlparse

from openpyxl import load_workbook
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

XLSX = Path(__file__).with_name("companies.xlsx")
START_URL = "https://www.allabolag.se"
BRANCHES_URL = "https://www.allabolag.se/branscher"

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+46|0046|\b0)[\s\-()]*\d(?:[\s\-()]*\d){6,10}\b")

# Pages we never record from (search engines, social media)
SKIP_HOSTS = ("google.", "bing.com", "duckduckgo.com", "facebook.com", "instagram.com",
              "linkedin.com", "youtube.com", "twitter.com", "x.com", "tiktok.com")
# allabolag's sister sites and data sources, linked from every page - never a company's own website
NOT_COMPANY_SITES = ("allabolag", "proff.", "enento", "asiakastieto", "uc.se", "bolagsverket", "scb.se",
                     "skatteverket", "apple.com", "play.google")
BAD_EMAIL_DOMAINS = ("example.com", "example.se", "sentry.io", "wixpress.com", "domain.com",
                     "allabolag.se", "proff.se", "uc.se")
CONTACT_LINK_RE = re.compile(r"kontakt|contact|om-oss|about", re.I)
MAX_PER_PAGE = 6  # more contacts than this = a list page, skip it
EMPLOYEES_RE = re.compile(r"Antal anställda\s*(\d[\d\s]*)")
MIN_EMPLOYEES = 300  # companies with fewer employees (or no figure shown) are not saved; --min-employees

READ_PAGE_JS = """() => ({
    h1: (document.querySelector('h1')?.innerText || '').trim(),
    siteName: document.querySelector('meta[property="og:site_name"]')?.content || '',
    title: document.title || '',
    text: document.body ? document.body.innerText : '',
    mailto: [...document.querySelectorAll('a[href^="mailto:"]')].map(a => a.getAttribute('href')),
    tel: [...document.querySelectorAll('a[href^="tel:"]')].map(a => a.getAttribute('href')),
    links: [...document.querySelectorAll('a[href]')].map(a => ({href: a.href, text: a.innerText.trim()})),
})"""


def norm(s):
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


LEGAL_FORMS_RE = re.compile(r"\b(aktiebolag|ab|publ|handelsbolag|hb|kommanditbolag|kb|ekonomisk förening)\b")


def company_key(name):
    """Name without legal form and punctuation, so "Foo AB" and "Foo Aktiebolag (publ)" match."""
    return re.sub(r"\s+", " ", re.sub(r"[^\w\s]", " ", LEGAL_FORMS_RE.sub("", norm(name)))).strip()


def phone_digits(phone):
    d = re.sub(r"\D", "", phone)
    if d.startswith("0046"):
        d = "0" + d[4:]
    elif d.startswith("46"):
        d = "0" + d[2:]
    return d


def valid_phone(phone):
    d = phone_digits(phone)
    return d.startswith("0") and not d.startswith("00") and 8 <= len(d) <= 10


def valid_email(email):
    e = email.lower()
    return (not e.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp"))
            and not e.split("@")[1].endswith(BAD_EMAIL_DOMAINS))


def unique(items, key):
    seen, out = set(), []
    for item in items:
        k = key(item)
        if k not in seen:
            seen.add(k)
            out.append(item)
    return out


def extract_contacts(data):
    """Emails and phones on the page; mailto:/tel: links first since they are most reliable."""
    emails = [m.split(":", 1)[1].split("?")[0].strip() for m in data["mailto"]]
    emails += EMAIL_RE.findall(data["text"])
    emails = unique([e for e in emails if EMAIL_RE.fullmatch(e) and valid_email(e)], str.lower)

    phones = [t.split(":", 1)[1].strip() for t in data["tel"]]
    phones += [m.group(0) for m in PHONE_RE.finditer(data["text"])]
    phones = unique([re.sub(r"\s+", " ", p).strip() for p in phones if valid_phone(p)], phone_digits)
    return emails, phones



def find_website(links):
    """The company's own website link on an allabolag company page."""
    for link in links:
        host = urlparse(link["href"]).hostname or ""
        if not host or any(s in host for s in SKIP_HOSTS + NOT_COMPANY_SITES):
            continue
        text = link["text"].lower()
        if "hemsida" in text or "webbplats" in text or text.startswith("www.") \
                or text == host.removeprefix("www."):
            return link["href"]
    return None


class Sheet:
    def __init__(self, path):
        self.path = path
        self.wb = load_workbook(path)
        self.ws = self.wb.active
        self.dirty = False
        # Drop the example row from the template
        if norm(self.ws["A2"].value) == "example ab":
            for c in "ABC":
                self.ws[f"{c}2"].value = None
            self.dirty = True
        self.last = 1
        while any(self.ws.cell(self.last + 1, c).value for c in (1, 2, 3)):
            self.last += 1
        self.remove_duplicates()

    def rows(self):
        return range(2, self.last + 1)

    def index(self, r):
        """Remember row r under its name, phone and email so find() needs no scan."""
        name, phone, email = (self.ws.cell(r, i).value for i in (1, 2, 3))
        if name:
            self.by_name.setdefault(company_key(name), r)
        if phone:
            self.by_phone.setdefault(phone_digits(str(phone)), r)
        if email:
            self.by_email.setdefault(norm(email), r)

    def reindex(self):
        self.by_name, self.by_phone, self.by_email = {}, {}, {}
        for r in self.rows():
            self.index(r)

    def find(self, name, phone, email):
        return (name and self.by_name.get(company_key(name))) \
            or (email and self.by_email.get(norm(email))) \
            or (phone and self.by_phone.get(phone_digits(phone))) \
            or None

    def remove_duplicates(self):
        """Merge rows for the same company (same name, phone or email) into the first one."""
        self.by_name, self.by_phone, self.by_email = {}, {}, {}
        keep = []
        for r in self.rows():
            name, phone, email, contacted = (self.ws.cell(r, i).value for i in range(1, 5))
            first = self.find(name, str(phone or ""), str(email or ""))
            if first is None:
                keep.append(r)
                self.index(r)
                continue
            for c, v in ((2, phone), (3, email)):
                if v and not self.ws.cell(first, c).value:
                    self.ws.cell(first, c).value = v
            if contacted == "Yes":
                self.ws.cell(first, 4).value = "Yes"
            self.index(first)
            print(f"  merged duplicate: {name}")
        rows = list(self.rows())
        if len(keep) == len(rows):
            return
        # Move the kept rows up so there are no gaps, then clear the leftover rows
        values = [[self.ws.cell(r, c).value for c in range(1, 5)] for r in keep]
        for i, r in enumerate(rows):
            for c in range(1, 5):
                self.ws.cell(r, c).value = values[i][c - 1] if i < len(values) else None
        self.last = len(keep) + 1
        self.reindex()
        self.dirty = True

    def has_email(self, name):
        r = self.find(name, "", "")
        return bool(r and self.ws.cell(r, 3).value)

    def format_row(self, r):
        """Rows past the template's pre-formatted ones get the same look.
        The Yes/No dropdown and colours already cover the whole D column (D2:D100000)."""
        if self.ws.row_dimensions[r].height:
            return
        self.ws.row_dimensions[r].height = 24
        for c in range(1, 5):
            src, dst = self.ws.cell(2, c), self.ws.cell(r, c)
            dst.font, dst.border, dst.alignment = copy(src.font), copy(src.border), copy(src.alignment)
        self.ws.auto_filter.ref = f"A1:D{r}"

    def add(self, name, phone, email):
        r = self.find(name, phone, email)
        if r:
            changed = False
            if phone and not self.ws.cell(r, 2).value:
                self.ws.cell(r, 2).value, changed = phone, True
            if email and not self.ws.cell(r, 3).value:
                self.ws.cell(r, 3).value, changed = email, True
            if changed:
                self.index(r)
            self.dirty |= changed
            return "updated" if changed else None
        self.last += 1
        r = self.last
        self.format_row(r)
        for c, v in zip("ABCD", (name, phone, email, "No")):
            self.ws[f"{c}{r}"].value = v
        self.index(r)
        self.dirty = True
        return "added"

    def save(self):
        if not self.dirty:
            return
        try:
            self.wb.save(self.path)
            self.dirty = False
        except PermissionError:
            print("  ! Couldn't save - close companies.xlsx in Excel/LibreOffice. Retrying...")


def lookup_website(browser, url):
    """Visit a company's website (home + contact page) and return emails and phones found."""
    emails, phones = [], []
    page = browser.new_page()
    try:
        page.goto(url, timeout=15000, wait_until="domcontentloaded")
        data = page.evaluate(READ_PAGE_JS)
        e, p = extract_contacts(data)
        emails += e
        phones += p
        host = urlparse(page.url).hostname
        contact_pages = unique([l["href"].split("#")[0] for l in data["links"]
                                if urlparse(l["href"]).hostname == host
                                and CONTACT_LINK_RE.search(l["href"] + " " + l["text"])], str)
        for href in contact_pages[:2]:
            time.sleep(1)
            page.goto(urljoin(url, href), timeout=15000, wait_until="domcontentloaded")
            e, p = extract_contacts(page.evaluate(READ_PAGE_JS))
            emails += e
            phones += p
    except PlaywrightError as err:
        print(f"  ! Couldn't read {url}: {err.message.splitlines()[0]}")
    finally:
        page.close()
    return unique(emails, str.lower), unique(phones, phone_digits)


def is_company_page(url):
    parsed = urlparse(url)
    return (parsed.hostname or "") in ("allabolag.se", "www.allabolag.se") \
        and unquote(parsed.path).startswith(("/foretag/", "/företag/"))


def is_branch_list(url):
    parsed = urlparse(url)
    return (parsed.hostname or "") in ("allabolag.se", "www.allabolag.se") \
        and unquote(parsed.path).startswith("/bransch-sök")


def pause(page, low=1.0, high=2.5):
    """Random wait so the clicking looks like a person, not a flood of requests."""
    page.wait_for_timeout(random.uniform(low, high) * 1000)


def wait_for_cookie_popup(page):
    """The cookie popup blocks clicks - wait until the user has answered it."""
    popup = page.locator("#qc-cmp2-container button").first
    if popup.is_visible():
        print("  Answer the cookie popup in the browser to continue...")
        popup.wait_for(state="hidden", timeout=0)


def list_companies(page):
    """(relative href, name, employees or None) of every company on a branch list page, in page order.
    Only the result headings - sole traders also link to the owner's personal page - and not
    the "Träffar i företagsnamn" box on page 1, which lists companies outside the branch."""
    return page.evaluate("""() => {
        const inNameHitsBox = a => {
            for (let el = a.parentElement; el && el !== document.body; el = el.parentElement)
                if (el.querySelector('a[href*="retagsnamn"]'))
                    return el.querySelectorAll('h2').length <= 10;
            return false;
        };
        const employees = h2 => {
            let card = h2;  // widest ancestor that still holds only this result
            while (card.parentElement && card.parentElement.querySelectorAll('h2').length === 1)
                card = card.parentElement;
            const m = card.innerText.match(/ANSTÄLLDA\\s*([\\d\\s\\u00a0]+)/i);
            return m ? parseInt(m[1].replace(/\\D/g, ''), 10) : null;
        };
        const seen = new Set(), out = [];
        for (const a of document.querySelectorAll('h2 a[href^="/foretag/"]')) {
            if (inNameHitsBox(a)) continue;
            const href = a.getAttribute('href'), name = a.innerText.trim();
            if (name && !seen.has(href)) { seen.add(href); out.push([href, name, employees(a.closest('h2'))]); }
        }
        return out;
    }""")


def sorted_by_employees(url):
    """The same branch list, biggest companies first."""
    parsed = urlparse(url)
    query = [(k, v) for k, v in parse_qsl(parsed.query) if k != "sort"] + [("sort", "numEmployeesDesc")]
    return parsed._replace(query=urlencode(query, quote_via=quote)).geturl()


def read_html(context, url):
    """Fetch a page without showing it and return the same data READ_PAGE_JS gives."""
    body = context.request.get(url, timeout=15000).text()
    body = re.sub(r"<(script|style)\b.*?</\1>", " ", body, flags=re.S | re.I)
    links = [{"href": urljoin(url, html.unescape(href)),
              "text": html.unescape(re.sub(r"<[^>]+>", "", text)).strip()}
             for href, text in re.findall(r'<a\b[^>]*?href="([^"]*)"[^>]*>(.*?)</a>', body, re.S)]
    h1 = re.search(r"<h1\b[^>]*>(.*?)</h1>", body, re.S)
    return {
        "h1": html.unescape(re.sub(r"<[^>]+>", "", h1.group(1))).strip() if h1 else "",
        "text": html.unescape(re.sub(r"<[^>]+>", "\n", body)),
        "mailto": [l["href"] for l in links if l["href"].startswith("mailto:")],
        "tel": [l["href"] for l in links if l["href"].startswith("tel:")],
        "links": links,
    }


def open_company(page, href, sheet, helper, looked_up):
    """Visibly click a company on the list, save its contacts, go back to the list."""
    list_url = page.url
    wait_for_cookie_popup(page)
    page.locator(f'a[href="{href}"]').first.click()
    try:
        page.wait_for_url(lambda u: is_company_page(u), timeout=15000)
        page.wait_for_selector("h1", timeout=15000)
        process(page, sheet, helper, looked_up)
    finally:
        page.go_back()
        page.wait_for_url(list_url, timeout=15000)
        page.wait_for_selector('a[href^="/foretag/"]', timeout=15000)


def crawl_branch(page, sheet, helper, looked_up, visited, done_lists, click=False):
    """Save the contacts of every company on the list, then go to the next page.
    With an employee minimum, the list is sorted biggest first and stops at the first smaller one."""
    if MIN_EMPLOYEES and "sort=numEmployeesDesc" not in page.url:
        page.goto(sorted_by_employees(page.url), wait_until="domcontentloaded")
    list_url = page.url
    page_no = 1
    while True:
        page.wait_for_selector('a[href^="/foretag/"]', timeout=15000)
        companies = list_companies(page)
        print(f"\n== {unquote(list_url)} ({len(companies)} companies)")
        for href, name, employees in companies:
            if page.url != list_url:
                print("  (you left the list - stopping this branch)")
                return
            if MIN_EMPLOYEES and employees is not None and employees < MIN_EMPLOYEES:
                sheet.save()
                done_lists.add(list_url)
                print(f"\nReached companies under {MIN_EMPLOYEES} employees ({name}: {employees}) - branch done."
                      "\nPick another branch or close the window.")
                return
            if href in visited or sheet.has_email(name):
                continue
            visited.add(href)
            try:
                if click:
                    open_company(page, href, sheet, helper, looked_up)
                else:
                    data = read_html(page.context, urljoin(list_url, href))
                    data["h1"] = data["h1"] or name
                    record(data, sheet, helper, looked_up)
                    pause(page, 0.2, 0.5)
            except PlaywrightError:
                print(f"  ! {name} didn't load, skipping")
        sheet.save()

        next_link = page.locator('a[aria-label="Go to next page"]')
        if not next_link.count():
            done_lists.add(list_url)
            print(f"\nBranch done ({page_no} pages). Pick another branch or close the window.")
            return
        page.goto(urljoin(list_url, next_link.get_attribute("href")), wait_until="domcontentloaded")
        list_url = page.url
        page_no += 1


def process(page, sheet, helper, looked_up):
    if is_company_page(page.url):
        record(page.evaluate(READ_PAGE_JS), sheet, helper, looked_up)


def employee_count(text):
    """The "Antal anställda" figure on a company page, or None if the page doesn't show one."""
    m = EMPLOYEES_RE.search(text)
    return int(re.sub(r"\D", "", m.group(1))) if m else None


def record(data, sheet, helper, looked_up):
    """Save the name, phone and email from one company page's data."""
    emails, phones = extract_contacts(data)
    if len(emails) > MAX_PER_PAGE or len(phones) > MAX_PER_PAGE:
        return
    name = data["h1"]
    if not name:
        return
    employees = employee_count(data["text"])
    if MIN_EMPLOYEES and (employees is None or employees < MIN_EMPLOYEES):
        return

    if helper and not emails and not sheet.has_email(name):
        site = find_website(data["links"])
        if site and site not in looked_up:
            looked_up.add(site)
            print(f"  looking up {site} for {name} ...")
            e, p = lookup_website(helper, site)
            emails += e
            phones += [x for x in p if phone_digits(x) not in map(phone_digits, phones)]

    phone, email = (phones[0] if phones else ""), (emails[0] if emails else "")
    result = sheet.add(name, phone, email)
    if result:
        print(f"[{result}] {name} | {phone or '-'} | {email or '-'}")


def main():
    global MIN_EMPLOYEES
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--lookup", action="store_true",
                        help="visit company websites to find emails allabolag doesn't show")
    parser.add_argument("--auto", action="store_true",
                        help="open allabolag.se/branscher; when you click a branch, read every company in it")
    parser.add_argument("--click", action="store_true",
                        help="with --auto: visibly click into each company instead of reading it in the background")
    parser.add_argument("--min-employees", type=int, default=MIN_EMPLOYEES, metavar="N",
                        help=f"only save companies with at least N employees (default {MIN_EMPLOYEES}, 0 = save all)")
    args = parser.parse_args()
    MIN_EMPLOYEES = args.min_employees

    if not XLSX.exists():
        raise SystemExit(f"{XLSX} not found - create it first.")
    sheet = Sheet(XLSX)
    looked_up = set()
    visited = set()  # company links already opened in auto mode
    done_lists = set()  # branch list pages already crawled

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(locale="sv-SE", viewport=None)
        context.new_page().goto(BRANCHES_URL if args.auto else START_URL)
        helper = p.chromium.launch() if args.lookup else None
        if args.auto:
            print("Click a branch in the Chrome window - the script then opens every company in it.")
        else:
            print("Browse in the Chrome window. Contacts are saved to companies.xlsx.")
        print("Close the window (or press Ctrl+C here) to stop.\n")

        try:
            while browser.is_connected() and context.pages:
                for page in list(context.pages):
                    try:
                        if args.auto and is_branch_list(page.url) and page.url not in done_lists:
                            crawl_branch(page, sheet, helper, looked_up, visited, done_lists, args.click)
                        else:
                            process(page, sheet, helper, looked_up)
                    except PlaywrightError:
                        pass  # page was navigating or closed mid-read
                sheet.save()
                try:
                    context.pages[-1].wait_for_timeout(1500)
                except (PlaywrightError, IndexError):
                    time.sleep(0.5)
        except KeyboardInterrupt:
            pass
        finally:
            sheet.save()
            if helper:
                helper.close()
            if browser.is_connected():
                browser.close()
    print("\nDone. Open companies.xlsx to see your list.")


if __name__ == "__main__":
    main()
