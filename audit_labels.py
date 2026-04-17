"""
Audit utility: visit every Palisades Fire update page and dump a CSV of all
unique labels/headings seen, flagging which ones are currently NOT mapped
(and therefore would or should become new columns per the extra-credit rule).

Outputs:
  - audit_labels.csv : (label, kind, mapped_column, occurrences, first_source_url)
    where kind is one of dt / heading / inline_candidate.

Run:  python3 audit_labels.py
"""
from __future__ import annotations

import csv
import re
from collections import Counter, defaultdict
from urllib.parse import urljoin

from bs4 import BeautifulSoup, Tag
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException


BASE_URL = "https://www.fire.ca.gov"
UPDATES_URL = f"{BASE_URL}/incidents/2025/1/7/palisades-fire/updates"
INITIAL_WAIT = 3
REFRESH_WAIT = 20


def normalize_label(value: str) -> str:
    cleaned = value.replace("\xa0", " ").strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return cleaned.rstrip(":")


def wait_for(driver, css: str) -> None:
    try:
        WebDriverWait(driver, INITIAL_WAIT).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, css))
        )
        return
    except TimeoutException:
        driver.refresh()
    WebDriverWait(driver, REFRESH_WAIT).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, css))
    )


def load_mapping() -> tuple[set[str], set[str], set[str]]:
    """Return (label_keys_lower, inline_keys_lower, ignored_lower)."""
    import importlib.util
    spec = importlib.util.spec_from_file_location("project3_ai", "project3-ai.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)  # type: ignore
    label_keys = {normalize_label(k).lower() for k in module.LABEL_MAPPING.keys()}
    inline_keys = set(module.INLINE_SECTION_MAPPING.keys())
    ignored = set(module.IGNORED_UNMAPPED_SECTIONS)
    return label_keys, inline_keys, ignored


def get_update_links(driver) -> list[str]:
    driver.get(UPDATES_URL)
    wait_for(driver, "div.detail-page")
    soup = BeautifulSoup(driver.page_source, "html.parser")
    detail = soup.find("div", class_="detail-page")
    if detail is None:
        return []
    out: list[str] = []
    seen: set[str] = set()
    for a in detail.find_all("a", href=True):
        href = a["href"].strip()
        if not href:
            continue
        full = urljoin(BASE_URL, href)
        if full in seen:
            continue
        seen.add(full)
        if "/palisades-fire/updates/" not in full:
            continue
        out.append(full)
    return out


def collect_labels(driver, url: str) -> list[tuple[str, str]]:
    """Return list of (label, kind) tuples for one report page."""
    driver.get(url)
    wait_for(driver, "main.main-content")
    soup = BeautifulSoup(driver.page_source, "html.parser")
    content = soup.find("main", class_="main-content")
    if content is None:
        return []
    for hidden in content.select(".visually-hidden"):
        hidden.decompose()

    results: list[tuple[str, str]] = []

    for dt in content.find_all("dt"):
        text = dt.get_text(" ", strip=True)
        if text:
            results.append((normalize_label(text), "dt"))

    for section in content.find_all("div", class_="p-3"):
        for h in section.find_all(["h2", "h3", "h4", "h5", "h6"]):
            text = h.get_text(" ", strip=True)
            if text:
                results.append((normalize_label(text), f"heading-{h.name}"))

    # Heading-like lines INSIDE sections: <strong>/<b> at start of <p>, or label:
    for section in content.find_all("div", class_="p-3"):
        for p in section.find_all("p"):
            first = p.find(["strong", "b"])
            if first is not None:
                text = first.get_text(" ", strip=True)
                if text and len(text) < 80:
                    results.append((normalize_label(text), "inline-strong"))
            # "Label:" pattern at very start of paragraph text
            raw = p.get_text(" ", strip=True)
            m = re.match(r"^([A-Z][^:\n]{2,60}):", raw)
            if m:
                results.append((normalize_label(m.group(1)), "inline-colon"))

    return results


def main() -> None:
    label_keys, inline_keys, ignored = load_mapping()

    options = Options()
    driver = webdriver.Chrome(options=options)
    try:
        links = get_update_links(driver)
        print(f"Found {len(links)} update links.")
        counts: Counter[tuple[str, str]] = Counter()
        first_seen: dict[tuple[str, str], str] = {}

        for idx, link in enumerate(links, 1):
            try:
                items = collect_labels(driver, link)
            except Exception as exc:  # noqa
                print(f"[{idx}/{len(links)}] error {link}: {exc}")
                continue
            for key in items:
                counts[key] += 1
                first_seen.setdefault(key, link)
            if idx % 25 == 0 or idx == len(links):
                print(f"  scanned {idx}/{len(links)}")

    finally:
        driver.quit()

    rows = []
    for (label, kind), n in sorted(counts.items(), key=lambda x: (-x[1], x[0][0])):
        lc = label.lower()
        if lc in label_keys:
            status = "mapped"
        elif lc in inline_keys:
            status = "inline-mapped"
        elif lc in ignored:
            status = "explicitly-ignored"
        else:
            status = "UNMAPPED"
        rows.append({
            "label": label,
            "kind": kind,
            "status": status,
            "occurrences": n,
            "first_source_url": first_seen[(label, kind)],
        })

    with open("audit_labels.csv", "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=["label", "kind", "status", "occurrences", "first_source_url"])
        writer.writeheader()
        writer.writerows(rows)

    unmapped = [r for r in rows if r["status"] == "UNMAPPED"]
    ignored_rows = [r for r in rows if r["status"] == "explicitly-ignored"]
    print("\n=== summary ===")
    print(f"total unique (label, kind) entries: {len(rows)}")
    print(f"mapped:              {sum(1 for r in rows if r['status']=='mapped')}")
    print(f"inline-mapped:       {sum(1 for r in rows if r['status']=='inline-mapped')}")
    print(f"explicitly-ignored:  {len(ignored_rows)}")
    print(f"UNMAPPED (candidates for new columns): {len(unmapped)}")
    print("\nTop 30 UNMAPPED:")
    for r in unmapped[:30]:
        print(f"  {r['occurrences']:4d}x  [{r['kind']}]  {r['label']}")
    print("\nExplicitly-ignored entries (for reference):")
    for r in ignored_rows:
        print(f"  {r['occurrences']:4d}x  [{r['kind']}]  {r['label']}")
    print(f"\nFull report saved to audit_labels.csv")


if __name__ == "__main__":
    main()
