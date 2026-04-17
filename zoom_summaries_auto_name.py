#!/usr/bin/env python3
"""
Zoom Meeting Summaries Downloader
Downloads DOCX summaries from Zoom and renames them based on Topic and Date Created.

Modes:
  all   — download every summary on every list page (one full pass).
  daily — only download rows whose Date Created falls on a calendar day
          (local midnight through end of day); default day is today.
"""

import argparse
import os
import re
import sys
import tempfile
import time
from datetime import date, datetime
from pathlib import Path
from typing import Optional

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout


def log(msg: str):
    """Print with immediate flush for real-time output."""
    print(msg, flush=True)


def _safe_close_page(tab) -> None:
    if tab is None:
        return
    try:
        if not tab.is_closed():
            tab.close()
    except Exception:
        pass


def _running_on_aws_lambda() -> bool:
    return bool(os.environ.get("AWS_LAMBDA_FUNCTION_NAME") or os.environ.get("LAMBDA_TASK_ROOT"))


def _resolve_downloads_folder() -> Path:
    """
    Writable directory for exported DOCX files.

    ``ZOOM_DOWNLOAD_DIR`` overrides everything if set.

    On AWS Lambda, defaults to ``/tmp/zoom_downloads`` (``~/Downloads`` is usually missing).
    Otherwise tries ``~/Downloads``, then the process temp directory.
    """
    override = os.environ.get("ZOOM_DOWNLOAD_DIR", "").strip()
    if override:
        p = Path(override).expanduser()
        p.mkdir(parents=True, exist_ok=True)
        return p.resolve()
    if _running_on_aws_lambda():
        p = Path("/tmp/zoom_downloads")
        p.mkdir(parents=True, exist_ok=True)
        return p.resolve()
    primary = Path.home() / "Downloads"
    try:
        primary.mkdir(parents=True, exist_ok=True)
        return primary.resolve()
    except OSError:
        fallback = Path(tempfile.gettempdir()) / "zoom_agent_downloads"
        fallback.mkdir(parents=True, exist_ok=True)
        log(
            f"Note: could not use {primary}; saving to {fallback.resolve()}"
        )
        return fallback.resolve()


def _return_to_summaries_list(docs_tab, detail_tab, list_url: str):
    """
    Close the Zoom Docs tab when it was opened in a new window, then open the
    summaries list on the original tab. Prevents unbounded tab growth on large runs.
    """
    _safe_close_page(docs_tab)
    try:
        if detail_tab is not None and not detail_tab.is_closed():
            detail_tab.goto(list_url)
    except Exception:
        pass
    time.sleep(2)
    return detail_tab


def clean_filename(name: str, max_length: int = 150) -> str:
    r"""
    Clean a string to be safe for use as a filename.
    - Remove invalid characters: / \ : * ? " < > |
    - Replace spaces with underscores
    - Replace colons in time with dashes
    - Limit length
    """
    # Replace colons (common in times like 02:32) with dashes
    name = name.replace(":", "-")
    # Replace spaces with underscores
    name = name.replace(" ", "_")
    # Remove invalid filename characters
    invalid_chars = r'[/\\:*?"<>|]'
    name = re.sub(invalid_chars, "", name)
    # Remove any double underscores that might have been created
    name = re.sub(r"_+", "_", name)
    # Strip leading/trailing underscores
    name = name.strip("_")
    # Limit length
    if len(name) > max_length:
        name = name[:max_length]
    return name


def get_unique_filepath(filepath: Path) -> Path:
    """
    If filepath exists, add _2, _3, etc. until we find a unique name.
    """
    if not filepath.exists():
        return filepath
    
    stem = filepath.stem
    suffix = filepath.suffix
    parent = filepath.parent
    
    counter = 2
    while True:
        new_path = parent / f"{stem}_{counter}{suffix}"
        if not new_path.exists():
            return new_path
        counter += 1


def parse_summary_datetime(text: str) -> Optional[datetime]:
    """
    Parse Date Created strings from the Zoom summaries table into a naive datetime.
    Uses common Zoom / locale formats; returns None if parsing fails.
    """
    text = (text or "").strip()
    if not text:
        return None
    text = re.sub(r"\s+", " ", text)
    formats = (
        "%b %d, %Y %I:%M %p",
        "%b %d, %Y %I:%M%p",
        "%b %d, %Y",
        "%B %d, %Y %I:%M %p",
        "%B %d, %Y %I:%M%p",
        "%B %d, %Y",
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %I:%M%p",
        "%m/%d/%Y",
        "%d/%m/%Y %I:%M %p",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    )
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    # Month name + day + year, optional time after
    m = re.search(
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+"
        r"(\d{1,2}),\s*(\d{4})"
        r"(?:\s+(\d{1,2}):(\d{2})\s*(AM|PM))?",
        text,
        re.I,
    )
    if m:
        mon, d, y = m.group(1), int(m.group(2)), int(m.group(3))
        base = f"{mon[:3].title()} {d}, {y}"
        if m.group(4):
            base += f" {int(m.group(4)):d}:{m.group(5)} {m.group(6).upper()}"
            try:
                return datetime.strptime(base, "%b %d, %Y %I:%M %p")
            except ValueError:
                return None
        try:
            return datetime.strptime(base, "%b %d, %Y")
        except ValueError:
            return None
    return None


def download_zoom_summaries(mode: str = "all", target_date: Optional[date] = None):
    """
    Download Zoom meeting summaries.

    mode ``all``: every row on every page.
    mode ``daily``: only rows whose parsed Date Created is on ``target_date``
    (calendar day in local time). ``target_date`` is required when mode is daily.
    """
    if mode == "daily" and target_date is None:
        raise ValueError("target_date is required when mode is 'daily'")

    # Zoom credentials (set in environment — do not commit secrets)
    ZOOM_EMAIL = os.environ.get("ZOOM_EMAIL", "").strip()
    ZOOM_PASSWORD = os.environ.get("ZOOM_PASSWORD", "").strip()

    downloads_folder = _resolve_downloads_folder()

    log("=" * 60)
    log("Zoom Meeting Summaries Downloader")
    log("=" * 60)
    if mode == "all":
        log("Mode: ALL — download every summary on every page (one pass).")
    else:
        log(
            f"Mode: DAILY — only summaries with Date Created on "
            f"{target_date.isoformat()} (local calendar day, midnight–end of day)."
        )
    log(f"Downloads will be saved to: {downloads_folder}")
    if not ZOOM_EMAIL or not ZOOM_PASSWORD:
        log(
            "Note: ZOOM_EMAIL and ZOOM_PASSWORD are not both set in the environment — "
            "auto-login is off; complete sign-in in the browser if prompted."
        )
    log("")

    with sync_playwright() as p:
        # Launch browser in visible mode
        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        
        # Navigate to Zoom summaries page
        url = "https://us06web.zoom.us/user/meeting/summary#/list"
        log(f"Navigating to: {url}")
        log("")
        
        page.goto(url)
        
        # Dismiss cookie popup if present
        time.sleep(2)
        try:
            cookie_selectors = [
                "button:has-text('ACCEPT COOKIES')",
                "button:has-text('Accept')",
                "button:has-text('Accept All')",
                "[class*='cookie'] button",
                "#onetrust-accept-btn-handler",
            ]
            for selector in cookie_selectors:
                try:
                    cookie_btn = page.query_selector(selector)
                    if cookie_btn and cookie_btn.is_visible():
                        cookie_btn.click()
                        log("   Dismissed cookie popup")
                        time.sleep(1)
                        break
                except:
                    continue
        except:
            pass
        
        # Handle login if redirected to sign-in page
        time.sleep(2)
        current_url = page.url
        
        if "signin" in current_url or "login" in current_url:
            log("Login page detected. Attempting auto-login...")
            
            try:
                # Step 1: Enter email
                email_selectors = [
                    "#email",
                    "input[type='email']",
                    "input[name='email']",
                    "input[placeholder*='email' i]",
                    "input[autocomplete='email']",
                ]
                
                email_filled = False
                for selector in email_selectors:
                    try:
                        email_input = page.wait_for_selector(selector, timeout=5000)
                        if email_input:
                            email_input.fill(ZOOM_EMAIL)
                            email_filled = True
                            log(f"   Email entered: {ZOOM_EMAIL}")
                            break
                    except:
                        continue
                
                if not email_filled:
                    log("   Could not find email field. Please enter manually.")
                
                # Step 2: Click Next button (Zoom uses multi-step login)
                if email_filled:
                    next_selectors = [
                        "button:has-text('Next')",
                        "button:has-text('Continue')",
                        "button[type='submit']",
                        "#js_btn_login",
                        ".btn-primary",
                        "button[class*='next']",
                    ]
                    
                    for selector in next_selectors:
                        try:
                            next_btn = page.query_selector(selector)
                            if next_btn and next_btn.is_visible():
                                next_btn.click()
                                log("   Clicked Next button")
                                time.sleep(2)
                                break
                        except:
                            continue
                
                # Step 3: Wait for and fill password field
                password_selectors = [
                    "#password",
                    "input[type='password']",
                    "input[name='password']",
                    "input[placeholder*='password' i]",
                ]
                
                password_filled = False
                for selector in password_selectors:
                    try:
                        password_input = page.wait_for_selector(selector, timeout=10000)
                        if password_input:
                            password_input.fill(ZOOM_PASSWORD)
                            password_filled = True
                            log("   Password entered: ********")
                            break
                    except:
                        continue
                
                if not password_filled:
                    log("   Could not find password field. Please enter manually.")
                
                # Step 4: Click Sign In button
                if password_filled:
                    signin_selectors = [
                        "button:has-text('Sign In')",
                        "button:has-text('Log In')",
                        "button[type='submit']",
                        "#js_btn_login",
                        "#login-btn",
                        "input[type='submit']",
                        "[data-testid='login-button']",
                        "button[class*='submit']",
                    ]
                    
                    for selector in signin_selectors:
                        try:
                            signin_btn = page.query_selector(selector)
                            if signin_btn and signin_btn.is_visible():
                                signin_btn.click()
                                log("   Clicked Sign In button")
                                break
                        except:
                            continue
                    
                    # Wait for navigation after login
                    log("   Waiting for login to complete...")
                    time.sleep(5)
                    
                    # Check if we need to handle any additional prompts (2FA, etc.)
                    current_url = page.url
                    if "signin" in current_url or "login" in current_url:
                        log("\n>>> Login may require additional verification (2FA, captcha, etc.) <<<")
                        log(">>> Please complete the login manually <<<\n")
                
            except Exception as e:
                log(f"   Auto-login error: {e}")
                log("   Please log in manually.")
        
        # Wait for the table to load - wait for table rows to appear
        # Zoom uses various table structures, so we try multiple selectors
        table_selectors = [
            "table tbody tr",
            "[class*='summary'] table tr",
            "[class*='list'] table tr",
            ".zm-table tbody tr",
            "[data-testid*='table'] tr",
            "tr[class*='row']",
        ]
        
        log("Waiting for summaries table to load...")
        log("(This may take a moment after login)")
        
        # Wait up to 5 minutes for user to login and table to appear
        table_found = False
        rows_selector = None
        
        for selector in table_selectors:
            try:
                page.wait_for_selector(selector, timeout=300000)  # 5 minutes
                rows_selector = selector
                table_found = True
                log(f"Table found using selector: {selector}")
                break
            except PlaywrightTimeout:
                continue
        
        if not table_found:
            # Fallback: wait for any table row
            try:
                page.wait_for_selector("tr", timeout=60000)
                rows_selector = "tr"
                table_found = True
            except PlaywrightTimeout:
                log("ERROR: Could not find summaries table. Please ensure you're logged in.")
                browser.close()
                return
        
        total_downloaded_count = 0
        current_page_num = 1
        list_base_url = url  # Keep list URL for pagination

        while True:
            # Give the page a moment to fully render
            time.sleep(2)

            # Find all table rows on current page
            rows = page.query_selector_all(rows_selector)

            # Filter out header rows (usually first row or rows without download button)
            data_rows = []
            for row in rows:
                cells = row.query_selector_all("td")
                if len(cells) >= 2:
                    data_rows.append(row)

            if not data_rows:
                # Try alternative: look for rows with download icons
                all_rows = page.query_selector_all("tr")
                for row in all_rows:
                    download_elem = row.query_selector("button, [class*='download'], svg, [aria-label*='download' i]")
                    if download_elem:
                        data_rows.append(row)

            if len(data_rows) == 0:
                if current_page_num == 1:
                    log("No data rows found. The table might be empty or use a different structure.")
                    log("Taking a screenshot for debugging...")
                    page.screenshot(path=str(downloads_folder / "zoom_debug_screenshot.png"))
                    log(f"Screenshot saved to: {downloads_folder / 'zoom_debug_screenshot.png'}")
                break

            total_rows = len(data_rows)
            log(f"\n{'=' * 60}")
            log(f"PAGE {current_page_num} — {total_rows} summary rows")
            log("-" * 60)

            # Process each row on this page
            for idx in range(1, total_rows + 1):
                try:
                    # Re-query rows fresh each time (page may have reloaded)
                    rows = page.query_selector_all(rows_selector)
                    data_rows = [r for r in rows if len(r.query_selector_all("td")) >= 2]
                    
                    if idx > len(data_rows):
                        log(f"\n[{idx}/{total_rows}] Row no longer exists, skipping...")
                        continue
                    
                    row = data_rows[idx - 1]  # 0-indexed
                    log(f"\n[{idx}/{total_rows}] Processing row...")
                    
                    # Extract text from cells
                    cells = row.query_selector_all("td")
                    
                    # Try to get Topic (usually first meaningful cell)
                    topic = ""
                    date_created = ""
                    
                    if len(cells) >= 2:
                        # First column is usually checkbox, topic is in second column
                        # Try to find the topic cell (skip checkbox columns)
                        for i, cell in enumerate(cells[:3]):
                            cell_text = cell.inner_text().strip()
                            # Skip cells that look like checkboxes or empty
                            if cell_text.lower().startswith("select") or not cell_text or len(cell_text) < 3:
                                continue
                            # Skip cells that are just dates
                            if re.match(r'^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$', cell_text):
                                continue
                            # Found a good topic candidate
                            topic = cell_text
                            # Clean up: remove "Select" prefix if present
                            topic = re.sub(r'^Select\s*', '', topic, flags=re.I).strip()
                            break
                        
                        # Date is typically after the topic column
                        for cell in cells[1:5]:  # Check a few columns
                            cell_text = cell.inner_text().strip()
                            # Look for date-like patterns (contains numbers and common date separators)
                            if re.search(r'\d{4}[-/]\d{2}[-/]\d{2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4}', cell_text):
                                date_created = cell_text
                                break
                            # Also check for month names
                            if re.search(r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)', cell_text, re.I):
                                date_created = cell_text
                                break
                    
                    # If we couldn't parse cells, try getting all row text
                    if not topic:
                        row_text = row.inner_text().strip()
                        parts = row_text.split('\n')
                        if parts:
                            topic = parts[0].strip()
                    
                    # Skip disabled/placeholder rows
                    skip_keywords = ["disabled", "no summary", "no data", "empty"]
                    if any(kw in topic.lower() for kw in skip_keywords) or not topic:
                        log(f"   Skipping row (invalid/disabled): {topic[:40] if topic else 'empty'}")
                        continue
                    
                    log(f"   Topic: {topic[:50]}..." if len(topic) > 50 else f"   Topic: {topic}")
                    log(f"   Date: {date_created}")

                    if mode == "daily":
                        created_dt = parse_summary_datetime(date_created)
                        if created_dt is None:
                            created_dt = parse_summary_datetime(
                                re.sub(r"\s+", " ", row.inner_text())
                            )
                        if created_dt is None:
                            log(
                                "   Skipping (daily mode: could not parse Date Created)"
                            )
                            continue
                        if created_dt.date() != target_date:
                            log(
                                "   Skipping (daily mode: created "
                                f"{created_dt.date().isoformat()}, "
                                f"target {target_date.isoformat()})"
                            )
                            continue

                    # Click on topic link to go to detail page
                    # The topic link might be a button or anchor with class 'topic-link'
                    detail_page_opened = False
                    original_url = page.url
                    
                    try:
                        # Try multiple selectors for the topic link
                        topic_link_selectors = [
                            ".topic-link",
                            "[class*='topic-link']",
                            "button[class*='topic']",
                            "td:nth-child(2) button",
                            "td:nth-child(2) a",
                            "td button[class*='link']",
                        ]
                        
                        for selector in topic_link_selectors:
                            topic_link = row.query_selector(selector)
                            if topic_link:
                                # Use JavaScript click which is more reliable
                                page.evaluate("el => el.click()", topic_link)
                                time.sleep(3)
                                
                                # Check if we navigated to a new page
                                if page.url != original_url:
                                    detail_page_opened = True
                                    log(f"   Opened detail page: {page.url[:60]}...")
                                    break
                    except Exception as e:
                        log(f"   DEBUG: Link click error: {e}")
                    
                    if not detail_page_opened:
                        log(f"   ✗ Could not open detail page for this row")
                        continue
                    
                    # Dismiss any cookie popups on detail page
                    try:
                        cookie_btn = page.query_selector("button:has-text('ACCEPT COOKIES')")
                        if cookie_btn and cookie_btn.is_visible():
                            cookie_btn.click()
                            time.sleep(0.5)
                    except:
                        pass
                    
                    # Step 1: Open in Docs — keep one list/detail tab; close a new Docs tab when done
                    detail_tab = page
                    docs_tab = None
                    open_docs_clicked = False
                    open_docs_selectors = [
                        "button:has-text('Open in Docs')",
                        "span:has-text('Open in Docs')",
                        "[class*='open-in-docs']",
                        "button:has-text('Open')",
                    ]

                    for selector in open_docs_selectors:
                        try:
                            open_docs_btn = detail_tab.query_selector(selector)
                            if not open_docs_btn or not open_docs_btn.is_visible():
                                continue
                            try:
                                with context.expect_page(timeout=15000) as new_page_info:
                                    detail_tab.evaluate("el => el.click()", open_docs_btn)
                                docs_tab = new_page_info.value
                                docs_tab.wait_for_load_state("domcontentloaded")
                                page = docs_tab
                                log("   Opened Docs in new tab")
                                open_docs_clicked = True
                                break
                            except PlaywrightTimeout:
                                time.sleep(1)
                                if "docs.zoom" in (detail_tab.url or "").lower():
                                    page = detail_tab
                                    log("   Opened Docs (same tab)")
                                    open_docs_clicked = True
                                    break
                        except Exception:
                            continue

                    if not open_docs_clicked:
                        for selector in open_docs_selectors:
                            try:
                                open_docs_btn = detail_tab.query_selector(selector)
                                if open_docs_btn and open_docs_btn.is_visible():
                                    open_docs_btn.click()
                                    time.sleep(3)
                                    page = detail_tab
                                    open_docs_clicked = True
                                    log("   Clicked 'Open in Docs' (fallback)")
                                    break
                            except Exception:
                                continue

                    if not open_docs_clicked:
                        log("   ✗ Could not find 'Open in Docs' button")
                        page = _return_to_summaries_list(None, detail_tab, original_url)
                        continue
                    
                    # Step 2: On Docs page, click the 3-dot menu (...)
                    time.sleep(2)  # Wait for docs page to load
                    
                    three_dot_clicked = False
                    three_dot_selectors = [
                        "button[aria-label*='more' i]",
                        "button[aria-label*='menu' i]",
                        "button:has-text('...')",
                        "[class*='more-menu']",
                        "[class*='dropdown'] button",
                        "button[class*='icon-more']",
                        "[aria-haspopup='true']",
                        "button:has(svg)",  # Often icon-only buttons
                    ]
                    
                    # Look in the top-right area for the menu button
                    for selector in three_dot_selectors:
                        try:
                            menu_btns = page.query_selector_all(selector)
                            for menu_btn in menu_btns:
                                if menu_btn.is_visible():
                                    menu_btn.click()
                                    time.sleep(1)
                                    # Check if a menu appeared
                                    export_visible = page.query_selector("text=Export")
                                    if export_visible:
                                        three_dot_clicked = True
                                        log(f"   Clicked 3-dot menu")
                                        break
                            if three_dot_clicked:
                                break
                        except:
                            continue
                    
                    if not three_dot_clicked:
                        log("   ✗ Could not find 3-dot menu")
                        page = _return_to_summaries_list(docs_tab, detail_tab, original_url)
                        continue
                    
                    # Step 3: Click "Export" in the menu
                    export_clicked = False
                    try:
                        export_btn = page.query_selector("text=Export")
                        if export_btn and export_btn.is_visible():
                            export_btn.hover()
                            time.sleep(0.5)
                            export_btn.click()
                            time.sleep(1)
                            export_clicked = True
                            log(f"   Clicked 'Export'")
                    except:
                        pass
                    
                    if not export_clicked:
                        log("   ✗ Could not click Export")
                        page = _return_to_summaries_list(docs_tab, detail_tab, original_url)
                        continue
                    
                    # Step 4: Select "Word" from the export submenu
                    download_clicked = False
                    try:
                        word_btn = page.query_selector("text=Word")
                        if word_btn and word_btn.is_visible():
                            # Start waiting for download before clicking
                            with page.expect_download(timeout=30000) as download_info:
                                word_btn.click()
                            
                            download = download_info.value
                            download_clicked = True
                            
                            # Get the suggested filename to preserve extension
                            suggested_name = download.suggested_filename
                            extension = Path(suggested_name).suffix or ".docx"
                            
                            # Build the new filename
                            if topic and date_created:
                                base_name = f"{topic}_{date_created}"
                            elif topic:
                                base_name = topic
                            else:
                                base_name = suggested_name.replace(extension, "")
                            
                            clean_name = clean_filename(base_name)
                            new_filename = f"{clean_name}{extension}"
                            
                            # Get unique filepath
                            output_path = get_unique_filepath(downloads_folder / new_filename)
                            
                            # Save the download
                            download.save_as(output_path)
                            
                            log(f"   ✓ Saved: {output_path.name}")
                            total_downloaded_count += 1
                    except PlaywrightTimeout:
                        log(f"   ✗ Download timed out")
                    except Exception as e:
                        log(f"   ✗ Error during Word export: {e}")
                    
                    if not download_clicked:
                        log("   ✗ Could not complete Word export")

                    # Close Docs tab if it was extra; always return list on the main tab
                    page = _return_to_summaries_list(docs_tab, detail_tab, original_url)
                    
                    # Re-dismiss cookie popup if it reappears
                    try:
                        cookie_btn = page.query_selector("button:has-text('ACCEPT COOKIES')")
                        if cookie_btn and cookie_btn.is_visible():
                            cookie_btn.click()
                            time.sleep(0.5)
                    except:
                        pass
                    
                    # Small delay between downloads to be respectful
                    time.sleep(1)
                    
                except Exception as e:
                    log(f"   ✗ Error processing row: {e}")
                    continue

            # After processing all rows on this page, try to go to next page
            next_btn = None
            # Prefer explicit "next" / right-arrow (avoid clicking page number)
            next_selectors = [
                "button[aria-label*='next' i]",
                "button[aria-label*='Next' i]",
                "[aria-label*='next' i]",
                "button[title*='next' i]",
                "[class*='pagination'] button[aria-label*='next' i]",
                "[class*='pagination'] [class*='next']:not([disabled])",
            ]
            for sel in next_selectors:
                try:
                    btn = page.query_selector(sel)
                    if btn and btn.is_visible() and btn.get_attribute("disabled") is None:
                        next_btn = btn
                        break
                except Exception:
                    continue
            # Fallback: last non-disabled button in pagination (usually "next" arrow)
            if not next_btn:
                try:
                    pagination = page.query_selector("[class*='pagination'], nav[aria-label*='pagination' i], [role='navigation']")
                    if pagination:
                        all_btns = pagination.query_selector_all("button:not([disabled])")
                        if len(all_btns) >= 2:
                            next_btn = all_btns[-1]  # last = next
                except Exception:
                    pass

            if next_btn:
                try:
                    next_btn.click()
                    log(f"\n>>> Going to next page...")
                    time.sleep(2)
                    current_page_num += 1
                except Exception as e:
                    log(f"   Could not click next: {e}")
                    break
            else:
                log(f"\n>>> No next page (last page reached).")
                break

        log("\n" + "=" * 60)
        log(f"Download complete! {total_downloaded_count} files saved (all pages).")
        log(f"Files saved to: {downloads_folder}")
        log("=" * 60)
        
        try:
            input("\nPress Enter to close the browser...")
        except EOFError:
            log("Running in non-interactive mode, closing browser automatically...")
            time.sleep(3)
        browser.close()


def _parse_cli() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Download Zoom meeting summaries as Word documents."
    )
    p.add_argument(
        "--mode",
        choices=("all", "daily"),
        default="all",
        help="all: every summary on every page once. daily: only summaries whose "
        "Date Created falls on the chosen calendar day (local date).",
    )
    p.add_argument(
        "--date",
        metavar="YYYY-MM-DD",
        dest="target_date",
        help="Calendar day for --mode daily (default: today on this machine).",
    )
    return p.parse_args()


if __name__ == "__main__":
    args = _parse_cli()
    td: Optional[date] = None
    if args.mode == "daily":
        td = date.fromisoformat(args.target_date) if args.target_date else date.today()
    download_zoom_summaries(mode=args.mode, target_date=td)
