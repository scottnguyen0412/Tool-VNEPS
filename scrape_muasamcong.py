import json
import time
import random
import pandas as pd
from playwright.sync_api import sync_playwright

import os

# Fix for PyInstaller: Tell Playwright to look for browsers in the system default location
# instead of looking inside the temporary _MEI folder.
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

def run(output_path=None, max_pages=None, ministry_filter="", search_keyword="", pause_event=None, stop_event=None):
    # Handle default arguments if not provided (CLI usage fallback)
    if output_path is None and max_pages is None:
        # User inputs for CLI mode
        print("--- Cấu hình Tool Scraping ---")
        output_path = input("Nhập tên file lưu (mặc định 'investors_data_detailed.xlsx'): ").strip()
        if not output_path:
            output_path = "investors_data_detailed.xlsx"
        
        limit_input = input("Nhập số trang muốn cào (nhập 'all' hoặc để trống để cào hết): ").strip()
        if not limit_input or limit_input.lower() == 'all':
            max_pages = float('inf')
        else:
            try:
                max_pages = int(limit_input)
            except ValueError:
                print("Đầu vào không hợp lệ. Mặc định cào tất cả.")
                max_pages = float('inf')

    # Ensure extension
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"
    
    if max_pages is None: # explicit None passed
         max_pages = float('inf')

    # Load existing data to avoid duplicates
    processed_items = set()
    all_data = []
    
    if os.path.exists(output_path):
        try:
            print(f"File '{output_path}' tồn tại. Đang đọc dữ liệu cũ để tránh trùng lặp...")
            existing_df = pd.read_excel(output_path)
            
            # Use 'Mã định danh' for checking if available, else 'Entity Name'
            check_col = "Mã định danh"
            if check_col not in existing_df.columns:
                 check_col = "Entity Name" # Fallback
            
            if check_col in existing_df.columns:
                processed_items = set(existing_df[check_col].dropna().astype(str).str.strip())
                # Load all fields to keep history
                all_data = existing_df.to_dict('records')
                
            print(f"Đã tải {len(processed_items)} nhà đầu tư đã cào trước đó (Check theo: {check_col}).")
        except Exception as e:
            print(f"Cảnh báo: Không đọc được file cũ ({e}). Sẽ tạo mới.")


    def check_status():
        # Check Stop
        if stop_event and stop_event.is_set():
            print(">>> STOP SIGNAL RECEIVED. Exiting...")
            raise InterruptedError("Process stopped by user.")

        # Check Pause
        if pause_event:
            if not pause_event.is_set():
                print(">>> PAUSED. Waiting for resume...")
                try:
                    pause_event.wait()
                    
                    # Re-check stop after waking up
                    if stop_event and stop_event.is_set():
                         raise InterruptedError("Process stopped by user.")
                         
                    print(">>> RESUMED.")
                    print("Checking connection after resume...")
                except Exception as e:
                    print(f"Pause error: {e}")
                    if isinstance(e, InterruptedError):
                        raise e

    def wait_for_internet(page):
        while True:
            try:
                is_online = page.evaluate("navigator.onLine")
                if is_online:
                    return
                print("Mất kết nối Internet! Đang chờ kết nối lại...")
            except:
                print("Lỗi khi kiểm tra kết nối. Đang thử lại...")
            time.sleep(5)

    with sync_playwright() as p:
        browser = None
        # Try using system browsers (Chrome/Edge) to avoid path issues in EXE
        try:
            print("Đang khởi động Google Chrome...")
            browser = p.chromium.launch(headless=False, channel="chrome")
        except Exception:
            try:
                print("Không tìm thấy Chrome, đang thử Microsoft Edge...")
                browser = p.chromium.launch(headless=False, channel="msedge")
            except Exception:
                print("Không tìm thấy Edge, đang thử trình duyệt tích hợp (có thể lỗi trong EXE)...")
                browser = p.chromium.launch(headless=False)
                
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={"width": 1280, "height": 720},
            ignore_https_errors=True
        )
        page = context.new_page()

        print("Navigating to page...")
        try:
            page.goto("https://muasamcong.mpi.gov.vn/web/guest/investors-approval-v2", timeout=60000)
        except:
             print("Lỗi tải trang ban đầu. Đang thử tải lại...")
             wait_for_internet(page)
             page.reload()

        
        # Wait for search bar
        search_input_selector = 'input[placeholder="Tìm kiếm chủ đầu tư"]'
        try:
            page.wait_for_selector(search_input_selector, timeout=30000)
            print("Found search bar.")
        except:
            print("Search bar not found. Vui lòng kiểm tra kết nối mạng.")
            wait_for_internet(page)
            # Try reload once
            page.reload()
            try:
                page.wait_for_selector(search_input_selector, timeout=30000)
            except:
                print("Vẫn không thấy search bar. Exiting.")
                browser.close()
                return

        # 1. Apply Ministry Filter (if provided)
        if ministry_filter:
            print(f"Applying Ministry Filter: {ministry_filter}...")
            try:
                # Open Filter Panel
                filter_btn = "button.content__title__searchbar__right"
                page.click(filter_btn)
                time.sleep(1)
                
                # Select Ministry
                sel_selector = 'select[placeholder="Chọn Bộ / ban ngành"]'
                page.wait_for_selector(sel_selector, timeout=5000)
                page.select_option(sel_selector, label=ministry_filter)
                
                # Click Apply
                apply_btn = page.locator("button.ant-btn-primary").filter(has_text="Áp dụng")
                if apply_btn.count() > 0:
                    apply_btn.click()
                else:
                    page.click("button.ant-btn-primary")
                
                print("Filter applied. Waiting for reload...")
                time.sleep(3)
                wait_for_internet(page)
            except Exception as e:
                print(f"Warning: Failed to apply ministry filter ({e}). Continuing...")

        # 2. Apply Search Keyword (if provided)
        # Even if filter is applied, we might want to narrow down by keyword
        if search_keyword:
             print(f"Triggering keyword search: '{search_keyword}'...")
             page.fill(search_input_selector, search_keyword)
             page.press(search_input_selector, "Enter")
        elif not ministry_filter:
             # Triggers empty search only if NO filter and NO keyword (initial state)
             # But if filter was applied, results are already loaded. 
             # However, to be safe and ensure "Search" event is fired if needed:
             pass 
             # Actually, if we applied filter, the list should update.
             # If we didn't apply filter (All mode), we need to trigger search to load initial list.
             print("Triggering empty search to load list...")
             page.fill(search_input_selector, "")
             page.press(search_input_selector, "Enter")
        
        # Wait for items
        item_selector = "h2.content__body__item__title"
        
        # Wait for items to appear
        item_selector = "h2.content__body__item__title"
        # Since we want to check ID, we should also look for the container or the title. 
        # The title selector is what we click on later.
        
        try:
            page.wait_for_selector(item_selector, timeout=30000)
            print("Items loaded.")
        except:
            print("No items found after search. Retrying...")
            wait_for_internet(page)
            # Robust Reload
            for _ in range(3):
                try:
                    page.reload(timeout=60000)
                    break 
                except Exception as e:
                    print(f"Reload failed: {e}. Retrying...")
                    wait_for_internet(page)
                    time.sleep(2)
            # Need to search again?
             # Trigger search again
            try:
                 page.fill(search_input_selector, "")
                 page.press(search_input_selector, "Enter")
                 page.wait_for_selector(item_selector, timeout=30000)
            except:
                 print("Failed to load empty list. Exiting.")
                 browser.close()
                 return

        page_num = 1
        
        while page_num <= max_pages:
            check_status() # Check Stop/Pause
            print(f"Processing Page {page_num}...")
            wait_for_internet(page)
            
            if max_pages != float('inf'):
                 print(f"(Target: {max_pages} pages)")
            
            # Re-query items to get count
            time.sleep(2) # Stabilize
            count = page.locator(item_selector).count()
            print(f"Found {count} items on this page.")
            
            if count == 0:
                print("No items found on this page. Ending.")
                break

            for i in range(count):
                check_status() # Check Stop/Pause (Item Level)
                print(f"  Scraping item {i+1}/{count}...")
                
                # Retry mechanism for stale elements
                retry = 0
                skipped = False
                while retry < 3:
                    try:
                        wait_for_internet(page)
                        # Locate item again
                        item_title_element = page.locator(item_selector).nth(i)
                        
                        # Find the ID for this item
                        # Strategy: item_title_element is h2. From h2, go up to common parent, then find h4 with ID
                        # Based on inspection: h2 and h4 are inside the same container.
                        # We can use locator chaining or xpath relative to the title
                        # But simpler: The list of IDs corresponds to the list of Titles in order usually.
                        # Proper way: Get parent then find h4.
                        # h2 class: content__body__item__title
                        # h4 class: content__body__item__heading__text
                        
                        # Using 'locator(..).locator(..)' finds descendants.
                        # We need sibling/cousin. 
                        # Let's select the card container first.
                        # Card container is likely .content__body__item based on inspection
                        card = page.locator(".content__body__item").nth(i)
                        
                        # Check for duplicate before clicking
                        item_id = ""
                        try:
                            # Try to get ID
                            id_el = card.locator("h4.content__body__item__heading__text")
                            if id_el.count() > 0:
                                item_id = id_el.first.inner_text().strip()
                        except:
                            pass
                        
                        # Fallback to Name if ID not found (unlikely)
                        if not item_id:
                             item_id = item_title_element.inner_text().strip() # Use Name as ID
                        
                        if item_id in processed_items:
                            print(f"    Skipping duplicate (ID/Name): {item_id}")
                            skipped = True
                            break 
                        
                        # Click the title to open
                        item_title_element.click(timeout=10000)
                        break
                    except Exception as e:
                        print(f"    Error clicking item: {e}. Retrying...")
                        wait_for_internet(page) # Check net
                        time.sleep(1)
                        retry += 1
                
                if skipped:
                    continue
                
                if retry == 3:
                     print("    Failed to click item. Skipping.")
                     continue

                # Wait for detail page
                back_btn_selector = "button.btn-back" 
                try:
                    page.wait_for_selector(back_btn_selector, timeout=10000)
                except:
                    print("    Detail page load failed. Trying to go back manually.")
                    wait_for_internet(page)
                    try:
                        page.go_back(timeout=60000, wait_until="domcontentloaded")
                    except Exception as e:
                        print(f"    Warning: go_back timed out or failed: {e}. Checking if we returned anyway...")
                    
                    # Wait for list to reappear, with retry
                    try:
                         page.wait_for_selector(item_selector, timeout=15000)
                    except:
                         print("    Timeout waiting for list. Reloading page...")
                         wait_for_internet(page)
                         page.reload()
                         page.wait_for_selector(item_selector, timeout=60000)
                    continue

                # Extract Detail Data
                entity_name = "Unknown"
                # Try multiple selectors for the name. Prioritize specific ones.
                name_selectors = [
                    ".content-body__header h5", 
                    "h5.content__body__heading", 
                    ".content-body__header", 
                    "h3.font-weight-bold", 
                    "h3", 
                    "h5", 
                    "h4", 
                    "h2", 
                    ".title"
                ]
                
                bad_names = ["Danh sách chủ đầu tư được phê duyệt", "EGP_v2.0", "Trang chủ", "Home", "Tra cứu", "Tình huống đấu thầu"]

                for sel in name_selectors:
                    elements = page.locator(sel)
                    count_sel = elements.count()
                    for k in range(count_sel):
                        el = elements.nth(k)
                        text = el.inner_text().strip()
                        
                        # Skip if part of sidebar/menu (often inside .card-header or mb-0)
                        # We can check class or parent class, but simpler to just check text for now.
                        # Advanced: Check if visible?
                        if not el.is_visible():
                             continue

                        # Check if text is valid and not a bad generic header
                        is_bad = False
                        for bad in bad_names:
                            if bad.lower() in text.lower():
                                is_bad = True
                                break
                        
                        if text and not is_bad:
                            entity_name = text
                            break
                    
                    if entity_name != "Unknown":
                        break
                
                item_data = {"Entity Name": entity_name}
                
                # Extract Fields
                # Try specific class first, found by inspection
                row_selector = ".infomation-course__content"
                if page.locator(row_selector).count() == 0:
                    row_selector = ".row"
                
                rows = page.locator(row_selector)
                row_count = rows.count()
                
                for r in range(row_count):
                    try:
                        row = rows.nth(r)
                        # Check if it has 2 clear columns
                        # Usually title is first child, value is second
                        if row_selector == ".infomation-course__content":
                             # Based on inspection: .infomation-course__content__title and .infomation-course__content__value
                             label_el = row.locator(".infomation-course__content__title")
                             if label_el.count() > 0:
                                 label = label_el.inner_text(timeout=5000).strip()
                                 
                                 # Try specific value class first
                                 value_el = row.locator(".infomation-course__content__value")
                                 value = ""
                                 if value_el.count() > 0:
                                     value = value_el.first.inner_text(timeout=5000).strip()
                                 else:
                                     # Fallback: get all text and replace label
                                     # Or try generic sibling
                                     value_divs = row.locator("div")
                                     if value_divs.count() >= 2:
                                          try:
                                              value = value_divs.nth(1).inner_text(timeout=5000).strip()
                                          except:
                                              pass
                                     
                                     if not value:
                                         # Last resort: raw text parsing
                                         try:
                                             full_text = row.inner_text(timeout=5000).strip()
                                             if label in full_text:
                                                 value = full_text.replace(label, "", 1).strip()
                                                 if value.startswith(":"):
                                                     value = value[1:].strip()
                                         except:
                                             pass

                                 if label:
                                     item_data[label] = value
                        else:
                            # Fallback for generic .row
                            cols = row.locator("div")
                            if cols.count() >= 2:
                                label = cols.nth(0).inner_text(timeout=5000).strip()
                                value = cols.nth(1).inner_text(timeout=5000).strip()
                                if label and value:
                                    item_data[label] = value
                    except Exception as e:
                        # If a single row fails, don't crash the whole item scraping
                        # print(f"    Warning: Error scraping a detail row: {e}")
                        continue
                
                # Cross-fill Logic (Self-Healing)
                # 1. If Entity Name is Unknown, try to get from "Tên đơn vị (đầy đủ)" or "Tên dùng trong đấu thầu"
                if entity_name == "Unknown":
                    if "Tên đơn vị (đầy đủ)" in item_data and item_data["Tên đơn vị (đầy đủ)"]:
                        entity_name = item_data["Tên đơn vị (đầy đủ)"]
                    elif "Tên tiếng Việt" in item_data and item_data["Tên tiếng Việt"]: # Possible alias
                         entity_name = item_data["Tên tiếng Việt"]

                # 2. If "Tên đơn vị (đầy đủ)" is missing, fill from Entity Name
                if "Tên đơn vị (đầy đủ)" not in item_data or not item_data["Tên đơn vị (đầy đủ)"]:
                    if entity_name != "Unknown":
                        item_data["Tên đơn vị (đầy đủ)"] = entity_name

                # Update Entity Name in dict if we want it there explicitly as the main key
                item_data["Entity Name"] = entity_name
                
                all_data.append(item_data)
                
                # Add to processed list using logical ID
                # If we have "Mã định danh" in item_data, use that. Else use entity_name
                final_id = item_data.get("Mã định danh", entity_name) 
                
                # If we scraped successfully, let's prefer the ID we saw on the list if detail missed it? 
                # No, detail is source of truth. But for skip logic, list ID is used.
                # To be safe: Add BOTH the List ID (if we got it) AND the Detail ID.
                # But processed_items is a set of ONE key.
                # Recommendation: Use the same check key as on the list.
                # Check key is the one we extracted: item_id (which is ID from list)
                if item_id:
                     processed_items.add(item_id)
                elif final_id:
                     processed_items.add(final_id)
                    
                print(f"    Collected: {entity_name} (ID: {final_id})")
                
                # Sleep explicitly as requested to avoid block
                time.sleep(random.uniform(1, 2))

                # Go Back
                print("    Going back to list...")
                wait_for_internet(page)
                back_success = False
                for b_retry in range(3):
                    try:
                        # Try normal click
                        if b_retry == 0:
                            page.click(back_btn_selector, timeout=10000)
                        else:
                            # Try JS click if normal click fails (element detached/overlayed)
                            print("    Retry back button using JS...")
                            page.evaluate("document.querySelector('button.btn-back') && document.querySelector('button.btn-back').click()")
                            time.sleep(2)
                        
                        # Verify we actually went back? 
                        # We do this by waiting for list item which is done below.
                        back_success = True
                        break
                    except Exception as e:
                        print(f"    Back click failed (attempt {b_retry+1}): {e}")
                        time.sleep(2)
                
                if not back_success:
                     print("    Critical: Could not click Back button. Force reloading page to list...")
                     try:
                         page.goto("https://muasamcong.mpi.gov.vn/web/guest/investors-approval-v2", timeout=60000)
                     except:
                         pass

                # Wait for list to reappear
                try:
                    page.wait_for_selector(item_selector, timeout=20000)
                except:
                     print("    Timeout waiting for list after back. Reloading page...")
                     wait_for_internet(page)
                     try:
                         page.reload()
                         page.wait_for_selector(item_selector, timeout=60000)
                     except Exception as e:
                         print(f"    Critical error reloading list: {e}. Skipping to next iteration.")

            # Save progress after each page
            try:
                df = pd.DataFrame(all_data)
                # Ensure columns order if wanted
                df.to_excel(output_path, index=False)
                print(f"  Page completed. Progress saved to {output_path}.")
            except PermissionError:
                backup_name = output_path.replace(".xlsx", "_backup.xlsx")
                print(f"  Error: Could not save to '{output_path}'. Is it open? Saving to '{backup_name}' instead.")
                df.to_excel(backup_name, index=False)
            except Exception as e:
                print(f"  Error saving excel: {e}")

            # Pagination
            print("Checking pagination...")
            wait_for_internet(page)
            # Scroll to bottom
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)

            # Look for "Next" button.
            # Selector: button.btn-next
            # Use .first to avoid strict mode violation if multiple exist (e.g. top and bottom)
            next_btn = page.locator("button.btn-next").first
            
            # Check if button exists and is visible
            if next_btn.count() > 0 and next_btn.is_visible():
                # Check if disabled
                if not next_btn.is_enabled():
                    print("Next button found but disabled. Reached last page.")
                    break
                
                print("Clicking Next page...")
                try:
                    next_btn.click()
                    page_num += 1
                    
                    # Wait for items to refresh
                    time.sleep(3) 
                    page.wait_for_selector(item_selector, timeout=20000)
                except Exception as e:
                    print(f"Error clicking next or waiting for load: {e}")
                    wait_for_internet(page)
                    # Retry once
                    try: 
                        next_btn.click()
                        page.wait_for_selector(item_selector, timeout=20000)
                    except:
                        break
            else:
                print("No Next button found. Last page reached.")
                break

        browser.close()
        print("Scraping completed.")

def run_contractor_selection(output_path=None, max_pages=None, keywords="", exclude_words="", from_date="", to_date="", pause_event=None, stop_event=None):
    """
    Function to scrape Contractor Selection Results (Kết quả lựa chọn nhà thầu).
    Specific logic for:
    - URL: https://muasamcong.mpi.gov.vn/web/guest/contractor-selection?render=search
    - Match Type: "Khớp từ hoặc một số từ"
    - Search By: "Thông báo mời thầu thuốc, dược liệu..."
    - Field: "Hàng hóa"
    - Date Range: "Thời gian đăng tải"
    """
    import os
    
    # 1. Setup Defaults
    if output_path is None:
        output_path = "contractor_results.xlsx"
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"
    if max_pages is None:
        max_pages = float('inf')

    # Default Keywords if empty (as per user request reference)
    default_keywords = "thuốc, generic, tân dược, biệt dược, bệnh viện, chữa bệnh, vật tư y tế, điều trị, bệnh nhân, thiết bị y tế, khám chữa bệnh, khám bệnh, chữa bệnh, dược liệu, dược"
    # Note: "Thông báo mời thầu thuốc..." is a Search By option, not a keyword.
    
    default_exclude = "linh kiện, xây dựng, cải tạo, lắp đặt, thi công"
    
    if not keywords:
        keywords = default_keywords
    if not exclude_words:
        exclude_words = default_exclude

    print(f"--- Starting Contractor Selection Scrape ---")
    print(f"File: {output_path}")
    print(f"Keywords: {keywords[:50]}...")
    print(f"Exclude: {exclude_words}")
    if from_date or to_date:
        print(f"Date Range: {from_date} - {to_date}")

    # 2. Check Existing Data
    processed_items = set()
    all_data = []
    if os.path.exists(output_path):
        try:
            df = pd.read_excel(output_path)
            # Use 'Mã TBMT/Gói thầu' or check col "Số TBMT"
            check_col = "Số TBMT"
            if check_col not in df.columns:
                 # Try to find a unique column
                 possible_cols = ["Mã KQLCNT", "Số TBMT", "Tên gói thầu"]
                 for c in possible_cols:
                     if c in df.columns:
                         check_col = c
                         break
            
            if check_col in df.columns:
                processed_items = set(df[check_col].dropna().astype(str).str.strip())
                all_data = df.to_dict('records')
            print(f"Loaded {len(processed_items)} existing items (Check col: {check_col}).")
        except Exception as e:
            print(f"Warning: Could not read existing file: {e}")

    # 3. Helpers
    def check_status():
        if stop_event and stop_event.is_set():
            raise InterruptedError("Stopped by user.")
        if pause_event and not pause_event.is_set():
            print(">>> PAUSED...")
            pause_event.wait()
            if stop_event and stop_event.is_set():
                 raise InterruptedError("Stopped by user.")
            print(">>> RESUMED.")

    def wait_for_internet(page):
        while True:
            try:
                if page.evaluate("navigator.onLine"): return
            except: pass
            print("Waiting for internet...")
            time.sleep(5)

    # 4. Playwright Execution
    with sync_playwright() as p:
        # Browser Launch
        browser = None
        try:
            browser = p.chromium.launch(headless=False, channel="chrome")
        except:
            try:
                browser = p.chromium.launch(headless=False, channel="msedge")
            except:
                 browser = p.chromium.launch(headless=False)
        
        context = browser.new_context(viewport={"width": 1366, "height": 768}, ignore_https_errors=True)
        page = context.new_page()

        # Navigate
        print("Navigating to Contractor Selection Search...")
        url = "https://muasamcong.mpi.gov.vn/web/guest/contractor-selection?render=search"
        try:
            page.goto(url, timeout=60000)
        except:
            wait_for_internet(page)
            page.reload()

        # 5. Fill Search Form
        try:
            print("Configuring search criteria...")
            check_status()
            
            # Wait for form to load
            page.wait_for_selector('input[placeholder*="TBMT"]', timeout=30000)

            # 1. Select "Tìm theo": "Thông báo mời thầu thuốc, dược liệu..."
            # Click Dropdown
            dropdown = page.locator('.ant-select-selection--single').first
            dropdown.click()
            time.sleep(1)
            # Click Option
            option = page.locator("li.ant-select-dropdown-menu-item").filter(has_text="Thông báo mời thầu thuốc, dược liệu")
            if option.count() > 0:
                option.first.click()
                print("Selected Search By: Medicine")
            else:
                print("Warning: Could not find 'Medicine' search option.")
            
            time.sleep(1)

            # 2. Select Field: "Hàng hóa"
            # It's a checkbox with id="HH" usually, or find by label
            hh_cb = page.locator('input#HH')
            if hh_cb.count() > 0:
                if not hh_cb.is_checked():
                     hh_cb.click()
            else:
                # Try by text
                page.locator("label").filter(has_text="Hàng hóa").click()
            print("Selected Field: Goods")
            
            time.sleep(1)

            # 3. Enter Keywords
            # Input 1: Keywords
            key_input = page.locator('input[placeholder*="TBMT"]') 
            key_input.fill(keywords)

            # Input 2: Exclude
            exclude_input = page.locator('input[placeholder*="Áp dụng cho tất cả"]') 
            exclude_input.fill(exclude_words)
            
            # 4. Use Date Range if provided (Fix for Readonly)
            if from_date or to_date:
                print(f"Setting Date Range: {from_date} - {to_date}")
                try:
                    # Find inputs with placeholder 'dd/mm/yyyy'
                    dates = page.locator('input[placeholder="dd/mm/yyyy"]')
                    
                    if dates.count() >= 2:
                        # Start Date
                        if from_date:
                            start_inp = dates.nth(0)
                            start_inp.click() # Focus
                            time.sleep(0.5)
                            # Use keyboard type instead of fill because element is readonly
                            page.keyboard.type(from_date, delay=100)
                            page.keyboard.press("Enter")
                            time.sleep(0.5)
                            
                        # End Date
                        if to_date:
                            end_inp = dates.nth(1)
                            end_inp.click() # Focus 
                            time.sleep(0.5)
                            page.keyboard.type(to_date, delay=100)
                            page.keyboard.press("Enter")
                            time.sleep(0.5)
                    else:
                        print("Warning: Could not find date inputs with placeholder 'dd/mm/yyyy'.")

                except Exception as e:
                    print(f"Error setting dates: {e}")

            # 5. Select Match Type: "Khớp từ hoặc một số từ" (Radio 3)
            # Do this LAST to ensure it doesn't get reset by other changes
            try:
                # Find label containing text and click it
                match_label = page.locator("label").filter(has_text="Khớp từ hoặc một số từ (Phân biệt dấu)")
                if match_label.count() > 0:
                    match_label.click()
                    print("Selected Match Type: 'Khớp từ hoặc một số từ (Phân biệt dấu)'")
                else:
                    # Retry with xpath specific to text
                    print("Label not found with simple filter, trying xpath...")
                    page.locator('//label[contains(text(), "Khớp từ hoặc một số từ (Phân biệt dấu)")]').click()
            except Exception as e:
                print(f"Error selecting match type: {e}. Trying index...")
                try: page.locator('input[type="radio"]').nth(2).click()
                except: pass

            # E. Click Search

            # E. Click Search
            print("Clicking Search...")
            search_btn = page.locator("button.content__footer__btn").filter(has_text="Tìm kiếm")
            search_btn.click()
            
            # Wait for results
            print("Waiting for results...")
            time.sleep(3) # Initial wait
            
            item_selector = ".content__body__list__item"
            try:
                page.wait_for_selector(item_selector, timeout=30000)
            except:
                print("No results found or timeout.")
                browser.close()
                return

        except Exception as e:
            print(f"Error setting up search: {e}")
            browser.close()
            return

        # 6. Scraping Loop
        page_num = 1
        while page_num <= max_pages:
            check_status()
            print(f"--- Processing Page {page_num} ---")
            wait_for_internet(page)
            
            items = page.locator(item_selector)
            count = items.count()
            print(f"Found {count} items.")
            
            if count == 0:
                break

            for i in range(count):
                check_status()
                print(f"  Scraping item {i+1}/{count}...")
                
                retry = 0
                skipped = False
                while retry < 3:
                    try:
                        # Re-locate
                        item = page.locator(item_selector).nth(i)
                        
                        # Get Title to check duplicate
                        title_el = item.locator("h5 a, a.content__body__list__item__title")
                        if title_el.count() == 0:
                             title_el = item.locator("a").first 
                        
                        title_text = title_el.inner_text().strip()
                        item_id = title_text 
                        
                        if item_id in processed_items:
                            print(f"    Skipping duplicate: {item_id}")
                            skipped = True
                            break
                        
                        # Click to Open
                        title_el.click()
                        
                        # Wait for detail
                        try:
                            page.wait_for_selector(".btn-back, .content-body__header", timeout=15000)
                        except:
                            print("    Timeout loading detail. Retrying click...")
                            page.reload() 
                            raise Exception("Detail load timeout")

                        # --- SCRAPE DETAIL ---
                        detail_data = {"Entity Name": title_text, "Link": page.url}
                        
                        # Use a generic scraper for rows/cols similar to existing function
                        rows = page.locator(".row, .infomation-course__content")
                        for r_idx in range(rows.count()):
                            row = rows.nth(r_idx)
                            try:
                                txt = row.inner_text().strip()
                                parts = txt.split(":", 1)
                                if len(parts) == 2:
                                    lab = parts[0].strip()
                                    val = parts[1].strip()
                                    if lab and val:
                                        detail_data[lab] = val
                            except:
                                pass

                        all_data.append(detail_data)
                        print(f"    Collected: {title_text}")
                        
                        processed_items.add(item_id)

                        # GO BACK
                        back_btn = page.locator("button.btn-back")
                        if back_btn.count() > 0:
                            back_btn.click()
                        else:
                            page.go_back()
                        
                        # Wait for list
                        page.wait_for_selector(item_selector, timeout=15000)
                        break

                    except Exception as e:
                        print(f"    Error: {e}. Retrying...")
                        retry += 1
                        if "Detail" in str(e) or page.url != url:
                            page.go_back()
                            try: page.wait_for_selector(item_selector, timeout=10000)
                            except: page.goto(url); 
                        time.sleep(2)
                
                time.sleep(1)

            # Save
            try:
                pd.DataFrame(all_data).to_excel(output_path, index=False)
                print(f"  Saved progress to {output_path}")
            except Exception as e:
                print(f"  Save error: {e}")

            # Next Page
            next_btn = page.locator("button.btn-next").first
            if next_btn.count() > 0 and not next_btn.is_disabled():
                print("Next page...")
                next_btn.click()
                time.sleep(3)
                page.wait_for_selector(item_selector, timeout=20000)
                page_num += 1
            else:
                print("End of pages.")
                break
        
        browser.close()

if __name__ == "__main__":
    run()
