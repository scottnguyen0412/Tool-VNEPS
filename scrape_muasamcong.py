import json
import time
import sys
from datetime import datetime
import random
import pandas as pd
import urllib.parse
from playwright.sync_api import sync_playwright
import os
import requests
import ssl
from requests.adapters import HTTPAdapter
from urllib3.util.ssl_ import create_urllib3_context
from urllib3.poolmanager import PoolManager

os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"

def fetch_bid_detail(api_context, token, bid_id):
    url = f"https://muasamcong.mpi.gov.vn/o/egp-portal-contractor-selection-v2/services/lcnt_tbmt_ttc_ldt?token={token}"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
        "Origin": "https://muasamcong.mpi.gov.vn",
        "Referer": "https://muasamcong.mpi.gov.vn/",
    }
    payload = {"id": bid_id}
    
    try:
        response = api_context.post(url, data=payload, headers=headers)
        if response.ok:
            return response.json()
        else:
            # print(f"Detail API Error {response.status}: {response.status_text}")
            return None
    except Exception as e:
        # print(f"Detail Fetch Error: {e}")
        return None
    return None

def process_detail_data(detail_json):
    if not detail_json: return None
    
    # Root object helper
    m = detail_json.get("bidoNotifyContractorM", {}) or {}
    plan = detail_json.get("bidpPlanDetail", {}) or {}
    contractor = detail_json.get("bidInvContractorOfflineDTO", {}) or {}
    status_obj = detail_json.get("bidoBidStatus", {}) or {}
    
    # Helpers
    def get_date(iso):
        if not iso: return ""
        try:
            return iso.split("T")[0]
        except: return iso
        
    def get_datetime(iso):
         if not iso: return ""
         try:
            return iso.replace("T", " ").split(".")[0]
         except: return iso

    def map_val(val, mapping):
        s = str(val) if val is not None else ""
        return mapping.get(s, s)
        
    def format_currency(val):
        if val is None or val == "": return ""
        try:
            # Format: 9.422.000 VNĐ
            s = "{:,.0f}".format(float(val))
            return s.replace(",", ".") + " VNĐ"
        except: return str(val)

    unit_map = {"D": "Ngày", "M": "Tháng", "Y": "Năm", "W": "Tuần", "Q": "Quý"}
    def format_period(val, unit):
        v = str(val) if val is not None else ""
        u = unit_map.get(unit, unit) if unit else ""
        return f"{v} {u}".strip()

    # Location
    locs = detail_json.get("bidpBidLocationList", [])
    loc_str = ""
    if locs:
        l_parts = []
        for l in locs:
            d = l.get("districtName", "")
            p = l.get("provName", "")
            l_parts.append(f"{p}, {d}")
        loc_str = "; ".join(l_parts)
    
    # Lot Count
    lot_list = []
    if m.get("lotDTOList"): lot_list = m.get("lotDTOList")
    if not lot_list:
        resp = detail_json.get("bidNoContractorResponse", {}) or {}
        notif = resp.get("bidNotification", {}) or {}
        if notif.get("lotDTOList"): lot_list = notif.get("lotDTOList")
    if not lot_list and "lotDTOList" in detail_json:
        lot_list = detail_json["lotDTOList"]

    # isMultiLot Logic (Fallback to status or plan if missing in m)
    is_multi = m.get("isMultiLot")
    if is_multi is None: is_multi = status_obj.get("isMultiLot")
    if is_multi is None: is_multi = plan.get("isMultiLot")
    multi_lot_str = "Có" if str(is_multi) == "1" else "Không"

    # Mapping
    return {
        "Mã TBMT": m.get("notifyNo"),
        "Ngày đăng tải": get_datetime(m.get("publicDate")),
        "Mã KHLCNT": m.get("planNo"),
        "Phân loại KHLCNT": map_val(m.get("planType"), {"TX": "Chi thường xuyên", "DT": "Đầu tư phát triển"}),
        "Tên dự toán mua sắm": m.get("planName") or m.get("projectName"),
        "Quy trình áp dụng": map_val(m.get("processApply"), {"LDT": "Luật Đấu thầu", "LDT2023": "Luật đấu thầu 2023"}),
        "Tên gói thầu": m.get("bidName"),
        "Chủ đầu tư": m.get("investorName"),
        "Chi tiết nguồn vốn": m.get("capitalDetail"),
        "Lĩnh vực": map_val(m.get("investField"), {"HH": "Hàng hóa", "XL": "Xây lắp", "PTV": "Phi tư vấn", "TV": "Tư vấn"}),
        "Hình thức lựa chọn nhà đầu": map_val(m.get("bidForm"), {"DTRR": "Đấu thầu rộng rãi", "CHCT": "Chào hàng cạnh tranh"}),
        "Loại hợp đồng": map_val(m.get("contractType"), {"DGCD": "Đơn giá cố định", "TRGO": "Trọn gói"}),
        "Trong nước/ Quốc tế": "Trong nước" if str(m.get("isDomestic")) == "1" else "Quốc tế",
        "Thời gian thực hiện gói thầu": format_period(m.get('contractPeriod'), m.get('contractPeriodUnit')),
        "Gói thầu có nhiều phần/ lô": multi_lot_str,
        "Số lượng phần (lô)": len(lot_list) if lot_list else 0,
        "Hình thức dự thầu": "Qua mạng" if str(m.get("isInternet")) == "1" else "Không qua mạng",
        "Địa điểm phát hành e-HSMT": m.get("issueLocation"),
        "Địa điểm nhận e-HSDT": m.get("receiveLocation"),
        "Địa điểm thực hiện gói thầu": loc_str,
        "Thời điểm đóng thầu": get_datetime(m.get("bidCloseDate")),
        "Thời điểm mở thầu": get_datetime(m.get("bidOpenDate")),
        "Hiệu lực hồ sơ dự thầu": format_period(m.get('bidValidityPeriod'), m.get('bidValidityPeriodUnit')),
        "Số tiền bảo đảm dự thầu": format_currency(m.get("guaranteeValue")),
        "Hình thức đảm bảo dự thầu": m.get("guaranteeForm"),
        "Số quyêt định phê duyệt": contractor.get("decisionNo"),
        "Ngày phê duyệt": get_date(contractor.get("decisionDate")),
        "Cơ quan ban hành quyết định": contractor.get("decisionAgency")
    }

def fetch_lot_open_detail(api_context, token, notify_no, notify_id):
    url = f"https://muasamcong.mpi.gov.vn/o/egp-portal-contractor-selection-v2/services/expose/ldtkqmt/bid-notification-p/lotOpenDetail?token={token}"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
        "Origin": "https://muasamcong.mpi.gov.vn",
        "Referer": "https://muasamcong.mpi.gov.vn/",
    }
    payload = {
        "notifyNo": notify_no,
        "notifyId": notify_id,
        "type": "TBMT",
        "packType": 0
    }
    try:
        response = api_context.post(url, data=payload, headers=headers)
        if response.ok:
            return response.json()
    except Exception as e:
        print(f"Error fetching lotOpenDetail: {e}")
    return None

def fetch_bid_open(api_context, token, notify_no, notify_id):
    url = f"https://muasamcong.mpi.gov.vn/o/egp-portal-contractor-selection-v2/services/expose/ldtkqmt/bid-notification-p/bid-open?token={token}"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
        "Origin": "https://muasamcong.mpi.gov.vn",
        "Referer": "https://muasamcong.mpi.gov.vn/",
    }
    payload = {
        "notifyNo": notify_no,
        "notifyId": notify_id,
        "type": "TBMT",
        "packType": 0
    }
    try:
         response = api_context.post(url, data=payload, headers=headers)
         if response.ok:
             return response.json()
    except Exception as e:
         print(f"Error fetching bid-open: {e}")
    return None

def fetch_contractor_input_result(api_context, token, bid_id):
    url = f"https://muasamcong.mpi.gov.vn/o/egp-portal-contractor-selection-v2/services/expose/contractor-input-result/get?token={token}"
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
        "Origin": "https://muasamcong.mpi.gov.vn",
        "Referer": "https://muasamcong.mpi.gov.vn/",
    }
    payload = {"id": bid_id}
    try:
        response = api_context.post(url, data=payload, headers=headers)
        if response.ok:
            return response.json()
    except Exception as e:
        print(f"Error fetching contractor-input-result: {e}")
    return None


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
        except Exception as e:
            print(f"Error setting up search (Form Fill): {e}")
            browser.close()
            return
            
        # 6. API Scraping Loop
        # We need to capture the API token and the base payload from the initial search
        # The easiest way is to wait for the request after clicking search
        
        api_url = None
        base_payload = None
        
        try:
             # Wait for the specific API request
             with page.expect_request(lambda request: "services/smart/search" in request.url and request.method == "POST", timeout=10000) as first_req:
                 # Check if request happens naturally or if we need to wait
                 pass
             
             api_url = first_req.value.url
             base_payload = first_req.value.post_data_json
             print(f"Captured API URL: {api_url}")
             
        except:
             # If we missed the event (because it happened too fast), try to search again or just look at last requests
             print("Missed initial API capture, clicking search again...")
             try:
                 search_btn.click()
                 with page.expect_request(lambda request: "services/smart/search" in request.url and request.method == "POST") as first_req:
                     pass
                 api_url = first_req.value.url
                 base_payload = first_req.value.post_data_json
             except Exception as e:
                 print(f"Could not capture API: {e}")
                 browser.close()
                 return

        if not api_url or not base_payload:
            print("Failed to configure API scraper.")
            browser.close()
            return
            
        print("API Scraper Configured. Starting batch processing...")
        
        # Override page number and size
        
        # Modify payload for high volume
        # The payload is typically a List: [{"pageSize": 10, "pageNumber": 0, "query": [...]}]
        
        current_payload_obj = base_payload[0] if isinstance(base_payload, list) else base_payload
        current_payload_obj["pageSize"] = 50
        
        page_num = 0
        total_fetched = 0
        
        # Setup API Context
        api_context = context.request
        
        while True:
            check_status() # Pause/Stop support
            
            # Update Page Number
            current_payload_obj["pageNumber"] = page_num
            final_payload = [current_payload_obj] if isinstance(base_payload, list) else current_payload_obj
            
            print(f"--- Fetching API Page {page_num} (Size: 50) ---")
            
            retry = 0
            response_json = None
            
            while retry < 3:
                try:
                    # POST request
                    resp = api_context.post(api_url, data=final_payload)
                    if resp.ok:
                        response_json = resp.json()
                        break
                    else:
                        print(f"API Error {resp.status}: {resp.status_text}")
                        time.sleep(2)
                        retry += 1
                except Exception as e:
                     print(f"Request failed: {e}")
                     time.sleep(2)
                     retry += 1
            
            if not response_json:
                print("Failed to get response after retries. Stop.")
                break
            
            # Extract Items
            items = []
            try:
                # Handle possible structures
                if "page" in response_json and "content" in response_json["page"]:
                    items = response_json["page"]["content"]
                elif "content" in response_json:
                    items = response_json["content"]
                elif isinstance(response_json, list):
                     items = response_json
            except: pass
            
            if not items:
                print("No items in response. End of items.")
                break
                
            print(f"  Got {len(items)} items from API.")
            
            # Process Items
            batch_data = []
            for item in items:
                # Mapping
                # Mã TBMT = notifyNo
                # Tên gói thầu = bidName (can be array or string)
                # Chủ đầu tư = investorName
                # Ngày đăng tải thông báo = originalPublicDate
                # Lĩnh vực = investField (Array)
                # Địa điểm = locations (Array of objects? subagent said distName + provName)
                # Thời điểm đóng thầu = bidCloseDate
                # Hình thức dự thầu = isInternet (1=QM)
                # Trạng thái = status
                
                try:
                    # Bid Name
                    bid_name = item.get("bidName", "")
                    if isinstance(bid_name, list):
                        bid_name = "; ".join(bid_name)
                    
                    # Invest Field
                    inv_field = item.get("investField", "")
                    if isinstance(inv_field, list):
                        inv_field = ", ".join(inv_field)
                    if inv_field == "HH": inv_field = "Hàng hóa"
                    elif inv_field == "XL": inv_field = "Xây lắp"
                    elif inv_field == "PTV": inv_field = "Phi tư vấn"
                    
                    # Location
                    loc_str = ""
                    locs = item.get("locations", [])
                    if locs and isinstance(locs, list):
                        # Assuming objects with distName, provName or just strings
                        l_parts = []
                        for l in locs:
                            if isinstance(l, dict):
                                d = l.get("districtName", "")
                                p = l.get("provName", "")
                                l_parts.append(f"{d} - {p}")
                            else:
                                l_parts.append(str(l))
                        loc_str = "; ".join(l_parts)
                    
                    # Internet
                    # isInternet can be 1 or 0
                    is_net = item.get("isInternet", 0)
                    hinh_thuc = "Qua mạng" if str(is_net) == "1" else "Không qua mạng"
                    
                    # Status
                    st_code = item.get("status", "")
                    st_map = {
                        "01": "Chưa đóng thầu",
                        "OPEN_BID": "Đang xét thầu", 
                        "IS_PUBLISH": "Có nhà trúng thầu",
                        "CANCEL_BID": "Đã hủy thầu"
                    }
                    trang_thai = st_map.get(str(st_code), str(st_code))

                    # Date Formatting
                    def format_date_str(iso_str):
                        if not iso_str: return ""
                        try:
                            # Handle ISO format. "2026-01-16T11:21:49.875"
                            # We can just split T and take string parts if simple, or use datetime
                            # Let's use string manipulation for speed and robustness against variations
                            # Or strict parsing if preferred.
                            # Standard ISO: YYYY-MM-DDTHH:MM:SS
                            
                            if "T" in iso_str:
                                p1, p2 = iso_str.split("T")
                                y, m, d = p1.split("-")
                                # p2 is HH:MM:SS.ms
                                time_part = p2.split(".")[0]
                                h, mi = time_part.split(":")[:2]
                                return f"{d}/{m}/{y} {h}:{mi}"
                            return iso_str
                        except:
                            return iso_str

                    row = {
                        "Mã TBMT": item.get("notifyNo", ""),
                        "Chủ đầu tư": item.get("investorName", ""),
                        "Địa điểm": loc_str,
                        "Thời điểm đóng thầu": format_date_str(item.get("bidCloseDate", "")),
                        "Thời điểm mở thầu": format_date_str(item.get("bidOpenDate", "")), # Added for Phase 2 Fallback
                        "Trạng thái": trang_thai,
                        "id": item.get("id", ""),
                        "bidID": item.get("bidId", ""), 
                        "inputResultId": item.get("inputResultId", "") 
                    }
                    batch_data.append(row)
                except Exception as e:
                    print(f"Error parse item: {e}")
            
            # Save Batch
            if batch_data:
                all_data.extend(batch_data)
                try:
                    pd.DataFrame(all_data).to_excel(output_path, index=False)
                    print(f"  Saved {len(batch_data)} API items to {output_path}")
                except Exception as e:
                    print(f"  Save error: {e}")

            total_fetched += len(items)
            
            # Check Max Pages
            if page_num >= max_pages:
                print(f"Reached max pages limit ({max_pages}).")
                break
                
            # Next Page
            page_num += 1
            time.sleep(1) # Gentle rate limit
            
        print(f"Scraping completed. Total items: {total_fetched}")
        
        # --- PHASE 2: Fetch Details ---
        token = None
        if api_url:
             try:
                 parsed = urllib.parse.urlparse(api_url)
                 token = urllib.parse.parse_qs(parsed.query).get('token', [None])[0]
             except: pass
        
        # Define New API Helper for Sheet 2
        def fetch_bid_pack_detail(api_context, bid_id):
            url = "https://muasamcong.mpi.gov.vn/api/unau/portal/ebid/bid-pack-info/get-detail"
            headers = {
                "Content-Type": "application/json",
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36",
            }
            payload = {"body": {"id": bid_id}}
            try:
                # Note: This is a different host/endpoint structure, usually doesn't need token if unau/portal?
                # But headers usually needed.
                resp = api_context.post(url, data=payload, headers=headers)
                if resp.ok:
                    return resp.json()
            except Exception as e:
                print(f"Error fetching bid-pack-info: {e}")
            return None

        phase3_inputs = []
        if token and all_data:
            print(f"--- Starting Detail Scraping (Token: {token[:10]}...) ---")
            detail_output_path = output_path.replace(".xlsx", " Detail.xlsx")
            
            detail_all = [] # Sheet 1
            hsmtt_all = []  # Sheet 2
            
            total_d = len(all_data)
            print(f"Total items to detail: {total_d}")
            
            for idx, row in enumerate(all_data):
                bid_id = row.get("id")
                # Use bidID from Phase 1 if available? 
                # User said: "use bidID from phase 1" for the new API.
                # Phase 1 extraction: "bidID": item.get("bidId", "")
                phase1_bid_id = row.get("bidID") 
                
                # Check which ID to use for which API?
                # Existing fetch_bid_detail uses 'id' (bid_id here).
                # New fetch_bid_pack_detail uses 'bidID' (phase1_bid_id).
                # If phase1_bid_id is missing, maybe fallback to bid_id?
                target_pack_id = phase1_bid_id if phase1_bid_id else bid_id
                
                check_status()
                print(f"  Fetching detail {idx+1}/{total_d}: {row.get('Mã TBMT', bid_id)}")
                
                # --- SHEET 1: General Info ---
                d_json = fetch_bid_detail(api_context, token, bid_id)
                if d_json:
                    d_row = process_detail_data(d_json)
                    if d_row:
                        detail_all.append(d_row)
                    
                    try:
                        plan = d_json.get("bidpPlanDetail", {}) or {}
                        link_info_str = plan.get("linkNotifyInfo")
                        if link_info_str:
                             phase3_inputs.append(link_info_str)
                    except: pass
                
                # --- SHEET 2: Ho so moi thau ---
                if target_pack_id:
                    pack_json = fetch_bid_pack_detail(api_context, target_pack_id)
                    if pack_json and "body" in pack_json:
                        body = pack_json["body"]
                        notif = body.get("bidNotification", {}) or {}
                        
                        # Formatting helpers
                        def fmt_ts(iso):
                            if not iso: return ""
                            try:
                                # 2026-01-15T15:05:00 -> dd/mm/yyyy HH:MM:SS
                                # Simple string replace
                                T_split = iso.split("T")
                                date_part = T_split[0]
                                time_part = T_split[1] if len(T_split) > 1 else ""
                                y, m, d = date_part.split("-")
                                return f"{d}/{m}/{y} {time_part.split('.')[0]}"
                            except: return iso
                        
                        def fmt_num(v): 
                             if v is None or v == "": return ""
                             try: return "{:,.0f}".format(float(v)).replace(",", ".")
                             except: return str(v)

                        # Lot List Logic
                        lots = notif.get("lotDTOList")
                        if not lots:
                            lots = body.get("detailLotList")
                        
                        if lots and isinstance(lots, list):
                            for lot in lots:
                                # Data Extraction Logic with multiple fallbacks
                                # 1. Bid Name
                                b_name = notif.get("bidName")
                                if not b_name: b_name = body.get("bidName")
                                if not b_name: b_name = row.get("Tên gói thầu", "")
                                
                                # 2. Dates
                                d_open = notif.get("bidOpenDate")
                                d_close = notif.get("bidCloseDate")
                                
                                if not d_close: d_close = row.get("Thời điểm đóng thầu", "")
                                final_open = fmt_ts(d_open)
                                final_close = fmt_ts(d_close) 
                                if not final_close and row.get("Thời điểm đóng thầu"):
                                     final_close = row.get("Thời điểm đóng thầu")
                                if not final_open and row.get("Thời điểm mở thầu"):
                                     final_open = row.get("Thời điểm mở thầu")

                                r2 = {
                                    "Mã TBMT": notif.get("notifyNo") if notif.get("notifyNo") else body.get("linkNotifyInfo", {}).get("notifyNo", row.get("Mã TBMT")),
                                    "Tên gói thầu": b_name,
                                    "Mã phần (Lô)": lot.get("lotNo"), 
                                    "Mã Thuốc": lot.get("medicineCode"),
                                    "Tên hoạt chất/ Tên thành phần thuốc": lot.get("lotName"),
                                    "Nồng độ/ hàm lượng": lot.get("nongDo"),
                                    "Đường dùng": lot.get("duongDung"),
                                    "Dạng bào chế": lot.get("dangBaoChe"),
                                    "Đơn vị tính": lot.get("uom"),
                                    "Số lượng": fmt_num(lot.get("quantity")),
                                    "Giá trị ước tính từng phần (VND)": fmt_num(lot.get("lotPrice")),
                                    "Giá kế hoạch": fmt_num(lot.get("pricePlan")),
                                    "Nhóm thuốc": lot.get("groupMedicine"),
                                    "Thời điểm mở thầu": final_open,
                                    "Thời điểm đóng thầu": final_close
                                }
                                hsmtt_all.append(r2)
                
                # Auto-Save (Dual Sheets)
                if (idx + 1) % 10 == 0:
                    try:
                        with pd.ExcelWriter(detail_output_path) as writer:
                             pd.DataFrame(detail_all).to_excel(writer, sheet_name='Thông tin chung', index=False)
                             pd.DataFrame(hsmtt_all).to_excel(writer, sheet_name='Hồ sơ mời thầu', index=False)
                        print(f"  [Auto-Save] Saved {len(detail_all)} details so far.")
                    except Exception as e:
                        print(f"  [Auto-Save Error] {e}")

                time.sleep(0.5)
            
            # Final Save
            if detail_all or hsmtt_all:
                try:
                    with pd.ExcelWriter(detail_output_path) as writer:
                         pd.DataFrame(detail_all).to_excel(writer, sheet_name='Thông tin chung', index=False)
                         pd.DataFrame(hsmtt_all).to_excel(writer, sheet_name='Hồ sơ mời thầu', index=False)
                    print(f"Successfully saved details to: {detail_output_path}")
                except Exception as e:
                    print(f"Error saving details: {e}")
            else:
                print("No detailed data collected.")
        else:
            print("Skipping details: No token found or no data.")

        # --- PHASE 3: Bid Opening Details ---
        # --- PHASE 3: Bid Opening Details ---
        if token and phase3_inputs:
             print(f"--- Starting Phase 3: Bid Opening Details (Count: {len(phase3_inputs)}) ---")
             
             # Save to same directory as output_path
             dir_name = os.path.dirname(output_path)
             phase3_output_path = os.path.join(dir_name, "Bien ban mo thau detail.xlsx")

             
             phase3_data = []
             
             for i, info_str in enumerate(phase3_inputs):
                 check_status()
                 try:
                     # Parse info
                     info_obj = json.loads(info_str)
                     notify_no = info_obj.get("notifyNo")
                     notify_id = info_obj.get("notifyId")
                     
                     if not notify_no or not notify_id: continue
                     
                     print(f"  Phase 3 Processing {i+1}/{len(phase3_inputs)}: {notify_no}")
                     
                     # Call APIs
                     # API 1
                     lot_details = fetch_lot_open_detail(api_context, token, notify_no, notify_id)
                     
                     # API 2
                     bid_opens = fetch_bid_open(api_context, token, notify_no, notify_id)
                     
                     if not lot_details: 
                         # Try API 2 alone? Join requires API 1 usually.
                         continue

                     # Process & Join
                     # API 1 list
                     lot_list = lot_details if isinstance(lot_details, list) else []
                     
                     # API 2 Map
                     bid_map = {} # Map bid_id (id in API 2) -> object
                     if bid_opens and "bidSubmissionByContractorViewResponse" in bid_opens:
                         sub_list = bid_opens["bidSubmissionByContractorViewResponse"].get("bidSubmissionDTOList", [])
                         if sub_list:
                             for sub in sub_list:
                                 if "id" in sub:
                                     bid_map[sub["id"]] = sub
                     
                     # Join
                     # Helper for formatting: 13545000 -> 13.545.000
                     def fmt_val(v):
                        if v is None or v == "": return ""
                        try:
                            return "{:,.0f}".format(float(v)).replace(",", ".")
                        except:
                            return str(v)

                     for lot in lot_list:
                          # Create Row
                          bid_open_id = lot.get("bidOpenId")
                          linked_bid = bid_map.get(bid_open_id, {})
                          
                          row = {
                              "Mã TBMT": notify_no,
                              "Mã phân/ lô": lot.get("lotNo"),
                              "Tên thành phần thuốc": lot.get("lotName"),
                              "Mã định danh": lot.get("contractorCode"),
                              "Tên nhà thầu": lot.get("contractorName"),
                              "Tỷ lệ phần trăm giảm giá (nếu có)": lot.get("discountPercent"),
                              "Giá dự thầu": fmt_val(lot.get("lotFinalPrice")),
                              
                              # Joined fields
                              "Hiệu lực HSDT": linked_bid.get("bidValidityNum"),
                              "Bảo đảm dự thầu cho các phần tham dự (VND)": fmt_val(linked_bid.get("bidGuarantee")),
                              "Hiệu lực của BĐDT": linked_bid.get("bidGuaranteeValidity")
                          }
                          phase3_data.append(row)
                     
                     time.sleep(0.5)

                 except Exception as e:
                     print(f"  Error Phase 3 item: {e}")
             
             # Save
             if phase3_data:
                 try:
                     pd.DataFrame(phase3_data).to_excel(phase3_output_path, index=False)
                     print(f"Successfully saved Phase 3 data to: {phase3_output_path}")
                 except Exception as e:
                     print(f"Error saving Phase 3 excel: {e}")
             else:
                 print("No data collected for Phase 3.")

        # --- PHASE 4: Contractor Input Result (Danh Sach Nha Thau & Hang Hoa) ---
        if token and all_data:
             print(f"--- Starting Phase 4: Contractor & Goods Lists ---")
             dir_name = os.path.dirname(output_path)
             nha_thau_path = os.path.join(dir_name, "Danh Sach Nha Thau.xlsx")
             hang_hoa_path = os.path.join(dir_name, "Danh Sach Hang Hoa.xlsx")
             
             list_nha_thau = []
             list_hang_hoa = []
             
             total_p4 = len(all_data)
             
             for i, row in enumerate(all_data):

                 # Phase 4 ID Logic
                 phase4_id = row.get("inputResultId")
                 if not phase4_id: 
                     # print(f"  Phase 4: Skipping {row.get('Mã TBMT')} (No inputResultId)")
                     continue
                 
                 check_status()
                 print(f"  Phase 4: Processing {i+1}/{total_p4} | InputResultID: {phase4_id}...")
             
                 try:
                     res = fetch_contractor_input_result(api_context, token, phase4_id)
                     if not res: 
                         print(f"    -> [Skipped] API returned None for {phase4_id}")
                         continue
                 
                     root_dto = res.get("bideContractorInputResultDTO", {})
                     if not root_dto: 
                         print(f"    -> [Skipped] No bideContractorInputResultDTO for {bid_id}")
                         continue
                 
                     notify_no = root_dto.get("notifyNo")
                     lot_results = root_dto.get("lotResultDTO") or []
                     lot_items = root_dto.get("lotResultItems") or []
                    #  Check if lot_items is empty
                     if not lot_items:
                         d_vers = root_dto.get("decisionVersions")
                         if d_vers and isinstance(d_vers, list):
                             for v in reversed(d_vers):
                                 items = v.get("lotResultItems")
                                 if items:
                                     lot_items = items
                                     break
                     
                     if not lot_results and not lot_items:
                         print(f"    -> [Info] No lot results or items.")
                     
                     # 1. Process Danh Sach Nha Thau
                     # Strategy: Iterate Lots -> ContractorList -> Link to LotItems
                     for lot in lot_results:
                         l_no = lot.get("lotNo")
                         l_name = lot.get("lotName")
                         c_list = lot.get("contractorList") or []
                         
                         for cntr in c_list:
                             # Link: item_result["listLotResultId"] == cntr["id"]
                             cntr_id = cntr.get("id")
                             linked_item = None
                             for it in lot_items:
                                 if it.get("listLotResultId") == cntr_id:
                                     linked_item = it
                                     break
                            
                             don_gia = None
                             qty = None
                             
                             if linked_item:
                                 fv_str = linked_item.get("formValue")
                                 if fv_str:
                                     try:
                                         fv_json = json.loads(fv_str)
                                         if fv_json and isinstance(fv_json, list):
                                              # Take first item for summary?
                                              first = fv_json[0]
                                              don_gia = first.get("donGia")
                                              qty = first.get("quantity")
                                     except: pass
                             
                             # Mapping
                             result_status = "Không trúng thầu"
                             if don_gia: # Rule: If has price -> Won? Or "Nếu có Giá dự thầu" - user request
                                 result_status = "Trúng thầu"
                             
                             # Formatting
                             def fmt(x):
                                 if x is None or x == "": return ""
                                 try: return "{:,.0f}".format(float(x)).replace(",", ".")
                                 except: return str(x)

                             row_nt = {
                                 "Mã TBMT": notify_no,
                                 "Mã phần (lô)": l_no,
                                 "Tên hoạt chất/ Tên thành phần thuốc": l_name,
                                 "Mã định danh": cntr.get("orgCode"),
                                 "Mã số thuế": cntr.get("taxCode"),
                                 "Tên nhà thầu": cntr.get("orgFullname"),
                                 "Giá Dự thầu": fmt(cntr.get("lotPrice")),
                                 "Đơn giá trúng thầu (VND)": fmt(don_gia),
                                 "Giá trúng thầu của từng phần đã bao gồm giảm giá (VND) (đã bao gồm các hạng mục của phần đó)": fmt(cntr.get("lotFinalPrice")),
                                 "Số lượng trúng thầu": fmt(qty),
                                 "Kết quả": result_status,
                                 "Thời gian thực hiện gói thầu": cntr.get("cperiodText"),
                                 "Thời gian thực hiện hợp đồng": cntr.get("bidExecutionTime")
                             }
                             list_nha_thau.append(row_nt)

                     # 2. Process Danh Sach Hang Hoa
                     # Strategy: Iterate LotItems -> formValue
                     for it in lot_items:
                         fv_str = it.get("formValue")
                         if not fv_str: continue
                         try:
                             goods = json.loads(fv_str)
                             if not goods or not isinstance(goods, list): continue
                             
                             for g in goods:
                                 def fmt(x):
                                     if x is None or x == "": return ""
                                     try: return "{:,.0f}".format(float(x)).replace(",", ".")
                                     except: return str(x)

                                 row_hh = {
                                     "Mã TBMT": notify_no,
                                     "Mã Phần/lô": g.get("lotNo"),
                                     "Mã thuốc": g.get("medicineCode"),
                                     "Tên thuốc": g.get("name"),
                                     "Tên hoạt chất/ Tên thành phần của thuốc": g.get("tenHoatChat"),
                                     "Nồng độ, hàm lượng": g.get("nongDo"),
                                     "Đường dùng": g.get("duongDung"),
                                     "Dạng bào chế": g.get("dangBaoChe"),
                                     "Quy cách": g.get("quyCach"),
                                     "Nhóm thuốc": g.get("groupMedicine"),
                                     "Hạn dùng (Tuổi thọ)": g.get("hanDung"),
                                     "GĐKLH hoặc GPNK": g.get("gdklh"),
                                     "Cơ sở sản xuất": g.get("csSanXuat"),
                                     "Xuất xứ": g.get("nuocSanXuat"),
                                     "Đơn vị tính": g.get("uom"),
                                     "Số lượng": fmt(g.get("quantity")),
                                     "Đơn giá trúng thầu (VND)": fmt(g.get("donGia")),
                                     "Thành tiền": fmt(g.get("amount")),
                                     "Nhà thầu trúng thầu": g.get("contractorName"),
                                     "Tiến độ cung cấp": g.get("tienDo")
                                 }
                                 list_hang_hoa.append(row_hh)
                         except: pass

                     time.sleep(0.3)
                 
                 except Exception as e:
                     print(f"  Error Phase 4 {bid_id}: {e}")
             
             # Save Files
             if list_nha_thau:
                 try:
                     pd.DataFrame(list_nha_thau).to_excel(nha_thau_path, index=False)
                     print(f"Successfully saved: {nha_thau_path}")
                 except Exception as e:
                     print(f"Error saving Nha Thau: {e}")
             
             if list_hang_hoa:
                 try:
                     pd.DataFrame(list_hang_hoa).to_excel(hang_hoa_path, index=False)
                     print(f"Successfully saved: {hang_hoa_path}")
                 except Exception as e:
                     print(f"Error saving Hang Hoa: {e}")

        browser.close()


def run_drug_price_scrape(output_path=None, pause_event=None, stop_event=None):
    if output_path is None:
        output_path = "CongBoGiaThuoc.xlsx"
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"

    print(f"--- Bắt đầu cào Công bố giá thuốc ---")
    print(f"File lưu: {output_path}")

    url = "https://dichvucong.dav.gov.vn/api/services/app/quanLyGiaThuoc/GetListCongBoPublicPaging"
    
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    # Default payload structure
    payload = {
        "CongBoGiaThuoc": {},
        "KichHoat": True,
        "skipCount": 0,
        # maxResultCount: Có thể tùy chỉnh: 15, 20, 50,100 
        "maxResultCount": 100,
        "sorting": None
    }

    all_data = []
    existing_ids = set()

    # Load existing data if file exists
    if os.path.exists(output_path):
        try:
            print(f"Loading existing data from {output_path}...")
            df_old = pd.read_excel(output_path)
            
            # Check for 'id' column
            if "id" in df_old.columns:
                existing_ids = set(df_old["id"].dropna().astype(str))
                all_data = df_old.to_dict('records')
            else:
                 # Fallback: keep existing data but can't check duplicates by ID effectively
                 # Just load it so we don't lose it.
                 all_data = df_old.to_dict('records')
                 
            print(f"Loaded {len(all_data)} existing records. ({len(existing_ids)} unique IDs)")
        except Exception as e:
            print(f"Warning: could not read existing file: {e}")

    skip_count = 0
    # max_result_count: Có thể tùy chỉnh: 15, 20, 50,100 
    max_result_count = 100 
    total_count = None
    processed_count = 0

    # formatting helper
    def format_money(val):
        if val is None or val == "":
            return ""
        try:
             # 999900.0 -> 999,900
            return "{:,.0f}".format(float(val))
        except:
            return str(val)

    def format_date(iso_str):
        if not iso_str: return ""
        s = str(iso_str)
        try:
            # 1. Handle ISO T split
            if "T" in s:
                s = s.split("T")[0]
            
            # 2. Handle YYYY-MM-DD
            if "-" in s:
                parts = s.split("-")
                if len(parts) == 3:
                     # YYYY-MM-DD
                     return f"{parts[2]}/{parts[1]}/{parts[0]}"
            
            # 3. If already has /, check usage? 
            # Assuming output needed is specific, but input could be anything.
            # If input is already correct, return it.
            return s
        except:
            return s

    while True:
        # Check Thread Events
        if stop_event and stop_event.is_set():
            print(">>> STOPPED by user.")
            break
        if pause_event:
            while not pause_event.is_set():
                time.sleep(1)
                if stop_event and stop_event.is_set():
                    break
        
        # Update payload
        payload["skipCount"] = skip_count
        
        print(f"Fetching skipCount={skip_count}...")
        
        try:
            resp = requests.post(url, json=payload, headers=headers, timeout=30)
            if resp.status_code != 200:
                print(f"Error: API returned status {resp.status_code}")
                time.sleep(5)
                continue
            
            data = resp.json()
            result = data.get("result", {})
            items = result.get("items", [])
            
            if total_count is None:
                total_count = result.get("totalCount", 0)
                print(f"Total items found: {total_count}")
            
            if not items:
                print("No more items.")
                break
                
            # Process items
            batch_new = 0
            for item in items:
                item_id = item.get("id")
                if not item_id: item_id = ""
                
                # Check duplicate
                if str(item_id) in existing_ids:
                    continue
                
                # Mapping
                row = {
                    "Ngày Công bố": format_date(item.get("ngayTiepNhan")),
                    "Tên thuốc": item.get("tenThuoc"),
                    "Tên HC": item.get("hoatChat"),
                    "NĐ/HL": item.get("hamLuong"),
                    "Số GPLH/GPNK": item.get("soDangKy"),
                    "Dạng bào chế": item.get("dangBaoChe"),
                    "Quy Cách Đóng gói": item.get("quyCachDongGoi"),
                    "ĐVT": item.get("donViTinh"),
                    "Giá Bán Buôn Dự Kiến (VNĐ) (VAT)": format_money(item.get("giaBanBuonDuKien")),
                    "Cơ sở SX": item.get("doanhNghiepSanXuat"),
                    "Nước Sản Xuất": item.get("nuocSanXuat"),
                    "Đối tượng thực hiện công bố": item.get("donViKeKhai"),
                    "id": item_id
                }
                all_data.append(row)
                existing_ids.add(str(item_id))
                batch_new += 1
            
            processed_count += len(items)
            print(f"  Fetched {len(items)} items. (New: {batch_new}) Total processed: {processed_count}/{total_count}")
            
            # Save every batch (or every few batches)
            try:
                pd.DataFrame(all_data).to_excel(output_path, index=False)
            except Exception as e:
                print(f"  Save error: {e}")

            # Prepare next page
            skip_count += max_result_count
            
            if skip_count >= total_count:
                print("All items fetched.")
                break
                
            time.sleep(1) # Polite delay

        except Exception as e:
            print(f"Request failed: {e}")
            time.sleep(5)
            #retry if fail
            pass

    print(f"Completed! Data saved to {output_path}")

def run_investor_scan_api(output_path=None, pause_event=None, stop_event=None, ministries=None):
    """
    New API-based scanning for Investors.
    Modes:
      1. 'Toàn Bộ' (All): ministries=None. Target: 'y tế', 'bệnh viện' (global).
      2. 'Theo Bộ Ngành' (Ministry): ministries=[list of names]. 
         - Logic: Look up code from CQCQ.
         - For 'Bộ Công an', 'Bộ Quốc phòng': Target 'bệnh viện', 'y tế' with agencyName filter.
         - For others: Target "" (all) with agencyName filter.
    """
    print("--- Starting API-based Investor Scan ---")
    
    if output_path is None:
        output_path = "investors_data_api.xlsx"
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"
    
    # 1. Load Mappings
    country_map = {}
    cqcq_map = {}
    
    try:

        # Robust Resource Path Finder
        def get_data_path(filename):
            candidates = []
            
            # 1. Bundled Path (PyInstaller)
            if getattr(sys, 'frozen', False):
                base_mei = sys._MEIPASS
                candidates.append(os.path.join(base_mei, "DATA_SAMPLE", filename))
                candidates.append(os.path.join(base_mei, filename)) # Check root of bundle
            
            # 2. Development / Script Path
            else:
                base_script = os.path.dirname(os.path.abspath(__file__))
                candidates.append(os.path.join(base_script, "DATA_SAMPLE", filename))
                candidates.append(os.path.join(base_script, filename))
            
            # 3. External Path (Next to Exe or Script)
            if getattr(sys, 'frozen', False):
                base_exe = os.path.dirname(sys.executable)
                candidates.append(os.path.join(base_exe, "DATA_SAMPLE", filename))
                candidates.append(os.path.join(base_exe, filename))
            
            # 4. Current Working Directory
            candidates.append(os.path.join(os.getcwd(), "DATA_SAMPLE", filename))
            candidates.append(os.path.join(os.getcwd(), filename))
            
            # Check all candidates
            for path in candidates:
                if os.path.exists(path):
                    return path
            
            # Return first candidate as default for error msg
            return candidates[0] if candidates else filename

        # Load Country
        c_path = get_data_path("Data-Country.json")
        if os.path.exists(c_path):
            with open(c_path, "r", encoding="utf-8") as f:
                c_data = json.load(f)
                for item in c_data:
                    code = item.get("code")
                    name = item.get("name")
                    if code: country_map[code] = name
        else:
             print(f"Warning: {c_path} not found.")

        # Load CQCQ
        q_path = get_data_path("Data-CQCQ.json")
        if os.path.exists(q_path):
            with open(q_path, "r", encoding="utf-8") as f:
                q_data = json.load(f)
                for item in q_data:
                    code = item.get("code")
                    name = item.get("name")
                    if code: cqcq_map[code] = name
        else:
             print(f"Warning: {q_path} not found.")
                
        print(f"Loaded {len(country_map)} countries and {len(cqcq_map)} agencies.")
    except Exception as e:
        print(f"Warning: Could not load mapping files: {e}")

    # SSL FIX for DH_KEY_TOO_SMALL
    class LegacyAdapter(HTTPAdapter):
        def init_poolmanager(self, connections, maxsize, block=False):
            ctx = create_urllib3_context()
            # Lower security level to allow smaller DH keys (legacy servers)
            try:
                ctx.set_ciphers('DEFAULT@SECLEVEL=1')
            except Exception:
                pass 
            self.poolmanager = PoolManager(
                num_pools=connections,
                maxsize=maxsize,
                block=block,
                ssl_context=ctx
            )

    session = requests.Session()
    adapter = LegacyAdapter()
    session.mount('https://', adapter)

    # 2. Pre-fetch Business Types
    business_type_map = {}
    try:
        url_bt = "https://muasamcong.mpi.gov.vn/o/egp-portal-investor-approved-v2/services/get-business-type-list"
        payload_bt = {"queryParams":{"categoryTypeCode":{"equals":"BUSINESS_TYPE"}}}
        headers = {
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        # Use session
        resp = session.post(url_bt, json=payload_bt, headers=headers, timeout=10)
        if resp.ok:
            bt_list = resp.json()
            for item in bt_list:
                code = item.get("code")
                name = item.get("name")
                if code: business_type_map[code] = name
            print(f"Loaded {len(business_type_map)} business types.")
        else:
            print("Failed to load business types.")
    except Exception as e:
        print(f"Error loading business types: {e}")

    # 3. Helpers
    def check_status():
        if stop_event and stop_event.is_set():
            raise InterruptedError("Stopped by user")
        if pause_event and not pause_event.is_set():
            print(">>> PAUSED...")
            try:
                pause_event.wait()
                if stop_event and stop_event.is_set():
                    raise InterruptedError("Stopped by user")
                print(">>> RESUMED")
            except Exception as e:
                pass

    area_cache = {} 
    def get_area_name(code):
        if not code: return ""
        if code in area_cache: return area_cache[code]
        
        url = "https://muasamcong.mpi.gov.vn/o/egp-portal-investor-approved-v2/services/get-area-by-code"
        payload = {"queryParams":{"code":{"equals": code}}}
        try:
            r = session.post(url, json=payload, headers=headers, timeout=10)
            if r.ok:
                data = r.json()
                if data and isinstance(data, list) and len(data) > 0:
                    name = data[0].get("name", "")
                    area_cache[code] = name
                    return name
        except:
            pass
        return code

    processed_org_codes = set()
    all_rows = []
    
    if os.path.exists(output_path):
        try:
            print(f"Reading existing file: {output_path}")
            df_old = pd.read_excel(output_path)
            # Normalize column reading
            if "Mã định danh" in df_old.columns:
                processed_org_codes = set(df_old["Mã định danh"].dropna().astype(str))
                all_rows = df_old.to_dict('records')
            print(f"Loaded {len(all_rows)} existing records.")
        except Exception as e:
            print(f"Warning: could not read existing file: {e}")
    
    keywords = ["y tế", "bệnh viện"]
    
    headers = {
        "Content-Type": "application/json",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Origin": "https://muasamcong.mpi.gov.vn",
        "Referer": "https://muasamcong.mpi.gov.vn/"
    }

    item_buffer = []

    # Prepare list of tasks: [(MinistryName, MinistryCode, Keyword)]
    # If ministries=None (All Mode): [("All", None, "y tế"), ("All", None, "bệnh viện")]
    
    tasks = []
    
    if not ministries:
        # All Mode
        tasks.append(("All", None, "y tế"))
        tasks.append(("All", None, "bệnh viện"))
    else:
        # Ministry Mode
        # Build Name->Code map with validation
        name_to_code = {}
        for c, n in cqcq_map.items():
            if n: 
                # Normalize spaces? API usually exact match or contains.
                # User input comes from GUI which likely matches JSON name exactly.
                name_to_code[n] = c
                
        for m_name in ministries:
            m_code = name_to_code.get(m_name)
            if not m_code:
                # Try simple fuzzy or stripped?
                # For now assume exact match from GUI
                print(f"Warning: Could not find code for '{m_name}'. Skipping.")
                continue
                
            if m_name in ["Bộ Công an", "Bộ Quốc phòng"]:
                tasks.append((m_name, m_code, "bệnh viện"))
                tasks.append((m_name, m_code, "y tế"))
            else:
                tasks.append((m_name, m_code, "")) # Empty keyword for others

    for m_name, m_code, kw in tasks:
        print(f"Scanning Ministry: {m_name} | Code: {m_code} | Keyword: '{kw}'")
        page_idx = 0
        total_pages = 1 
        
        while page_idx < total_pages:
            check_status()
            
            url_1 = "https://muasamcong.mpi.gov.vn/o/egp-portal-investor-approved-v2/services/um/lookup-orgInfo"
            
            # Construct Payload
            q_params = {
                "roleType": {"equals": "CDT"},
                "orgName": {"contains": None},
                "orgCode": {"contains": None},
                "agencyName": {"in": None}, # Default None
                "effRoleDate": {
                    "greaterThanOrEqual": None,
                    "lessThanOrEqual": None
                }
            }
            
            # Set Keyword
            if kw:
                q_params["orgNameOrOrgCode"] = {"contains": kw}
            else:
                 q_params["orgNameOrOrgCode"] = {"contains": ""}
                 
            # Set Agency Code
            if m_code:
                q_params["agencyName"] = {"in": [m_code]}
            
            payload_1 = {
                "pageSize": 10,
                "pageNumber": page_idx,
                "queryParams": q_params
            }
            
            try:
                print(f"  Fetching page {page_idx + 1}...")
                resp = session.post(url_1, json=payload_1, headers=headers, timeout=30)
                if not resp.ok:
                    print(f"  Error fetching page {page_idx}: {resp.status_code}. Retrying...")
                    time.sleep(5)
                    # Simple retry once
                    try:
                        resp = session.post(url_1, json=payload_1, headers=headers, timeout=30)
                    except: pass
                    
                    if not resp.ok:
                         print("  Skip this page due to error.")
                         page_idx += 1
                         continue

                data = resp.json()
                content_obj = data.get("ebidOrgInfos", {})
                items = content_obj.get("content", [])
                
                # Update total pages from response
                if "totalPages" in content_obj:
                    total_pages = content_obj["totalPages"]
                
                if not items:
                    print("  No items found on this page.")
                    if page_idx >= total_pages - 1:
                        break
                    else:
                        page_idx += 1
                        continue
                    
                for item in items:
                    check_status()
                    
                    org_code = item.get("orgCode")
                    if not org_code: continue
                    
                    if str(org_code) in processed_org_codes:
                        # print(f"    Skipping duplicate {org_code}")
                        continue
                    
                    # API 2: Detail
                    url_2 = "https://muasamcong.mpi.gov.vn/o/egp-portal-investor-approved-v2/services/um/org/get-detail-info"
                    payload_2 = {"orgCode": org_code}
                    
                    detail_info = {}
                    try:
                        r2 = session.post(url_2, json=payload_2, headers=headers, timeout=10)
                        if r2.ok:
                            detail_json = r2.json()
                            detail_info = detail_json.get("orgInfo", {})
                    except Exception as e:
                        print(f"    Error fetching detail for {org_code}: {e}")
                    
                    # Mapping Helpers
                    def get_d(key, default=""):
                        val = detail_info.get(key)
                        if val: return val
                        # Fallback to item (API 1)
                        if key == "orgFullName": return item.get("orgFullname")
                        if key == "orgEnName": return item.get("orgEnName")
                        if key == "repName": return item.get("repFullname")
                        return item.get(key, default)
                    
                    def fmt_date_arr(arr):
                        if not arr or not isinstance(arr, list) or len(arr) < 3: return ""
                        # Adjust for year, month, day. Usually [y, m, d, h, m, s]
                        return f"{arr[2]:02}/{arr[1]:02}/{arr[0]}" 
                    
                    def fmt_timestamp(ts):
                        if not ts: return ""
                        try:
                            if isinstance(ts, str): return ts # already string?
                            dt = datetime.fromtimestamp(int(ts) / 1000)
                            return dt.strftime("%d/%m/%Y")
                        except: return str(ts)
                    
                    row = {}
                    row["Tên đơn vị (đầy đủ)"] = get_d("orgFullName")
                    row["Tên đơn vị (Tiếng Anh)"] = get_d("orgEnName")
                    row["Mã định danh"] = org_code
                    
                    b_type = get_d("businessType")
                    row["Tình hình pháp lý"] = business_type_map.get(b_type, b_type)
                    
                    row["Mã số thuế"] = get_d("taxCode")
                    
                    tax_date = detail_info.get("taxDate")
                    row["Ngày cấp"] = fmt_timestamp(tax_date)
                    
                    # Tax nation: API 2 often has it, or API 1
                    tax_nation = detail_info.get("taxNation") 
                    if not tax_nation: tax_nation = item.get("taxNation")
                    row["Quốc gia cấp"] = country_map.get(tax_nation, tax_nation)
                    
                    eff_date = item.get("effRoleDate")
                    row["Ngày phê duyệt yêu cầu đăng ký"] = fmt_date_arr(eff_date)
                    
                    st = item.get("status")
                    row["Trạng thái vai trò"] = "Đang hoạt động" if str(st) == "1" else str(st)
                    
                    ag_code = get_d("agencyName") 
                    row["Tên cơ quan chủ quản"] = cqcq_map.get(ag_code, ag_code)
                    
                    row["Mã quan hệ ngân sách"] = get_d("budgetCode")
                    
                    off_pro = item.get("officePro") or detail_info.get("officePro")
                    off_dis = item.get("officeDis") or detail_info.get("officeDis")
                    
                    row["Tỉnh/ thành phố"] = get_area_name(off_pro)
                    row["Phường/ Xã/ Thị Trấn"] = get_area_name(off_dis)
                    
                    row["Số nhà, đường phố/ Xóm/ Ấp/ Thôn"] = item.get("officeAdd") or detail_info.get("officeAdd")
                    
                    # New columns requested: Phone and Email
                    row["Số điện thoại"] = item.get("officePhone") or detail_info.get("officePhone")
                    row["Email"] = item.get("recEmail") or detail_info.get("recEmail") or detail_info.get("email")

                    row["Web"] = item.get("officeWeb") or detail_info.get("officeWeb")
                    
                    row["Họ và tên"] = get_d("repName")
                    row["Chức vụ"] = detail_info.get("repPosition")
                    
                    all_rows.append(row)
                    processed_org_codes.add(str(org_code))
                    item_buffer.append(row)
                    
                    # Simplified log with Ministry context if applicable
                    log_prefix = f"[{m_name}] " if m_name != "All" else ""
                    print(f"    {log_prefix}Collected: {row['Tên đơn vị (đầy đủ)']}")
                    
                    # Save every 10 items
                    if len(item_buffer) >= 10:
                        df = pd.DataFrame(all_rows)
                        try:
                            df.to_excel(output_path, index=False)
                            print(f"    Auto-saved {len(all_rows)} items to {output_path}")
                            item_buffer = []
                        except PermissionError:
                             # Auto-backup logic
                             # import time  <-- REMOVED to avoid shadowing
                             ts = int(time.time())
                             base, ext = os.path.splitext(output_path)
                             # If already a backup, just keep using it or make new one? 
                             # Simpler: Update output_path to a new name and retry ONCE
                             if "_backup_" not in output_path:
                                 new_path = f"{base}_backup_{ts}{ext}"
                             else:
                                 # Already a backup, maybe make another or overwrite?
                                 # Let's make unique
                                 new_path = f"{base}_{ts}{ext}"
                             
                             print(f"    WARNING: File {output_path} is open/locked. Switching to backup: {new_path}")
                             output_path = new_path
                             
                             try:
                                 df.to_excel(output_path, index=False)
                                 print(f"    Auto-saved to backup {output_path}")
                                 item_buffer = []
                             except Exception as e2:
                                 print(f"    Backup save failed: {e2}")

                        except Exception as ex:
                            print(f"    Save error: {ex}")
            
            except Exception as e:
                print(f"  Page error: {e}")
                
            page_idx += 1
            time.sleep(1)

    # Final Save
    if item_buffer or all_rows:
        df = pd.DataFrame(all_rows)
        try:
            df.to_excel(output_path, index=False)
            print(f"Final save completed. Total {len(all_rows)} items to {output_path}")
        except PermissionError:
             # Final save backup logic
             # import time <-- REMOVED
             ts = int(time.time())
             base, ext = os.path.splitext(output_path)
             if "_backup_" not in output_path:
                 new_path = f"{base}_backup_{ts}{ext}"
             else:
                 new_path = f"{base}_{ts}{ext}"
             
             print(f"    WARNING: File {output_path} is open/locked. Saving final to backup: {new_path}")
             try:
                 df.to_excel(new_path, index=False)
                 print(f"Final save completed to {new_path}")
             except Exception as ex:
                 print(f"Final save backup failed: {ex}")
        except Exception as e:
            print(f"Final save error: {e}")




