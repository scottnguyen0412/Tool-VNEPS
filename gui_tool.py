import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import threading
import sys
import os
import json
import urllib.request
import subprocess
import time
import shutil
import scrape_muasamcong

# Configuration
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Resource helper for PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Sửa mỗi khi release
CURRENT_VERSION = "v1.3.4"
REPO_OWNER = "scottnguyen0412"
REPO_NAME = "Tool-VNEPS"

class ScraperApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Configuration
        self.title(f"VN-EPS SCRAPER ({CURRENT_VERSION})")
        self.geometry("900x700")
        
        # Main Grid Layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 1. Header Frame (Top Bar)
        self.header_frame = ctk.CTkFrame(self, height=80, corner_radius=0, fg_color=("gray90", "#2B2B2B"))
        self.header_frame.grid(row=0, column=0, sticky="ew")
        
        # Logo Logic
        self.logo_img = None 
        try:
            logo_path = resource_path("Image/BSTPharma_Logo.png")
            if os.path.exists(logo_path):
                pil_img = Image.open(logo_path)
                # Resize to height 50, maintain aspect ratio
                base_height = 50
                w_percent = (base_height / float(pil_img.size[1]))
                w_size = int((float(pil_img.size[0]) * float(w_percent)))
                
                self.logo_img = ctk.CTkImage(light_image=pil_img, size=(w_size, base_height))
                self.logo_label = ctk.CTkLabel(self.header_frame, text="", image=self.logo_img)
                self.logo_label.pack(side="left", padx=20, pady=10)
        except Exception as e:
            print(f"Logo load error: {e}")

        # Title Group
        self.header_text = ctk.CTkFrame(self.header_frame, fg_color="transparent")
        self.header_text.pack(side="left", pady=10)

        self.header_label = ctk.CTkLabel(self.header_text, text="VN-EPS DATA SCRAPER", 
                                       font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"))
        self.header_label.pack(anchor="w")
        
        self.version_label = ctk.CTkLabel(self.header_text, text=f"{CURRENT_VERSION}", 
                                        text_color="gray", font=ctk.CTkFont(size=12))
        self.version_label.pack(anchor="w")

        self.update_btn = ctk.CTkButton(self.header_frame, text="Check Updates", width=120, height=32,
                                        fg_color="#2980B9", hover_color="#2471A3", text_color="white",
                                        command=self.check_for_updates_thread)
        self.update_btn.pack(side="right", padx=20)

        # 2. Main Content Area
        self.content_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew", padx=25, pady=20)
        self.content_frame.grid_columnconfigure(0, weight=1)
        self.content_frame.grid_rowconfigure(2, weight=1)

        # --- Settings Card ---
        self.settings_card = ctk.CTkFrame(self.content_frame, corner_radius=10)
        self.settings_card.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        self.settings_card.grid_columnconfigure(1, weight=1)

        # Path Input
        self.lbl_path = ctk.CTkLabel(self.settings_card, text="File Save Path:", font=ctk.CTkFont(weight="bold"))
        self.lbl_path.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        
        default_path = os.path.join(os.getcwd(), "investors_data_detailed.xlsx")
        self.path_entry = ctk.CTkEntry(self.settings_card, placeholder_text="Path to save .xlsx file", height=35)
        self.path_entry.grid(row=0, column=1, padx=10, pady=20, sticky="ew")
        self.path_entry.insert(0, default_path)

        self.browse_btn = ctk.CTkButton(self.settings_card, text="Browse", width=100, height=35, 
                                      fg_color="#008A80", hover_color="#006960", text_color="white",
                                      command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, padx=20, pady=20)

        # --- TABS for Filter ---
        self.tab_view = ctk.CTkTabview(self.settings_card, height=130)
        self.tab_view.grid(row=1, column=0, columnspan=3, padx=20, pady=(0, 10), sticky="ew")
        
        self.tab_all = self.tab_view.add("Cào Toàn Bộ") # Tab 1
        self.tab_filter = self.tab_view.add("Cào Theo Bộ") # Tab 2
        
        # Tab 1 Content
        ctk.CTkLabel(self.tab_all, text="Chế độ này sẽ cào tất cả dữ liệu (Không lọc theo Bộ).", text_color="gray").pack(pady=20)
        
        # Tab 2 Content
        self.ministry_frame = ctk.CTkFrame(self.tab_filter, fg_color="transparent")
        self.ministry_frame.pack(fill="x", padx=20, pady=10)
        
        self.ministries_list = ["Bộ Y tế", "Bộ Quốc phòng", "Bộ Công an"]
        
        ctk.CTkLabel(self.ministry_frame, text="Chọn Bộ bắt đầu:", font=ctk.CTkFont(weight="bold")).pack(side="left")
        
        self.combo_ministry = ctk.CTkComboBox(self.ministry_frame, values=self.ministries_list, width=200, state="readonly")
        self.combo_ministry.pack(side="left", padx=10)
        self.combo_ministry.set("Bộ Y tế")
        
        self.chk_sequential = ctk.CTkCheckBox(self.ministry_frame, text="Cào tuần tự các bộ tiếp theo")
        self.chk_sequential.pack(side="left", padx=20)
        self.chk_sequential.select() # Default selected

        # Limit Input (Global)
        self.limit_frame = ctk.CTkFrame(self.settings_card, fg_color="transparent")
        self.limit_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=20, pady=(0, 20))
        
        self.lbl_limit = ctk.CTkLabel(self.limit_frame, text="Giới hạn trang:", font=ctk.CTkFont(weight="bold"))
        self.lbl_limit.pack(side="left", padx=(20, 10))
        
        self.limit_entry = ctk.CTkEntry(self.limit_frame, width=100, height=30)
        self.limit_entry.pack(side="left")
        self.limit_entry.insert(0, "0")
        
        self.lbl_hint = ctk.CTkLabel(self.limit_frame, text="(0 = Scrape ALL)", text_color="gray")
        self.lbl_hint.pack(side="left", padx=10)

        # --- Action Section (Button & Progress) ---
        self.action_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.action_frame.grid(row=1, column=0, sticky="ew", pady=(0, 20))
        
        self.start_btn = ctk.CTkButton(self.action_frame, text="START SCRAPING", height=55,
                                       font=ctk.CTkFont(size=18, weight="bold"),
                                       fg_color="#008A80", hover_color="#006960", text_color="white", corner_radius=8,
                                       command=self.start_scraping)
        self.start_btn.pack(fill="x")

        self.progress_bar = ctk.CTkProgressBar(self.action_frame, height=12, mode="indeterminate", progress_color="#4CC144")
        # Pack later when running

        # --- Log Section ---
        self.log_frame = ctk.CTkFrame(self.content_frame, corner_radius=10)
        self.log_frame.grid(row=2, column=0, sticky="nsew")
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        self.log_header = ctk.CTkLabel(self.log_frame, text="Execution Logs", font=ctk.CTkFont(size=14, weight="bold"))
        self.log_header.grid(row=0, column=0, padx=15, pady=10, sticky="w")
        
        self.log_area = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12), fg_color=("gray95", "#1E1E1E"))
        self.log_area.grid(row=1, column=0, padx=15, pady=(0, 15), sticky="nsew")
        self.log_area.configure(state="disabled")

        # 3. Footer Area
        self.footer_frame = ctk.CTkFrame(self, height=30, fg_color="transparent")
        self.footer_frame.grid(row=2, column=0, sticky="ew", padx=25, pady=10)
        
        self.status_label = ctk.CTkLabel(self.footer_frame, text="Status: Ready", anchor="w", font=ctk.CTkFont(weight="bold"))
        self.status_label.pack(side="left")
        
        self.footer_branding = ctk.CTkLabel(self.footer_frame, text="Made with ❤️ by IT Boston", 
                                          font=ctk.CTkFont(size=12, slant="italic"), text_color="gray")
        self.footer_branding.pack(side="right")

        # System Redirects
        sys.stdout = self
        sys.stderr = self
        
        # Startup Tasks
        self.after(2000, self.check_for_updates_thread)

    def browse_file(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            initialfile="investors_data_detailed.xlsx",
            title="Save Output As"
        )
        if filename:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, filename)

    def write(self, text):
        self.after(0, self._safe_write, text)

    def _safe_write(self, text):
        try:
            self.log_area.configure(state="normal")
            self.log_area.insert(tk.END, text)
            self.log_area.see(tk.END)
            self.log_area.configure(state="disabled")
        except:
            pass

    def flush(self):
        pass

    def start_scraping(self):
        output_path = self.path_entry.get().strip()
        limit_str = self.limit_entry.get().strip()

        if not output_path:
            messagebox.showerror("Error", "Please specify a save file path!")
            return

        try:
            max_pages = int(limit_str)
            if max_pages <= 0:
                max_pages = float('inf')
        except ValueError:
            messagebox.showerror("Error", "Page Limit must be an integer!")
            return

        # UI Updates
        self.start_btn.configure(state="disabled", text="RUNNING...", fg_color="#e67e22")
        self.path_entry.configure(state="disabled")
        self.limit_entry.configure(state="disabled")
        self.status_label.configure(text="Status: Scraping in progress... Please do not close the browser window.", text_color="#E67E22")
        
        # Show Progress Bar
        self.progress_bar.pack(fill="x", pady=(15, 0))
        self.progress_bar.start()

        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", tk.END)
        self.log_area.configure(state="disabled")

        # Get Filter (Ministry)
        current_tab = self.tab_view.get()
        start_ministry = ""
        is_sequential = False
        
        if current_tab == "Cào Theo Bộ":
            start_ministry = self.combo_ministry.get()
            is_sequential = True if self.chk_sequential.get() == 1 else False

        # Threading
        t = threading.Thread(target=self.run_process, args=(output_path, max_pages, start_ministry, is_sequential))
        t.daemon = True
        t.start()
    
    def run_process(self, output_path, max_pages, start_ministry, is_sequential):
        try:
            print(">>> INITIALIZING SCRAPER...")
            print(f"> Target File: {output_path}")
            print(f"> Page Limit:  {'Unlimited' if max_pages == float('inf') else max_pages}")
            
            ministries_to_scrape = []
            
            # Custom Sequences
            ministry_sequences = {
                "Bộ Y tế": ["Bộ Y tế", "Bộ Quốc phòng", "Bộ Công an"],
                "Bộ Quốc phòng": ["Bộ Quốc phòng", "Bộ Công an", "Bộ Y tế"],
                "Bộ Công an": ["Bộ Công an", "Bộ Quốc phòng", "Bộ Y tế"]
            }
            
            # Determine list of ministries
            if start_ministry: # Filter Mode
                if start_ministry in ministry_sequences:
                    if is_sequential:
                        ministries_to_scrape = ministry_sequences[start_ministry]
                        print(f"> Mode: SEQUENTIAL starting from {start_ministry}")
                        print(f"> Queue: { ' -> '.join(ministries_to_scrape) }")
                    else:
                        ministries_to_scrape = [start_ministry]
                        print(f"> Mode: SINGLE Ministry ({start_ministry})")
                else:
                    ministries_to_scrape = [start_ministry]
            else:
                # ALL Mode
                print("> Mode: ALL DATA (No Filter)")
                ministries_to_scrape = [""] 
            
            print("-" * 50)
            
            # Main Loop
            for m_idx, current_m in enumerate(ministries_to_scrape):
                if current_m:
                    print(f"\n[{m_idx+1}/{len(ministries_to_scrape)}] Processing Ministry: {current_m}")
                
                # Special Logic for Keywords
                keywords_for_ministry = [""]
                if current_m in ["Bộ Công an", "Bộ Quốc phòng"]:
                    keywords_for_ministry = ["bệnh viện", "y tế"]
                
                for kw in keywords_for_ministry:
                    if current_m:
                        if kw:
                             print(f"   > Filter: {current_m} | Keyword: '{kw}'")
                        else:
                             print(f"   > Filter: {current_m} (No extra keyword)")
                    
                    scrape_muasamcong.run(
                        output_path=output_path, 
                        max_pages=max_pages, 
                        ministry_filter=current_m,
                        search_keyword=kw
                    )
                    
                    if kw != keywords_for_ministry[-1]:
                        time.sleep(2)
                
                if m_idx < len(ministries_to_scrape) - 1:
                    print(f"--- Finished {current_m}. Moving to next ministry in 3s... ---")
                    time.sleep(3)
            
            print("\n>>> COMPLETED SUCCESSFULLY!")
            self.status_label.configure(text="Status: Completed ✅")
            messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
            
        except Exception as e:
            print(f"\n>>> ERROR: {e}")
            self.status_label.configure(text="Status: Error ❌", text_color="red")
            messagebox.showerror("Error", f"An error occurred:\n{e}")
        finally:
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.start_btn.configure(state="normal", text="START SCRAPING", fg_color="#008A80")
        self.path_entry.configure(state="normal")
        self.limit_entry.configure(state="normal")
        self.tab_view.configure(state="normal") # Enable tabs
        self.tab_view.configure(state="normal") # Enable tabs
        
        # Stop Progress
        self.progress_bar.stop()
        self.progress_bar.pack_forget()

    # --- UPDATE LOGIC ---
    def check_for_updates_thread(self):
        t = threading.Thread(target=self.check_update)
        t.daemon = True
        t.start()

    def check_update(self):
        print("Checking for updates...")
        try:
            api_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
            with urllib.request.urlopen(api_url) as response:
                data = json.loads(response.read().decode())
                latest_tag = data.get("tag_name", "")
                
                # Check version
                if latest_tag and latest_tag != CURRENT_VERSION:
                    print(f"New version found: {latest_tag}")
                    
                    exe_asset = None
                    for asset in data.get("assets", []):
                        if asset["name"].endswith(".exe"):
                            exe_asset = asset
                            break
                    
                    if exe_asset:
                        # VISUAL NOTIFICATION
                        url = exe_asset["browser_download_url"]
                        self.after(0, lambda: self.indicate_update_found(latest_tag, url))
                    else:
                        print("No EXE asset found in release.")
                else:
                    self.after(0, lambda: self.update_btn.configure(text="Latest Version ✅", fg_color="#27AE60"))
                    print("You are using the latest version.")
        except Exception as e:
            print(f"Update check failed: {e}")

    def indicate_update_found(self, version, url):
        # Change button to Red Alert
        self.update_btn.configure(text=f"UPDATE {version} 🔔", fg_color="#C0392B", hover_color="#922B21")
        # Prompt user
        self.prompt_update(version, url)

    def prompt_update(self, version, url):
        ans = messagebox.askyesno("Update Available", f"A new version {version} is available.\nDo you want to update now?")
        if ans:
            self.start_update(url)

    def start_update(self, url):
        self.status_label.configure(text="Updating... DO NOT CLOSE")
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, "\n>>> STARTING UPDATE...\n")
        
        t = threading.Thread(target=self.download_and_swap, args=(url,))
        t.daemon = True
        t.start()

    def download_and_swap(self, url):
        try:
            print(f"Downloading update from {url}...")
            new_exe_name = "Tool_Scrape_Muasamcong_new.exe"
            
            # Use chunks
            req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req) as response, open(new_exe_name, 'wb') as out_file:
                shutil.copyfileobj(response, out_file)
                
            print("Download completed.")
            
            # Create Update Script
            current_exe = sys.executable
            
            if getattr(sys, 'frozen', False):
                exe_name = os.path.basename(current_exe)
                bat_script = f"""
@echo off
timeout /t 3 /nobreak
del "{exe_name}"
move "{new_exe_name}" "{exe_name}"
start "" "{exe_name}"
del "%~f0"
"""
                with open("update_tool.bat", "w") as f:
                    f.write(bat_script)
                
                print("restarting app...")
                subprocess.Popen("update_tool.bat", shell=True)
                os._exit(0)
            else:
                 print("Dev mode: Downloaded new exe but cannot auto-replace python script.")
                 messagebox.showinfo("Dev Mode", f"Downloaded {new_exe_name}. Cannot auto-update .py file.")
                 
        except Exception as e:
            print(f"Update failed: {e}")
            messagebox.showerror("Update Failed", str(e))

if __name__ == "__main__":
    app = ScraperApp()
    app.mainloop()
