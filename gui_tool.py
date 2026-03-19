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
import requests
import scrape_muasamcong
from gui_icons import IconLib

# Configuration
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

# === DESIGN SYSTEM (Light Mode) ===
COLORS = {
    "primary": "#00897B",
    "primary_dark": "#00695C",
    "primary_light": "#B2DFDB",
    "accent": "#00897B",
    "surface": "#EDF2F7",
    "surface_light": "#FFFFFF",
    "card": "#FFFFFF",
    "card_hover": "#E8EDF2",
    "success": "#2E7D32",
    "warning": "#E65100",
    "danger": "#C62828",
    "text": "#1A202C",
    "text_secondary": "#4A5568",
    "border": "#CBD5E0",
    "sidebar_bg": "#00695C",
    "sidebar_text": "#FFFFFF",
    "sidebar_accent": "#B2DFDB",
    "input_bg": "#EDF2F7",
    "log_bg": "#F7FAFC",
    "log_text": "#1A202C",
}

# Resource helper for PyInstaller
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Sửa mỗi khi release
CURRENT_VERSION = "v2.0.7"
REPO_OWNER = "scottnguyen0412"
REPO_NAME = "Tool-VNEPS"

class AnimatedGradientBorderFrame(ctk.CTkFrame):
    def __init__(self, master, border_width=3, animation_speed=20, colors=None, **kwargs):
        super().__init__(master, **kwargs)
        self.border_width = border_width
        self.animation_speed = animation_speed
        
        if colors:
            self.colors = colors
        else:
            # Professional Teal Gradient
            self.colors = [
                (0, 137, 123), (0, 105, 92), (38, 166, 154),
                (77, 182, 172), (0, 121, 107), (0, 150, 136)
            ]
        self.current_idx = 0
        self.t = 0.0 # Interpolation factor 0.0 to 1.0
        
        self.configure(fg_color=self._rgb_to_hex(self.colors[0]))
        self.animate()

    def _rgb_to_hex(self, rgb):
        return '#{:02x}{:02x}{:02x}'.format(int(rgb[0]), int(rgb[1]), int(rgb[2]))

    def animate(self):
        c1 = self.colors[self.current_idx]
        c2 = self.colors[(self.current_idx + 1) % len(self.colors)]
        
        # Interpolate
        r = c1[0] + (c2[0] - c1[0]) * self.t
        g = c1[1] + (c2[1] - c1[1]) * self.t
        b = c1[2] + (c2[2] - c1[2]) * self.t
        
        color_hex = self._rgb_to_hex((r, g, b))
        try:
             self.configure(fg_color=color_hex)
        except: pass # Handle window closed
        
        self.t += 0.02
        if self.t >= 1.0:
            self.t = 0.0
            self.current_idx = (self.current_idx + 1) % len(self.colors)
            
        self.after(self.animation_speed, self.animate)

class LoginWindow(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("")
        self.geometry("440x420")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["surface_light"])
        
        # Center Window
        self.update_idletasks()
        try:
            s_w = self.winfo_screenwidth()
            s_h = self.winfo_screenheight()
            x = int((s_w - 440) / 2)
            y = int((s_h - 420) / 2)
            self.geometry(f"+{x}+{y}")
        except: pass
        
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.attributes("-topmost", True)
        
        # Logo
        try:
            logo_path = resource_path("Image/BSTPharma_Logo.png")
            if os.path.exists(logo_path):
                pil_img = Image.open(logo_path)
                base_height = 50
                w_percent = (base_height / float(pil_img.size[1]))
                w_size = int((float(pil_img.size[0]) * float(w_percent)))
                self.login_logo = ctk.CTkImage(light_image=pil_img, size=(w_size, base_height))
                ctk.CTkLabel(self, text="", image=self.login_logo).pack(pady=(30, 5))
        except: pass
        
        # Title
        ctk.CTkLabel(self, text="ĐĂNG NHẬP HỆ THỐNG", 
                     font=ctk.CTkFont(family="Segoe UI", size=20, weight="bold"),
                     text_color=COLORS["primary"]).pack(pady=(10, 5))
        ctk.CTkLabel(self, text="Muasamcong Data Scraper",
                     font=ctk.CTkFont(size=12), text_color=COLORS["text_secondary"]).pack(pady=(0, 20))
        
        # Form container
        form = ctk.CTkFrame(self, fg_color="transparent")
        form.pack(padx=50, fill="x")
        
        icon_lib = IconLib(resource_path("Image/MaterialIcons-Regular.ttf"))
        icon_person = icon_lib.get_icon("person", size=18, light_color=COLORS["text_secondary"], dark_color=COLORS["text_secondary"])
        icon_lock = icon_lib.get_icon("lock", size=18, light_color=COLORS["text_secondary"], dark_color=COLORS["text_secondary"])
        icon_login = icon_lib.get_icon("login", size=20, light_color="white", dark_color="white")
        
        ctk.CTkLabel(form, text=" Tên đăng nhập", image=icon_person, compound="left", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLORS["text_secondary"], anchor="w").pack(fill="x", pady=(0,4))
        self.entry_user = ctk.CTkEntry(form, placeholder_text="Nhập username", height=42,
                                       corner_radius=8, fg_color=COLORS["input_bg"],
                                       border_color=COLORS["border"], border_width=1)
        self.entry_user.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(form, text=" Mật khẩu", image=icon_lock, compound="left", font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COLORS["text_secondary"], anchor="w").pack(fill="x", pady=(0,4))
        self.entry_pass = ctk.CTkEntry(form, placeholder_text="Nhập mật khẩu", show="•", height=42,
                                       corner_radius=8, fg_color=COLORS["input_bg"],
                                       border_color=COLORS["border"], border_width=1)
        self.entry_pass.pack(fill="x", pady=(0, 20))
        self.entry_pass.bind("<Return>", self.login_event)
        
        self.btn_login = ctk.CTkButton(form, text="ĐĂNG NHẬP", image=icon_login, compound="right", height=44, corner_radius=8,
                                        font=ctk.CTkFont(size=14, weight="bold"),
                                        fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"],
                                        command=self.check_login)
        self.btn_login.pack(fill="x", pady=(0, 15))
        
        ctk.CTkLabel(self, text="Chưa có tài khoản? Liên hệ IT Boston Pharma", 
                     text_color=COLORS["text_secondary"], font=ctk.CTkFont(size=11),
                     wraplength=360).pack(pady=(0, 15))
        
        self.entry_user.focus()

    def login_event(self, event):
        self.check_login()

    def check_login(self):
        u = self.entry_user.get().strip()
        p = self.entry_pass.get().strip()
        
        if not u or not p:
            messagebox.showwarning("Input", "Please enter username and password", parent=self)
            return

        # API Auth
        api_url = "https://dashboard-vneps.vercel.app/api/auth/login"
        try:
            # Attempt Login via API
            resp = requests.post(api_url, json={"username": u, "password": p}, timeout=5)
            
            if resp.status_code == 200:
                # Success
                data = resp.json()
                # Optional: Store token? self.master.token = data.get("token")
                self.master.deiconify() 
                self.destroy()
                return
            else:
                # Login Failed
                try: 
                    msg = resp.json().get("detail") or resp.json().get("message") or "Login failed"
                except: msg = "Invalid credentials"
                messagebox.showerror("Login Failed", f"{msg}", parent=self)
                return

        except requests.exceptions.ConnectionError:
            # Fallback for offline/admin mode
            print("Auth Server unreachable. Checking local admin...")
            if u == "admin" and p == "admin123":
                 if messagebox.askyesno("Offline Mode", "Could not connect to Auth Server.\nContinue in local Admin mode?", parent=self):
                     self.master.deiconify()
                     self.destroy()
                     return
            
            messagebox.showerror("Connection Error", "Could not connect to Authentication Server (localhost:3000).", parent=self)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}", parent=self)

    def on_close(self):
        self.master.destroy()
        sys.exit()

class ScraperApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Login Logic
        self.withdraw()
        LoginWindow(self)

        # Window Configuration
        self.title(f"Tool Muasamcong Scrape Data ({CURRENT_VERSION})")
        self.geometry("1350x800")
        self.configure(fg_color=COLORS["surface"])
        
        # Set Icon
        try:
             icon_path = resource_path("Image/BST_Pharma_ICO.ico")
             if os.path.exists(icon_path):
                 self.iconbitmap(icon_path)
        except Exception as e:
             print(f"Warning: Could not set icon ({e})")

        # Control Logic
        self.icon_lib = IconLib(resource_path('Image/MaterialIcons-Regular.ttf'))
        self.pause_event = threading.Event()
        self.pause_event.set()
        self.stop_event = threading.Event()
        
        # ═══ MAIN LAYOUT: Left Panel | Right Panel (Logs) ═══
        self.grid_columnconfigure(0, weight=3)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(0, weight=1)

        # ═══ LEFT PANEL (Unified: Header + Settings + Actions) ═══
        self.left_panel = ctk.CTkFrame(self, fg_color=COLORS["card"], corner_radius=0)
        self.left_panel.grid(row=0, column=0, sticky="nsew")
        self.left_panel.grid_rowconfigure(3, weight=1)  # Tabs expand
        self.left_panel.grid_columnconfigure(0, weight=1)

        # --- Header Bar (White) ---
        self.header_bar = ctk.CTkFrame(self.left_panel, height=56, corner_radius=0,
                                        fg_color=COLORS["card"], border_width=0)
        self.header_bar.grid(row=0, column=0, sticky="ew")
        self.header_bar.grid_propagate(False)
        
        # Bottom Separator for Header
        header_separator = ctk.CTkFrame(self.left_panel, height=1, fg_color=COLORS["border"], corner_radius=0)
        header_separator.grid(row=1, column=0, sticky="ew")
        
        self.logo_img = None 
        try:
            logo_path = resource_path("Image/BSTPharma_Logo.png")
            if os.path.exists(logo_path):
                pil_img = Image.open(logo_path)
                base_height = 42
                w_percent = (base_height / float(pil_img.size[1]))
                w_size = int((float(pil_img.size[0]) * float(w_percent)))
                self.logo_img = ctk.CTkImage(light_image=pil_img, size=(w_size, base_height))
                ctk.CTkLabel(self.header_bar, text="", image=self.logo_img).pack(side="left", padx=(15, 12))
        except Exception as e:
            print(f"Logo load error: {e}")

        title_frame = ctk.CTkFrame(self.header_bar, fg_color="transparent")
        title_frame.pack(side="left")
        ctk.CTkLabel(title_frame, text="MUASAMCONG SCRAPER",
                     font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
                     text_color=COLORS["primary_dark"]).pack(anchor="w")
        ctk.CTkLabel(title_frame, text=f"{CURRENT_VERSION}",
                     font=ctk.CTkFont(size=10), text_color=COLORS["text_secondary"]).pack(anchor="w")

        # Status in header
        status_frame = ctk.CTkFrame(self.header_bar, fg_color="transparent")
        status_frame.pack(side="right", padx=15)
        
        self.status_dot = ctk.CTkLabel(status_frame, text="●", font=ctk.CTkFont(size=12),
                                       text_color=COLORS["success"])
        self.status_dot.pack(side="left", padx=(0, 4))
        self.status_label = ctk.CTkLabel(status_frame, text="Ready",
                                         font=ctk.CTkFont(size=11, weight="bold"),
                                         text_color=COLORS["text"])
        self.status_label.pack(side="left", padx=(0, 10))
        self.timer_label = ctk.CTkLabel(status_frame, text="00:00",
                                        font=ctk.CTkFont(family="Consolas", size=11),
                                        text_color=COLORS["text_secondary"])
        self.timer_label.pack(side="left")

        self.update_btn = ctk.CTkButton(self.header_bar, text="", height=28, width=28,
                                        image=self.icon_lib.get_icon("update", 16, COLORS["text"], COLORS["text"]),
                                        corner_radius=6,
                                        fg_color=COLORS["surface"], hover_color=COLORS["border"],
                                        text_color=COLORS["text"], font=ctk.CTkFont(size=12),
                                        command=self.check_for_updates_thread)
        self.update_btn.pack(side="right", padx=(0, 5), pady=10)

        # --- File Path Row ---
        path_container = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        path_container.grid(row=2, column=0, sticky="ew", padx=12, pady=(10, 5))
        
        self.lbl_path = ctk.CTkLabel(path_container, text="Save Path:",
                                     font=ctk.CTkFont(size=12, weight="bold"),
                                     text_color=COLORS["text"])
        self.lbl_path.pack(side="left", padx=(0, 8))
        
        default_path = os.path.join(os.getcwd(), "investors_data_detailed.xlsx")
        self.path_entry = ctk.CTkEntry(path_container, placeholder_text="Path to .xlsx",
                                       height=34, corner_radius=8,
                                       fg_color=COLORS["input_bg"], border_color=COLORS["border"],
                                       border_width=1, text_color=COLORS["text"],
                                       font=ctk.CTkFont(size=11))
        self.path_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.path_entry.insert(0, default_path)

        self.browse_btn = ctk.CTkButton(path_container, text=" Browse", width=75, height=34,
                                      image=self.icon_lib.get_icon("folder", 16, "white", "white"),
                                      corner_radius=8,
                                      fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"],
                                      text_color="white", font=ctk.CTkFont(size=11, weight="bold"),
                                      command=self.browse_file)
        self.browse_btn.pack(side="right")

        # --- Settings Tabs (Expanding Area) ---
        # Dark teal tab bar → white text readable on all states, active tab pops
        self.tab_view = ctk.CTkTabview(self.left_panel, command=self.on_tab_change,
                                        fg_color=COLORS["card"],
                                        segmented_button_fg_color="#004D40",
                                        segmented_button_selected_color=COLORS["primary"],
                                        segmented_button_selected_hover_color="#26A69A",
                                        segmented_button_unselected_color="#004D40",
                                        segmented_button_unselected_hover_color="#00695C",
                                        corner_radius=8)
        self.tab_view.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 5))
        
        self.tab_view._segmented_button.configure(
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), height=38,
            text_color="#FFFFFF",
            selected_color=COLORS["primary"],
            selected_hover_color="#26A69A",
            unselected_color="#004D40",
            unselected_hover_color="#00695C",
            corner_radius=6)
        
        self.tab_all = self.tab_view.add("Thông Tin Nhà Đầu Tư") # Merged Tab
        # self.tab_filter Removed
        self.tab_contractor = self.tab_view.add("Kết Quả Đấu Thầu") # Tab 2
        self.tab_rfq = self.tab_view.add("Yêu cầu báo giá") # Tab 4
        self.tab_drug = self.tab_view.add("Công bố giá thuốc") # Tab 3
        
        # --- Tab 2 Content (Contractor Results) ---
        self.contractor_frame = ctk.CTkFrame(self.tab_contractor, fg_color="transparent")
        self.contractor_frame.pack(fill="both", expand=True)

        # Tab Segmented Button
        top_bar = ctk.CTkFrame(self.contractor_frame, fg_color="transparent")
        top_bar.pack(fill="x", padx=12, pady=(10, 5))
        self.contractor_mode_seg = ctk.CTkSegmentedButton(top_bar, values=["Tìm theo bộ lọc", "Tìm theo danh sách IB (Excel/Nhập)"], command=self.on_contractor_mode_change)
        self.contractor_mode_seg.pack(fill="x")

        # --- Mode 1: Filter Frame ---
        self.filter_frame = ctk.CTkFrame(self.contractor_frame, fg_color="transparent")
        
        # Section 1: Search Configuration
        search_section = ctk.CTkFrame(self.filter_frame, fg_color=COLORS["surface"], corner_radius=8)
        search_section.pack(fill="x", padx=12, pady=(5, 8))
        
        ctk.CTkLabel(search_section, text="Loại tìm theo:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text"]).pack(anchor="w", padx=12, pady=(10, 0))
        self.search_type_combo = ctk.CTkComboBox(
            search_section,
            values=["Thông báo mời thầu thuốc, dược liệu, vị thuốc cổ truyền", "Thông báo mời thầu"],
            state="readonly", width=400, height=34, corner_radius=6,
            fg_color=COLORS["card"], border_color=COLORS["border"], button_color=COLORS["primary"]
        )
        self.search_type_combo.pack(fill="x", padx=12, pady=(5, 10))
        self.search_type_combo.set("Thông báo mời thầu thuốc, dược liệu, vị thuốc cổ truyền")
        
        # Keywords Header with Auto-fill Button
        kw_header = ctk.CTkFrame(search_section, fg_color="transparent")
        kw_header.pack(fill="x", padx=12, pady=(0, 2))
        
        ctk.CTkLabel(kw_header, text="Từ khóa:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text"]).pack(side="left")
        
        def _fill_default_keywords():
            self.entry_keywords.delete(0, "end")
            self.entry_keywords.insert(0, "thuốc, generic, tân dược, biệt dược, bệnh viện, chữa bệnh, vật tư y tế, điều trị, bệnh nhân, thiết bị y tế, khám chữa bệnh, khám bệnh, chữa bệnh, dược liệu, dược")
            self.entry_exclude.delete(0, "end")
            self.entry_exclude.insert(0, "linh kiện, xây dựng, cải tạo, lắp đặt, thi công")

        self.btn_default_kw = ctk.CTkButton(
            kw_header, text=" Bộ từ khóa mẫu", image=self.icon_lib.get_icon("bolt", 16, COLORS["primary_dark"], COLORS["primary_dark"]),
            width=140, height=24, font=ctk.CTkFont(size=11, weight="bold"), corner_radius=4,
            fg_color="#E0F2F1", hover_color="#B2DFDB",
            text_color=COLORS["primary_dark"],
            command=_fill_default_keywords
        )
        self.btn_default_kw.pack(side="right")

        self.entry_keywords = ctk.CTkEntry(search_section, height=34, corner_radius=6, border_color=COLORS["border"], fg_color=COLORS["card"],
            placeholder_text="Ví dụ: thuốc, generic, tân dược, bệnh viện...")
        self.entry_keywords.pack(fill="x", padx=12, pady=(0, 10))
        
        ctk.CTkLabel(search_section, text="Loại trừ:", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text"]).pack(anchor="w", padx=12)
        self.entry_exclude = ctk.CTkEntry(search_section, height=34, corner_radius=6, border_color=COLORS["border"], fg_color=COLORS["card"],
            placeholder_text="Ví dụ: linh kiện, xây dựng, thi công...")
        self.entry_exclude.pack(fill="x", padx=12, pady=(2, 12))

        # Section 2: Date Range
        date_section = ctk.CTkFrame(self.filter_frame, fg_color=COLORS["surface"], corner_radius=8)
        date_section.pack(fill="x", padx=12, pady=(0, 8))
        
        ctk.CTkLabel(date_section, text="Khoảng thời gian", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLORS["text"]).pack(anchor="w", padx=12, pady=(10, 8))
        
        self.date_frame = ctk.CTkFrame(date_section, fg_color="transparent")
        self.date_frame.pack(fill="x", padx=12, pady=(0, 10))
        
        ctk.CTkLabel(self.date_frame, text="Từ ngày:", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text"]).pack(side="left", padx=(0, 5))
        self.entry_from_date = ctk.CTkEntry(self.date_frame, width=120, height=32, corner_radius=6, placeholder_text="dd/mm/yyyy", fg_color=COLORS["card"], border_color=COLORS["border"])
        self.entry_from_date.pack(side="left", padx=(0, 15))
        
        ctk.CTkLabel(self.date_frame, text="Đến ngày:", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text"]).pack(side="left", padx=(0, 5))
        self.entry_to_date = ctk.CTkEntry(self.date_frame, width=120, height=32, corner_radius=6, placeholder_text="dd/mm/yyyy", fg_color=COLORS["card"], border_color=COLORS["border"])
        self.entry_to_date.pack(side="left")

        # Note
        ctk.CTkLabel(self.filter_frame, text="* Tự động chọn Field: Hàng hóa, Search By: Thuốc/Dược liệu", 
                     text_color=COLORS["text_secondary"], font=ctk.CTkFont(size=11)).pack(anchor="w", padx=15, pady=(0, 5))

        # --- Mode 2: IB List Frame ---
        self.ib_frame = ctk.CTkFrame(self.contractor_frame, fg_color="transparent")
        
        ib_section = ctk.CTkFrame(self.ib_frame, fg_color=COLORS["surface"], corner_radius=8)
        ib_section.pack(fill="x", padx=12, pady=(5, 8))

        self.ib_action_frame = ctk.CTkFrame(ib_section, fg_color="transparent")
        self.ib_action_frame.pack(fill="x", padx=12, pady=(12, 5))
        ctk.CTkLabel(self.ib_action_frame, text="Danh sách mã IB (ngăn cách bởi dấu phẩy):", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text"]).pack(side="left")
        self.btn_upload_excel = ctk.CTkButton(self.ib_action_frame, text=" Tải lên Excel", width=150, height=30,
                                                image=self.icon_lib.get_icon("upload_file", 18, "white", "white"),
                                                command=self.upload_ib_excel,
                                                fg_color=COLORS["success"], hover_color="#1B5E20",
                                                corner_radius=6, font=ctk.CTkFont(size=12, weight="bold"))
        self.btn_upload_excel.pack(side="right")

        self.entry_ib_list = ctk.CTkTextbox(ib_section, height=120, fg_color=COLORS["card"],
                                              border_width=1, border_color=COLORS["border"], corner_radius=6, text_color=COLORS["text"])
        self.entry_ib_list.pack(fill="x", padx=12, pady=(0, 5))
        ctk.CTkLabel(ib_section, text="* Hỗ trợ tải lên file Excel có chứa cột mang tên 'IB' và lấy dữ liệu từng hàng.", 
                     text_color=COLORS["text_secondary"], font=ctk.CTkFont(size=11), justify="left").pack(anchor="w", padx=12, pady=(0, 12))

        self.contractor_mode_seg.set("Tìm theo bộ lọc")
        self.filter_frame.pack(fill="both", expand=True)
        
        # Tab 3 Content (Drug Price)
        self.drug_frame = ctk.CTkFrame(self.tab_drug, fg_color="transparent")
        self.drug_frame.pack(fill="both", expand=True)

        drug_banner = ctk.CTkFrame(self.drug_frame, fg_color=COLORS["sidebar_bg"], corner_radius=8, height=40)
        drug_banner.pack(fill="x", padx=12, pady=(10, 5))
        ctk.CTkLabel(drug_banner, text="Thu Thập Dữ Liệu Giá Thuốc (dichvucong.dav.gov.vn)", 
                     font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color="#FFFFFF").pack(padx=15, pady=10, fill="x")

        drug_section = ctk.CTkFrame(self.drug_frame, fg_color=COLORS["surface"], corner_radius=8)
        drug_section.pack(fill="x", padx=12, pady=(5, 8))
        
        ctk.CTkLabel(drug_section, text="Thông tin thu thập bao gồm:", font=ctk.CTkFont(weight="bold"), text_color=COLORS["text"]).pack(anchor="w", padx=15, pady=(15, 2))
        ctk.CTkLabel(drug_section, text="• Tên thuốc  • Hoạt chất  • Giá kê khai  • Cơ sở sản xuất/đăng ký", 
                     text_color=COLORS["text_secondary"], justify="left").pack(anchor="w", padx=25, pady=(0, 10))
                     
        ctk.CTkLabel(drug_section, text="Lưu ý: Quá trình sẽ chạy tuần tự từ trang đầu đến hết.", 
                     font=ctk.CTkFont(size=11, slant="italic", weight="bold"), text_color=COLORS["warning"]).pack(anchor="w", padx=15, pady=(0, 15))
        
        # Tab 4 Content (RFQ)
        self.rfq_frame = ctk.CTkFrame(self.tab_rfq, fg_color="transparent")
        self.rfq_frame.pack(fill="both", expand=True)

        rfq_banner = ctk.CTkFrame(self.rfq_frame, fg_color=COLORS["sidebar_bg"], corner_radius=8, height=40)
        rfq_banner.pack(fill="x", padx=12, pady=(10, 5))
        ctk.CTkLabel(rfq_banner, text="CHẾ ĐỘ: Thu thập dữ liệu Yêu Cầu Báo Giá", 
                     font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"), text_color="#FFFFFF").pack(padx=15, pady=10, fill="x")

        # Section 1: Keywords
        rfq_search_section = ctk.CTkFrame(self.rfq_frame, fg_color=COLORS["surface"], corner_radius=8)
        rfq_search_section.pack(fill="x", padx=12, pady=(5, 8))
        
        ctk.CTkLabel(rfq_search_section, text="Mã/Tên yêu cầu báo giá (hoặc từ khóa):", font=ctk.CTkFont(size=12, weight="bold"), text_color=COLORS["text"]).pack(anchor="w", padx=12, pady=(12, 2))
        self.entry_rfq_keywords = ctk.CTkEntry(rfq_search_section, height=34, corner_radius=6, border_color=COLORS["border"], fg_color=COLORS["card"])
        self.entry_rfq_keywords.pack(fill="x", padx=12, pady=(0, 15))
        self.entry_rfq_keywords.insert(0, "thuốc, generic, genegic")
        
        # Section 2: Date Range
        rfq_date_section = ctk.CTkFrame(self.rfq_frame, fg_color=COLORS["surface"], corner_radius=8)
        rfq_date_section.pack(fill="x", padx=12, pady=(0, 8))
        
        ctk.CTkLabel(rfq_date_section, text="Khoảng thời gian (Ngày hết hạn báo giá)", font=ctk.CTkFont(size=13, weight="bold"), text_color=COLORS["text"]).pack(anchor="w", padx=12, pady=(10, 8))
        
        self.rfq_date_frame = ctk.CTkFrame(rfq_date_section, fg_color="transparent")
        self.rfq_date_frame.pack(fill="x", padx=12, pady=(0, 10))
        
        ctk.CTkLabel(self.rfq_date_frame, text="Từ ngày:", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text"]).pack(side="left", padx=(0, 5))
        self.entry_rfq_from_date = ctk.CTkEntry(self.rfq_date_frame, width=120, height=32, corner_radius=6, placeholder_text="dd/mm/yyyy", fg_color=COLORS["card"], border_color=COLORS["border"])
        self.entry_rfq_from_date.pack(side="left", padx=(0, 15))
        
        ctk.CTkLabel(self.rfq_date_frame, text="Đến ngày:", font=ctk.CTkFont(size=11, weight="bold"), text_color=COLORS["text"]).pack(side="left", padx=(0, 5))
        self.entry_rfq_to_date = ctk.CTkEntry(self.rfq_date_frame, width=120, height=32, corner_radius=6, placeholder_text="dd/mm/yyyy", fg_color=COLORS["card"], border_color=COLORS["border"])
        self.entry_rfq_to_date.pack(side="left")
        
        ctk.CTkLabel(self.rfq_frame, text="Hệ thống tự động chọn:\n• Tìm Theo: Yêu Cầu Báo Giá\n• Khớp từ hoặc một số từ (Phân biệt dấu)", 
                     text_color=COLORS["text_secondary"], font=ctk.CTkFont(size=11), justify="left").pack(anchor="w", padx=15, pady=(0, 5))
        
        # Merged Tab Content (Investor Search) - Modern Design
        # Info Banner (Static, cleaner than animated)
        self.info_banner = ctk.CTkFrame(self.tab_all, fg_color=COLORS["sidebar_bg"],
                                         corner_radius=8, height=40)
        self.info_banner.pack(fill="x", padx=12, pady=(8, 10))
        
        self.lbl_investor_desc = ctk.CTkLabel(self.info_banner, text="", 
                                            text_color="#FFFFFF",
                                            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
                                            wraplength=600, justify="center")
        self.lbl_investor_desc.pack(padx=15, pady=8, fill="x")
        
        # --- Section: Bộ Ngành Selection ---
        ministry_section = ctk.CTkFrame(self.tab_all, fg_color=COLORS["surface"],
                                         corner_radius=8)
        ministry_section.pack(fill="x", padx=12, pady=(0, 8))
        
        ministry_header = ctk.CTkFrame(ministry_section, fg_color="transparent")
        ministry_header.pack(fill="x", padx=12, pady=(10, 8))
        
        ctk.CTkLabel(ministry_header, text="Bộ Ngành",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left")
        
        self.chk_sequential = ctk.CTkCheckBox(ministry_header, text="Chạy tuần tự tiếp theo",
                                               command=self.update_mode_desc,
                                               font=ctk.CTkFont(size=11),
                                               text_color=COLORS["text"],
                                               fg_color=COLORS["primary"],
                                               hover_color=COLORS["primary_dark"],
                                               border_color=COLORS["border"])
        self.chk_sequential.pack(side="right")
        
        self.ministries_list = ["Tất cả (Chạy toàn bộ)", "Bộ Y tế", "Bộ Quốc phòng", "Bộ Công an"]
        
        self.ministry_frame = ctk.CTkFrame(ministry_section, fg_color="transparent")
        self.ministry_frame.pack(fill="x", padx=12, pady=(0, 10))
        
        self.combo_ministry = ctk.CTkComboBox(self.ministry_frame, values=self.ministries_list, 
                                            width=280, state="readonly", command=self.update_mode_desc,
                                            height=34, corner_radius=8,
                                            fg_color=COLORS["card"], border_color=COLORS["border"],
                                            border_width=1, text_color=COLORS["text"],
                                            button_color=COLORS["primary"],
                                            button_hover_color=COLORS["primary_dark"],
                                            dropdown_fg_color=COLORS["card"],
                                            dropdown_text_color=COLORS["text"],
                                            dropdown_hover_color=COLORS["primary_light"],
                                            font=ctk.CTkFont(size=12))
        self.combo_ministry.pack(side="left")
        self.combo_ministry.set("Tất cả (Chạy toàn bộ)")

        # --- Section: Date Range ---
        date_section = ctk.CTkFrame(self.tab_all, fg_color=COLORS["surface"],
                                     corner_radius=8)
        date_section.pack(fill="x", padx=12, pady=(0, 8))
        
        ctk.CTkLabel(date_section, text="Khoảng thời gian",
                     font=ctk.CTkFont(size=13, weight="bold"),
                     text_color=COLORS["text"]).pack(anchor="w", padx=12, pady=(10, 8))
        
        self.investor_date_frame = ctk.CTkFrame(date_section, fg_color="transparent")
        self.investor_date_frame.pack(fill="x", padx=12, pady=(0, 10))
        
        ctk.CTkLabel(self.investor_date_frame, text="Từ ngày:",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left", padx=(0, 5))
        self.entry_investor_from = ctk.CTkEntry(self.investor_date_frame, width=120, height=32,
                                                 corner_radius=6, placeholder_text="dd/mm/yyyy",
                                                 fg_color=COLORS["card"], border_color=COLORS["border"],
                                                 border_width=1, text_color=COLORS["text"],
                                                 font=ctk.CTkFont(size=11))
        self.entry_investor_from.pack(side="left", padx=(0, 15))
        
        ctk.CTkLabel(self.investor_date_frame, text="Đến ngày:",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=COLORS["text"]).pack(side="left", padx=(0, 5))
        self.entry_investor_to = ctk.CTkEntry(self.investor_date_frame, width=120, height=32,
                                               corner_radius=6, placeholder_text="dd/mm/yyyy",
                                               fg_color=COLORS["card"], border_color=COLORS["border"],
                                               border_width=1, text_color=COLORS["text"],
                                               font=ctk.CTkFont(size=11))
        self.entry_investor_to.pack(side="left")
        
        # Initial UI State
        self.update_mode_desc()



        # --- Action Buttons (Bottom Left) ---
        self.action_frame = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        self.action_frame.grid(row=4, column=0, sticky="ew", padx=15, pady=(10, 15))
        
        self.action_frame.grid_columnconfigure(0, weight=2)
        self.action_frame.grid_columnconfigure(1, weight=1)
        self.action_frame.grid_columnconfigure(2, weight=1)

        self.start_btn = ctk.CTkButton(self.action_frame, text=" BẮT ĐẦU SCRAPING", height=46,
                                       image=self.icon_lib.get_icon("play_arrow", 18, "white", "white"),
                                       font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                                       fg_color=COLORS["primary"], hover_color=COLORS["primary_dark"],
                                       text_color="white", corner_radius=10,
                                       command=self.start_scraping)
        self.start_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.pause_btn = ctk.CTkButton(self.action_frame, text=" TẠM DỪNG", height=46,
                                       image=self.icon_lib.get_icon("pause", 18, "white", "white"),
                                       font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                                       fg_color=COLORS["warning"], hover_color="#E65100",
                                       text_color="white", corner_radius=10,
                                       state="disabled",
                                       command=self.toggle_pause)
        self.pause_btn.grid(row=0, column=1, sticky="ew", padx=5)
        
        self.reset_btn = ctk.CTkButton(self.action_frame, text=" LÀM LẠI", height=46,
                                       image=self.icon_lib.get_icon("autorenew", 18, COLORS["text"], COLORS["text"]),
                                       font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
                                       fg_color=COLORS["border"], hover_color=COLORS["card_hover"],
                                       text_color=COLORS["text"], corner_radius=10,
                                       command=self.reset_click_handler)
        self.reset_btn.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.action_frame, height=8, corner_radius=4, 
                                             progress_color=COLORS["accent"], 
                                             fg_color=COLORS["primary_light"], 
                                             border_width=0)
        # Grid it later when running

        # === RIGHT PANEL (Execution Logs) ===
        self.log_frame = ctk.CTkFrame(self, corner_radius=0,
                                      fg_color=COLORS["card"], border_width=0)
        self.log_frame.grid(row=0, column=1, sticky="nsew")
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)
        
        # Log header bar
        log_header_frame = ctk.CTkFrame(self.log_frame, fg_color="transparent")
        log_header_frame.grid(row=0, column=0, padx=15, pady=(12, 5), sticky="ew")
        log_header_frame.grid_columnconfigure(0, weight=1)
        
        self.log_header = ctk.CTkLabel(log_header_frame, text="Execution Logs",
                                       font=ctk.CTkFont(size=14, weight="bold"),
                                       text_color=COLORS["text"])
        self.log_header.grid(row=0, column=0, sticky="w")
        
        self.clear_log_btn = ctk.CTkButton(log_header_frame, text=" Clear", width=60, height=26,
                                            image=self.icon_lib.get_icon("delete", 14, COLORS["text_secondary"], COLORS["text_secondary"]),
                                            corner_radius=6,
                                            fg_color=COLORS["surface"], hover_color=COLORS["border"],
                                            text_color=COLORS["text_secondary"],
                                            font=ctk.CTkFont(size=11),
                                            command=self.clear_logs)
        self.clear_log_btn.grid(row=0, column=1, sticky="e")
        
        self.log_area = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12),
                                       fg_color=COLORS["log_bg"],
                                       text_color=COLORS["log_text"],
                                       corner_radius=8)
        self.log_area.grid(row=1, column=0, padx=12, pady=(0, 12), sticky="nsew")
        self.log_area.configure(state="disabled")

        # System Redirects
        sys.stdout = self
        sys.stderr = self
        
        # Startup Tasks
        self.after(2000, self.check_for_updates_thread)

    def update_mode_desc(self, _=None):
        choice = self.combo_ministry.get()
        is_seq = self.chk_sequential.get() == 1
        
        if "Tất cả" in choice:
            txt = "CHẾ ĐỘ: Thu thập toàn bộ thông tin nhà đầu tư liên quan đến ngành Dược (Không giới hạn Bộ Ngành)"
            self.chk_sequential.deselect()
            self.chk_sequential.configure(state="disabled", text="Chạy tuần tự (Không khả dụng)", text_color="gray")
        else:
            self.chk_sequential.configure(state="normal", text="Chạy tuần tự các bộ tiếp theo", text_color="black")
            
            # Base text
            if choice in ["Bộ Công an", "Bộ Quốc phòng"]:
                txt = f"CHẾ ĐỘ: Thu thập toàn bộ thông tin nhà đầu tư thuộc {choice} có liên quan đến Ngành Dược"
            else:
                txt = f"CHẾ ĐỘ: Thu thập toàn bộ thông tin nhà đầu tư thuộc {choice}"
            
            # Sequential suffix
            if is_seq:
                txt += "\n➤ Lưu ý: Sau khi hoàn thành, hệ thống sẽ TỰ ĐỘNG quét tiếp các Bộ còn lại"
                
        self.lbl_investor_desc.configure(text=txt)

    def on_tab_change(self):
        current = self.tab_view.get()
        if current in ["Kết Quả Đấu Thầu", "Yêu cầu báo giá"]:
            self.lbl_path.configure(text="Folder Save Path:")
            # Remove filename if present (User request: No initialFile)
            curr_val = self.path_entry.get()
            if curr_val.lower().endswith(".xlsx"):
                new_val = os.path.dirname(curr_val)
                if not new_val: new_val = os.getcwd()
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, new_val)
        else:
            self.lbl_path.configure(text="File Save Path:")
            # Restore filename if missing
            curr_val = self.path_entry.get()
            if not curr_val.lower().endswith(".xlsx"):
                # Append default filename
                if not curr_val or not os.path.exists(curr_val):
                    curr_val = os.getcwd()
                
                if os.path.isdir(curr_val):
                     new_val = os.path.join(curr_val, "investors_data_detailed.xlsx")
                     self.path_entry.delete(0, tk.END)
                     self.path_entry.insert(0, new_val)

    def on_contractor_mode_change(self, value):
        if value == "Tìm theo bộ lọc":
            self.ib_frame.pack_forget()
            self.filter_frame.pack(fill="both", expand=True)
        else:
            self.filter_frame.pack_forget()
            self.ib_frame.pack(fill="both", expand=True)

    def upload_ib_excel(self):
        import pandas as pd
        from tkinter import filedialog, messagebox
        file_path = filedialog.askopenfilename(title="Chọn file Excel", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                if "IB" not in df.columns:
                    messagebox.showerror("Lỗi", "File Excel không chứa cột 'IB'. Vui lòng kiểm tra lại!")
                    return
                # Extract non-null IB strings
                ib_list = df["IB"].dropna().astype(str).str.strip().tolist()
                ib_list = [ib for ib in ib_list if ib]
                if not ib_list:
                    messagebox.showwarning("Cảnh báo", "Cột 'IB' trống!")
                    return
                
                # Append to current
                ib_text = ", ".join(ib_list)
                current_text = self.entry_ib_list.get("1.0", "end-1c").strip()
                if current_text:
                    if not current_text.endswith(","):
                        current_text += ", "
                    self.entry_ib_list.delete("1.0", "end")
                    self.entry_ib_list.insert("1.0", current_text + ib_text)
                else:
                    self.entry_ib_list.delete("1.0", "end")
                    self.entry_ib_list.insert("1.0", ib_text)
                
                messagebox.showinfo("Thành công", f"Đã nạp {len(ib_list)} mã IB từ file Excel.")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc file Excel: {str(e)}")

    def browse_file(self):
        current_tab = self.tab_view.get()
        
        if current_tab in ["Kết Quả Đấu Thầu", "Yêu cầu báo giá"]:
            # For Contractor and RFQ mode, we select a FOLDER
            folder_selected = filedialog.askdirectory(title="Select Folder to Save Data")
            if folder_selected:
                self.path_entry.delete(0, tk.END)
                self.path_entry.insert(0, folder_selected)
        else:
            # Normal mode: File Save Dialog
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

    def clear_logs(self):
        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", tk.END)
        self.log_area.configure(state="disabled")

    def toggle_pause(self):
        if self.pause_event.is_set():
            self.pause_event.clear()
            self.pause_btn.configure(text=" TIẾP TỤC", image=self.icon_lib.get_icon("play_arrow", 18, "white", "white"), fg_color=COLORS["danger"], hover_color="#B71C1C")
            self.status_label.configure(text="Paused", text_color=COLORS["warning"])
            self.status_dot.configure(text_color=COLORS["warning"])
            print(">>> Signal: PAUSE")
        else:
            self.pause_event.set()
            self.pause_btn.configure(text=" TẠM DỪNG", image=self.icon_lib.get_icon("pause", 18, "white", "white"), fg_color=COLORS["warning"], hover_color="#E65100")
            self.status_label.configure(text="Resumed", text_color=COLORS["success"])
            self.status_dot.configure(text_color=COLORS["success"])
            print(">>> Signal: RESUME")

    def start_scraping(self):
        output_path = self.path_entry.get().strip()

        if not output_path:
            messagebox.showerror("Error", "Please specify a save path!")
            return

        # Handle Folder Logic for Contractor Tab and RFQ Tab
        current_tab = self.tab_view.get()
        if current_tab in ["Kết Quả Đấu Thầu", "Yêu cầu báo giá"]:
             if current_tab == "Kết Quả Đấu Thầu":
                 if getattr(self, "contractor_mode_seg", None) and self.contractor_mode_seg.get() == "Tìm theo danh sách IB (Excel/Nhập)":
                     pass
                 else:
                     f_d = self.entry_from_date.get().strip()
                     t_d = self.entry_to_date.get().strip()  
                     import datetime
                     def check_date_strict(d):
                         if not d: return True
                         if "-" in d or "." in d: return False
                         try:
                             datetime.datetime.strptime(d, "%d/%m/%Y")
                             return True
                         except: return False
                     
                     if not check_date_strict(f_d) or not check_date_strict(t_d):
                         messagebox.showwarning("Sai định dạng ngày", "Vui lòng nhập ngày theo định dạng: dd/mm/yyyy\nVí dụ: 30/12/2025")
                         return

             import datetime
             now_str = datetime.datetime.now().strftime("%d_%m_%Y %H_%M_%S")
             if current_tab == "Kết Quả Đấu Thầu":
                 folder_name = f"Ket Qua Dau Thau {now_str}"
                 file_name = "Danh Sach Thong Bao Moi Thau.xlsx"
             else:
                 folder_name = f"YCBG_{now_str}"
                 file_name = "Yeu Cau Bao Gia.xlsx"
                 
             if os.path.isdir(output_path):
                 full_folder_path = os.path.join(output_path, folder_name)
                 try:
                     os.makedirs(full_folder_path, exist_ok=True)
                     output_path = os.path.join(full_folder_path, file_name)
                 except Exception as e:
                     messagebox.showerror("Error", f"Could not create folder: {e}")
                     return
             else:
                 if output_path.endswith(".xlsx"):
                      parent = os.path.dirname(output_path)
                      if os.path.isdir(parent):
                           full_folder_path = os.path.join(parent, folder_name)
                           try:
                               os.makedirs(full_folder_path, exist_ok=True)
                               output_path = os.path.join(full_folder_path, file_name)
                           except: pass




        # UI Updates
        self.start_btn.configure(state="disabled", text=" ĐANG CHẠY...", image=self.icon_lib.get_icon("autorenew", 18, "white", "white"), fg_color=COLORS["surface_light"])
        self.pause_btn.configure(state="normal", text=" TẠM DỪNG", image=self.icon_lib.get_icon("pause", 18, "white", "white"), fg_color=COLORS["warning"])
        self.path_entry.configure(state="disabled")
        self.status_label.configure(text="Scraping in progress...", text_color=COLORS["warning"])
        self.status_dot.configure(text_color=COLORS["warning"])
        
        # Start Timer
        self.timer_running = True
        self.job_start_time = time.time()
        self.update_timer()
        
        # Reset Pause/Stop Event
        self.pause_event.set()
        self.stop_event.clear()
        
        # Show Progress Bar
        self.progress_bar.context_menu = None 
        self.progress_bar.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(20, 0))
        self.progress_bar.start()

        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", tk.END)
        self.log_area.configure(state="disabled")

        # Get Filter (Ministry) or Contractor Mode
        current_tab = self.tab_view.get()
        start_ministry = ""
        is_sequential = False
        mode = "NORMAL"
        kw = ""
        exclude = ""
        from_date = ""
        to_date = ""
        search_type = ""
        
        if current_tab == "Thông Tin Nhà Đầu Tư":
            val = self.combo_ministry.get()
            if val and "Tất cả" not in val:
                start_ministry = val
                is_sequential = True if self.chk_sequential.get() == 1 else False
            # Get dates for Investor
            from_date = self.entry_investor_from.get().strip()
            to_date = self.entry_investor_to.get().strip()
        if current_tab == "Kết Quả Đấu Thầu":
            if getattr(self, "contractor_mode_seg", None) and self.contractor_mode_seg.get() == "Tìm theo danh sách IB (Excel/Nhập)":
                mode = "CONTRACTOR_IB"
                kw = self.entry_ib_list.get("1.0", "end-1c").strip()
                exclude = ""
                from_date = ""
                to_date = ""
                if not kw:
                    messagebox.showerror("Lỗi", "Vui lòng nhập hoặc upload danh sách IB!")
                    self.reset_ui()
                    return
            else:
                mode = "CONTRACTOR"
                kw = self.entry_keywords.get().strip()
                exclude = self.entry_exclude.get().strip()
                from_date = self.entry_from_date.get().strip()
                to_date = self.entry_to_date.get().strip()
                search_type = self.search_type_combo.get() if hasattr(self, 'search_type_combo') else ""
                
                # Only use defaults if user entered something (field is NOT empty)
                # If field is completely empty → user wants to skip default
                use_default_keywords = (kw != "")
                use_default_exclude = (exclude != "") 
        elif current_tab == "Công bố giá thuốc":
            mode = "DRUG_PRICE" 
        elif current_tab == "Yêu cầu báo giá":
            mode = "RFQ"
            kw = self.entry_rfq_keywords.get().strip()
            from_date = self.entry_rfq_from_date.get().strip()
            to_date = self.entry_rfq_to_date.get().strip()
            if not kw:
                messagebox.showerror("Lỗi", "Vui lòng nhập từ khóa tìm kiếm!", parent=self)
                self.reset_ui()
                return 

        # Threading
        use_default_kw = use_default_keywords if mode == "CONTRACTOR" else True
        use_default_ex = use_default_exclude if mode == "CONTRACTOR" else True
        t = threading.Thread(target=self.run_process, args=(output_path, start_ministry, is_sequential, mode, kw, exclude, from_date, to_date, search_type, use_default_kw, use_default_ex))
        t.daemon = True
        t.start()
    
    def run_process(self, output_path, start_ministry, is_sequential, mode="NORMAL", kw="", exclude="", from_date="", to_date="", search_type="", use_default_keywords=True, use_default_exclude=True):
        start_time = time.time()
        try:
            print(">>> INITIALIZING SCRAPER...")
            print(f"> Target File: {output_path}")
            
            if mode == "CONTRACTOR":
                scrape_muasamcong.run_contractor_selection(
                    output_path=output_path,
                    keywords=kw,
                    exclude_words=exclude,
                    from_date=from_date,
                    to_date=to_date,
                    search_type=search_type,
                    use_default_keywords=use_default_keywords,
                    use_default_exclude=use_default_exclude,
                    pause_event=self.pause_event,
                    stop_event=self.stop_event
                )
                print("\n>>> COMPLETED SUCCESSFULLY!")
                self.timer_running = False
                self.status_label.configure(text="Completed", text_color=COLORS["success"])
                self.status_dot.configure(text_color=COLORS["success"])
                messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
                return
            elif mode == "CONTRACTOR_IB":
                scrape_muasamcong.run_contractor_selection(
                    output_path=output_path,
                    ib_list=kw,
                    pause_event=self.pause_event,
                    stop_event=self.stop_event
                )
                print("\n>>> COMPLETED SUCCESSFULLY!")
                self.timer_running = False
                self.status_label.configure(text="Completed", text_color=COLORS["success"])
                self.status_dot.configure(text_color=COLORS["success"])
                messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
                return
            elif mode == "DRUG_PRICE":
                scrape_muasamcong.run_drug_price_scrape(
                    output_path=output_path,
                    pause_event=self.pause_event,
                    stop_event=self.stop_event
                )
                print("\n>>> COMPLETED SUCCESSFULLY!")
                self.timer_running = False
                self.status_label.configure(text="Completed", text_color=COLORS["success"])
                self.status_dot.configure(text_color=COLORS["success"])
                messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
                return
            elif mode == "RFQ":
                scrape_muasamcong.run_rfq_scrape(
                    output_path=output_path,
                    pause_event=self.pause_event,
                    stop_event=self.stop_event,
                    keywords=kw,
                    from_date=from_date,
                    to_date=to_date
                )
                print("\n>>> COMPLETED SUCCESSFULLY!")
                self.timer_running = False
                self.status_label.configure(text="Completed", text_color=COLORS["success"])
                self.status_dot.configure(text_color=COLORS["success"])
                messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
                return
            else:
                ministries_to_scrape = []
            
            # Custom Sequences
            ministry_sequences = {
                "Bộ Y tế": ["Bộ Y tế", "Bộ Quốc phòng", "Bộ Công an"],
                "Bộ Quốc phòng": ["Bộ Quốc phòng", "Bộ Công an", "Bộ Y tế"],
                "Bộ Công an": ["Bộ Công an", "Bộ Quốc phòng", "Bộ Y tế"]
            }
            
            # Determine list of ministries
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
                
                # --- MINISTRY MODE (API) ---
                print(f"> Using API Scan for Ministries: {ministries_to_scrape}")
                scrape_muasamcong.run_investor_scan_api(
                     output_path=output_path,
                     pause_event=self.pause_event,
                     stop_event=self.stop_event,
                     ministries=ministries_to_scrape,
                     from_date_str=from_date,
                     to_date_str=to_date
                )

            else:
                # ALL Mode - USE NEW API SCAN
                print("> Mode: ALL DATA (API SCAN - New Logic)")
                scrape_muasamcong.run_investor_scan_api(
                     output_path=output_path,
                     pause_event=self.pause_event,
                     stop_event=self.stop_event,
                     ministries=None, # All mode
                     from_date_str=from_date,
                     to_date_str=to_date
                )
                
            print("\n>>> COMPLETED SUCCESSFULLY!")
            self.timer_running = False
            self.status_label.configure(text="Completed", text_color=COLORS["success"])
            self.status_dot.configure(text_color=COLORS["success"])
            messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
            return 
            
            # (Old scraping loop removed here)
            print("-" * 50)
            
            print("\n>>> COMPLETED SUCCESSFULLY!")
            self.timer_running = False # Stop timer
            self.status_label.configure(text="Completed", text_color=COLORS["success"])
            self.status_dot.configure(text_color=COLORS["success"])
            messagebox.showinfo("Success", f"Data scraped successfully to:\n{output_path}")
            
        except InterruptedError:
            self.timer_running = False # Stop timer
            print("\n>>> PROCESS STOPPED BY USER.")
            self.status_label.configure(text="Stopped 🛑", text_color=COLORS["text_secondary"])
            self.status_dot.configure(text_color=COLORS["text_secondary"])
            # No error popup for manual stop
            
        except Exception as e:
            self.timer_running = False # Stop timer
            print(f"\n>>> ERROR: {e}")
            self.status_label.configure(text="Error", text_color=COLORS["danger"])
            self.status_dot.configure(text_color=COLORS["danger"])
            messagebox.showerror("Error", f"An error occurred:\n{e}")
        finally:
            elapsed_time = time.time() - start_time
            m, s = divmod(int(elapsed_time), 60)
            h, m = divmod(m, 60)
            time_str = f"{h}h {m}m {s}s" if h > 0 else f"{m}m {s}s"
            print(f">>> TOTAL EXECUTION TIME: {time_str}")
            
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.start_btn.configure(state="normal", text=" BẮT ĐẦU SCRAPING", image=self.icon_lib.get_icon("play_arrow", 18, "white", "white"), fg_color=COLORS["primary"])
        self.pause_btn.configure(state="disabled", text=" TẠM DỪNG", image=self.icon_lib.get_icon("pause", 18, "white", "white"), fg_color=COLORS["warning"])
        self.pause_event.set()
        self.path_entry.configure(state="normal")
        self.tab_view.configure(state="normal")
        
        self.progress_bar.stop()
        self.progress_bar.grid_forget()
        
        self.timer_running = False

    def update_timer(self):
        if self.timer_running:
            elapsed = time.time() - self.job_start_time
            m, s = divmod(int(elapsed), 60)
            h, m = divmod(m, 60)
            if h > 0:
                time_str = f"{h:02}:{m:02}:{s:02}"
            else:
                time_str = f"{m:02}:{s:02}"
            
            self.timer_label.configure(text=f"{time_str}")
            self.after(1000, self.update_timer)

    def reset_click_handler(self):
        # If running, stop first
        if self.start_btn._state == "disabled" or self.progress_bar.winfo_ismapped():
            ans = messagebox.askyesno("Confirm Reset", "Scraping is running. Do you want to STOP and RESET?")
            if ans:
                # Signal Stop
                self.stop_event.set()
                self.pause_event.set() # Unpause to allow exit
                self.status_label.configure(text="Stopping...", text_color=COLORS["danger"])
                self.status_dot.configure(text_color=COLORS["danger"])
                
                # We can't immediately reset UI because thread is still winding down.
                # But for UX, we can just reset inputs. Thread will die eventually.
                self.reset_inputs()
                self.reset_ui() # Force UI reset
        else:
            self.reset_inputs()

    def reset_inputs(self):
        # 1. Reset Path
        default_path = os.path.join(os.getcwd(), "investors_data_detailed.xlsx")
        self.path_entry.delete(0, tk.END)
        self.path_entry.insert(0, default_path)
        
        # 2. Reset Limit
        self.limit_entry.delete(0, tk.END)
        self.limit_entry.insert(0, "0")
        
        # 3. Reset Tabs & Ministry
        self.tab_view.set("Thông Tin Nhà Đầu Tư") # Default tab (need to check name)
        # Tab names are defined in init. 
        # Check logic: self.tab_all = self.tab_view.add("Toàn Bộ")
        # To switch tab, use tab name string.
        
        self.combo_ministry.set("Bộ Y tế")
        self.chk_sequential.select()
        
        # 4. Clear Logs
        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", tk.END)
        self.log_area.configure(state="disabled")
        
        self.status_label.configure(text="Reset to Defaults", text_color=COLORS["text_secondary"])
        self.status_dot.configure(text_color=COLORS["text_secondary"])

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
                    self.after(0, lambda: self.update_btn.configure(text="OK", fg_color=COLORS["success"]))
                    print("You are using the latest version.")
        except Exception as e:
            print(f"Update check failed: {e}")

    def indicate_update_found(self, version, url):
        # Change button to Red Alert
        self.update_btn.configure(text=f"🔔 UPDATE {version}", fg_color=COLORS["danger"], hover_color="#B71C1C")
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
