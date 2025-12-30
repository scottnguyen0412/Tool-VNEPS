import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import sys
import os
import scrape_muasamcong

# Configuration
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class ScraperApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("Tool Cào Dữ Liệu Mua Sắm Công - VNEPS")
        self.geometry("800x600")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # Log area expands

        # 1. Header Section
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 10))
        
        self.title_label = ctk.CTkLabel(self.header_frame, text="VN-EPS DATA SCRAPER", 
                                      font=ctk.CTkFont(size=24, weight="bold"))
        self.title_label.pack(anchor="w")
        
        self.subtitle_label = ctk.CTkLabel(self.header_frame, text="Automation Tool for Mua Sam Cong Data", 
                                         font=ctk.CTkFont(size=14, slant="italic"), text_color="gray")
        self.subtitle_label.pack(anchor="w")

        # 2. Controls Section (Card-like look)
        self.controls_frame = ctk.CTkFrame(self)
        self.controls_frame.grid(row=1, column=0, sticky="ew", padx=20, pady=10)
        self.controls_frame.grid_columnconfigure(1, weight=1)

        # File Path Input
        self.path_label = ctk.CTkLabel(self.controls_frame, text="File Save Path:", font=ctk.CTkFont(weight="bold"))
        self.path_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        default_path = os.path.join(os.getcwd(), "investors_data_detailed.xlsx")
        self.path_entry = ctk.CTkEntry(self.controls_frame, placeholder_text="Path to save .xlsx file")
        self.path_entry.grid(row=0, column=1, padx=(0, 10), pady=(20, 10), sticky="ew")
        self.path_entry.insert(0, default_path)
        
        self.browse_btn = ctk.CTkButton(self.controls_frame, text="Browse", text_color="white", width=80, command=self.browse_file)
        self.browse_btn.grid(row=0, column=2, padx=20, pady=(20, 10))

        # Page Limit Input
        self.limit_label = ctk.CTkLabel(self.controls_frame, text="Page Limit:", font=ctk.CTkFont(weight="bold"))
        self.limit_label.grid(row=1, column=0, padx=20, pady=(10, 20), sticky="w")
        
        self.limit_entry = ctk.CTkEntry(self.controls_frame, placeholder_text="0 for all")
        self.limit_entry.grid(row=1, column=1, padx=(0, 20), pady=(10, 20), sticky="w")
        self.limit_entry.insert(0, "0")
        
        self.hint_label = ctk.CTkLabel(self.controls_frame, text="(Enter 0 to scrape ALL pages)", text_color="gray")
        self.hint_label.grid(row=1, column=1, padx=(150, 0), pady=(10, 20), sticky="w")

        # Action Buttons
        self.start_btn = ctk.CTkButton(self.controls_frame, text="START SCRAPING", 
                                     font=ctk.CTkFont(size=15, weight="bold"),
                                     text_color="white",
                                     height=40, fg_color="#2ecc71", hover_color="#27ae60",
                                     command=self.start_scraping)
        self.start_btn.grid(row=2, column=0, columnspan=3, padx=20, pady=(10, 20), sticky="ew")

        # 3. Console/Log Section
        self.log_frame = ctk.CTkFrame(self)
        self.log_frame.grid(row=2, column=0, sticky="nsew", padx=20, pady=(10, 20))
        self.log_frame.grid_rowconfigure(1, weight=1)
        self.log_frame.grid_columnconfigure(0, weight=1)

        self.log_label = ctk.CTkLabel(self.log_frame, text="Real-time Logs", font=ctk.CTkFont(weight="bold"))
        self.log_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")

        self.log_area = ctk.CTkTextbox(self.log_frame, font=("Consolas", 12))
        self.log_area.grid(row=1, column=0, padx=10, pady=(0, 10), sticky="nsew")
        self.log_area.configure(state="disabled")

        # 4. Footer
        self.status_label = ctk.CTkLabel(self, text="Ready", anchor="w")
        self.status_label.grid(row=3, column=0, sticky="ew", padx=20, pady=(0, 10))

        # Redirect Stdout
        sys.stdout = self
        sys.stderr = self

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
        self.log_area.configure(state="normal")
        self.log_area.insert(tk.END, text)
        self.log_area.see(tk.END)
        self.log_area.configure(state="disabled")
        # CustomTkinter needs update calls sometimes for thread safety in UI updates
        # self.update_idletasks() # Careful with this in threads

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
        self.status_label.configure(text="Status: Scraping in progress... Please do not close the browser window.")
        
        self.log_area.configure(state="normal")
        self.log_area.delete("0.0", tk.END)
        self.log_area.configure(state="disabled")

        # Threading
        t = threading.Thread(target=self.run_process, args=(output_path, max_pages))
        t.daemon = True
        t.start()
    
    def run_process(self, output_path, max_pages):
        try:
            print(">>> INITIALIZING SCRAPER...")
            print(f"> Target File: {output_path}")
            print(f"> Page Limit:  {'Unlimited' if max_pages == float('inf') else max_pages}")
            print("-" * 50)
            
            scrape_muasamcong.run(output_path=output_path, max_pages=max_pages)
            
            print("\n>>> COMPLETED SUCCESSFULLY!")
            self.status_label.configure(text="Status: Completed ✅")
            messagebox.showinfo("Success", f"Data scraping finished!\nSaved to: {output_path}")
        except Exception as e:
            print(f"\n>>> ERROR: {e}")
            self.status_label.configure(text="Status: Error ❌")
            messagebox.showerror("Error", f"An error occurred: {e}")
        finally:
            # Safely reset UI
            self.after(0, self.reset_ui)

    def reset_ui(self):
        self.start_btn.configure(state="normal", text="START SCRAPING", fg_color="#2ecc71")
        self.path_entry.configure(state="normal")
        self.limit_entry.configure(state="normal")

if __name__ == "__main__":
    app = ScraperApp()
    app.mainloop()
