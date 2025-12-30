import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import sys
import scrape_muasamcong

class ScraperGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Tool Scrape Muasamcong")
        self.root.geometry("600x500")
        
        # Style
        style = ttk.Style()
        style.configure('TButton', font=('Helvetica', 10))
        style.configure('TLabel', font=('Helvetica', 10))

        # Main Frame
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Title
        title_label = ttk.Label(main_frame, text="TOOL CÀO DỮ LIỆU NHÀ ĐẦU TƯ", font=('Helvetica', 16, 'bold'))
        title_label.pack(pady=(0, 20))

        # Input Frame
        input_frame = ttk.Frame(main_frame)
        input_frame.pack(fill=tk.X, pady=5)

        # Filename Input
        ttk.Label(input_frame, text="Tên file kết quả (.xlsx):").grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.filename_var = tk.StringVar(value="investors_data_detailed.xlsx")
        self.filename_entry = ttk.Entry(input_frame, textvariable=self.filename_var, width=40)
        self.filename_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

        # Page Limit Input
        ttk.Label(input_frame, text="Số trang muốn cào (0 = Tất cả):").grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.limit_var = tk.StringVar(value="0")
        self.limit_entry = ttk.Entry(input_frame, textvariable=self.limit_var, width=10)
        self.limit_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        input_frame.columnconfigure(1, weight=1)

        # Button Frame
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=20)

        self.start_btn = ttk.Button(btn_frame, text="Bắt đầu chạy", command=self.start_scraping)
        self.start_btn.pack(side=tk.LEFT, padx=10)
        
        self.stop_btn = ttk.Button(btn_frame, text="Thoát", command=root.destroy)
        self.stop_btn.pack(side=tk.LEFT, padx=10)

        # Log Area
        ttk.Label(main_frame, text="Log hoạt động:").pack(anchor=tk.W)
        self.log_area = scrolledtext.ScrolledText(main_frame, height=15, state='disabled', font=('Consolas', 9))
        self.log_area.pack(fill=tk.BOTH, expand=True, pady=5)

        # Redirect stdout
        sys.stdout = self
        sys.stderr = self

    def write(self, text):
        self.log_area.configure(state='normal')
        self.log_area.insert(tk.END, text)
        self.log_area.see(tk.END)
        self.log_area.configure(state='disabled')
        # Force update idle tasks to keep UI responsive
        self.root.update_idletasks()

    def flush(self):
        pass

    def start_scraping(self):
        filename = self.filename_var.get().strip()
        limit_str = self.limit_var.get().strip()

        if not filename:
            messagebox.showerror("Lỗi", "Vui lòng nhập tên file!")
            return

        try:
            limit = int(limit_str)
            if limit <= 0:
                max_pages = float('inf')
            else:
                max_pages = limit
        except ValueError:
            messagebox.showerror("Lỗi", "Số trang phải là số nguyên!")
            return

        # Disable button
        self.start_btn.config(state='disabled')
        self.log_area.configure(state='normal')
        self.log_area.delete(1.0, tk.END)
        self.log_area.configure(state='disabled')

        # Run in thread
        t = threading.Thread(target=self.run_process, args=(filename, max_pages))
        t.daemon = True
        t.start()

    def run_process(self, filename, max_pages):
        try:
            print("Đang khởi động trình duyệt...")
            scrape_muasamcong.run(output_path=filename, max_pages=max_pages)
            print("\nHOÀN THÀNH!")
            messagebox.showinfo("Thành công", "Đã cào dữ liệu xong!")
        except Exception as e:
            print(f"\nLỖI: {e}")
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {e}")
        finally:
            self.root.after(0, lambda: self.start_btn.config(state='normal'))

if __name__ == "__main__":
    root = tk.Tk()
    app = ScraperGUI(root)
    root.mainloop()
