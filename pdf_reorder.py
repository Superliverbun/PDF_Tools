import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import io

class SplitPDFDialog(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.parent = parent
        self.callback = callback
        self.title("分割 PDF")
        self.geometry("400x300")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        
        # 創建選項框架
        options_frame = ttk.LabelFrame(self, text="分割選項")
        options_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 選擇分割模式
        self.split_mode = tk.StringVar(value="range")
        
        # 按頁面範圍分割
        ttk.Radiobutton(
            options_frame, 
            text="按頁面範圍分割", 
            variable=self.split_mode, 
            value="range",
            command=self.update_ui
        ).pack(anchor=tk.W, padx=10, pady=(10, 5))
        
        # 範圍輸入框架
        self.range_frame = ttk.Frame(options_frame)
        self.range_frame.pack(fill=tk.X, padx=20, pady=5)
        
        ttk.Label(self.range_frame, text="頁面範圍 (例如: 1-3,5,7-9):").pack(anchor=tk.W)
        self.range_entry = ttk.Entry(self.range_frame)
        self.range_entry.pack(fill=tk.X, pady=5)
        ttk.Label(
            self.range_frame, 
            text="注意: 頁碼從 1 開始，每個範圍將生成一個獨立的 PDF 檔案",
            wraplength=350,
            font=("Arial", 8),
            foreground="gray"
        ).pack(anchor=tk.W)
        
        # 每頁分割為單獨 PDF
        ttk.Radiobutton(
            options_frame, 
            text="每頁分割為單獨 PDF", 
            variable=self.split_mode, 
            value="each",
            command=self.update_ui
        ).pack(anchor=tk.W, padx=10, pady=5)
        
        # 目標資料夾選擇
        folder_frame = ttk.Frame(options_frame)
        folder_frame.pack(fill=tk.X, padx=10, pady=(15, 5))
        
        ttk.Label(folder_frame, text="輸出資料夾:").pack(side=tk.LEFT, padx=5)
        self.folder_var = tk.StringVar()
        folder_entry = ttk.Entry(folder_frame, textvariable=self.folder_var, width=25)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(folder_frame, text="瀏覽...", command=self.browse_folder).pack(side=tk.LEFT, padx=5)
        
        # 檔名前綴
        prefix_frame = ttk.Frame(options_frame)
        prefix_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(prefix_frame, text="檔名前綴:").pack(side=tk.LEFT, padx=5)
        self.prefix_var = tk.StringVar(value="split_")
        prefix_entry = ttk.Entry(prefix_frame, textvariable=self.prefix_var, width=20)
        prefix_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        # 按鈕框架
        button_frame = ttk.Frame(self)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(button_frame, text="取消", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        ttk.Button(button_frame, text="分割", command=self.do_split).pack(side=tk.RIGHT, padx=5)
        
        # 設定預設輸出資料夾
        self.folder_var.set(os.path.expanduser("~/Documents"))
        
        # 更新 UI
        self.update_ui()
        
        # 設置焦點
        self.range_entry.focus_set()
    
    def update_ui(self):
        # 根據分割模式更新 UI
        if self.split_mode.get() == "range":
            self.range_frame.pack(fill=tk.X, padx=20, pady=5)
        else:
            self.range_frame.pack_forget()
    
    def browse_folder(self):
        # 選擇輸出資料夾
        folder = filedialog.askdirectory(
            title="選擇輸出資料夾",
            initialdir=self.folder_var.get()
        )
        if folder:
            self.folder_var.set(folder)
    
    def do_split(self):
        # 檢查輸出資料夾
        output_dir = self.folder_var.get()
        if not output_dir or not os.path.isdir(output_dir):
            messagebox.showerror("錯誤", "請選擇有效的輸出資料夾")
            return
        
        # 檢查檔名前綴
        prefix = self.prefix_var.get()
        if not prefix:
            prefix = "split_"
        
        # 根據分割模式執行分割
        if self.split_mode.get() == "range":
            # 檢查頁面範圍
            range_text = self.range_entry.get().strip()
            if not range_text:
                messagebox.showerror("錯誤", "請輸入有效的頁面範圍")
                return
            
            # 解析頁面範圍
            try:
                page_ranges = self.parse_range(range_text)
                if not page_ranges:
                    messagebox.showerror("錯誤", "無法解析頁面範圍，請檢查格式是否正確")
                    return
            except Exception as e:
                messagebox.showerror("錯誤", f"解析頁面範圍時出錯: {str(e)}")
                return
            
            # 執行分割操作
            self.callback(output_dir, prefix, "range", page_ranges)
        else:
            # 每頁分割
            self.callback(output_dir, prefix, "each", None)
        
        # 關閉對話框
        self.destroy()
    
    def parse_range(self, range_text):
        # 解析頁面範圍字串，例如 "1-3,5,7-9"
        page_ranges = []
        parts = range_text.split(",")
        
        for part in parts:
            part = part.strip()
            if "-" in part:
                # 處理範圍，例如 "1-3"
                start, end = part.split("-")
                start = int(start.strip())
                end = int(end.strip())
                if start < 1 or end < start:
                    raise ValueError(f"無效的頁面範圍: {part}")
                page_ranges.append((start-1, end-1))  # 轉換為從 0 開始的索引
            else:
                # 處理單頁，例如 "5"
                page = int(part.strip())
                if page < 1:
                    raise ValueError(f"無效的頁面: {part}")
                page_ranges.append((page-1, page-1))  # 轉換為從 0 開始的索引
        
        return page_ranges


class PDFReorderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 頁面順序調整工具")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # 設定變數
        self.current_pdf = None
        self.pages = []
        self.page_images = []
        self.selected_index = None
        
        # 建立界面
        self.create_ui()
        
    def create_ui(self):
        # 創建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 上方工具列
        toolbar = ttk.Frame(main_frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        # 開啟 PDF 按鈕
        ttk.Button(toolbar, text="開啟 PDF", command=self.open_pdf).pack(side=tk.LEFT, padx=5)
        
        # 儲存 PDF 按鈕
        ttk.Button(toolbar, text="儲存 PDF", command=self.save_pdf).pack(side=tk.LEFT, padx=5)
        
        # 分割 PDF 按鈕 (新增)
        ttk.Button(toolbar, text="分割 PDF", command=self.show_split_dialog).pack(side=tk.LEFT, padx=5)
        
        # 側邊欄和預覽區域的框架
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # 頁面列表框架（左側）
        pages_frame = ttk.LabelFrame(content_frame, text="PDF 頁面")
        pages_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10), expand=True)
        
        # 頁面列表和滾動條
        list_frame = ttk.Frame(pages_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.page_list = tk.Listbox(list_frame, selectmode=tk.SINGLE, activestyle='none',
                                   yscrollcommand=scrollbar.set, font=("Arial", 10))
        self.page_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.page_list.yview)
        
        # 頁面列表選擇事件
        self.page_list.bind('<<ListboxSelect>>', self.on_page_select)
        
        # 頁面操作按鈕
        buttons_frame = ttk.Frame(pages_frame)
        buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_frame, text="上移", command=self.move_page_up).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(buttons_frame, text="下移", command=self.move_page_down).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(buttons_frame, text="刪除", command=self.delete_page).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # 頁面預覽框架（右側）
        preview_frame = ttk.LabelFrame(content_frame, text="頁面預覽")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 預覽畫布和滾動條
        canvas_frame = ttk.Frame(preview_frame)
        canvas_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 垂直滾動條
        vscrollbar = ttk.Scrollbar(canvas_frame)
        vscrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 水平滾動條
        hscrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL)
        hscrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 預覽畫布
        self.preview_canvas = tk.Canvas(canvas_frame, 
                                      yscrollcommand=vscrollbar.set, 
                                      xscrollcommand=hscrollbar.set,
                                      bg="light gray")
        self.preview_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        vscrollbar.config(command=self.preview_canvas.yview)
        hscrollbar.config(command=self.preview_canvas.xview)
        
        # 狀態欄
        self.status_var = tk.StringVar()
        self.status_var.set("請開啟 PDF 檔案")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill=tk.X, pady=(10, 0))
    
    def open_pdf(self):
        # 開啟檔案對話框
        pdf_file = filedialog.askopenfilename(
            title="選擇 PDF 檔案",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        
        if not pdf_file:
            return
        
        try:
            # 開啟 PDF 檔案
            self.current_pdf = fitz.open(pdf_file)
            self.pages = list(range(len(self.current_pdf)))
            self.page_images = []
            
            # 更新狀態
            self.status_var.set(f"已開啟: {os.path.basename(pdf_file)} ({len(self.current_pdf)} 頁)")
            
            # 更新頁面列表
            self.update_page_list()
            
            # 清空預覽
            self.preview_canvas.delete("all")
            
            # 加載頁面預覽
            self.load_page_previews()
            
        except Exception as e:
            messagebox.showerror("錯誤", f"無法開啟 PDF 檔案: {str(e)}")
    
    def load_page_previews(self):
        # 清空現有的頁面預覽
        self.page_images = []
        
        if not self.current_pdf:
            return
        
        # 為每頁創建預覽圖
        for page_num in self.pages:
            try:
                page = self.current_pdf[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(0.3, 0.3))  # 縮小預覽圖
                img_data = pix.tobytes("ppm")
                img = Image.open(io.BytesIO(img_data))
                self.page_images.append(ImageTk.PhotoImage(img))
            except Exception as e:
                # 如果無法創建預覽，添加一個空白圖像
                self.page_images.append(None)
                print(f"無法創建頁面 {page_num+1} 的預覽: {str(e)}")
    
    def update_page_list(self):
        # 更新頁面列表
        self.page_list.delete(0, tk.END)
        for i, page_num in enumerate(self.pages):
            self.page_list.insert(tk.END, f"頁面 {page_num+1}")
    
    def on_page_select(self, event):
        # 處理頁面選擇事件
        selection = self.page_list.curselection()
        if selection:
            self.selected_index = selection[0]
            self.show_page_preview(self.selected_index)
        else:
            self.selected_index = None
    
    def show_page_preview(self, index):
        # 顯示選中頁面的預覽
        if not self.current_pdf or index is None or index >= len(self.pages):
            return
        
        # 清空預覽畫布
        self.preview_canvas.delete("all")
        
        # 獲取頁面號
        page_num = self.pages[index]
        
        # 檢查是否有預覽圖
        if index < len(self.page_images) and self.page_images[index]:
            # 使用已加載的預覽圖
            img = self.page_images[index]
            self.preview_canvas.config(scrollregion=(0, 0, img.width(), img.height()))
            self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=img)
        else:
            # 直接創建預覽
            try:
                page = self.current_pdf[page_num]
                pix = page.get_pixmap(matrix=fitz.Matrix(1, 1))
                img_data = pix.tobytes("ppm")
                img = Image.open(io.BytesIO(img_data))
                img_tk = ImageTk.PhotoImage(img)
                
                # 保存引用以防垃圾回收
                self.current_preview = img_tk
                
                # 設定畫布大小
                self.preview_canvas.config(scrollregion=(0, 0, img.width, img.height))
                self.preview_canvas.create_image(0, 0, anchor=tk.NW, image=img_tk)
            except Exception as e:
                # 顯示錯誤訊息
                self.preview_canvas.create_text(
                    200, 200, text=f"無法顯示頁面 {page_num+1} 的預覽\n{str(e)}"
                )
    
    def move_page_up(self):
        # 將選中頁面上移
        if self.selected_index is None or self.selected_index <= 0:
            return
        
        # 交換頁面順序
        self.pages[self.selected_index], self.pages[self.selected_index-1] = \
            self.pages[self.selected_index-1], self.pages[self.selected_index]
        
        # 更新預覽圖順序
        if self.page_images:
            self.page_images[self.selected_index], self.page_images[self.selected_index-1] = \
                self.page_images[self.selected_index-1], self.page_images[self.selected_index]
        
        # 更新頁面列表
        self.update_page_list()
        
        # 更新選中項
        self.selected_index -= 1
        self.page_list.selection_clear(0, tk.END)
        self.page_list.selection_set(self.selected_index)
        self.page_list.see(self.selected_index)
        
        # 更新預覽
        self.show_page_preview(self.selected_index)
    
    def move_page_down(self):
        # 將選中頁面下移
        if self.selected_index is None or self.selected_index >= len(self.pages) - 1:
            return
        
        # 交換頁面順序
        self.pages[self.selected_index], self.pages[self.selected_index+1] = \
            self.pages[self.selected_index+1], self.pages[self.selected_index]
        
        # 更新預覽圖順序
        if self.page_images:
            self.page_images[self.selected_index], self.page_images[self.selected_index+1] = \
                self.page_images[self.selected_index+1], self.page_images[self.selected_index]
        
        # 更新頁面列表
        self.update_page_list()
        
        # 更新選中項
        self.selected_index += 1
        self.page_list.selection_clear(0, tk.END)
        self.page_list.selection_set(self.selected_index)
        self.page_list.see(self.selected_index)
        
        # 更新預覽
        self.show_page_preview(self.selected_index)
    
    def delete_page(self):
        # 刪除選中頁面
        if self.selected_index is None or not self.pages:
            return
        
        # 確認刪除
        if not messagebox.askyesno("確認刪除", "確定要刪除選中的頁面嗎？"):
            return
        
        # 刪除頁面
        del self.pages[self.selected_index]
        
        # 刪除預覽圖
        if self.page_images:
            del self.page_images[self.selected_index]
        
        # 更新頁面列表
        self.update_page_list()
        
        # 更新選中項
        if self.selected_index >= len(self.pages):
            self.selected_index = len(self.pages) - 1 if self.pages else None
        
        if self.selected_index is not None:
            self.page_list.selection_clear(0, tk.END)
            self.page_list.selection_set(self.selected_index)
            self.page_list.see(self.selected_index)
            self.show_page_preview(self.selected_index)
        else:
            # 清空預覽
            self.preview_canvas.delete("all")
    
    def save_pdf(self):
        # 儲存修改後的 PDF
        if not self.current_pdf or not self.pages:
            messagebox.showinfo("提示", "沒有 PDF 檔案可儲存")
            return
        
        # 選擇儲存路徑
        save_path = filedialog.asksaveasfilename(
            title="儲存 PDF 檔案",
            defaultextension=".pdf",
            filetypes=[("PDF 檔案", "*.pdf"), ("所有檔案", "*.*")]
        )
        
        if not save_path:
            return
        
        try:
            # 創建新的 PDF 文件
            new_pdf = fitz.open()
            
            # 按照新的順序加入頁面
            for page_idx in self.pages:
                new_pdf.insert_pdf(self.current_pdf, from_page=page_idx, to_page=page_idx)
            
            # 儲存 PDF
            new_pdf.save(save_path)
            new_pdf.close()
            
            # 更新狀態
            self.status_var.set(f"已儲存: {os.path.basename(save_path)}")
            
            messagebox.showinfo("成功", f"PDF 檔案已儲存至:\n{save_path}")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"儲存 PDF 時發生錯誤: {str(e)}")
    
    # 新增的分割 PDF 相關方法
    def show_split_dialog(self):
        # 檢查是否已開啟 PDF
        if not self.current_pdf:
            messagebox.showinfo("提示", "請先開啟 PDF 檔案")
            return
        
        # 顯示分割 PDF 對話框
        SplitPDFDialog(self.root, self.split_pdf)
    
    def split_pdf(self, output_dir, prefix, mode, page_ranges):
        """
        分割 PDF 檔案
        
        Args:
            output_dir (str): 輸出資料夾
            prefix (str): 檔名前綴
            mode (str): 分割模式，"range" 或 "each"
            page_ranges (list): 頁面範圍列表，只在 mode="range" 時使用
        """
        if not self.current_pdf:
            messagebox.showinfo("提示", "沒有 PDF 檔案可分割")
            return
        
        try:
            # 取得原始檔案名稱（不含路徑和副檔名）
            if hasattr(self.current_pdf, 'name'):
                original_name = os.path.splitext(os.path.basename(self.current_pdf.name))[0]
            else:
                original_name = "document"
            
            # 根據模式分割 PDF
            if mode == "each":
                # 每頁分割為單獨 PDF
                success_count = 0
                for i, page_idx in enumerate(self.pages):
                    # 創建新的 PDF 文件
                    new_pdf = fitz.open()
                    
                    # 加入單頁
                    new_pdf.insert_pdf(self.current_pdf, from_page=page_idx, to_page=page_idx)
                    
                    # 儲存 PDF
                    output_path = os.path.join(output_dir, f"{prefix}{original_name}_page_{page_idx+1}.pdf")
                    new_pdf.save(output_path)
                    new_pdf.close()
                    
                    success_count += 1
                
                # 顯示成功訊息
                messagebox.showinfo("成功", f"已將 PDF 分割為 {success_count} 個單頁檔案\n儲存至: {output_dir}")
                
            else:  # mode == "range"
                # 按頁面範圍分割
                success_count = 0
                for i, (start, end) in enumerate(page_ranges):
                    # 檢查頁面範圍
                    valid_start = max(0, min(start, len(self.pages)-1))
                    valid_end = max(valid_start, min(end, len(self.pages)-1))
                    
                    # 創建新的 PDF 文件
                    new_pdf = fitz.open()
                    
                    # 加入頁面
                    for page_idx in range(valid_start, valid_end + 1):
                        real_idx = self.pages[page_idx]
                        new_pdf.insert_pdf(self.current_pdf, from_page=real_idx, to_page=real_idx)
                    
                    # 儲存 PDF
                    range_str = f"{start+1}-{end+1}" if start != end else f"{start+1}"
                    output_path = os.path.join(output_dir, f"{prefix}{original_name}_pages_{range_str}.pdf")
                    new_pdf.save(output_path)
                    new_pdf.close()
                    
                    success_count += 1
                
                # 顯示成功訊息
                messagebox.showinfo("成功", f"已將 PDF 分割為 {success_count} 個檔案\n儲存至: {output_dir}")
            
            # 更新狀態
            self.status_var.set(f"已分割 PDF 至: {output_dir}")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"分割 PDF 時發生錯誤: {str(e)}")

if __name__ == "__main__":
    # 創建主視窗
    root = tk.Tk()
    app = PDFReorderApp(root)
    root.mainloop()