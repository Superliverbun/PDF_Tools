import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import win32gui
import win32con
import threading
from pathlib import Path
import sys

# 設置程式圖標
def resource_path(relative_path):
    """ 獲取資源的絕對路徑，兼容普通運行和 PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 添加一個全局變量來控制轉換過程
conversion_running = False
total_files = 0
processed_files = 0

def count_docx_files(root_folder):
    count = 0
    for _, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".docx") and not filename.startswith("~$"):
                count += 1
    return count

def convert_docx_to_pdf_custom(root_folder, output_folder, log_callback, progress_callback, stop_check):
    global processed_files
    processed_files = 0
    
    # 啟動 Word 應用程式並完全隱藏
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False  # 確保 Word 完全隱藏，不顯示任何視窗
    word.DisplayAlerts = 0  # 禁止顯示任何警告訊息
    
    # 設置 Word 窗口狀態為最小化
    if hasattr(word, 'WindowState'):
        word.WindowState = 2  # 2 = wdWindowStateMinimize
    
    # 隱藏任務欄圖標
    try:
        # 嘗試找到並隱藏 Word 的主窗口
        hwnd = win32gui.FindWindow("OpusApp", None)
        if hwnd:
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    except:
        pass
    
    try:
        for foldername, _, filenames in os.walk(root_folder):
            for filename in filenames:
                # 檢查是否應該停止進程
                if stop_check():
                    log_callback("⚠️ 轉換過程已被用戶停止")
                    return
                    
                if filename.endswith(".docx") and not filename.startswith("~$"):
                    docx_path = os.path.abspath(os.path.join(foldername, filename))
                    rel_path = os.path.relpath(foldername, root_folder)
                    top_level_folder = rel_path.split(os.sep)[0] if rel_path != "." else ""
                    output_subfolder = os.path.join(output_folder, top_level_folder)
                    os.makedirs(output_subfolder, exist_ok=True)
                    pdf_name = os.path.splitext(filename)[0] + ".pdf"
                    pdf_path = os.path.join(output_subfolder, pdf_name)
                    try:
                        # 打開文檔時使用 ReadOnly 和 NoEnvelopeVis 參數
                        doc = word.Documents.Open(
                            docx_path, 
                            ReadOnly=True,
                            Visible=False,
                            AddToRecentFiles=False
                        )
                        # 如果有文檔窗口，也嘗試隱藏它
                        if hasattr(doc, 'ActiveWindow') and doc.ActiveWindow is not None:
                            doc.ActiveWindow.Visible = False
                            
                        doc.SaveAs(pdf_path, FileFormat=17)
                        doc.Close(SaveChanges=False)
                        processed_files += 1
                        progress_callback(processed_files)
                        log_callback(f"✅ {docx_path} → {pdf_path}")
                    except Exception as e:
                        log_callback(f"❌ 無法轉換 {docx_path}，原因：{e}")
    finally:
        # 確保無論如何都關閉 Word
        try:
            word.Quit()
        except:
            pass
            
    if not stop_check():  # 只有在正常完成時才顯示完成訊息
        log_callback("🎉 所有 Word 檔轉 PDF 完成！")

def browse_folder(entry, is_output=False):
    if is_output:
        folder = filedialog.askdirectory(title="選擇輸出資料夾")
    else:
        folder = filedialog.askdirectory(title="選擇包含 Word 文件的資料夾")
    if folder:
        entry.delete(0, tk.END)
        entry.insert(0, folder)

def update_progress(value):
    def _update():
        if total_files > 0:
            progress_percent = (value / total_files) * 100
            progress_bar["value"] = progress_percent
            progress_label.config(text=f"進度: {value}/{total_files} ({progress_percent:.1f}%)")
            window.update_idletasks()
    window.after(0, _update)

def start_conversion_thread():
    global conversion_running, total_files
    input_path = input_entry.get()
    output_path = output_entry.get()
    
    if not input_path.strip() or not output_path.strip():
        messagebox.showerror("錯誤", "請選擇輸入和輸出資料夾")
        return
    
    if not os.path.exists(input_path) or not os.path.exists(output_path):
        messagebox.showerror("錯誤", "請確認輸入與輸出資料夾都存在")
        return
    
    # 清空日誌
    log_text.delete(1.0, tk.END)
    
    # 計算總文件數
    log_text.insert(tk.END, "正在掃描文件...\n")
    window.update_idletasks()
    total_files = count_docx_files(input_path)
    if total_files == 0:
        log_text.insert(tk.END, "⚠️ 找不到任何 Word 文件\n")
        return
    
    log_text.insert(tk.END, f"找到 {total_files} 個 Word 文件\n")
    log_text.insert(tk.END, "開始轉換...\n")
    window.update_idletasks()
    
    # 開始轉換
    conversion_running = True
    
    # 更新按鈕狀態
    convert_button['state'] = tk.DISABLED
    stop_button['state'] = tk.NORMAL
    
    # 重置進度條
    progress_bar["value"] = 0
    progress_label.config(text=f"進度: 0/{total_files} (0%)")
    
    # 在新線程中執行轉換，避免阻塞主界面
    threading.Thread(
        target=convert_docx_to_pdf_custom, 
        args=(input_path, output_path, 
              lambda msg: window.after(0, lambda: add_log(msg)), 
              update_progress,
              lambda: not conversion_running)
    ).start()

def add_log(message):
    log_text.insert(tk.END, message + "\n")
    log_text.see(tk.END)  # 自動捲動到底部

def start_conversion():
    start_conversion_thread()

def stop_conversion():
    global conversion_running
    if conversion_running:
        conversion_running = False
        add_log("⏱️ 正在停止轉換過程，請稍候...")
        stop_button['state'] = tk.DISABLED  # 禁用停止按鈕，防止重複點擊
        # 轉換線程將自行結束，在結束後恢復按鈕狀態
        window.after(100, check_conversion_status)  # 定期檢查轉換狀態

def check_conversion_status():
    global conversion_running
    if not conversion_running:
        convert_button['state'] = tk.NORMAL
        stop_button['state'] = tk.DISABLED
        add_log("⚡ 轉換過程已停止")
    else:
        window.after(100, check_conversion_status)  # 繼續檢查

def open_output_folder():
    output_path = output_entry.get()
    if os.path.exists(output_path):
        os.startfile(output_path)
    else:
        messagebox.showerror("錯誤", "輸出資料夾不存在")

def toggle_dark_mode():
    current_mode = dark_mode_var.get()
    if current_mode:
        # 深色模式
        style.configure("TFrame", background="#2E2E2E")
        style.configure("TLabel", background="#2E2E2E", foreground="#FFFFFF")
        style.configure("TButton", background="#3D3D3D", foreground="#FFFFFF")
        style.configure("Green.TButton", background="#2E7D32", foreground="#FFFFFF")
        style.configure("Red.TButton", background="#C62828", foreground="#FFFFFF")
        style.configure("Blue.TButton", background="#1565C0", foreground="#FFFFFF")
        style.configure("TEntry", fieldbackground="#3D3D3D", foreground="#FFFFFF")
        style.configure("Horizontal.TProgressbar", background="#4CAF50")
        log_text.config(bg="#424242", fg="#FFFFFF", insertbackground="#FFFFFF")
        window.configure(background="#2E2E2E")
        title_label.config(background="#2E2E2E", foreground="#4CAF50")
        main_frame.configure(background="#2E2E2E")
        footer_frame.configure(background="#2E2E2E")
        mode_checkbutton.config(text="🌙", background="#2E2E2E", foreground="#FFFFFF",
                              activebackground="#2E2E2E", activeforeground="#FFFFFF", selectcolor="#2E2E2E")
    else:
        # 亮色模式
        style.configure("TFrame", background="#F0F0F0")
        style.configure("TLabel", background="#F0F0F0", foreground="#000000")
        style.configure("TButton", background="#E0E0E0", foreground="#000000")
        style.configure("Green.TButton", background="#4CAF50", foreground="#FFFFFF")
        style.configure("Red.TButton", background="#F44336", foreground="#FFFFFF")
        style.configure("Blue.TButton", background="#2196F3", foreground="#FFFFFF")
        style.configure("TEntry", fieldbackground="#FFFFFF", foreground="#000000")
        style.configure("Horizontal.TProgressbar", background="#4CAF50")
        log_text.config(bg="#FFFFFF", fg="#000000", insertbackground="#000000")
        window.configure(background="#F0F0F0")
        title_label.config(background="#F0F0F0", foreground="#4CAF50")
        main_frame.configure(background="#F0F0F0")
        footer_frame.configure(background="#F0F0F0")
        mode_checkbutton.config(text="☀️", background="#F0F0F0", foreground="#000000",
                              activebackground="#F0F0F0", activeforeground="#000000", selectcolor="#F0F0F0")

def on_closing():
    global conversion_running
    if conversion_running:
        if messagebox.askyesno("退出確認", "轉換過程正在進行中，確定要退出嗎？"):
            conversion_running = False
            window.after(1000, window.destroy)  # 給轉換過程一些時間來清理
    else:
        window.destroy()

# GUI 介面
window = tk.Tk()
window.title("Word 批次轉 PDF 工具")
window.minsize(700, 500)  # 設置最小視窗大小
window.geometry("800x600")  # 設置預設視窗大小
window.protocol("WM_DELETE_WINDOW", on_closing)  # 處理窗口關閉事件

# 設置主題風格
style = ttk.Style()
style.theme_use("clam")

# 建立主框架
main_frame = ttk.Frame(window, padding=(20, 10))
main_frame.pack(fill=tk.BOTH, expand=True)

# 應用程式標題
title_label = tk.Label(main_frame, text="Word 批次轉 PDF 工具", font=("Arial", 18, "bold"))
title_label.pack(pady=(0, 20))

# 輸入/輸出框架
io_frame = ttk.Frame(main_frame)
io_frame.pack(fill=tk.X, pady=10)

# 輸入資料夾
input_frame = ttk.Frame(io_frame)
input_frame.pack(fill=tk.X, pady=5)

ttk.Label(input_frame, text="來源資料夾：", width=12).pack(side=tk.LEFT, padx=(0, 5))
input_entry = ttk.Entry(input_frame)
input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
ttk.Button(input_frame, text="選擇", command=lambda: browse_folder(input_entry)).pack(side=tk.LEFT)

# 輸出資料夾
output_frame = ttk.Frame(io_frame)
output_frame.pack(fill=tk.X, pady=5)

ttk.Label(output_frame, text="輸出資料夾：", width=12).pack(side=tk.LEFT, padx=(0, 5))
output_entry = ttk.Entry(output_frame)
output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
ttk.Button(output_frame, text="選擇", command=lambda: browse_folder(output_entry, True)).pack(side=tk.LEFT)

# 進度條
progress_frame = ttk.Frame(main_frame)
progress_frame.pack(fill=tk.X, pady=10)

progress_bar = ttk.Progressbar(progress_frame, orient="horizontal", mode="determinate", length=100)
progress_bar.pack(fill=tk.X, padx=(0, 10), side=tk.LEFT, expand=True)

progress_label = ttk.Label(progress_frame, text="進度: 0/0 (0%)")
progress_label.pack(side=tk.LEFT, padx=(0, 10))

# 按鈕框架
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=10)

# 開始按鈕
convert_button = ttk.Button(button_frame, text="開始轉換", command=start_conversion, style="Green.TButton", width=15)
convert_button.pack(side=tk.LEFT, padx=5)

# 停止按鈕
stop_button = ttk.Button(button_frame, text="停止轉換", command=stop_conversion, style="Red.TButton", width=15, state=tk.DISABLED)
stop_button.pack(side=tk.LEFT, padx=5)

# 打開輸出資料夾按鈕
open_output_button = ttk.Button(button_frame, text="打開輸出資料夾", command=open_output_folder, style="Blue.TButton", width=15)
open_output_button.pack(side=tk.LEFT, padx=5)

# 日誌區域標題
log_label = ttk.Label(main_frame, text="轉換日誌")
log_label.pack(anchor=tk.W, pady=(10, 5))

# 日誌區域
log_frame = ttk.Frame(main_frame)
log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

log_text = tk.Text(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# 添加滾動條
scrollbar = ttk.Scrollbar(log_frame, command=log_text.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
log_text.config(yscrollcommand=scrollbar.set)

# 底部框架
footer_frame = ttk.Frame(window, padding=(20, 5))
footer_frame.pack(fill=tk.X)

# 版權信息
copyright_label = ttk.Label(footer_frame, text="©️ 2025 Word 批次轉 PDF 工具 - 版本 1.0")
copyright_label.pack(side=tk.LEFT)

# 深色/淺色模式切換
dark_mode_var = tk.BooleanVar(value=False)
mode_checkbutton = tk.Checkbutton(footer_frame, text="☀️", variable=dark_mode_var, 
                               command=toggle_dark_mode, cursor="hand2",
                               bd=0, relief=tk.FLAT, highlightthickness=0)
mode_checkbutton.pack(side=tk.RIGHT)

# 初始化視覺樣式
style.configure("Green.TButton", background="#4CAF50", foreground="#FFFFFF")
style.configure("Red.TButton", background="#F44336", foreground="#FFFFFF")
style.configure("Blue.TButton", background="#2196F3", foreground="#FFFFFF")

# 添加一些初始提示
log_text.insert(tk.END, "📄 歡迎使用 Word 批次轉 PDF 工具！\n")
log_text.insert(tk.END, "請選擇來源資料夾（包含 Word 文件的資料夾）和輸出資料夾，然後點擊「開始轉換」按鈕。\n")
log_text.insert(tk.END, "✨ 提示：轉換過程中可以隨時點擊「停止轉換」按鈕來中斷操作。\n")

window.mainloop()