import os
import fitz  # PyMuPDF
import time

def compress_pdf_safe(input_pdf, output_pdf=None, compression_level="medium"):
    """
    使用更安全的方式壓縮 PDF 檔案
    
    Args:
        input_pdf (str): 輸入 PDF 檔案路徑
        output_pdf (str, optional): 輸出 PDF 檔案路徑
        compression_level (str): 壓縮等級，可選 "low", "medium", "high"
    """
    # 如果沒有指定輸出路徑，自動生成
    if output_pdf is None:
        file_name, file_ext = os.path.splitext(input_pdf)
        output_pdf = f"{file_name}_compressed{file_ext}"
    
    # 獲取原始檔案大小
    original_size = os.path.getsize(input_pdf)
    original_size_mb = original_size / (1024 * 1024)
    
    print(f"開始壓縮 PDF: {input_pdf}")
    print(f"原始檔案大小: {original_size_mb:.2f} MB")
    
    # 設定壓縮參數
    if compression_level == "low":
        compress_params = {
            "deflate": True,
            "garbage": 2,
            "clean": True,
            "pretty": False,
        }
    elif compression_level == "medium":
        compress_params = {
            "deflate": True,
            "garbage": 3,
            "clean": True,
            "pretty": False,
        }
    elif compression_level == "high":
        compress_params = {
            "deflate": True,
            "deflate_images": True,
            "deflate_fonts": True,
            "garbage": 4,
            "clean": True,
            "pretty": False,
        }
    else:
        raise ValueError("壓縮等級必須是 'low', 'medium' 或 'high'")
    
    start_time = time.time()
    
    # 開啟 PDF 檔案
    pdf_document = fitz.open(input_pdf)
    
    # 建立新的 PDF 文件來存放壓縮後的內容
    new_pdf = fitz.open()
    
    # 逐頁複製到新文件
    for page_num in range(len(pdf_document)):
        # 從原始 PDF 複製頁面到新 PDF
        new_pdf.insert_pdf(pdf_document, from_page=page_num, to_page=page_num)
        print(f"已處理頁面 {page_num + 1}/{len(pdf_document)}")
    
    # 保存壓縮後的 PDF
    new_pdf.save(output_pdf, **compress_params)
    
    # 關閉文件
    pdf_document.close()
    new_pdf.close()
    
    # 獲取壓縮後的檔案大小
    compressed_size = os.path.getsize(output_pdf)
    compressed_size_mb = compressed_size / (1024 * 1024)
    compression_ratio = (1 - compressed_size / original_size) * 100
    
    print(f"\n壓縮完成！")
    print(f"原始檔案大小: {original_size_mb:.2f} MB")
    print(f"壓縮後大小: {compressed_size_mb:.2f} MB")
    print(f"壓縮率: {compression_ratio:.2f}%")
    print(f"耗時: {time.time() - start_time:.2f} 秒")
    print(f"輸出檔案: {output_pdf}")
    
    return output_pdf, original_size, compressed_size

if __name__ == "__main__":
    # PDF 檔案路徑
    input_pdf = r"C:\Users"
    
    # 可選：指定輸出路徑
    # output_pdf = r"C:\Users\"
    
    # 執行壓縮（可選 "low", "medium", "high"）
    compress_pdf_safe(input_pdf, compression_level="high")