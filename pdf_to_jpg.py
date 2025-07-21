import fitz  # PyMuPDF
import os

def convert_pdf_to_jpg(pdf_path, output_folder=None):
    """
    Convert a PDF file to JPG images, one per page.

    Args:
        pdf_path (str): Path to the PDF file.
        output_folder (str, optional): Folder to save images. Defaults to PDF's folder.
    """

    # 如果沒有指定輸出資料夾，使用 PDF 所在的資料夾
    if output_folder is None:
        output_folder = os.path.dirname(pdf_path)
    
    # 建立輸出資料夾（如果不存在）
    os.makedirs(output_folder, exist_ok=True)
    
    # 取得不含副檔名的 PDF 檔名
    pdf_filename = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # 開啟 PDF 檔案
    pdf_document = fitz.open(pdf_path)
    
    # 遍歷每一頁
    for page_number in range(len(pdf_document)):
        # 取得頁面
        page = pdf_document.load_page(page_number)
        
        # 將頁面轉換為圖片 (設定解析度為 300dpi)
        pix = page.get_pixmap(matrix=fitz.Matrix(300/72, 300/72))
        
        # 定義輸出路徑
        image_path = os.path.join(output_folder, f"{pdf_filename}_page_{page_number+1}.jpg")
        
        # 儲存圖片
        pix.save(image_path)
        print(f"已儲存 {image_path}")
    
    # 關閉 PDF 檔案
    pdf_document.close()

# 使用範例
if __name__ == "__main__":
    # 替換為你的 PDF 檔案路徑
    pdf_file = r"C:\Users\"
    convert_pdf_to_jpg(pdf_file)
    print("轉換完成！")