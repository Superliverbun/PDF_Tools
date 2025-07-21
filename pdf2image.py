from pdf2image import convert_from_path
import os

def convert_pdf_to_jpg(pdf_path, output_folder=None):
    """將 PDF 檔案轉換為 JPG 圖片
    
    Args:
        pdf_path (str): PDF 檔案路徑
        output_folder (str): 輸出資料夾 (預設: 與 PDF 相同資料夾)
    """
    # 如果沒有指定輸出資料夾，使用 PDF 所在的資料夾
    if output_folder is None:
        output_folder = os.path.dirname(pdf_path)
    
    # 建立輸出資料夾（如果不存在）
    os.makedirs(output_folder, exist_ok=True)
    
    # 取得不含副檔名的 PDF 檔名
    pdf_filename = os.path.splitext(os.path.basename(pdf_path))[0]
    
    # 轉換 PDF 為圖片
    # Windows 用戶: 如需指定 poppler 路徑，加入 poppler_path="C:\\path\\to\\poppler\\bin" 參數
    images = convert_from_path(pdf_path, dpi=300)
    
    # 儲存圖片
    for i, image in enumerate(images):
        image_path = os.path.join(output_folder, f"{pdf_filename}_page_{i+1}.jpg")
        image.save(image_path, "JPEG")
        print(f"已儲存 {image_path}")

# 使用範例
if __name__ == "__main__":
    # 替換為你的 PDF 檔案路徑
    pdf_file = "example.pdf"
    
    # 可選：指定輸出資料夾
    # output_folder = "output_images"
    
    convert_pdf_to_jpg(pdf_file)
    print("轉換完成！")