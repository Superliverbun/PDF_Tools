import os
import fitz  # PyMuPDF
from PIL import Image
from glob import glob

def combine_to_pdf(input_folder, output_pdf, image_types=("*.jpg", "*.jpeg", "*.png"), include_pdf=True):
    """
    將資料夾中的圖片和 PDF 檔案合併為單一 PDF 檔案
    
    Args:
        input_folder (str): 輸入檔案所在資料夾路徑
        output_pdf (str): 輸出 PDF 檔案的路徑
        image_types (tuple): 要包含的圖片類型
        include_pdf (bool): 是否包含資料夾中的 PDF 檔案
    """
    # 建立新的 PDF 文件
    pdf_output = fitz.open()
    
    # 處理所有圖片檔案
    image_files = []
    for image_type in image_types:
        files = glob(os.path.join(input_folder, image_type))
        image_files.extend(files)
    
    # 排序圖片檔案
    image_files.sort()
    
    # 加入圖片到 PDF
    for img_path in image_files:
        try:
            # 使用 PyMuPDF 插入圖片
            img_doc = fitz.open()
            rect = fitz.Rect(0, 0, 595, 842)  # A4 大小
            
            # 創建新頁面
            page = img_doc.new_page(width=rect.width, height=rect.height)
            
            # 插入圖片
            page.insert_image(rect, filename=img_path)
            
            # 將頁面加入輸出 PDF
            pdf_output.insert_pdf(img_doc)
            print(f"已加入圖片: {os.path.basename(img_path)}")
        except Exception as e:
            print(f"處理圖片 {img_path} 時發生錯誤: {e}")
    
    # 處理 PDF 檔案
    if include_pdf:
        pdf_files = glob(os.path.join(input_folder, "*.pdf"))
        pdf_files.sort()
        
        for pdf_path in pdf_files:
            # 避免處理輸出檔案本身
            if os.path.abspath(pdf_path) == os.path.abspath(output_pdf):
                continue
            
            try:
                # 開啟 PDF 檔案
                pdf_doc = fitz.open(pdf_path)
                
                # 將整個 PDF 加入輸出文件
                pdf_output.insert_pdf(pdf_doc)
                print(f"已加入 PDF ({pdf_doc.page_count} 頁): {os.path.basename(pdf_path)}")
            except Exception as e:
                print(f"處理 PDF {pdf_path} 時發生錯誤: {e}")
    
    # 儲存合併後的 PDF
    if pdf_output.page_count > 0:
        pdf_output.save(output_pdf)
        print(f"\n成功將 {pdf_output.page_count} 頁合併為 PDF: {output_pdf}")
    else:
        print(f"沒有找到任何檔案可以合併")
    
    # 關閉 PDF 文件
    pdf_output.close()

if __name__ == "__main__":
    # 輸入檔案所在資料夾
    input_folder = r"C:\Users\"
    
    # 輸出 PDF 檔案路徑
    output_pdf = r"C:\Users\"
    
    # 執行合併
    combine_to_pdf(input_folder, output_pdf)