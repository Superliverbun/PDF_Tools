from PIL import Image
import os
from glob import glob

def convert_images_to_pdf(image_folder, output_pdf, image_types=("*.jpg", "*.jpeg", "*.png")):
    """
    使用 Pillow 將資料夾中的圖片合併為單一 PDF 檔案
    
    Args:
        image_folder (str): 圖片所在資料夾路徑
        output_pdf (str): 輸出 PDF 檔案的路徑
        image_types (tuple): 要包含的圖片類型
    """
    # 獲取所有符合條件的圖片路徑
    image_files = []
    for image_type in image_types:
        files = glob(os.path.join(image_folder, image_type))
        image_files.extend(files)
    
    # 按名稱排序
    image_files.sort()
    
    if not image_files:
        print(f"在 {image_folder} 中找不到圖片檔案")
        return
    
    # 開啟所有圖片
    images = []
    for image_file in image_files:
        img = Image.open(image_file)
        # 如果不是 RGB 模式 (例如 PNG 的 RGBA)，轉換為 RGB
        if img.mode != "RGB":
            img = img.convert("RGB")
        images.append(img)
    
    # 使用第一張圖片的格式儲存所有圖片為 PDF
    if images:
        first_image = images[0]
        other_images = images[1:] if len(images) > 1 else []
        first_image.save(
            output_pdf, 
            "PDF", 
            resolution=100.0, 
            save_all=True,
            append_images=other_images
        )
        print(f"已將 {len(images)} 張圖片合併為 PDF: {output_pdf}")
        print(f"已包含的圖片: {[os.path.basename(f) for f in image_files]}")

if __name__ == "__main__":
    # 圖片所在資料夾
    image_folder = r"C:\Users\"
    
    # 輸出 PDF 檔案路徑
    output_pdf = r"C:\Users\"
    
    # 執行轉換
    convert_images_to_pdf(image_folder, output_pdf)