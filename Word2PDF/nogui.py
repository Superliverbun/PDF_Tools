import os
import win32com.client

def convert_docx_to_pdf_custom(root_folder, output_folder):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    for foldername, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith(".docx") and not filename.startswith("~$"):
                docx_path = os.path.join(foldername, filename)

                # 找出 root_folder 的下一層資料夾（例如：hw1）
                rel_path = os.path.relpath(foldername, root_folder)
                top_level_folder = rel_path.split(os.sep)[0]  # 例如 hw1

                # 建立輸出資料夾：Output/hw1
                output_subfolder = os.path.join(output_folder, top_level_folder)
                os.makedirs(output_subfolder, exist_ok=True)

                # 原始 Word 檔案名稱（不含中間資料夾）
                pdf_name = os.path.splitext(filename)[0] + ".pdf"
                pdf_path = os.path.join(output_subfolder, pdf_name)

                try:
                    doc = word.Documents.Open(docx_path)
                    doc.SaveAs(pdf_path, FileFormat=17)
                    doc.Close()
                    print(f"✅ {docx_path} → {pdf_path}")
                except Exception as e:
                    print(f"❌ 無法轉換 {docx_path}，原因：{e}")

    word.Quit()
    print("🎉 所有 Word 檔轉 PDF 完成！")

# 修改這兩個變數為你自己的路徑：
root_folder = r"C:\Users\檔案路徑"
output_folder = r"C:\Users\輸出路徑"

convert_docx_to_pdf_custom(root_folder, output_folder)
