import os
import subprocess
import time

def compress_pdf_with_ghostscript(input_pdf, output_pdf=None, compression_level="high", gs_path=None):
    """
    使用 Ghostscript 壓縮 PDF 檔案
    
    Args:
        input_pdf (str): 輸入 PDF 檔案路徑
        output_pdf (str, optional): 輸出 PDF 檔案路徑
        compression_level (str): 壓縮等級，可選 "screen", "ebook", "printer", "prepress", "high", "extreme"
        gs_path (str): Ghostscript 執行檔的完整路徑
    """
    if output_pdf is None:
        file_name, file_ext = os.path.splitext(input_pdf)
        output_pdf = f"{file_name}_compressed{file_ext}"
    
    # 獲取原始檔案大小
    original_size = os.path.getsize(input_pdf)
    original_size_mb = original_size / (1024 * 1024)
    
    print(f"開始壓縮 PDF: {input_pdf}")
    print(f"原始檔案大小: {original_size_mb:.2f} MB")
    
    # 找到 Ghostscript 執行檔
    if gs_path is None:
        # 嘗試幾個常見的安裝位置
        possible_paths = [
            "gswin64c",  # 如果在 PATH 中
            r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe",
            r"C:\Program Files\gs\gs10.00.0\bin\gswin64c.exe",
            r"C:\Program Files\gs\gs9.56.1\bin\gswin64c.exe",
            r"C:\Program Files\gs\gs9.55.0\bin\gswin64c.exe",
            # 檢查 Program Files (x86) 路徑
            r"C:\Program Files (x86)\gs\gs10.05.1\bin\gswin32c.exe",
            r"C:\Program Files (x86)\gs\gs10.00.0\bin\gswin32c.exe",
            r"C:\Program Files (x86)\gs\gs9.56.1\bin\gswin32c.exe",
            r"C:\Program Files (x86)\gs\gs9.55.0\bin\gswin32c.exe",
        ]
        
        # 檢查 Ghostscript 可能的安裝路徑
        gs_found = False
        for path in possible_paths:
            try:
                # 測試是否可執行
                if os.path.isfile(path):
                    subprocess.run([path, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    gs_path = path
                    gs_found = True
                    print(f"找到 Ghostscript: {path}")
                    break
                elif " " not in path:  # 命令可能在 PATH 中
                    subprocess.run([path, "--version"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    gs_path = path
                    gs_found = True
                    print(f"找到 Ghostscript: {path} (在 PATH 中)")
                    break
            except (subprocess.SubprocessError, FileNotFoundError):
                continue
        
        if not gs_found:
            raise FileNotFoundError("找不到 Ghostscript 執行檔。請安裝 Ghostscript 或提供完整路徑。")
    
    start_time = time.time()
    
    # 設定 Ghostscript 參數
    if compression_level == "screen":
        # 螢幕查看用 (72 dpi)
        gs_params = ["-dPDFSETTINGS=/screen"]
    elif compression_level == "ebook":
        # 電子書用 (150 dpi)
        gs_params = ["-dPDFSETTINGS=/ebook"]
    elif compression_level == "printer":
        # 列印用 (300 dpi)
        gs_params = ["-dPDFSETTINGS=/printer"]
    elif compression_level == "prepress":
        # 印刷用 (300 dpi, 保留更多資訊)
        gs_params = ["-dPDFSETTINGS=/prepress"]
    elif compression_level == "high":
        # 高壓縮自定義設定
        gs_params = [
            "-dColorImageResolution=72",
            "-dGrayImageResolution=72",
            "-dColorImageDownsampleType=/Bicubic",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dMonoImageDownsampleType=/Bicubic",
            "-dOptimize=true",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dAutoRotatePages=/None",
            "-dColorImageFilter=/DCTEncode",
            "-dGrayImageFilter=/DCTEncode",
            "-dCompatibilityLevel=1.5",
            "-dDetectDuplicateImages=true",
            "-dPDFA=false",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            "-dSAFER"
        ]
    elif compression_level == "extreme":
        # 極度壓縮設定，可能影響品質
        gs_params = [
            "-dColorImageResolution=50",
            "-dGrayImageResolution=50",
            "-dMonoImageResolution=50",
            "-dColorImageDownsampleType=/Bicubic",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dMonoImageDownsampleType=/Bicubic",
            "-dOptimize=true",
            "-dEmbedAllFonts=true",
            "-dSubsetFonts=true",
            "-dAutoRotatePages=/None",
            #"-dColorConversionStrategy=/Gray",  # 轉換為灰階
            "-dColorImageFilter=/DCTEncode",
            "-dGrayImageFilter=/DCTEncode",
            "-dCompatibilityLevel=1.5",
            "-dDetectDuplicateImages=true",
            "-dPDFA=false",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            "-dSAFER"
        ]
    else:
        raise ValueError("壓縮等級必須是 'screen', 'ebook', 'printer', 'prepress', 'high' 或 'extreme'")
    
    # 執行 Ghostscript 命令
    try:
        # 構建 Ghostscript 命令
        gs_command = [gs_path, "-sDEVICE=pdfwrite", f"-sOutputFile={output_pdf}"] + gs_params + [input_pdf]
        
        # 執行命令
        process = subprocess.run(gs_command, check=True, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
        
        # 檢查輸出檔案是否存在
        if not os.path.exists(output_pdf):
            raise Exception("壓縮失敗，未生成輸出檔案")
        
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
        
        # 如果壓縮後的檔案比原始檔案大，發出警告
        if compressed_size > original_size:
            print("\n警告: 壓縮後的檔案比原始檔案更大！")
        
        return output_pdf, original_size, compressed_size
        
    except subprocess.CalledProcessError as e:
        print(f"Ghostscript 執行失敗: {e.stderr.decode() if e.stderr else str(e)}")
        raise
    except Exception as e:
        print(f"壓縮失敗: {str(e)}")
        raise

if __name__ == "__main__":
    # PDF 檔案路徑
    input_pdf = r"C:\Users\"
    
    # 手動指定 Ghostscript 路徑（如果您知道確切位置）
    # 替換為您電腦上 Ghostscript 的實際路徑
    gs_path = r"C:\Program Files\gs\gs10.05.1\bin\gswin64c.exe"  
    
    # 如果您不確定 Ghostscript 位置，可以不指定 gs_path，讓程式自動搜尋
    # gs_path = None
    
    # 執行壓縮 - 從低壓縮開始嘗試，如果壓縮後仍大於 10MB，則使用更激進的壓縮設定
    try:
        # 先嘗試高壓縮
        output_pdf, _, compressed_size = compress_pdf_with_ghostscript(
            input_pdf, 
            compression_level="high",
            gs_path=gs_path
        )
        
        # 如果還是大於 10MB，嘗試極度壓縮
        if compressed_size > 10 * 1024 * 1024:
            print("\n文件仍大於 10MB，嘗試極度壓縮設定...")
            compress_pdf_with_ghostscript(
                input_pdf, 
                output_pdf, 
                compression_level="extreme",
                gs_path=gs_path
            )
    except Exception as e:
        print(f"壓縮過程中發生錯誤: {str(e)}")