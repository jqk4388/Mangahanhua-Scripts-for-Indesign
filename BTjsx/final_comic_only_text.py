# 說明 python final_comic_only_text.py ./comic_folder/ -c "255,0,0" -o 0.8


import sys
import os
import argparse
import numpy as np
from PIL import Image, ImageFilter
import scipy.ndimage as ndimage

# ==========================================
# ⚙️ 默認配置
# ==========================================
DEFAULT_COLOR = "255,0,255"  # 默認洋红色（Magenta）
DEFAULT_OPACITY = 0.4        # 默認半透明
MASK_EXPANSION = 5           # 擴張程度
SATURATION_THRESHOLD = 30    # 去彩噪
KEEP_BLACK_THRESHOLD = 180   # 白底保留閾值
KEEP_WHITE_THRESHOLD = 100   # 黑底保留閾值

# 支持的圖片格式
VALID_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.webp'}
# ==========================================

def parse_color_string(color_str):
    color_str = color_str.strip()
    if color_str.startswith("#"): color_str = color_str[1:]
    if len(color_str) == 6 and "," not in color_str:
        try:
            return (int(color_str[0:2], 16), int(color_str[2:4], 16), int(color_str[4:6], 16))
        except ValueError: pass
    parts = color_str.split(",")
    if len(parts) == 3:
        try:
            return (int(parts[0]), int(parts[1]), int(parts[2]))
        except ValueError: pass
    return (255, 255, 0)

def process_single_image(original_path, target_color_rgb, target_opacity):
    """處理單張圖片的核心邏輯"""
    # 1. 路徑準備
    dir_name = os.path.dirname(original_path)
    full_file_name = os.path.basename(original_path)
    file_name_no_ext, _ = os.path.splitext(full_file_name)
    
    # Mask 路徑
    mask_path = os.path.join(dir_name, "mask", f"{file_name_no_ext}.png")
    
    # 輸出路徑
    output_dir = os.path.join(dir_name, "only_text")
    output_filename = f"{file_name_no_ext}.png"
    output_path = os.path.join(output_dir, output_filename)

    # 2. 檢查必要文件
    if not os.path.exists(mask_path):
        print(f"⚠️ 跳過: {full_file_name} (找不到對應 Mask: {mask_path})")
        return

    # 3. 建立輸出資料夾
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    try:
        print(f"🔄 正在處理: {full_file_name} ...", end="\r")

        img_orig = Image.open(original_path).convert("RGBA")
        img_mask = Image.open(mask_path).convert("L")

        if img_orig.size != img_mask.size:
            img_mask = img_mask.resize(img_orig.size)
        if MASK_EXPANSION > 0:
            img_mask = img_mask.filter(ImageFilter.MaxFilter(MASK_EXPANSION))

        data_orig = np.array(img_orig)
        data_mask = np.array(img_mask)
        data_gray = np.array(img_orig.convert("L"))
        data_sat = np.array(img_orig.convert("HSV"))[:, :, 1]

        binary_mask = data_mask > 128
        labeled_array, num_features = ndimage.label(binary_mask)
        final_alpha = np.zeros_like(data_mask, dtype=np.uint8)

        for i in range(1, num_features + 1):
            blob_mask = (labeled_array == i)
            blob_pixels = data_gray[blob_mask]
            if len(blob_pixels) == 0: continue

            avg_brightness = np.mean(blob_pixels)
            is_white_bg = avg_brightness > 100
            is_achromatic = (data_sat < SATURATION_THRESHOLD)

            if is_white_bg:
                cond = (data_gray < KEEP_BLACK_THRESHOLD) & is_achromatic
            else:
                cond = (data_gray > KEEP_WHITE_THRESHOLD) & is_achromatic

            current_alpha = np.where(cond & blob_mask, 255, 0).astype(np.uint8)
            final_alpha = np.maximum(final_alpha, current_alpha)

        final_alpha = np.minimum(final_alpha, data_mask)

        final_data = np.zeros_like(data_orig)
        text_zone = final_alpha > 0

        final_data[text_zone, 0] = target_color_rgb[0]
        final_data[text_zone, 1] = target_color_rgb[1]
        final_data[text_zone, 2] = target_color_rgb[2]

        weighted_alpha = final_alpha.astype(np.float32) * target_opacity
        final_data[:, :, 3] = np.clip(weighted_alpha, 0, 255).astype(np.uint8)

        Image.fromarray(final_data).save(output_path, "PNG")
        print(f"✅ 完成: {output_filename}         ") # 空格是為了覆蓋上面的進度條文字

    except Exception as e:
        print(f"\n❌ 錯誤 {full_file_name}: {e}")

def main():
    parser = argparse.ArgumentParser(description="漫畫文字提取與上色工具 (支持文件夾批次處理)")
    parser.add_argument("input_path", help="原圖路徑 或 文件夾路徑")
    parser.add_argument("-c", "--color", default=DEFAULT_COLOR, help="文字顏色 (Hex或RGB)")
    parser.add_argument("-o", "--opacity", type=float, default=DEFAULT_OPACITY, help="文字透明度")
    args = parser.parse_args()

    # 解析顏色參數
    target_color_rgb = parse_color_string(args.color)
    input_path = args.input_path

    if not os.path.exists(input_path):
        print(f"❌ 錯誤: 路徑不存在 -> {input_path}")
        return

    # 判斷是文件還是文件夾
    if os.path.isfile(input_path):
        # --- 單文件模式 ---
        print("📂 檢測到單一文件，開始處理...")
        process_single_image(input_path, target_color_rgb, args.opacity)
    
    elif os.path.isdir(input_path):
        # --- 文件夾模式 (批次) ---
        print(f"📂 檢測到文件夾，掃描圖片中: {input_path}")
        
        # 獲取所有圖片文件 (不遞歸)
        image_files = [
            f for f in os.listdir(input_path) 
            if os.path.isfile(os.path.join(input_path, f)) 
            and os.path.splitext(f)[1].lower() in VALID_EXTENSIONS
        ]
        
        total_files = len(image_files)
        if total_files == 0:
            print("⚠️ 該文件夾下沒有發現圖片文件 (.jpg, .png, etc)")
            return

        print(f"🚀 找到 {total_files} 張圖片，開始批次處理...")
        print(f"🎨 設定顏色: {target_color_rgb} | 透明度: {args.opacity}")
        print("-" * 40)

        for i, filename in enumerate(image_files):
            full_path = os.path.join(input_path, filename)
            process_single_image(full_path, target_color_rgb, args.opacity)
            
        print("-" * 40)
        print(f"🎉 批次處理完成！請查看 {os.path.join(input_path, 'only_text')}")

if __name__ == "__main__":
    main()