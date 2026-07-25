#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试转换功能（命令行版本）
"""

from PIL import Image
from pathlib import Path

def convert_png_to_jpg(png_path, jpg_path):
    """将 PNG 转换为 JPG"""
    with Image.open(png_path) as img:
        # 如果图片有透明通道，需要转换为 RGB
        if img.mode in ('RGBA', 'LA', 'P'):
            # 创建白色背景
            background = Image.new('RGB', img.size, (255, 255, 255))
            if img.mode == 'P':
                img = img.convert('RGBA')
            background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')
        
        # 保存为 JPG
        img.save(jpg_path, 'JPEG', quality=95)

def test_converter():
    """测试转换器"""
    input_folder = Path("test_images")
    output_folder = input_folder / "jpg_output"
    output_folder.mkdir(exist_ok=True)
    
    # 查找所有 PNG 文件
    png_files = list(input_folder.glob("*.png"))
    
    print(f"找到 {len(png_files)} 个 PNG 文件")
    print(f"输入文件夹: {input_folder}")
    print(f"输出文件夹: {output_folder}\n")
    
    for png_file in png_files:
        try:
            jpg_file = output_folder / (png_file.stem + ".jpg")
            convert_png_to_jpg(png_file, jpg_file)
            
            # 验证
            with Image.open(jpg_file) as img:
                print(f"✓ {png_file.name:25} -> {jpg_file.name:25} (大小: {img.size}, 模式: {img.mode})")
                
        except Exception as e:
            print(f"✗ {png_file.name}: {str(e)}")
    
    print(f"\n转换完成！JPG 文件保存在: {output_folder}")

if __name__ == "__main__":
    test_converter()
