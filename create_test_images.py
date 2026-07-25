#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
创建测试用的 PNG 图片
"""

from PIL import Image
import os

def create_test_png_images():
    """创建测试用的 PNG 图片"""
    output_dir = "test_images"
    os.makedirs(output_dir, exist_ok=True)
    
    # 创建几个不同颜色的 PNG 图片
    colors = [
        ((255, 0, 0), "red"),      # 红色
        ((0, 255, 0), "green"),    # 绿色
        ((0, 0, 255), "blue"),     # 蓝色
        ((255, 255, 0), "yellow"), # 黄色
    ]
    
    # 创建带透明度的图片
    transparent_colors = [
        ((255, 0, 0, 128), "red_transparent"),    # 半透明红色
        ((0, 255, 0, 128), "green_transparent"),  # 半透明绿色
    ]
    
    print("创建测试 PNG 图片...")
    
    # 创建不透明图片
    for color, name in colors:
        img = Image.new('RGB', (200, 200), color)
        filename = os.path.join(output_dir, f"{name}.png")
        img.save(filename)
        print(f"✓ 创建: {filename}")
    
    # 创建透明图片
    for color, name in transparent_colors:
        img = Image.new('RGBA', (200, 200), color)
        filename = os.path.join(output_dir, f"{name}.png")
        img.save(filename)
        print(f"✓ 创建: {filename}")
    
    print(f"\n完成！共创建 {len(colors) + len(transparent_colors)} 个测试图片")
    print(f"测试图片位置: {output_dir}/")

if __name__ == "__main__":
    create_test_png_images()
