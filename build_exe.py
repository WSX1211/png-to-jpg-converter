#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
打包脚本 - 将 PNG 转 JPG 转换器打包成 exe 文件
支持 Windows 和 macOS
"""

import PyInstaller.__main__
import os
import sys
import platform

# 解决 Windows 控制台编码问题
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def build_exe():
    """打包成 exe"""
    print("=" * 60)
    print("开始打包 PNG 转 JPG 转换器...")
    print(f"系统: {platform.system()} {platform.architecture()[0]}")
    print("=" * 60)
    
    # PyInstaller 基本参数
    args = [
        'png_to_jpg_converter.py',           # 主脚本
        '--name=PNG转JPG转换器',              # exe 名称
        '--onefile',                          # 打包成单个文件
        '--windowed',                         # GUI 程序不显示控制台窗口
        '--clean',                            # 清理临时文件
        '--noconfirm',                        # 不询问确认
        '--hidden-import=weekly_report',      # 显式包含周报合并模块
        '--collect-all=openpyxl',             # 包含 XLSX 读取/写入资源
        '--collect-all=xlrd',                 # 包含 XLS 读取资源
    ]
    
    # Windows 特定设置
    if platform.system() == 'Windows':
        # 如果有图标文件，添加图标
        if os.path.exists('icon.ico'):
            args.append('--icon=icon.ico')
    
    # macOS 特定设置
    elif platform.system() == 'Darwin':
        # macOS 使用冒号作为分隔符
        args = [arg.replace(';', ':') for arg in args]
        
        # macOS 的 windowed 模式建议使用 onedir
        print("\n⚠ 注意: macOS 上建议使用 onedir 模式")
        print("  将生成 .app 应用包")
    
    try:
        PyInstaller.__main__.run(args)
        print("\n" + "=" * 60)
        print("✓ 打包完成!")
        print("=" * 60)
        
        # 根据系统显示不同的输出信息
        if platform.system() == 'Windows':
            print(f"\n✓ exe 文件位置: dist/PNG转JPG转换器.exe")
        else:
            print(f"\n✓ 应用位置: dist/PNG转JPG转换器.app")
        
        print("\n使用方法:")
        print("1. 双击运行程序")
        print("2. 拖拽包含 PNG 图片的文件夹到窗口")
        print("3. 或点击「选择文件夹」按钮选择文件夹")
        print("4. 转换后的 JPG 图片保存在原文件夹的 jpg_output 子文件夹中")
        
        print("\n⚠ 重要提示:")
        print("- 此程序需要在对应系统上打包")
        print("- Windows: 在 Windows 上运行此脚本生成 .exe")
        print("- macOS: 在 macOS 上运行此脚本生成 .app")
        print("- 两者不能交叉使用")
        
    except Exception as e:
        print(f"\n✗ 打包失败: {str(e)}")
        print("\n请确保已安装 PyInstaller:")
        print("  pip install pyinstaller")


if __name__ == "__main__":
    build_exe()
