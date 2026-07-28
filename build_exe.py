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
    """打包成 Windows/macOS GUI 程序。"""
    app_name = "ImageWeeklyTool"
    print("=" * 60)
    print("开始打包图片转换与周报合并工具...")
    print(f"系统: {platform.system()} {platform.architecture()[0]}")
    print("=" * 60)

    args = [
        "png_to_jpg_converter.py",
        f"--name={app_name}",
        "--onefile",
        "--windowed",
        "--clean",
        "--noconfirm",
        "--hidden-import=weekly_report",
        "--hidden-import=analysis_template",
        "--collect-all=openpyxl",
        "--collect-all=xlrd",
    ]

    if platform.system() == "Windows" and os.path.exists("icon.ico"):
        args.append("--icon=icon.ico")
    elif platform.system() == "Darwin":
        print("\n⚠ macOS 将生成应用包；Windows EXE 由 GitHub Actions 构建。")

    try:
        PyInstaller.__main__.run(args)
        print("\n" + "=" * 60)
        print("✓ 打包完成!")
        print("=" * 60)
        if platform.system() == "Windows":
            print(f"\n✓ EXE 位置: dist/{app_name}.exe")
        else:
            print(f"\n✓ 应用位置: dist/{app_name}.app")
        print("\n功能:")
        print("1. PNG 单张、多张或文件夹批量转 JPG")
        print("2. XLSX、XLS、CSV 周报合并")
        print("3. 周报基础字段和追加指标可动态配置")
        print("4. XLSX、XLS 原始数据生成经营分析模板")
    except Exception as error:
        print(f"\n✗ 打包失败: {error}")
        raise


if __name__ == "__main__":
    build_exe()
