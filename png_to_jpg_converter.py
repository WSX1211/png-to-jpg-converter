#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PNG to JPG Converter
拖拽文件夹，将里面的 PNG 图片转换为 JPG 格式
"""

import os
import sys
from pathlib import Path
from PIL import Image
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading


class PNGtoJPGConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PNG 转 JPG 转换器")
        self.root.geometry("600x400")
        self.root.resizable(True, True)
        
        # 设置窗口图标（如果有的话）
        try:
            self.root.iconbitmap("icon.ico")
        except:
            pass
        
        self.setup_ui()
        self.setup_drag_drop()
        
    def setup_ui(self):
        """设置界面"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="PNG 转 JPG 转换器", 
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=(0, 10))
        
        # 说明文字
        info_label = ttk.Label(
            main_frame,
            text="拖拽包含 PNG 图片的文件夹到下方区域\n或点击按钮选择文件夹",
            font=("Arial", 10),
            justify=tk.CENTER
        )
        info_label.pack(pady=(0, 15))
        
        # 拖拽区域框架
        self.drop_frame = ttk.LabelFrame(
            main_frame, 
            text="拖拽文件夹到此处", 
            padding="30"
        )
        self.drop_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 拖拽提示标签
        self.drop_label = ttk.Label(
            self.drop_frame,
            text="📁\n拖拽文件夹到这里\n或",
            font=("Arial", 12),
            justify=tk.CENTER
        )
        self.drop_label.pack(expand=True)
        
        # 选择文件夹按钮
        self.select_btn = ttk.Button(
            self.drop_frame,
            text="选择文件夹",
            command=self.select_folder
        )
        self.select_btn.pack(pady=10)
        
        # 进度条框架
        progress_frame = ttk.Frame(main_frame)
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 进度条
        self.progress = ttk.Progressbar(
            progress_frame, 
            mode='determinate',
            length=400
        )
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # 进度标签
        self.progress_label = ttk.Label(progress_frame, text="就绪")
        self.progress_label.pack(side=tk.RIGHT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="转换日志", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        # 日志文本框
        self.log_text = tk.Text(log_frame, height=8, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 底部按钮框架
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(15, 0))
        
        # 清空日志按钮
        self.clear_btn = ttk.Button(
            button_frame,
            text="清空日志",
            command=self.clear_log
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 退出按钮
        self.quit_btn = ttk.Button(
            button_frame,
            text="退出",
            command=self.root.quit
        )
        self.quit_btn.pack(side=tk.RIGHT, padx=5)
        
    def setup_drag_drop(self):
        """设置拖拽功能"""
        # 尝试使用 tkinterdnd2 库实现拖拽
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # 重新创建支持拖拽的窗口
            self.root.destroy()
            self.root = TkinterDnD.Tk()
            self.root.title("PNG 转 JPG 转换器")
            self.root.geometry("600x400")
            
            # 重新设置界面
            self.setup_ui()
            
            # 绑定拖拽事件
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
            self.drop_frame.dnd_bind('<<DragEnter>>', self.on_drag_enter)
            self.drop_frame.dnd_bind('<<DragLeave>>', self.on_drag_leave)
            
            self.use_tkdnd = True
            self.log("✓ 已启用拖拽功能（使用 tkinterdnd2）")
            
        except ImportError:
            self.use_tkdnd = False
            self.log("⚠ 未安装 tkinterdnd2，拖拽功能不可用")
            self.log("💡 提示: 请点击「选择文件夹」按钮来选择包含 PNG 图片的文件夹")
            self.log("💡 安装拖拽支持: pip install tkinterdnd2")
    
    def on_drag_enter(self, event):
        """拖拽进入"""
        self.drop_frame.config(text="松开鼠标以处理文件夹")
        self.drop_label.config(text="📂\n正在处理...")
        
    def on_drag_leave(self, event):
        """拖拽离开"""
        self.drop_frame.config(text="拖拽文件夹到此处")
        self.drop_label.config(text="📁\n拖拽文件夹到这里\n或")
        
    def on_drop(self, event):
        """处理拖拽放下事件"""
        # 获取拖拽的路径
        paths = self.root.tk.splitlist(event.data)
        
        if paths:
            folder_path = paths[0].strip('{}')
            self.process_folder(folder_path)
            
        # 恢复界面
        self.on_drag_leave(None)
        
    def select_folder(self):
        """选择文件夹"""
        folder_path = filedialog.askdirectory(title="选择包含 PNG 图片的文件夹")
        if folder_path:
            self.process_folder(folder_path)
            
    def process_folder(self, folder_path):
        """处理文件夹中的 PNG 图片"""
        folder_path = Path(folder_path)
        
        if not folder_path.exists():
            messagebox.showerror("错误", f"文件夹不存在: {folder_path}")
            return
            
        if not folder_path.is_dir():
            messagebox.showerror("错误", f"不是有效的文件夹: {folder_path}")
            return
        
        # 在新线程中处理，避免界面卡顿
        thread = threading.Thread(target=self._process_folder_thread, args=(folder_path,))
        thread.daemon = True
        thread.start()
        
    def _process_folder_thread(self, folder_path):
        """在后台线程中处理文件夹"""
        try:
            # 查找所有 PNG 文件
            png_files = list(folder_path.glob("*.png")) + list(folder_path.glob("*.PNG"))
            
            if not png_files:
                self.root.after(0, lambda: messagebox.showinfo("提示", f"文件夹中没有找到 PNG 图片:\n{folder_path}"))
                return
            
            self.log(f"\n{'='*50}")
            self.log(f"开始处理文件夹: {folder_path}")
            self.log(f"找到 {len(png_files)} 个 PNG 文件")
            
            # 创建输出文件夹
            output_folder = folder_path / "jpg_output"
            output_folder.mkdir(exist_ok=True)
            self.log(f"输出文件夹: {output_folder}")
            
            # 设置进度条
            self.root.after(0, lambda: self.progress.config(maximum=len(png_files), value=0))
            
            success_count = 0
            error_count = 0
            
            for i, png_file in enumerate(png_files, 1):
                try:
                    # 转换图片
                    jpg_file = output_folder / (png_file.stem + ".jpg")
                    self.convert_png_to_jpg(png_file, jpg_file)
                    
                    success_count += 1
                    self.log(f"✓ [{i}/{len(png_files)}] 转换成功: {png_file.name} -> {jpg_file.name}")
                    
                except Exception as e:
                    error_count += 1
                    self.log(f"✗ [{i}/{len(png_files)}] 转换失败: {png_file.name} - {str(e)}")
                
                # 更新进度
                self.root.after(0, lambda v=i: self.progress.config(value=v))
                self.root.after(0, lambda v=i, t=len(png_files): self.progress_label.config(text=f"{v}/{t}"))
            
            # 完成
            self.log(f"\n{'='*50}")
            self.log(f"转换完成!")
            self.log(f"成功: {success_count} 个")
            self.log(f"失败: {error_count} 个")
            self.log(f"输出位置: {output_folder}")
            self.log(f"{'='*50}\n")
            
            # 显示完成消息
            self.root.after(0, lambda: messagebox.showinfo(
                "完成", 
                f"转换完成!\n\n成功: {success_count} 个\n失败: {error_count} 个\n\n输出文件夹:\n{output_folder}"
            ))
            
        except Exception as e:
            self.log(f"处理过程中出错: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"处理过程中出错:\n{str(e)}"))
            
    def convert_png_to_jpg(self, png_path, jpg_path):
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
            
    def log(self, message):
        """添加日志"""
        self.root.after(0, lambda: self.log_text.insert(tk.END, message + "\n"))
        self.root.after(0, lambda: self.log_text.see(tk.END))
        
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
        self.progress.config(value=0)
        self.progress_label.config(text="就绪")


def main():
    """主函数"""
    root = tk.Tk()
    app = PNGtoJPGConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
