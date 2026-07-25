#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PNG to JPG Converter
支持单张PNG图片或整个文件夹转换
"""

import os
import sys
from pathlib import Path
from PIL import Image
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import queue


class PNGtoJPGConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("PNG 转 JPG 转换器")
        self.root.geometry("700x550")
        self.root.resizable(True, True)
        
        # 创建消息队列用于线程间通信
        self.msg_queue = queue.Queue()
        
        # 初始化变量
        self.use_tkdnd = False
        
        try:
            self.setup_ui()
            self.setup_drag_drop()
            
            # 启动队列检查
            self.check_queue()
            
            self.log("✓ 程序启动成功")
            self.log("💡 使用说明:")
            self.log("  - 点击「选择文件」选择单个或多个PNG文件")
            self.log("  - 点击「选择文件夹」批量转换整个文件夹")
            self.log("  - 或直接拖拽文件/文件夹到窗口")
        except Exception as e:
            print(f"初始化失败: {e}")
            messagebox.showerror("初始化错误", f"程序初始化失败:\n{str(e)}")
    
    def check_queue(self):
        """检查消息队列，处理来自后台线程的消息"""
        try:
            while True:
                try:
                    msg = self.msg_queue.get_nowait()
                    msg_type = msg.get('type', '')
                    
                    if msg_type == 'log':
                        self.log_text.insert(tk.END, msg['text'] + "\n")
                        self.log_text.see(tk.END)
                    elif msg_type == 'progress':
                        self.progress['value'] = msg['value']
                    elif msg_type == 'progress_max':
                        self.progress['maximum'] = msg['value']
                    elif msg_type == 'progress_label':
                        self.progress_label.config(text=msg['text'])
                    elif msg_type == 'messagebox_info':
                        messagebox.showinfo("完成", msg['text'])
                    elif msg_type == 'messagebox_error':
                        messagebox.showerror("错误", msg['text'])
                        
                except queue.Empty:
                    break
        finally:
            # 每100ms检查一次队列
            self.root.after(100, self.check_queue)
        
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
            text="支持转换单个PNG文件或整个文件夹\n拖拽文件/文件夹到下方区域或点击按钮选择",
            font=("Arial", 10),
            justify=tk.CENTER
        )
        info_label.pack(pady=(0, 15))
        
        # 拖拽区域框架
        self.drop_frame = ttk.LabelFrame(
            main_frame, 
            text="拖拽文件或文件夹到此处", 
            padding="30"
        )
        self.drop_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 拖拽提示标签
        self.drop_label = ttk.Label(
            self.drop_frame,
            text="📁\n拖拽PNG文件或文件夹到这里\n或",
            font=("Arial", 12),
            justify=tk.CENTER
        )
        self.drop_label.pack(expand=True)
        
        # 按钮框架
        button_frame = ttk.Frame(self.drop_frame)
        button_frame.pack(pady=10)
        
        # 选择文件按钮
        self.select_file_btn = ttk.Button(
            button_frame,
            text="选择文件",
            command=self.select_file,
            width=15
        )
        self.select_file_btn.pack(side=tk.LEFT, padx=5)
        
        # 选择文件夹按钮
        self.select_folder_btn = ttk.Button(
            button_frame,
            text="选择文件夹",
            command=self.select_folder,
            width=15
        )
        self.select_folder_btn.pack(side=tk.LEFT, padx=5)
        
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
        self.log_text = tk.Text(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 底部按钮框架
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill=tk.X, pady=(15, 0))
        
        # 清空日志按钮
        self.clear_btn = ttk.Button(
            bottom_frame,
            text="清空日志",
            command=self.clear_log
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # 退出按钮
        self.quit_btn = ttk.Button(
            bottom_frame,
            text="退出",
            command=self.root.quit
        )
        self.quit_btn.pack(side=tk.RIGHT, padx=5)
        
    def setup_drag_drop(self):
        """设置拖拽功能"""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD
            
            # 重新创建支持拖拽的窗口
            self.root.destroy()
            self.root = TkinterDnD.Tk()
            self.root.title("PNG 转 JPG 转换器")
            self.root.geometry("700x550")
            
            # 重新设置界面
            self.setup_ui()
            
            # 绑定拖拽事件
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)
            self.drop_frame.dnd_bind('<<DragEnter>>', self.on_drag_enter)
            self.drop_frame.dnd_bind('<<DragLeave>>', self.on_drag_leave)
            
            self.use_tkdnd = True
            self.log("✓ 已启用拖拽功能")
            
            # 重新启动队列检查
            self.check_queue()
            
        except ImportError:
            self.use_tkdnd = False
            self.log("⚠ 未安装 tkinterdnd2，拖拽功能不可用")
        except Exception as e:
            self.use_tkdnd = False
            self.log(f"⚠ 拖拽功能初始化失败: {str(e)}")
    
    def on_drag_enter(self, event):
        """拖拽进入"""
        try:
            self.drop_frame.config(text="松开鼠标以处理")
            self.drop_label.config(text="📂\n正在处理...")
        except Exception as e:
            print(f"拖拽进入错误: {e}")
        
    def on_drag_leave(self, event):
        """拖拽离开"""
        try:
            self.drop_frame.config(text="拖拽文件或文件夹到此处")
            self.drop_label.config(text="📁\n拖拽PNG文件或文件夹到这里\n或")
        except Exception as e:
            print(f"拖拽离开错误: {e}")
        
    def on_drop(self, event):
        """处理拖拽放下事件"""
        try:
            # 获取拖拽的路径
            paths = self.root.tk.splitlist(event.data)
            
            if paths:
                path = paths[0].strip('{}')
                self.process_path(path)
                
            # 恢复界面
            self.on_drag_leave(None)
        except Exception as e:
            self.queue_log(f"处理拖拽错误: {str(e)}")
            self.queue_messagebox_error(f"处理拖拽失败:\n{str(e)}")
    
    def select_file(self):
        """选择单个文件"""
        try:
            file_paths = filedialog.askopenfilenames(
                title="选择PNG图片",
                filetypes=[("PNG 文件", "*.png *.PNG"), ("所有文件", "*.*")]
            )
            
            if file_paths:
                self.queue_log(f"选择了 {len(file_paths)} 个文件")
                
                # 在新线程中批量处理
                thread = threading.Thread(target=self._process_files_thread, args=(file_paths,))
                thread.daemon = True
                thread.start()
        except Exception as e:
            self.queue_log(f"选择文件错误: {str(e)}")
            self.queue_messagebox_error(f"选择文件失败:\n{str(e)}")
    
    def select_folder(self):
        """选择文件夹"""
        try:
            folder_path = filedialog.askdirectory(title="选择包含PNG图片的文件夹")
            if folder_path:
                self.queue_log(f"选择的文件夹: {folder_path}")
                
                # 在新线程中处理
                thread = threading.Thread(target=self._process_folder_thread, args=(folder_path,))
                thread.daemon = True
                thread.start()
        except Exception as e:
            self.queue_log(f"选择文件夹错误: {str(e)}")
            self.queue_messagebox_error(f"选择文件夹失败:\n{str(e)}")
    
    def process_path(self, path):
        """处理文件或文件夹路径"""
        try:
            path = Path(path)
            
            if not path.exists():
                self.queue_messagebox_error(f"路径不存在: {path}")
                return
            
            # 判断是文件还是文件夹
            if path.is_file():
                # 单个文件
                if path.suffix.lower() == '.png':
                    thread = threading.Thread(target=self._process_files_thread, args=([str(path)],))
                    thread.daemon = True
                    thread.start()
                else:
                    self.queue_messagebox_error(f"不支持的文件格式: {path.suffix}\n请选择PNG文件")
            elif path.is_dir():
                # 文件夹
                thread = threading.Thread(target=self._process_folder_thread, args=(str(path),))
                thread.daemon = True
                thread.start()
            else:
                self.queue_messagebox_error(f"未知的路径类型: {path}")
                
        except Exception as e:
            self.queue_log(f"处理路径错误: {str(e)}")
            self.queue_messagebox_error(f"处理路径失败:\n{str(e)}")
    
    def queue_log(self, message):
        """将日志消息放入队列"""
        print(message)  # 同时输出到控制台
        self.msg_queue.put({'type': 'log', 'text': message})
    
    def queue_progress(self, value):
        """将进度更新放入队列"""
        self.msg_queue.put({'type': 'progress', 'value': value})
    
    def queue_progress_max(self, value):
        """将进度最大值更新放入队列"""
        self.msg_queue.put({'type': 'progress_max', 'value': value})
    
    def queue_progress_label(self, text):
        """将进度标签更新放入队列"""
        self.msg_queue.put({'type': 'progress_label', 'text': text})
    
    def queue_messagebox_info(self, text):
        """将信息对话框放入队列"""
        self.msg_queue.put({'type': 'messagebox_info', 'text': text})
    
    def queue_messagebox_error(self, text):
        """将错误对话框放入队列"""
        self.msg_queue.put({'type': 'messagebox_error', 'text': text})
    
    def _process_files_thread(self, file_paths):
        """在后台线程中处理多个文件"""
        try:
            png_files = [Path(fp) for fp in file_paths if Path(fp).suffix.lower() == '.png']
            
            if not png_files:
                self.queue_messagebox_info("没有找到有效的PNG文件")
                return
            
            self.queue_log(f"\n{'='*60}")
            self.queue_log(f"开始处理 {len(png_files)} 个文件")
            
            # 确定输出文件夹（使用第一个文件的父目录）
            output_folder = png_files[0].parent / "jpg_output"
            output_folder.mkdir(exist_ok=True)
            self.queue_log(f"输出文件夹: {output_folder}")
            
            # 设置进度条
            self.queue_progress_max(len(png_files))
            self.queue_progress(0)
            
            success_count = 0
            error_count = 0
            
            for i, png_file in enumerate(png_files, 1):
                try:
                    # 转换图片
                    jpg_file = output_folder / (png_file.stem + ".jpg")
                    self.convert_png_to_jpg(png_file, jpg_file)
                    
                    success_count += 1
                    self.queue_log(f"✓ [{i}/{len(png_files)}] {png_file.name} -> {jpg_file.name}")
                    
                except Exception as e:
                    error_count += 1
                    self.queue_log(f"✗ [{i}/{len(png_files)}] {png_file.name} - {str(e)}")
                
                # 更新进度
                self.queue_progress(i)
                self.queue_progress_label(f"{i}/{len(png_files)}")
            
            # 完成
            self.queue_log(f"\n{'='*60}")
            self.queue_log(f"转换完成!")
            self.queue_log(f"成功: {success_count} 个")
            self.queue_log(f"失败: {error_count} 个")
            self.queue_log(f"输出位置: {output_folder}")
            self.queue_log(f"{'='*60}\n")
            
            # 显示完成消息
            self.queue_messagebox_info(
                f"转换完成!\n\n成功: {success_count} 个\n失败: {error_count} 个\n\n输出文件夹:\n{output_folder}"
            )
            
        except Exception as e:
            self.queue_log(f"处理过程中出错: {str(e)}")
            self.queue_messagebox_error(f"处理过程中出错:\n{str(e)}")
    
    def _process_folder_thread(self, folder_path):
        """在后台线程中处理文件夹"""
        try:
            folder_path = Path(folder_path)
            
            # 查找所有 PNG 文件
            png_files = list(folder_path.glob("*.png")) + list(folder_path.glob("*.PNG"))
            
            if not png_files:
                self.queue_messagebox_info(f"文件夹中没有找到PNG图片:\n{folder_path}")
                return
            
            self.queue_log(f"\n{'='*60}")
            self.queue_log(f"开始处理文件夹: {folder_path}")
            self.queue_log(f"找到 {len(png_files)} 个PNG文件")
            
            # 创建输出文件夹
            output_folder = folder_path / "jpg_output"
            output_folder.mkdir(exist_ok=True)
            self.queue_log(f"输出文件夹: {output_folder}")
            
            # 设置进度条
            self.queue_progress_max(len(png_files))
            self.queue_progress(0)
            
            success_count = 0
            error_count = 0
            
            for i, png_file in enumerate(png_files, 1):
                try:
                    # 转换图片
                    jpg_file = output_folder / (png_file.stem + ".jpg")
                    self.convert_png_to_jpg(png_file, jpg_file)
                    
                    success_count += 1
                    self.queue_log(f"✓ [{i}/{len(png_files)}] {png_file.name} -> {jpg_file.name}")
                    
                except Exception as e:
                    error_count += 1
                    self.queue_log(f"✗ [{i}/{len(png_files)}] {png_file.name} - {str(e)}")
                
                # 更新进度
                self.queue_progress(i)
                self.queue_progress_label(f"{i}/{len(png_files)}")
            
            # 完成
            self.queue_log(f"\n{'='*60}")
            self.queue_log(f"转换完成!")
            self.queue_log(f"成功: {success_count} 个")
            self.queue_log(f"失败: {error_count} 个")
            self.queue_log(f"输出位置: {output_folder}")
            self.queue_log(f"{'='*60}\n")
            
            # 显示完成消息
            self.queue_messagebox_info(
                f"转换完成!\n\n成功: {success_count} 个\n失败: {error_count} 个\n\n输出文件夹:\n{output_folder}"
            )
            
        except Exception as e:
            self.queue_log(f"处理过程中出错: {str(e)}")
            self.queue_messagebox_error(f"处理过程中出错:\n{str(e)}")
    
    def convert_png_to_jpg(self, png_path, jpg_path):
        """将PNG转换为JPG"""
        with Image.open(png_path) as img:
            # 如果图片有透明通道，需要转换为RGB
            if img.mode in ('RGBA', 'LA', 'P'):
                # 创建白色背景
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')
            
            # 保存为JPG
            img.save(jpg_path, 'JPEG', quality=95)
    
    def log(self, message):
        """添加日志（主线程调用）"""
        try:
            print(message)
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
        except Exception as e:
            print(f"日志记录错误: {e}")
    
    def clear_log(self):
        """清空日志"""
        try:
            self.log_text.delete(1.0, tk.END)
            self.progress['value'] = 0
            self.progress_label.config(text="就绪")
            self.log("✓ 日志已清空")
        except Exception as e:
            print(f"清空日志错误: {e}")


def main():
    """主函数"""
    try:
        print("=" * 60)
        print("PNG 转 JPG 转换器启动...")
        print("=" * 60)
        
        root = tk.Tk()
        app = PNGtoJPGConverter(root)
        
        print("程序初始化完成，进入主循环...")
        root.mainloop()
        
        print("程序正常退出")
    except Exception as e:
        print(f"程序运行错误: {e}")
        import traceback
        traceback.print_exc()
        input("按回车键退出...")


if __name__ == "__main__":
    main()
