# Design Document

## Overview
以最小侵入方式在主窗口增加入口，并用独立模块完成周报读取、校验、合并和子窗口交互。

## Architecture
- `png_to_jpg_converter.py`：保持图片流程，仅增加入口按钮。
- `weekly_report.py`：纯合并服务、格式读取器和独立 Tk 子窗口。
- `requirements.txt`：增加 openpyxl 与 xlrd。

## Components and Interfaces
### Reader
`read_rows(path)` 将 XLSX、XLS、CSV 转为二维列表；CSV 尝试 UTF-8 BOM、UTF-8、GB18030。

### Header Detector
`detect_header(rows, fields, file_name)` 在前100行匹配动态字段。表头映射保留同名字段的全部物理列位置。

### Merge Service
`merge_reports(main_path, source_paths, base_fields, metrics, output_path, log, progress)` 保留主表、扩展角色指标列并纵向追加数据。

### Weekly Report Window
独立维护文件路径、字段文本、消息队列、进度和日志。后台线程只写队列，Tk 主线程更新界面。

## Data Models
- `TableData`：文件路径、二维行、表头索引、字段到列索引列表。
- `MergeResult`：输出路径和各来源追加行数。
- `source_paths`：角色名到输入路径的映射。

## Correctness Properties
### Property 1: 主表保持不变
**Validates: Requirements 3.1**
主表原有单元格在输出中保持原位置和值。

### Property 2: 重复指标选择
**Validates: Requirements 2.5**
重复指标始终读取最大列索引，即最后一个同名列。

### Property 3: 来源隔离
**Validates: Requirements 3.2, 3.4**
每条来源行只写入其角色对应的指标列。

### Property 4: 行数守恒
**Validates: Requirements 3.3, 3.5**
输出追加行数等于所有有效来源数据行之和。

## Error Handling
缺少文件、格式不支持、编码失败、字段缺失、输出覆盖输入时抛出可读错误；界面恢复合并按钮并显示错误。

## Testing Strategy
执行语法检查和最小合并验证：XLSX 主文件、CSV 追加文件、多行表头、重复财务存列、主行保留及追加行位置。