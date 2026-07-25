# Implementation Plan: Weekly Report Merge

## Overview
在不修改图片转换行为的前提下，实现可配置的周报合并服务和独立子窗口。

## Tasks
- [x] 1. 新增 XLSX、XLS、CSV 读取与表头识别
- [x] 2. 实现动态基础字段和追加指标配置
- [x] 3. 实现重复同名指标取最后一列
- [x] 4. 实现主表保留、角色列扩展和数据纵向追加
- [x] 5. 新增独立周报合并子窗口和线程安全消息队列
- [x] 6. 在现有图片工具增加周报入口按钮
- [x] 7. 增加 Excel 依赖
- [x] 8. 执行语法检查和最小合并验证

## Task Dependency Graph
```json
{
  "waves": [
    {"wave": 1, "tasks": ["1", "2", "3"]},
    {"wave": 2, "tasks": ["4"]},
    {"wave": 3, "tasks": ["5", "6", "7"]},
    {"wave": 4, "tasks": ["8"]}
  ]
}
```

## Notes
输出统一为新 XLSX；主文件为 XLSX 时直接在其工作簿副本上追加，以尽量保留样式。