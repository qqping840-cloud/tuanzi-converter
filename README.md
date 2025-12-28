# 团子转换器 (Markdown -> Word)

基于 Pandoc 的 Markdown 转 Word GUI 工具，支持拖拽文件或粘贴文本导出 .docx。

## 功能
- 拖拽导入 `.md` / `.txt`
- 文本粘贴导入
- 字体模板、加粗预设、页码控制

## 运行

```bash
python app_qt.py
```

## Pandoc 依赖
- 方式一：系统已安装 Pandoc（在 PATH 中可用）
- 方式二：将 `pandoc.exe` 放入 `bin/` 目录

## 日志
默认写入：`%LocalAppData%\\TuanziConverter\\logs`
