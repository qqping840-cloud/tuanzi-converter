<div align="center">
  <img src="assets/logo.png" alt="Tuanzi Converter" width="220" />

  <h1>团子转换器</h1>

  <p>
    把 <b>AI 的回答</b> 一键转成 <b>排版干净的 Word 文档</b><br/>
    省掉格式整理时间，不用在线网站
  </p>
</div>

---

## 这工具解决什么问题？（讨论加QQ群：982849511）
你的 AI 回答通常是 Markdown（标题/列表/代码/引用）。直接粘到 Word 经常格式全乱，整理很费时间。

团子转换器的核心目的就是：

**把 AI 的回答“原样排版”导出为 Word（.docx），让你直接交付。**

它基于 Pandoc 做了魔改，把命令行能力封装成傻瓜 GUI：复制/粘贴 → 选模板 → 导出 Word。

---

## 为什么不用在线网站？
- 效果不全：标题/列表/代码块经常丢格式
- 操作麻烦：步骤多、选项多
- 限制/收费：次数、大小、高清导出
- 内容上传：隐私不安心

本工具完全本地运行：免费、离线、不上传。

---

## 下载与使用（推荐）
1. 打开本仓库的 **Releases**
2. 下载 `tuanzi-converter.exe`
3. 双击运行 → 粘贴 AI 回答 → 导出 `.docx`

> ? Release 版本已内置 Pandoc，无需安装 Python / Pandoc

---

## 日志位置
`C:\Users\Administrator\AppData\Local\TuanziConverter\logs`

---

## 开发者运行（可选）
```bash
python app_qt.py
```

