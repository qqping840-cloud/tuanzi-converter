<div align="center">
  <img src="assets/logo.png" alt="Tuanzi Converter" width="220" />

  <h1>团子转换器</h1>

  <p>
    把 <b>AI 的回答</b> 一键转成 <b>排版干净的 Word 文档</b><br/>
    不折腾格式、不写命令、不用在线网站
  </p>
</div>

---

## 这个工具是干嘛的？
团子转换器是一个 **极度傻瓜化** 的 Markdown → Word（.docx）桌面工具。

它解决的是一个非常具体、但被长期忽视的问题：

?? **如何把 AI（ChatGPT / Claude / DeepSeek 等）的回答，快速、干净、可控地变成 Word 文件？**

在线工具要么效果差、要么步骤复杂、要么收费。Pandoc 很强，但命令行太劝退。

**这个项目，就是把 Pandoc 魔改成“点一下就能用”的 GUI。**

---

## 下载与使用（推荐）
1. 打开本仓库的 **Releases** 页面
2. 下载 `Markdown转Word.exe`
3. 双击运行即可

> ? Release 版本已内置 Pandoc，无需安装 Python / Pandoc / 任何依赖

---

## 基本使用流程
1. 从 AI 工具中复制回答
2. 打开团子转换器
3. 粘贴 / 拖拽 Markdown 内容
4. 选择模板
5. 导出 `.docx`

---

## 日志位置
默认路径：
`C:\Users\Administrator\AppData\Local\TuanziConverter\logs`

---

## 开发者运行（可选）
```bash
python app_qt.py
```
