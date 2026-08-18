# 日记写作程序

一个 Windows 桌面日记软件：写日记 → 自动套用公文排版 → 导出 Word。纯本地运行，无后端，数据保存在本机 SQLite。

## 功能

- 按日期管理日记（新建 / 查看 / 编辑 / 删除）
- 列表视图 + 日历视图（可翻月）
- 富文本：加粗 / 斜体 / 下划线 / 插入图片
- 停止输入 5 秒自动保存
- 标题 + 正文关键词搜索
- 统计：总篇数 / 连续写作天数 / 累计字数
- 一键导出 Word：标题黑体二号居中，正文宋体小四、两端对齐、首行缩进 2 字符、1.5 倍行距，`一、` / `1.` 小标题自动识别
- 标题默认填充「日期 + 城市 + 天气」（中央气象台），可手动设置城市

## 运行

需要 Windows 10/11，导出 Word 依赖已安装的 Microsoft Word（COM）。

```bash
python Python版/日记.pyw
```

依赖：Python 3.11（自带 tkinter）、pywin32、Pillow。

也可以按 [PROJECT_HANDOFF.md](PROJECT_HANDOFF.md) 中的命令用 PyInstaller 打包成单文件 exe。

## 目录结构

- `Python版/日记.pyw` — 唯一源码文件（约 1100 行）
- `Python版/app.ico` — 程序图标（二进制，未包含在仓库中，打包时可选用无 --icon 的命令）
- `PROJECT_HANDOFF.md` — 完整项目文档（目录结构、打包命令、关键技术坑）
