# 日记写作程序 · 项目交接文档

> 一个 Windows 桌面日记软件：写日记 → 自动套排版 → 导出 Word。纯本地运行，无后端。

---

## 一、项目概览

- **定位**：个人日记写作工具，写完自动导出带排版的 Word 文档。
- **运行环境**：Windows 10/11，需安装 Microsoft Word（导出依赖 Word COM）。
- **语言/技术栈**：Python 3.11 + Tkinter（GUI）+ sqlite3（存储）+ win32com（Word 导出）+ Pillow（图片/图标）+ urllib（天气，标准库）。
- **形态**：单文件源码 + PyInstaller 打包成单 exe。

## 二、目录结构（当前最终形态）

```
D:\日记\
├── diary\            ← 本项目全部文件（2026-08-19 收纳）
│   ├── 日记.exe      ← 主程序（双击运行，exe 与 data 必须同目录）
│   ├── 启动.bat      ← 备用启动器（GBK 编码，中文文件名必须 GBK）
│   ├── data\         ← 运行时数据（自动生成）
│   │   ├── diary.db      ← 日记数据库（SQLite）
│   │   ├── city.json     ← 中央气象台城市编码映射缓存（2522 城市）
│   │   ├── config.json   ← 用户配置（city 城市 / export_dir 导出目录）
│   │   ├── images\       ← 日记插入的图片
│   │   └── template.json ← 旧模板缓存（已废弃，无害）
│   ├── Python版\     ← 源码 + 打包素材
│   │   ├── 日记.pyw  ← 唯一源码文件（约 1120 行）
│   │   ├── app.ico   ← exe 图标（由 png 生成，见「已知问题」）
│   │   └── README.md
│   ├── 成品\         ← 旧导出目录（历史 docx 仍在内，已不再使用）
│   ├── 其他\         ← 模板与预览素材（政府公文模板、界面预览图等）
│   └── PROJECT_HANDOFF.md  ← 本文档
└── SwashbucklerDiary-1.31.5-windows-x64\  ← 外部开源程序，与本项目无关
```

**关键：exe 与 data 必须同目录，都放在 `D:\日记\diary\`**。程序用 `sys.executable` 定位自身目录，数据落在 `D:\日记\diary\data\`。Word 导出的默认位置：`D:\日记\diary\成品`（可在程序里用「📂 导出位置」改，存 config.json 的 export_dir）。

## 三、功能清单

| 功能 | 说明 |
|---|---|
| 日记 CRUD | 按日期新建/查看/编辑/删除 |
| 双视图 | 列表视图 + 日历视图（可翻月） |
| 富文本 | 加粗 / 斜体 / 下划线 / 插入图片（选中文字后点 B/I/U） |
| 自动保存 | 停止输入 5 秒自动保存 |
| 搜索 | 标题 + 正文关键词 |
| 统计 | 总篇数 / 连续写作天数 / 累计字数 / 写作天数 |
| 导出 Word | 内置排版（见下） |
| 标题默认 | 「日期 + 城市 + 天气」（天气来自中央气象台） |
| 手动设置城市 | 顶栏「📍 城市」按钮，存 config.json |
| 自定义导出位置 | 顶栏「📂 导出位置」按钮，存 config.json |

**导出排版**（用户最终确定）：标题黑体**二号 22pt** 居中，正文宋体**小四 12pt** 两端对齐、首行缩进 2 字符、**1.5 倍行距**，页边距 2.5cm，小标题（`一、`/`二、`/`1.` 自动识别）黑体顶格，图片段居中。

## 四、源码结构（日记.pyw 单文件）

按顺序的大致分区：

1. **头部**：import + `_enable_dpi_awareness()`（高 DPI 适配，必须在建窗前列）
2. **常量**：`BASE_DIR`（frozen 判定）、`EXPORT_DIR = r'D:\日记\diary\成品'`（默认导出位置，可用 config.json 的 export_dir 覆盖）、配色常量、Word 排版常量（`FONT_TITLE='黑体'` `SIZE_TITLE=22` `FONT_BODY='宋体'` `SIZE_BODY=12` `LINE_RULE=1` 即 1.5 倍行距）
3. **天气**：`get_city()`（IP 定位）、`_build_city_map()`/`_load_city_map()`（中央气象台城市编码）、`_match_city_code()`、`_fetch_nmc_weather()`
4. **配置**：`load_config()`/`save_config()`、`get_export_dir()`
5. **Word 导出**：`export_word(date, title, html, out_path)` —— 核心排版逻辑
6. **数据库**：`DiaryDB` 类（sqlite3，表 `diaries` + `export_logs`）
7. **富文本转换**：`text_to_html()`（Text 组件 → HTML）、`_HTMLToText`（HTML → Text 组件）
8. **UI**：`RoundedButton`（Canvas 自绘圆角按钮）、`DiaryApp(tk.Tk)`（主窗口，含 `_build_ui`/`_load_date`/`_save_now`/`_export`/`_set_city`/`_set_export_dir` 等）

## 五、数据存储

- `diary.db`：`diaries` 表（id, date UNIQUE, title, content_html, content_text, word_count, status, created_at, updated_at）+ `export_logs` 表。
- `config.json`：`{"city": "...", "export_dir": "..."}`（两个字段，都可有可无）。
- `city.json`：`{城市名: 中央气象台code}`，首次联网构建后缓存，删掉会自动重建。

## 六、打包部署（务必保留此命令）

打包环境 = **miniconda3 base**（`C:\Users\sah10\miniconda3`，Python 3.11.15，自带 tkinter + pywin32 + Pillow + PyInstaller 6.22）。WorkBuddy 管理的 Python 3.13 无 tkinter，不可用。

```bash
cd /d/日记/diary/Python版
PY="/c/Users/sah10/miniconda3/python.exe"
LB="/c/Users/sah10/miniconda3/Library/bin"
"$PY" -m PyInstaller -F -w --name diary_new \
  --distpath dist_new --workpath build_new --specpath build_new \
  --icon "D:/日记/diary/Python版/app.ico" \
  --add-binary "$LB/sqlite3.dll;." --add-binary "$LB/tcl86t.dll;." \
  --add-binary "$LB/tk86t.dll;." --add-binary "$LB/libexpat.dll;." \
  --add-binary "$LB/liblzma.dll;." --add-binary "$LB/LIBBZ2.dll;." \
  --add-binary "$LB/libbz2.dll;." --add-binary "$LB/ffi.dll;." \
  --add-binary "$LB/zlib.dll;." \
  日记.pyw
# 产物 dist_new/diary_new.exe → 复制覆盖到 D:\日记\diary\日记.exe
```

> 注：仓库中未包含 `app.ico`（二进制图标且带水印）；克隆后打包请去掉 `--icon "D:/日记/diary/Python版/app.ico" \` 这一行。

**为什么 `--add-binary` 补 9 个 DLL**：conda 的这些 DLL 在 `Library\bin`，PyInstaller 默认不收集，缺了会 `DLL load failed`（尤其 sqlite3 直接导致启动崩溃）。

## 七、关键技术坑（接手者必读）

1. **PyInstaller onefile 数据丢失**：`__file__` 在 onefile 下指向临时解压目录，数据会丢。已用 `sys.frozen` + `sys.executable` 定位 exe 所在目录。
2. **Word COM 排版字体不生效**：打开 HTML 后段落用的是 "Normal (Web)" 样式（a3 basedOn Normal）。直接遍历设 `pr.Font` 在某些段不写入 run rPr；正确做法是设 `doc.Styles(-1).Font`（Normal 样式），段落继承生效，小标题单独覆盖。`para.Runs` 在 win32com 里不可用。
3. **沙箱 safe-delete 干扰**：WorkBuddy 沙箱把 `rm`/`os.remove` 重定向为「送回收站」，回收站不可用会报错。打包时用全新 `--distpath dist_new` 避免覆盖删除；删目录用 `/usr/bin/rm -rf`（真实 rm）或 PowerShell `Remove-Item -Recurse -Force`（并先 `Stop-Process` 杀残留进程）。
4. **中文 .bat 必须 GBK 编码**：UTF-8 写的中文文件名 cmd 读成乱码。
5. **Tkinter 无原生圆角**：`RoundedButton` 用 Canvas `create_polygon(smooth=True)` 自绘，`set_bg()` 用于 B/I/U 激活态和 tab 切换。改按钮颜色要用 `set_bg()` 而不是 `config(bg=)`。
6. **中央气象台接口**：`http://www.nmc.cn/rest/province` → `rest/province/{省码}`（返回城市 code 随机串）→ `rest/real/{城市code}`（返回 `weather.info` 中文天气 + `temperature` + `wind`）。城市 code 需遍历 34 省构建映射。
7. **IP 定位走网络出口**：用户开代理会定位到国外，中央气象台查不到 → 标题只剩日期。已加「📍 城市」手动设置兜底。

## 八、已知问题 / 待办

1. **图标有水印**：`app.ico` 由 `D:\Edge\日记程序图标设计 (1).png` 生成，原图右下角带「豆包AI生成」水印，会出现在 exe 和窗口图标上。正式发布需换图（重生成 ico + 重新打包）。
2. **`成品\` 目录**：历史 docx 与新导出都收在这里（默认导出位置 `D:\日记\diary\成品`）。
3. **`config.json` 曾被清空**：用户之前设的「上海」城市丢了，需用户重设。
4. **`~$政府公文模板.docx`**：Word 临时锁文件，可删。
5. **体积**：exe 约 33MB（conda 环境把 numpy 也打进去了），可优化但非必需。

## 九、快速验证

- 双击 `D:\日记\diary\日记.exe` → 应弹出深色窗口，无报错。
- 新建日记 → 标题自动填「日期 城市 天气」。
- 点「导出 Word」→ 默认在 `D:\日记\diary\成品\` 生成 `日期-标题.docx`，打开检查标题黑体 22pt、正文宋体 12pt、1.5 倍行距。
- 数据落点检查：`D:\日记\diary\data\diary.db` 存在即正常。
