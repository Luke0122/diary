# -*- coding: utf-8 -*-
"""
日记写作程序（Python + Tkinter + Word COM）— 纯黑主题
双击运行；写日记 → 按 D:\\格式.doc 模板排版 → 另存 Word 到 D:\\日记\\成品
"""

import os
import re
import sys
import json
import sqlite3
import datetime
import tempfile
import ctypes
import threading
import queue
import urllib.request
import urllib.parse

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont
from html.parser import HTMLParser

try:
    import win32com.client
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except Exception:
    HAS_PIL = False

# ---------------- 高 DPI 适配（解决界面模糊/分辨率低） ----------------
def _enable_dpi_awareness():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)   # Per-Monitor V2
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)  # System DPI Aware
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

_enable_dpi_awareness()

# ---------------- 路径与常量 ----------------
# PyInstaller onefile 下 __file__ 指向临时解压目录，改用 exe 所在目录，避免数据丢失
BASE_DIR     = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) \
               else os.path.dirname(os.path.abspath(__file__))
DATA_DIR     = os.path.join(BASE_DIR, 'data')
DB_PATH      = os.path.join(DATA_DIR, 'diary.db')
IMG_DIR      = os.path.join(DATA_DIR, 'images')
EXPORT_DIR   = r'D:\日记\diary\成品'

EDITOR_FONT = 'Microsoft YaHei'
EDITOR_SIZE = 13

# ---------------- 柔和深色配色（现代记事本风格，低饱和蓝灰） ----------------
C_BG           = '#1e2024'   # 主背景（深灰，不死黑）
C_PANEL        = '#26292f'   # 左栏面板
C_CARD         = '#2c3037'   # 卡片
C_TEXT         = '#d4d7dc'   # 主文字（浅灰，非纯白）
C_DIM          = '#8a8f98'   # 次要文字
C_BORDER       = '#33373f'   # 边框/分割线
C_BTN          = '#2e3239'   # 普通按钮
C_BTN_ACT      = '#3d4654'   # 按钮激活态（B/I/U 高亮）
C_BTN_HOVER    = '#3a3f48'   # 普通按钮 hover
C_ACCENT       = '#5d7fa3'   # 强调色（低饱和蓝灰）
C_ACCENT_HOVER = '#6f92b8'   # 强调色 hover
C_DANGER       = '#a9605a'   # 删除红（低饱和）
C_DANGER_HOVER = '#bd726c'   # 删除红 hover
C_OK           = '#6a9b78'   # 成功绿（柔和）
C_SELECT       = '#3a4556'   # 选中高亮（柔和蓝灰）
C_CURSOR       = '#c7cbd1'   # 光标

# ---------------- Word 排版常量（标题黑体小二 + 正文宋体小四 + 1.5 倍行距） ----------------
FONT_TITLE   = '黑体'
FONT_BODY    = '宋体'
FONT_HEADING = '黑体'
SIZE_TITLE   = 22          # 二号
SIZE_BODY    = 12          # 小四
LINE_RULE    = 1           # wdLineSpace1pt5 = 1.5 倍行距
MARGIN_TOP, MARGIN_BOTTOM, MARGIN_LEFT, MARGIN_RIGHT = 2.5, 2.5, 2.5, 2.5  # 厘米（常规文档）
_HEADING_RE = re.compile(r'^(?:[一二三四五六七八九十]+、|\d+[、.．]|（[一二三四五六七八九十]+）)')

def cm_to_points(cm):
    return round(cm * 28.35, 1)

def _is_heading(text):
    return bool(_HEADING_RE.match((text or '').strip()))

for _d in (DATA_DIR, IMG_DIR, EXPORT_DIR):
    try:
        os.makedirs(_d, exist_ok=True)
    except Exception:
        pass


# ---------------- 工具 ----------------
def today_str():
    return datetime.date.today().strftime('%Y-%m-%d')


def now_str():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')


def safe_filename(title):
    t = re.sub(r'[\\/:*?"<>|]', '_', title or '').strip()
    return (t or '无标题')[:60]


def strip_html(html):
    return re.sub(r'<[^>]+>', ' ', html or '')


def count_words(text):
    return len(re.sub(r'\s+', '', text or ''))


# ---------------- 地址与天气（IP定位 + 中央气象台，失败降级） ----------------
CITY_JSON = os.path.join(DATA_DIR, 'city.json')
CONFIG_JSON = os.path.join(DATA_DIR, 'config.json')

def load_config():
    if os.path.exists(CONFIG_JSON):
        try:
            with open(CONFIG_JSON, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(cfg):
    try:
        with open(CONFIG_JSON, 'w', encoding='utf-8') as f:
            json.dump(cfg, f, ensure_ascii=False)
    except Exception:
        pass

def get_export_dir():
    """导出目录：优先用 config.json 里用户设置的，否则用默认 EXPORT_DIR。"""
    d = (load_config().get('export_dir') or '').strip()
    return d if d else EXPORT_DIR


def _http_get(url, timeout=5):
    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode('utf-8', 'ignore')

def get_city():
    """通过 IP 定位获取城市（中文），失败返回空串。"""
    try:
        data = json.loads(_http_get('http://ip-api.com/json/?lang=zh-CN&fields=status,city,regionName', 4))
        if data.get('status') == 'success':
            return (data.get('city') or '').strip(), (data.get('regionName') or '').strip()
    except Exception:
        pass
    return '', ''

def _build_city_map():
    """遍历中央气象台 34 省城市列表，构建 {城市名: code} 映射并缓存。"""
    city_map = {}
    try:
        provinces = json.loads(_http_get('http://www.nmc.cn/rest/province', 8))
        for p in provinces:
            code = p.get('code')
            if not code:
                continue
            try:
                cities = json.loads(_http_get('http://www.nmc.cn/rest/province/%s' % code, 8))
                for c in cities:
                    name, ccode = c.get('city'), c.get('code')
                    if name and ccode:
                        city_map[name] = ccode
            except Exception:
                continue
    except Exception:
        pass
    if city_map:
        try:
            with open(CITY_JSON, 'w', encoding='utf-8') as f:
                json.dump(city_map, f, ensure_ascii=False)
        except Exception:
            pass
    return city_map

def _load_city_map():
    if os.path.exists(CITY_JSON):
        try:
            with open(CITY_JSON, 'r', encoding='utf-8') as f:
                m = json.load(f)
            if m:
                return m
        except Exception:
            pass
    return _build_city_map()

def _match_city_code(city_map, city):
    """城市名匹配 code，尝试去'市/省/自治区'等后缀。"""
    if not city:
        return None
    for name in (city, city.rstrip('市'), city.rstrip('省'), city.rstrip('地区')):
        if name in city_map:
            return city_map[name], name
    # 部分匹配：定位名可能含'区/县'，取前2字试一次
    if len(city) >= 3:
        head = city[:2]
        for k in city_map:
            if k.startswith(head):
                return city_map[k], k
    return None, city

def _fetch_nmc_weather(code):
    """中央气象台实况：返回 '多云 30°C 西南风'；失败返回空串。"""
    try:
        data = json.loads(_http_get('http://www.nmc.cn/rest/real/%s' % code, 6))
        w = data.get('weather') or {}
        info = (w.get('info') or '').strip()
        temp = w.get('temperature')
        wind = ((data.get('wind') or {}).get('power') or '').strip()
        parts = [info] if info else []
        if isinstance(temp, (int, float)):
            parts.append('%d°C' % round(temp))
        if wind and wind != '微风':
            parts.append(wind)
        return ' '.join(parts).strip()
    except Exception:
        return ''


# ---------------- Word COM 导出（内置公文排版） ----------------
def export_word(date, title, html, out_path):
    """内置公文排版：方正小标宋标题 + 仿宋正文 + 黑体小标题 + 固定行距 + 公文页边距。"""
    if not HAS_WIN32:
        return False, '未安装 pywin32，无法调用 Word'
    tmp = os.path.join(tempfile.gettempdir(), 'diary_%d.htm' % os.getpid())
    full = ('<!DOCTYPE html><html><head><meta charset="utf-8"></head><body>'
            + (html or '') + '</body></html>')
    word = None
    try:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.ScreenUpdating = False
        word.DisplayAlerts = 0
        with open(tmp, 'w', encoding='utf-8') as f:
            f.write(full)
        doc = word.Documents.Open(tmp, False, True)

        # 页边距（公文标准：上3.7 下3.5 左2.8 右2.6 厘米）
        ps = doc.PageSetup
        ps.TopMargin = cm_to_points(MARGIN_TOP)
        ps.BottomMargin = cm_to_points(MARGIN_BOTTOM)
        ps.LeftMargin = cm_to_points(MARGIN_LEFT)
        ps.RightMargin = cm_to_points(MARGIN_RIGHT)

        # 标题段：若标题已含日期（默认标题「日期 城市 天气」）则不重复拼接
        if title and title.startswith(date):
            title_text = title
        else:
            title_text = date + (('  ' + title) if title else '')
        rng = doc.Range(0, 0)
        rng.Text = title_text
        rng.Font.NameFarEast = FONT_TITLE
        rng.Font.NameAscii = 'Times New Roman'
        rng.Font.Size = SIZE_TITLE
        rng.Font.Bold = True
        rng.ParagraphFormat.Alignment = 1          # 居中
        rng.ParagraphFormat.LineSpacingRule = LINE_RULE    # 1.5 倍行距
        rng.ParagraphFormat.SpaceAfter = 6
        doc.Range(rng.End, rng.End).InsertParagraphAfter()

        # 正文默认：先设 Normal 样式为宋体小四（保证所有段继承正确字体字号）
        try:
            normal = doc.Styles(-1)                 # wdStyleNormal
            normal.Font.NameFarEast = FONT_BODY
            normal.Font.NameAscii = 'Times New Roman'
            normal.Font.Size = SIZE_BODY
        except Exception:
            pass

        # 遍历正文段落：行距 1.5 倍；小标题覆盖为黑体（保留 12pt）；正文继承 Normal
        for i in range(2, doc.Paragraphs.Count + 1):
            para = doc.Paragraphs(i)
            pr = para.Range
            txt = pr.Text.strip()
            pf = pr.ParagraphFormat
            pf.LineSpacingRule = LINE_RULE
            if not txt:
                continue
            if pr.InlineShapes.Count > 0:           # 图片段居中，不缩进
                pf.Alignment = 1
                continue
            is_h = _is_heading(txt)
            if is_h:                                    # 小标题：覆盖为黑体
                pr.Font.NameFarEast = FONT_HEADING
                pf.Alignment = 0                       # 左对齐
            else:                                       # 正文：继承 Normal（宋体 12pt），两端对齐 + 首行缩进 2 字符
                pf.Alignment = 3
                try:
                    pf.CharacterUnitFirstLineIndent = 2
                except Exception:
                    pf.FirstLineIndent = cm_to_points(1.12)

        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        doc.SaveAs2(out_path, 16)
        doc.Close(False)
        return True, out_path
    except Exception as e:
        return False, str(e)
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        try:
            os.remove(tmp)
        except Exception:
            pass


# ---------------- 数据库 ----------------
class DiaryDB:
    def __init__(self, path=DB_PATH):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript('''
        CREATE TABLE IF NOT EXISTS diaries (
            id INTEGER PRIMARY KEY,
            date TEXT UNIQUE NOT NULL,
            title TEXT NOT NULL DEFAULT '',
            content_html TEXT NOT NULL DEFAULT '',
            content_text TEXT NOT NULL DEFAULT '',
            word_count INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'draft',
            created_at TEXT, updated_at TEXT
        );
        CREATE TABLE IF NOT EXISTS export_logs (
            id INTEGER PRIMARY KEY, diary_id INTEGER, filename TEXT, exported_at TEXT
        );
        ''')

    def list(self):
        return [dict(r) for r in self.conn.execute(
            'SELECT id, date, title, status, word_count FROM diaries ORDER BY date DESC')]

    def get(self, date):
        r = self.conn.execute('SELECT * FROM diaries WHERE date=?', (date,)).fetchone()
        return dict(r) if r else None

    def upsert(self, date, title='', content_html='', status=None):
        text = strip_html(content_html)
        wc = count_words(text)
        now = now_str()
        ex = self.get(date)
        if ex:
            self.conn.execute(
                'UPDATE diaries SET title=?, content_html=?, content_text=?, word_count=?, status=?, updated_at=? WHERE date=?',
                (title, content_html, text, wc, status or ex['status'], now, date))
        else:
            self.conn.execute(
                'INSERT INTO diaries (date,title,content_html,content_text,word_count,status,created_at,updated_at) '
                'VALUES (?,?,?,?,?,?,?,?)',
                (date, title, content_html, text, wc, status or 'draft', now, now))
        self.conn.commit()
        return self.get(date)

    def delete(self, date):
        self.conn.execute('DELETE FROM diaries WHERE date=?', (date,))
        self.conn.commit()

    def search(self, keyword='', date_from=None, date_to=None):
        sql = 'SELECT id,date,title,status,word_count FROM diaries WHERE 1=1'
        args = []
        if keyword:
            sql += ' AND (title LIKE ? OR content_text LIKE ?)'
            kw = '%' + keyword + '%'
            args += [kw, kw]
        if date_from:
            sql += ' AND date >= ?'
            args.append(date_from)
        if date_to:
            sql += ' AND date <= ?'
            args.append(date_to)
        sql += ' ORDER BY date DESC LIMIT 200'
        return [dict(r) for r in self.conn.execute(sql, args)]

    def calendar_dates(self, year, month):
        prefix = '%04d-%02d-' % (year, month)
        return {r['date'] for r in self.conn.execute(
            'SELECT DISTINCT date FROM diaries WHERE date LIKE ?', (prefix + '%',))}

    def stats(self):
        row = self.conn.execute(
            'SELECT COUNT(*) AS total, COALESCE(SUM(word_count),0) AS words FROM diaries').fetchone()
        dates = [r['date'] for r in self.conn.execute('SELECT DISTINCT date FROM diaries ORDER BY date')]
        return {'total': row['total'], 'words': row['words'],
                'streak': self._streak(dates), 'days': len(dates)}

    @staticmethod
    def _streak(dates):
        if not dates:
            return 0
        s = set(dates)
        d = datetime.date.today()
        if d.isoformat() not in s:
            d -= datetime.timedelta(days=1)
            if d.isoformat() not in s:
                return 0
        streak = 0
        while d.isoformat() in s:
            streak += 1
            d -= datetime.timedelta(days=1)
        return streak

    def log_export(self, date, filename):
        r = self.get(date)
        if r:
            self.conn.execute('INSERT INTO export_logs (diary_id,filename,exported_at) VALUES (?,?,?)',
                              (r['id'], filename, now_str()))
            self.conn.commit()

    def close(self):
        self.conn.close()


# ---------------- 富文本：Text -> HTML ----------------
def text_to_html(widget, image_paths):
    from html import escape
    state = {'b': False, 'i': False, 'u': False}
    cur = {'b': False, 'i': False, 'u': False}
    out = ['<p>']

    def flush():
        if state['b'] and not cur['b']:
            out.append('<strong>'); cur['b'] = True
        elif not state['b'] and cur['b']:
            out.append('</strong>'); cur['b'] = False
        if state['i'] and not cur['i']:
            out.append('<em>'); cur['i'] = True
        elif not state['i'] and cur['i']:
            out.append('</em>'); cur['i'] = False
        if state['u'] and not cur['u']:
            out.append('<u>'); cur['u'] = True
        elif not state['u'] and cur['u']:
            out.append('</u>'); cur['u'] = False

    for key, val, _idx in widget.dump('1.0', 'end-1c', all=False):
        if key == 'tagon':
            if 'b' in val:
                state['b'] = True
            if 'i' in val:
                state['i'] = True
            if 'u' in val:
                state['u'] = True
        elif key == 'tagoff':
            if 'b' in val:
                state['b'] = False
            if 'i' in val:
                state['i'] = False
            if 'u' in val:
                state['u'] = False
        elif key == 'text':
            for ch in val:
                if ch == '\n':
                    flush()
                    out.append('</p><p>')
                else:
                    flush()
                    out.append(escape(ch))
        elif key == 'image':
            path = image_paths.get(val)
            if path:
                flush()
                out.append('<img src="file:///%s" />' % path.replace('\\', '/').lstrip('/'))

    out.append('</p>')
    s = ''.join(out)
    s = re.sub(r'<p>\s*</p>', '<p>&nbsp;</p>', s)
    return s


class _HTMLToText(HTMLParser):
    def __init__(self, widget, image_paths, insert_image):
        super().__init__()
        self.w = widget
        self.image_paths = image_paths
        self.insert_image = insert_image
        self.bits = {'b': False, 'i': False, 'u': False}

    def _combo(self):
        return ('b' if self.bits['b'] else '') + ('i' if self.bits['i'] else '') + ('u' if self.bits['u'] else '')

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in ('strong', 'b'):
            self.bits['b'] = True
        elif tag in ('em', 'i'):
            self.bits['i'] = True
        elif tag == 'u':
            self.bits['u'] = True
        elif tag in ('h1', 'h2', 'h3'):
            self.bits['b'] = True
            self.w.insert('end', '\n')
        elif tag in ('p', 'div', 'br'):
            self.w.insert('end', '\n')
        elif tag == 'li':
            self.w.insert('end', '\n• ')
        elif tag == 'img':
            src = a.get('src', '')
            if src.startswith('file://'):
                src = src[7:]
            if src and os.path.exists(src):
                self.insert_image(src)

    def handle_endtag(self, tag):
        if tag in ('strong', 'b', 'h1', 'h2', 'h3'):
            self.bits['b'] = False
        elif tag in ('em', 'i'):
            self.bits['i'] = False
        elif tag == 'u':
            self.bits['u'] = False
        elif tag in ('p', 'div'):
            self.w.insert('end', '\n')

    def handle_data(self, data):
        if data == '':
            return
        combo = self._combo()
        if combo:
            self.w.insert('end', data, (combo,))
        else:
            self.w.insert('end', data)


# ---------------- 圆角按钮（Canvas 自绘，支持 hover / 动态换色） ----------------
class RoundedButton(tk.Canvas):
    def __init__(self, parent, text, command, bg=C_BTN, fg=C_TEXT, hover=None,
                 font=('Microsoft YaHei', 10), radius=8, padx=14, pady=7):
        self._text = text
        self._command = command
        self._bg = bg
        self._fg = fg
        self._hover = hover if hover is not None else bg
        self._radius = radius
        self._font = font
        f = tkfont.Font(font=font)
        w = f.measure(text) + padx * 2
        h = f.metrics('linespace') + pady * 2
        super().__init__(parent, width=w, height=h, bg=parent.cget('bg'),
                         highlightthickness=0, bd=0, cursor='hand2')
        self._draw(bg)
        self.bind('<Enter>', lambda e: self._draw(self._hover))
        self.bind('<Leave>', lambda e: self._draw(self._bg))
        self.bind('<Button-1>', self._on_click)

    def _on_click(self, e):
        if self._command:
            self._command()

    def set_bg(self, bg, hover=None):
        """动态换色（用于 B/I/U 激活态、tab 切换）。"""
        self._bg = bg
        self._hover = hover if hover is not None else bg
        self._draw(bg)

    def _draw(self, bg):
        self.delete('all')
        w = int(self['width'])
        h = int(self['height'])
        self._round_rect(1, 1, w - 1, h - 1, self._radius, fill=bg, outline='')
        self.create_text(w // 2, h // 2, text=self._text, fill=self._fg, font=self._font)

    def _round_rect(self, x1, y1, x2, y2, r, **kw):
        pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
               x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
        return self.create_polygon(pts, smooth=True, **kw)


# ---------------- 主窗口 ----------------
class DiaryApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('日记')
        self.geometry('1180x760')
        self.minsize(920, 600)
        self.configure(bg=C_BG)

        self.db = DiaryDB()
        self.current_date = None
        self.image_paths = {}
        self._image_refs = []
        self._save_job = None
        self._search_on = False
        self._cal = (datetime.date.today().year, datetime.date.today().month)
        self._addr_weather = ''          # "城市 天气"（如 "北京 晴 28°C"）
        self._weather_q = queue.Queue()
        threading.Thread(target=self._fetch_weather_worker, daemon=True).start()
        self.after(300, self._poll_weather)

        self._build_fonts()
        self._build_ui()
        if HAS_PIL:
            try:
                _icon = Image.open(r'D:\Edge\日记程序图标设计 (1).png').resize((64, 64), Image.LANCZOS)
                self._app_icon = ImageTk.PhotoImage(_icon)
                self.iconphoto(False, self._app_icon)
            except Exception:
                self._app_icon = None
        self._refresh_list()
        self._refresh_stats()
        self.protocol('WM_DELETE_WINDOW', self._on_close)

    # ---------- 天气（后台获取地址 + 中央气象台天气） ----------
    def _fetch_weather_worker(self):
        try:
            manual_city = (load_config().get('city') or '').strip()
            if manual_city:
                city = manual_city
            else:
                city, _region = get_city()
            if not city:
                self._weather_q.put('')
                return
            city_map = _load_city_map()
            code, city_used = _match_city_code(city_map, city)
            if not code:
                self._weather_q.put(city_used)  # 只有城市，无天气
                return
            weather = _fetch_nmc_weather(code)
            self._weather_q.put(('%s %s' % (city_used, weather)).strip() if weather else city_used)
        except Exception:
            self._weather_q.put('')

    def _set_city(self):
        """手动设置城市，优先于 IP 定位。"""
        cur = load_config().get('city') or ''
        city = simpledialog.askstring('设置城市', '输入城市名（如：北京、上海、广州）\n留空则恢复 IP 自动定位：',
                                      initialvalue=cur, parent=self)
        if city is None:
            return
        city = city.strip()
        cfg = load_config()
        cfg['city'] = city
        save_config(cfg)
        self._addr_weather = ''
        threading.Thread(target=self._fetch_weather_worker, daemon=True).start()
        self.after(300, self._poll_weather)
        self.status_lbl.config(text='已设置城市：%s（正在获取天气…）' % (city or '自动定位'))

    def _poll_weather(self):
        try:
            info = self._weather_q.get_nowait()
            self._addr_weather = info
            # 若当前标题还是默认纯日期，补全为「日期 城市 天气」
            if self.current_date and info:
                cur = self.title_entry.get().strip()
                if cur == self.current_date:
                    self.title_entry.delete(0, 'end')
                    self.title_entry.insert(0, '%s  %s' % (self.current_date, info))
        except queue.Empty:
            self.after(300, self._poll_weather)

    def _default_title(self, date):
        if self._addr_weather:
            return '%s  %s' % (date, self._addr_weather)
        return date

    # ---------- 字体与格式标签 ----------
    def _build_fonts(self):
        self.f_normal = tkfont.Font(family=EDITOR_FONT, size=EDITOR_SIZE)
        self.f_b = tkfont.Font(family=EDITOR_FONT, size=EDITOR_SIZE, weight='bold')
        self.f_i = tkfont.Font(family=EDITOR_FONT, size=EDITOR_SIZE, slant='italic')
        self.f_bi = tkfont.Font(family=EDITOR_FONT, size=EDITOR_SIZE, weight='bold', slant='italic')

    def _init_text_tags(self, text):
        text.tag_configure('b', font=self.f_b, foreground=C_TEXT)
        text.tag_configure('i', font=self.f_i, foreground=C_TEXT)
        text.tag_configure('bi', font=self.f_bi, foreground=C_TEXT)
        text.tag_configure('u', underline=True, foreground=C_TEXT)
        text.tag_configure('bu', font=self.f_b, underline=True, foreground=C_TEXT)
        text.tag_configure('iu', font=self.f_i, underline=True, foreground=C_TEXT)
        text.tag_configure('biu', font=self.f_bi, underline=True, foreground=C_TEXT)

    # ---------- UI ----------
    def _btn(self, parent, text, command, accent=False, danger=False, small=False):
        if accent:
            bg, hover, fg = C_ACCENT, C_ACCENT_HOVER, '#eef2f6'
        elif danger:
            bg, hover, fg = C_DANGER, C_DANGER_HOVER, '#f3e9e7'
        else:
            bg, hover, fg = C_BTN, C_BTN_HOVER, C_TEXT
        font = ('Microsoft YaHei', 9 if small else 10)
        return RoundedButton(parent, text, command, bg=bg, fg=fg, hover=hover,
                             font=font, radius=8, padx=(10 if small else 16), pady=(4 if small else 7))

    def _build_ui(self):
        # 顶栏
        top = tk.Frame(self, bg=C_PANEL, height=58)
        top.pack(side='top', fill='x')
        top.pack_propagate(False)
        tk.Label(top, text='📔 日记', bg=C_PANEL, fg=C_TEXT,
                 font=('Microsoft YaHei', 16, 'bold')).pack(side='left', padx=(20, 8), pady=10)
        tk.Frame(top, bg=C_BORDER, width=1).pack(side='left', fill='y', padx=6, pady=14)
        self._btn(top, '＋ 新建', self._new_diary, accent=True).pack(side='left', padx=4, pady=12)
        self._btn(top, '打开目录', self._open_export_dir).pack(side='left', padx=4)
        self._btn(top, '📂 导出位置', self._set_export_dir).pack(side='left', padx=4)
        self._btn(top, '📍 城市', self._set_city).pack(side='left', padx=4)
        self.status_lbl = tk.Label(top, text='', bg=C_PANEL, fg=C_DIM, font=('Microsoft YaHei', 9))
        self.status_lbl.pack(side='left', padx=14)
        self._btn(top, '导出 Word', self._export, accent=True).pack(side='right', padx=4, pady=12)
        self._btn(top, '删除', self._delete, danger=True).pack(side='right', padx=(0, 8))

        # 顶栏与主体之间的分隔线
        tk.Frame(self, bg=C_BORDER, height=1).pack(side='top', fill='x')

        # 主体：左栏 + 右栏
        body = tk.Frame(self, bg=C_BG)
        body.pack(side='top', fill='both', expand=True)

        # 左栏（右侧加分割线）
        left = tk.Frame(body, bg=C_PANEL, width=272)
        left.pack(side='left', fill='y')
        left.pack_propagate(False)
        tk.Frame(body, bg=C_BORDER, width=1).pack(side='left', fill='y')

        # 视图切换（圆角 tab）
        tabbar = tk.Frame(left, bg=C_PANEL)
        tabbar.pack(side='top', fill='x', padx=10, pady=10)
        self.tab_list_btn = RoundedButton(tabbar, '列表', lambda: self._switch_view('list'),
                                          bg=C_ACCENT, fg='#eef2f6', hover=C_ACCENT_HOVER,
                                          font=('Microsoft YaHei', 10), padx=16, pady=5)
        self.tab_list_btn.pack(side='left')
        self.tab_cal_btn = RoundedButton(tabbar, '日历', lambda: self._switch_view('cal'),
                                         bg=C_BTN, fg=C_TEXT, hover=C_BTN_HOVER,
                                         font=('Microsoft YaHei', 10), padx=16, pady=5)
        self.tab_cal_btn.pack(side='left', padx=6)

        # 列表视图
        self.list_frame = tk.Frame(left, bg=C_PANEL)
        self.list_frame.pack(side='top', fill='both', expand=True)
        self.listbox = tk.Listbox(self.list_frame, bg=C_CARD, fg=C_TEXT, selectbackground=C_SELECT,
                                  selectforeground=C_TEXT, relief='flat', bd=0, highlightthickness=0,
                                  activestyle='none', font=('Microsoft YaHei', 10))
        self.listbox.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=(0, 8))
        sb = tk.Scrollbar(self.list_frame, command=self.listbox.yview, bg=C_DIM, troughcolor=C_PANEL,
                          activebackground=C_BTN_HOVER, bd=0, relief='flat', width=10)
        sb.pack(side='right', fill='y', pady=(0, 8))
        self.listbox.config(yscrollcommand=sb.set)
        self.listbox.bind('<<ListboxSelect>>', self._on_list_select)

        # 日历视图
        self.cal_frame = tk.Frame(left, bg=C_PANEL)
        self._build_calendar()

        # 搜索
        sframe = tk.Frame(left, bg=C_PANEL)
        sframe.pack(side='top', fill='x', padx=10, pady=(0, 8))
        self.search_var = tk.StringVar()
        ent = tk.Entry(sframe, textvariable=self.search_var, bg=C_CARD, fg=C_TEXT,
                       insertbackground=C_CURSOR, relief='flat', bd=0, font=('Microsoft YaHei', 10),
                       highlightthickness=1, highlightbackground=C_BORDER, highlightcolor=C_ACCENT)
        ent.pack(side='left', fill='x', expand=True, ipady=5)
        ent.bind('<Return>', lambda e: self._do_search())
        self._btn(sframe, '搜', self._do_search, small=True).pack(side='left', padx=(6, 0))

        # 统计（卡片）
        stat_card = tk.Frame(left, bg=C_CARD, highlightbackground=C_BORDER, highlightthickness=1)
        stat_card.pack(side='top', fill='x', padx=12, pady=(0, 14))
        self.stats_lbl = tk.Label(stat_card, text='', bg=C_CARD, fg=C_DIM, justify='left',
                                  font=('Microsoft YaHei', 9), anchor='w', padx=12, pady=10)
        self.stats_lbl.pack(fill='x')

        # 右栏：编辑器
        right = tk.Frame(body, bg=C_BG)
        right.pack(side='left', fill='both', expand=True)

        # 标题输入 + 分隔线（区分标题与正文）
        self.title_entry = tk.Entry(right, bg=C_BG, fg=C_TEXT, insertbackground=C_CURSOR,
                                    relief='flat', bd=0, font=('Microsoft YaHei', 20, 'bold'))
        self.title_entry.pack(side='top', fill='x', padx=28, pady=(18, 6))
        self.title_entry.bind('<KeyRelease>', lambda e: self._schedule_save())
        tk.Frame(right, bg=C_BORDER, height=1).pack(side='top', fill='x', padx=28)

        # 工具栏
        tb = tk.Frame(right, bg=C_BG)
        tb.pack(side='top', fill='x', padx=24, pady=(10, 4))
        self.bold_btn = RoundedButton(tb, 'B', lambda: self._toggle(0), bg=C_BTN, fg=C_TEXT,
                                      hover=C_BTN_HOVER, font=('Microsoft YaHei', 11, 'bold'), padx=13, pady=5)
        self.bold_btn.pack(side='left', padx=3)
        self.italic_btn = RoundedButton(tb, 'I', lambda: self._toggle(1), bg=C_BTN, fg=C_TEXT,
                                        hover=C_BTN_HOVER, font=('Microsoft YaHei', 11, 'italic'), padx=14, pady=5)
        self.italic_btn.pack(side='left', padx=3)
        self.under_btn = RoundedButton(tb, 'U', lambda: self._toggle(2), bg=C_BTN, fg=C_TEXT,
                                       hover=C_BTN_HOVER, font=('Microsoft YaHei', 11, 'underline'), padx=13, pady=5)
        self.under_btn.pack(side='left', padx=3)
        self._btn(tb, '🖼 图片', self._pick_image, small=True).pack(side='left', padx=(12, 3))
        tk.Label(tb, text='选中文字后点 B / I / U 设置格式', bg=C_BG, fg=C_DIM,
                 font=('Microsoft YaHei', 9)).pack(side='left', padx=10)

        # 正文（行距宽松）
        self.text = tk.Text(right, bg=C_BG, fg=C_TEXT, insertbackground=C_CURSOR,
                            relief='flat', bd=0, wrap='word', undo=True,
                            font=self.f_normal, padx=28, pady=14,
                            spacing1=4, spacing2=8, spacing3=4,
                            selectbackground=C_SELECT, selectforeground=C_TEXT)
        self._init_text_tags(self.text)
        tscroll = tk.Scrollbar(right, command=self.text.yview, bg=C_DIM, troughcolor=C_BG,
                               activebackground=C_BTN_HOVER, bd=0, relief='flat', width=10)
        tscroll.pack(side='right', fill='y')
        self.text.config(yscrollcommand=tscroll.set)
        self.text.pack(side='left', fill='both', expand=True)
        self.text.bind('<KeyRelease>', lambda e: self._schedule_save())
        self.text.bind('<<Selection>>', lambda e: self._update_toolbar())

    def _build_calendar(self):
        self.cal_nav = tk.Frame(self.cal_frame, bg=C_PANEL)
        self.cal_nav.pack(side='top', fill='x', padx=10, pady=(8, 4))
        RoundedButton(self.cal_nav, '‹', lambda: self._cal_shift(-1), bg=C_BTN, fg=C_TEXT,
                      hover=C_BTN_HOVER, font=('Microsoft YaHei', 11), padx=10, pady=3).pack(side='left')
        self.cal_title = tk.Label(self.cal_nav, text='', bg=C_PANEL, fg=C_TEXT,
                                  font=('Microsoft YaHei', 10, 'bold'))
        self.cal_title.pack(side='left', expand=True)
        RoundedButton(self.cal_nav, '›', lambda: self._cal_shift(1), bg=C_BTN, fg=C_TEXT,
                      hover=C_BTN_HOVER, font=('Microsoft YaHei', 11), padx=10, pady=3).pack(side='right')
        self.cal_grid = tk.Frame(self.cal_frame, bg=C_PANEL)
        self.cal_grid.pack(side='top', fill='x', padx=8, pady=(4, 0))

    # ---------- 列表/日历 ----------
    def _switch_view(self, which):
        if which == 'list':
            self.cal_frame.pack_forget()
            self.list_frame.pack(side='top', fill='both', expand=True)
            self.tab_list_btn.set_bg(C_ACCENT, C_ACCENT_HOVER)
            self.tab_cal_btn.set_bg(C_BTN, C_BTN_HOVER)
        else:
            self.list_frame.pack_forget()
            self.cal_frame.pack(side='top', fill='both', expand=True)
            self.tab_cal_btn.set_bg(C_ACCENT, C_ACCENT_HOVER)
            self.tab_list_btn.set_bg(C_BTN, C_BTN_HOVER)
            self._render_calendar()

    def _cal_shift(self, delta):
        y, m = self._cal
        m += delta
        if m < 1:
            y -= 1
            m = 12
        elif m > 12:
            y += 1
            m = 1
        self._cal = (y, m)
        self._render_calendar()

    def _render_calendar(self):
        y, m = self._cal
        for w in self.cal_grid.winfo_children():
            w.destroy()
        self.cal_title.config(text='%d 年 %d 月' % (y, m))
        days = self.db.calendar_dates(y, m)
        # 7 列等分列宽（适配左栏），每列 uniform 共享宽度
        for c in range(7):
            self.cal_grid.grid_columnconfigure(c, weight=1, uniform='day')
        for c, name in enumerate('日一二三四五六'):
            tk.Label(self.cal_grid, text=name, bg=C_PANEL, fg=C_DIM, font=('Microsoft YaHei', 9)
                     ).grid(row=0, column=c, sticky='nsew', padx=1, pady=2)
        first = datetime.date(y, m, 1).weekday()  # 周一=0
        ndays = (datetime.date(y + (m // 12), m % 12 + 1, 1) - datetime.timedelta(days=1)).day
        r, c = 1, first
        for d in range(1, ndays + 1):
            ds = '%04d-%02d-%02d' % (y, m, d)
            if ds == today_str():
                bg, fg = C_ACCENT, '#eef2f6'
            elif ds in days:
                bg, fg = C_SELECT, C_TEXT
            else:
                bg, fg = C_PANEL, C_DIM
            btn = tk.Button(self.cal_grid, text=str(d), bg=bg, fg=fg, font=('Microsoft YaHei', 9),
                            relief='flat', bd=0, cursor='hand2', activebackground=C_BTN_HOVER,
                            activeforeground=C_TEXT, highlightthickness=0,
                            command=lambda ds=ds: self._load_date(ds))
            btn.bind('<Enter>', lambda e, b=btn: b.config(bg=C_BTN_HOVER))
            btn.bind('<Leave>', lambda e, b=btn, bb=bg: b.config(bg=bb))
            btn.grid(row=r, column=c, sticky='nsew', padx=1, pady=1, ipadx=0, ipady=3)
            c += 1
            if c > 6:
                c = 0
                r += 1

    def _refresh_list(self):
        self.listbox.delete(0, 'end')
        items = self.db.list()
        selected_index = None
        for i, it in enumerate(items):
            mark = '✓' if it['status'] == 'published' else ' '
            self.listbox.insert('end', '%s %s  %s(%s字)' % (mark, it['date'], it['title'] or '无标题', it['word_count']))
            if it['date'] == self.current_date:
                selected_index = i
        self._list_items = items
        if selected_index is not None:
            self.listbox.selection_set(selected_index)
            self.listbox.see(selected_index)

    def _on_list_select(self, e):
        sel = self.listbox.curselection()
        if sel and self._list_items:
            date = self._list_items[sel[0]]['date']
            if date == self.current_date:
                return
            self._load_date(date)

    def _refresh_stats(self):
        s = self.db.stats()
        self.stats_lbl.config(text='总篇数：%d\n连续写作：%d 天\n累计字数：%d\n写作天数：%d'
                              % (s['total'], s['streak'], s['words'], s['days']))

    # ---------- 日记加载/保存 ----------
    def _load_date(self, date):
        self._flush_save()
        self.current_date = date
        r = self.db.get(date) or {'title': '', 'content_html': '', 'status': 'draft'}
        self.title_entry.delete(0, 'end')
        self.title_entry.insert(0, r['title'] or self._default_title(date))
        self.text.delete('1.0', 'end')
        self.image_paths.clear()
        self._image_refs.clear()
        if r['content_html']:
            try:
                p = _HTMLToText(self.text, self.image_paths, self._insert_image_at_end)
                p.feed(r['content_html'])
            except Exception:
                self.text.insert('end', strip_html(r['content_html']))
        self.status_lbl.config(text='%s  ·  %s' % (date, '已导出' if r['status'] == 'published' else '草稿'))
        self._update_toolbar()

    def _insert_image_at_end(self, path):
        self._insert_image(path, 'end')

    def _insert_image(self, path, index=None):
        if not HAS_PIL:
            messagebox.showwarning('提示', '未安装 Pillow，无法插入图片')
            return
        try:
            img = Image.open(path)
            img.thumbnail((700, 700))
            imgtk = ImageTk.PhotoImage(img)
            self._image_refs.append(imgtk)
            name = self.text.image_create(index or 'insert', image=imgtk)
            self.image_paths[name] = path
        except Exception as e:
            messagebox.showerror('错误', '插入图片失败：' + str(e))

    def _pick_image(self):
        path = filedialog.askopenfilename(
            title='选择图片', filetypes=[('图片', '*.png *.jpg *.jpeg *.gif *.bmp'), ('所有文件', '*.*')])
        if path:
            self._insert_image(path)

    def _flush_save(self):
        if self._save_job:
            self.after_cancel(self._save_job)
            self._save_job = None
        if self.current_date:
            self._save_now()

    def _schedule_save(self):
        if self._save_job:
            self.after_cancel(self._save_job)
        self._save_job = self.after(5000, self._save_now)

    def _save_now(self):
        self._save_job = None
        if not self.current_date:
            return
        title = self.title_entry.get()
        html = text_to_html(self.text, self.image_paths)
        self.db.upsert(self.current_date, title, html)
        self._refresh_list()
        self._refresh_stats()
        self.status_lbl.config(text='%s  ·  草稿（已自动保存）' % self.current_date)

    def _new_diary(self):
        self._load_date(today_str())

    def _delete(self):
        if not self.current_date:
            return
        if not messagebox.askyesno('确认', '确定删除 %s 的日记？' % self.current_date):
            return
        self.db.delete(self.current_date)
        self.current_date = None
        self.image_paths.clear()
        self._image_refs.clear()
        self.title_entry.delete(0, 'end')
        self.text.delete('1.0', 'end')
        self.status_lbl.config(text='')
        self._refresh_list()
        self._refresh_stats()

    def _export(self):
        if not self.current_date:
            messagebox.showinfo('提示', '请先新建或选择一篇日记')
            return
        self._save_now()
        r = self.db.get(self.current_date)
        title = r['title'] if r else ''
        html = r['content_html'] if r else ''
        # 文件名：标题已含日期（默认标题）则不重复拼接日期
        if title and title.startswith(self.current_date):
            filename = '%s.docx' % safe_filename(title)
        else:
            filename = '%s-%s.docx' % (self.current_date, safe_filename(title))
        out = os.path.join(get_export_dir(), filename)
        self.status_lbl.config(text='正在生成 Word ...')
        self.update_idletasks()
        ok, msg = export_word(self.current_date, title, html, out)
        if ok:
            self.db.upsert(self.current_date, title, html, status='published')
            self.db.log_export(self.current_date, filename)
            self._refresh_list()
            self._refresh_stats()
            self.status_lbl.config(text='已导出：' + filename)
            messagebox.showinfo('成功', '已导出到：\n' + out)
        else:
            self.status_lbl.config(text='导出失败')
            messagebox.showerror('导出失败', msg)

    def _open_export_dir(self):
        d = get_export_dir()
        os.makedirs(d, exist_ok=True)
        os.startfile(d)

    def _set_export_dir(self):
        """选择导出目录，保存到 config.json。"""
        cur = get_export_dir()
        d = filedialog.askdirectory(title='选择导出目录', initialdir=cur, parent=self)
        if not d:
            return
        cfg = load_config()
        cfg['export_dir'] = d
        save_config(cfg)
        self.status_lbl.config(text='导出目录：%s' % d)

    def _do_search(self):
        kw = self.search_var.get().strip()
        self._flush_save()
        res = self.db.search(keyword=kw)
        self.listbox.delete(0, 'end')
        self._list_items = res
        for it in res:
            self.listbox.insert('end', '%s  %s' % (it['date'], it['title'] or '无标题'))
        if kw:
            self.status_lbl.config(text='搜索“%s”：%d 条' % (kw, len(res)))
        else:
            self._refresh_list()

    # ---------- 富文本格式 ----------
    def _bits_at(self, index):
        names = self.text.tag_names(index)
        return [
            1 if any(t in ('b', 'bi', 'bu', 'biu') for t in names) else 0,
            1 if any(t in ('i', 'bi', 'iu', 'biu') for t in names) else 0,
            1 if any(t in ('u', 'bu', 'iu', 'biu') for t in names) else 0,
        ]

    def _apply_fmt(self, start, end, bits):
        b, i, u = bits
        key = ('b' if b else '') + ('i' if i else '') + ('u' if u else '')
        for t in ('b', 'i', 'u', 'bi', 'bu', 'iu', 'biu'):
            self.text.tag_remove(t, start, end)
        if key:
            self.text.tag_add(key, start, end)

    def _toggle(self, bit):
        sel = self.text.tag_ranges('sel')
        if sel:
            start, end = sel[0], sel[1]
        else:
            start = end = self.text.index('insert')
        bits = self._bits_at(start)
        bits[bit] = 1 - bits[bit]
        self._apply_fmt(start, end, bits)
        self._update_toolbar()

    def _update_toolbar(self):
        try:
            idx = self.text.index('sel.first') if self.text.tag_ranges('sel') else self.text.index('insert')
        except Exception:
            idx = 'insert'
        bits = self._bits_at(idx)
        self.bold_btn.set_bg(C_BTN_ACT if bits[0] else C_BTN, C_BTN_HOVER)
        self.italic_btn.set_bg(C_BTN_ACT if bits[1] else C_BTN, C_BTN_HOVER)
        self.under_btn.set_bg(C_BTN_ACT if bits[2] else C_BTN, C_BTN_HOVER)

    def _on_close(self):
        self._flush_save()
        self.db.close()
        self.destroy()


def main():
    if not HAS_WIN32:
        messagebox.showwarning('提示', '未检测到 pywin32，导出 Word 功能将不可用。\n请执行：pip install pywin32')
    app = DiaryApp()
    app.mainloop()


if __name__ == '__main__':
    main()
