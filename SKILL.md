---
name: momenta-ppt
description: 生成或润色 Momenta 风格 PPT。当用户要求制作新 PPT、对已有 PPT 进行格式修改/润色，或提供飞书文档链接要求转化为 PPT 时触发。
license: Proprietary
allowed-tools:
  - Bash
  - Read
  - Write
  - Edit
---

# Momenta PPT 制作规范

## 触发场景

以下情况均应主动应用本 Skill：
- 用户要求**新建 PPT**
- 用户要求对已有 PPT 进行**格式修改或润色**
- 用户提供**飞书文档链接**并要求转化为 PPT
- 用户要求生成某个主题的汇报/演示文稿

---

## 零、自动更新（每次触发必须首先执行）

**在执行任何任务之前，必须先运行以下命令拉取最新 Skill：**

```bash
# 定位 skill 目录并拉取更新
SKILL_DIR=$(python3 -c "import os; print(os.path.join(os.path.expanduser('~'), '.claude', 'skills', 'momenta-ppt'))" 2>/dev/null || echo "$USERPROFILE/.claude/skills/momenta-ppt")
git -C "$SKILL_DIR" pull --ff-only 2>&1
```

- 若输出 `Already up to date.`：继续执行任务
- 若输出有文件变更（如 `SKILL.md | N +++`）：告知用户"Skill 已更新到最新版本"，然后继续执行任务
- 若输出网络错误或非 git 仓库：**忽略错误，继续执行任务**（不因更新失败阻塞用户）

---

## 一、环境与路径

路径因用户电脑而异，脚本开头必须用以下代码**自动检测**，禁止硬编码任何用户名或绝对路径：

```python
import os, sys, platform, subprocess

HOME        = os.path.expanduser("~")
# 优先读取环境变量 MOMENTA_PPT_DIR，未设置则默认 ~/Desktop/Claude
SCRIPT_DIR  = os.environ.get("MOMENTA_PPT_DIR") or os.path.join(HOME, "Desktop", "Claude")
os.makedirs(SCRIPT_DIR, exist_ok=True)   # 目录不存在时自动创建
CN_TEMPLATE = os.path.join(SCRIPT_DIR, "PPT模板.pptx")
EN_TEMPLATE = os.path.join(SCRIPT_DIR, "Momenta PPT模板英文EN.pptx")
```

> 用户若想自定义目录，在 shell 中设置一次即永久生效：
> - macOS/Linux：`export MOMENTA_PPT_DIR="/your/path"` 加入 `~/.zshrc`
> - Windows：系统环境变量中添加 `MOMENTA_PPT_DIR`

运行命令（跨平台，使用当前 Python 解释器）：
```bash
# macOS / Linux
python3 '<脚本路径>'

# Windows（Git Bash / 命令提示符）
python '<脚本路径>'
```

生成后打开文件（跨平台）：
```bash
# macOS
open '<输出路径>'

# Windows
start '<输出路径>'
```

> 生成前确认目标 .pptx 文件已在 PowerPoint 中关闭，否则报 PermissionError。

---

## 二、Momenta 配色常量

```python
BLUE   = RGBColor(0x00, 0x68, 0xE9)   # 主色：标题 / 目录 / 表格表头
ORANGE = RGBColor(0xED, 0x7D, 0x31)   # 辅色：重要标注 / Before 列表头 / 徽章
DARK   = RGBColor(0x1F, 0x23, 0x29)   # 正文文字
GRAY   = RGBColor(0x75, 0x78, 0x7E)   # 辅助说明 / 页脚
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
```

---

## 三、安全区域（所有内容必须在此范围内）

```python
SAFE_L = Emu(468630)
SAFE_T = Emu(601345)
SAFE_W = Emu(11360785)
SAFE_H = Emu(5548630)
SAFE_R = SAFE_L + SAFE_W   # = 11829415
SAFE_B = SAFE_T + SAFE_H   # = 6149975
```

---

## 四、字号规范

| 用途 | 字号 | 说明 |
|------|------|------|
| 封面大标题 | **60 pt** | bold，白色，word_wrap=False |
| 内容页标题 | **24 pt** | bold，Momenta 蓝色 |
| 副标题 | 14 pt | 墨黑 |
| 正文内容 | **12–16 pt** | 墨黑，根据内容量调整 |
| 内容较多时最小 | 8–10 pt | 行距 1 磅 |
| 封面日期 | 24 pt | 白色 |
| 底部备注 / 页脚 | 10 pt | 灰色 |

---

## 五、字体规范（中英混排，每个 run 必须分段设置）

**原则**：
- 中文字符 → 随系统自动选择：Windows 用**微软雅黑**，macOS 用**苹方（PingFang SC）**
- 英文字符 → Arial（两平台均有）

**实现**：用 `CJK_FONT` 常量统一控制，将文本按 CJK / Latin 拆分成独立 run，分别设置 `a:latin` 和 `a:ea`。

```python
import platform, lxml.etree as etree
from pptx.oxml.ns import qn

# 中文字体：随平台自动选择
CJK_FONT = "PingFang SC" if platform.system() == "Darwin" else "微软雅黑"

def is_cjk_char(c):
    cp = ord(c)
    return (0x4E00 <= cp <= 0x9FFF or 0x3400 <= cp <= 0x4DBF or
            0x20000 <= cp <= 0x2A6DF or 0xF900 <= cp <= 0xFAFF or
            0x3000 <= cp <= 0x303F or 0xFF01 <= cp <= 0xFF60 or
            0xFE30 <= cp <= 0xFE4F)

def _set_run_font(run, face):
    rPr = run._r.get_or_add_rPr()
    for tag in (qn("a:latin"), qn("a:ea")):
        elem = rPr.find(tag)
        if elem is None: elem = etree.SubElement(rPr, tag)
        elem.set("typeface", face)

def _add_text_as_runs(p, text, size, bold, color):
    if not text: return
    segments, cur_cjk, buf = [], None, []
    for ch in text:
        cjk = is_cjk_char(ch)
        if cjk != cur_cjk:
            if buf: segments.append((cur_cjk, "".join(buf)))
            cur_cjk, buf = cjk, [ch]
        else: buf.append(ch)
    if buf: segments.append((cur_cjk, "".join(buf)))
    for seg_cjk, seg_text in segments:
        r = p.add_run()
        r.text = seg_text
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
        _set_run_font(r, CJK_FONT if seg_cjk else "Arial")
```

---

## 六、标准工具函数（每个脚本必须包含）

```python
import sys, os, platform, subprocess, urllib.parse, urllib.request

# ── 自动安装依赖（首次运行自动完成，无需手动 pip）──────────────
def _ensure_deps():
    try:
        import pptx  # noqa
    except ImportError:
        print("正在安装 python-pptx，请稍候...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "python-pptx", "-q"])
        print("安装完成。")
_ensure_deps()

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ── 跨平台路径 ─────────────────────────────────────────────
# 优先读取环境变量 MOMENTA_PPT_DIR，未设置则默认 ~/Desktop/Claude
HOME        = os.path.expanduser("~")
SCRIPT_DIR  = os.environ.get("MOMENTA_PPT_DIR") or os.path.join(HOME, "Desktop", "Claude")
os.makedirs(SCRIPT_DIR, exist_ok=True)   # 目录不存在时自动创建
CN_TEMPLATE = os.path.join(SCRIPT_DIR, "PPT模板.pptx")
EN_TEMPLATE = os.path.join(SCRIPT_DIR, "Momenta PPT模板英文EN.pptx")

GITHUB_RAW_BASE = "https://raw.githubusercontent.com/yutingzhou0202/momenta-ppt/main"
TEMPLATE_FILES = {
    CN_TEMPLATE: "PPT模板.pptx",
    EN_TEMPLATE: "Momenta PPT模板英文EN.pptx",
}

def _download_template(filename, dest_path):
    import urllib.request
    import urllib.parse
    url = f"{GITHUB_RAW_BASE}/{urllib.parse.quote(filename)}"
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    print(f"正在自动下载模板：{filename} ...")
    urllib.request.urlretrieve(url, dest_path)
    print(f"下载完成 → {dest_path}")

def check_templates(need_cn=True, need_en=False):
    """模板不存在时自动从 GitHub 下载，下载失败再提示用户"""
    needed = []
    if need_cn and not os.path.exists(CN_TEMPLATE): needed.append((CN_TEMPLATE, "PPT模板.pptx"))
    if need_en and not os.path.exists(EN_TEMPLATE): needed.append((EN_TEMPLATE, "Momenta PPT模板英文EN.pptx"))
    for dest, fname in needed:
        try:
            _download_template(fname, dest)
        except Exception as e:
            print("=" * 60)
            print(f"模板自动下载失败：{e}")
            print(f"请手动下载后放到：{SCRIPT_DIR}")
            print(f"下载地址：{GITHUB_RAW_BASE}/{urllib.parse.quote(fname)}")
            print("=" * 60)
            sys.exit(1)

def open_file(path):
    """跨平台打开文件"""
    if platform.system() == "Darwin":
        os.system(f'open "{path}"')
    elif platform.system() == "Windows":
        os.system(f'start "" "{path}"')
    else:
        os.system(f'xdg-open "{path}"')

from datetime import date
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
import lxml.etree as etree

# ── 配色 ──────────────────────────────────────────────────
BLUE   = RGBColor(0x00, 0x68, 0xE9)
ORANGE = RGBColor(0xED, 0x7D, 0x31)
DARK   = RGBColor(0x1F, 0x23, 0x29)
GRAY   = RGBColor(0x75, 0x78, 0x7E)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)

# ── 安全区域 ──────────────────────────────────────────────
SAFE_L = Emu(468630);  SAFE_T = Emu(601345)
SAFE_W = Emu(11360785); SAFE_H = Emu(5548630)
SAFE_R = SAFE_L + SAFE_W; SAFE_B = SAFE_T + SAFE_H

# ── 字体：随平台自动选择 ───────────────────────────────────
CJK_FONT = "PingFang SC" if platform.system() == "Darwin" else "微软雅黑"

def is_cjk_char(c):
    cp = ord(c)
    return (0x4E00<=cp<=0x9FFF or 0x3400<=cp<=0x4DBF or 0x20000<=cp<=0x2A6DF
            or 0xF900<=cp<=0xFAFF or 0x3000<=cp<=0x303F
            or 0xFF01<=cp<=0xFF60 or 0xFE30<=cp<=0xFE4F)

def _set_run_font(run, face):
    rPr = run._r.get_or_add_rPr()
    for tag in (qn("a:latin"), qn("a:ea")):
        elem = rPr.find(tag)
        if elem is None: elem = etree.SubElement(rPr, tag)
        elem.set("typeface", face)

def _add_text_as_runs(p, text, size, bold, color):
    if not text: return
    segments, cur_cjk, buf = [], None, []
    for ch in text:
        cjk = is_cjk_char(ch)
        if cjk != cur_cjk:
            if buf: segments.append((cur_cjk, "".join(buf)))
            cur_cjk, buf = cjk, [ch]
        else: buf.append(ch)
    if buf: segments.append((cur_cjk, "".join(buf)))
    for seg_cjk, seg_text in segments:
        r = p.add_run()
        r.text = seg_text
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color
        _set_run_font(r, CJK_FONT if seg_cjk else "Arial")

# ── 文本框 ────────────────────────────────────────────────
def add_textbox(slide, left, top, width, height,
                text, size=14, bold=False, color=DARK,
                align=PP_ALIGN.LEFT, bg=None, word_wrap=True):
    txb = slide.shapes.add_textbox(left, top, width, height)
    if bg: txb.fill.solid(); txb.fill.fore_color.rgb = bg
    tf = txb.text_frame; tf.word_wrap = word_wrap
    p = tf.paragraphs[0]; p.alignment = align
    _add_text_as_runs(p, text, size, bold, color)
    return txb

def add_multiline(slide, left, top, width, height,
                  lines, size=12, color=DARK, bold_first=False):
    txb = slide.shapes.add_textbox(left, top, width, height)
    tf = txb.text_frame; tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        _add_text_as_runs(p, line, size, bold_first and i == 0, color)
    return txb

# ── 矩形 ──────────────────────────────────────────────────
def add_rect(slide, left, top, width, height, fill_color=None):
    shape = slide.shapes.add_shape(1, left, top, width, height)
    if fill_color: shape.fill.solid(); shape.fill.fore_color.rgb = fill_color
    else: shape.fill.background()
    shape.line.fill.background()
    return shape

# ── Slide 管理 ────────────────────────────────────────────
def delete_slide(prs, idx):
    xml_slides = prs.slides._sldIdLst
    slide_el = xml_slides[idx]
    prs.part.drop_rel(slide_el.get(qn("r:id")))
    xml_slides.remove(slide_el)

def move_slide_to_end(prs, idx):
    xml_slides = prs.slides._sldIdLst
    slide_el = xml_slides[idx]
    xml_slides.remove(slide_el); xml_slides.append(slide_el)

def clean_layout(layout):
    to_remove = [s for s in layout.shapes
                 if s.has_text_frame and any(kw in s.text_frame.text for kw in
                    ("Agenda","agenda","©","All Rights Reserved",
                     "Text Content","Paragraph content","Title"))]
    for s in to_remove: layout.shapes._spTree.remove(s.element)

def _remove_placeholders(slide):
    sp_tree = slide.shapes._spTree
    for ph in list(slide.placeholders): sp_tree.remove(ph.element)
```

---

## 七、页面布局常量

```python
# 内容页 layout（使用 CN 模板 slide[3]）
CONTENT_LAYOUT = prs.slides[3].slide_layout
clean_layout(CONTENT_LAYOUT)

def make_slide():
    slide = prs.slides.add_slide(CONTENT_LAYOUT)
    _remove_placeholders(slide)
    return slide

# 内容区起点（副标题下方，统一距顶部 1020000）
CONTENT_T = SAFE_T + Emu(1020000)

# 标准页面元素
def page_title(slide, text):
    add_textbox(slide, SAFE_L, SAFE_T, SAFE_W, Emu(540000),
                text, size=24, bold=True, color=BLUE)

def page_subtitle(slide, text):
    add_textbox(slide, SAFE_L, SAFE_T + Emu(590000), SAFE_W, Emu(350000),
                text, size=14, color=DARK)

def page_footer(slide, text):
    add_textbox(slide, SAFE_L, SAFE_B - Emu(300000), SAFE_W, Emu(270000),
                text, size=10, color=GRAY)
```

---

## 八、封面（直接复用 CN 模板 slide[0]）

```python
s1 = prs.slides[0]
_remove_placeholders(s1)

# 主标题：60pt bold 白色，禁止换行
add_textbox(s1, Emu(970402), Emu(3086480), Emu(7886700), Emu(694184),
            "PPT 标题", size=60, bold=True, color=WHITE, word_wrap=False)

# 日期（月/日/年格式）：24pt 白色
date_str = date.today().strftime("%B %d, %Y")
add_textbox(s1, Emu(970402), Emu(3939468), Emu(10515601), Emu(1500188),
            date_str, size=24, color=WHITE)
```

---

## 九、致谢页（每次必须加在末尾）

```python
# 删除模板原始页 1–8（保留 slide[0] 封面 和 slide[9] Thanks）
for _ in range(8):
    delete_slide(prs, 1)
# Thanks 当前 index=1，移至末尾
move_slide_to_end(prs, 1)
# 最终顺序: [0=封面, 1..N-1=内容页, N=Thanks]
```

---

## 十、表格规范

- **表头**：Momenta 蓝为主色，Momenta 橙为辅色（如 Before/After 对比表，Before 列用橙色）
- **表头文字**：白色，13 pt bold
- **数据行**：正文 12 pt，交替浅蓝（`#F2F7FF`）/ 白色背景
- **示例**：

```python
LIGHT_BLUE = RGBColor(0xF2, 0xF7, 0xFF)
LIGHT_GRAY = RGBColor(0xF5, 0xF5, 0xF6)

def table_header_cell(slide, left, top, width, height, text, bg=BLUE, size=13):
    add_rect(slide, left, top, width, height, fill_color=bg)
    add_textbox(slide, left + Emu(80000), top, width - Emu(80000), height,
                text, size=size, bold=True, color=WHITE, align=PP_ALIGN.LEFT)

def table_data_cell(slide, left, top, width, height, lines, bg=WHITE, size=12):
    add_rect(slide, left, top, width, height - Emu(15000), fill_color=bg)
    if isinstance(lines, str): lines = [lines]
    add_multiline(slide, left + Emu(80000), top + Emu(60000),
                  width - Emu(160000), height - Emu(80000),
                  lines, size=size, color=DARK)
```

---

## 十一、强调徽章（可选，右上角）

```python
BADGE_W = Emu(3050000)
add_rect(slide, SAFE_R - BADGE_W, SAFE_T + Emu(30000), BADGE_W, Emu(480000),
         fill_color=ORANGE)
add_textbox(slide, SAFE_R - BADGE_W, SAFE_T + Emu(40000), BADGE_W, Emu(480000),
            "徽章文字", size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
```

---

## 十二、工作流程

### 新建 PPT
1. 读取内容（可来自用户描述或飞书文档）
2. 规划页面数量与每页主题
3. 在 `~/Desktop/Claude/` 创建 `make_<名称>.py`，脚本内用 `SCRIPT_DIR` 动态路径，禁止硬编码
4. 脚本 main 逻辑第一行调用 `check_templates(need_cn=True)` 确保模板存在
5. 使用上述所有规范编写脚本
6. 运行脚本生成 .pptx，调用 `open_file(output_path)` 打开查看（跨平台）

### 飞书文档转 PPT
1. 使用 feishu skill 读取文档内容（`feishu-sync-cli read_page_as_markdown <url>`）
2. 分析文档结构，规划幻灯片页数（通常 5–8 页内容页 + 封面 + Thanks）
3. 按本规范生成脚本并运行

### 格式润色已有脚本
1. 读取现有 `.py` 脚本
2. 对照本规范检查：字号、字体分段、表头配色、安全区域、内容起点
3. 修改不符合规范的地方并重新生成

---

## 十三、注意事项

- 生成前确认目标 .pptx 已在 PowerPoint 中关闭，否则报 PermissionError
- `chart.chart_area` 在当前版本不可用，勿使用
- 数据标签须显式开启：`series.data_labels.show_value = True`
- 柱状图逐柱着色：`series.points[i].format.fill.fore_color.rgb = color`
- CN 模板共 10 页：slide[0]=封面，slide[3]=内容页layout，slide[9]=Thanks
- EN 模板共 6 页：无 Thanks 页，用 `for _ in range(6): delete_slide(prs, 0)` 删除全部原始页
- 所有内容元素 left ≥ SAFE_L，top ≥ SAFE_T，right ≤ SAFE_R，bottom ≤ SAFE_B
