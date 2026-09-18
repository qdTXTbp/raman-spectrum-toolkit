"""拉曼光谱工具 · Raman Spectrum Toolkit
（JASCO .jws 转换 · 拉曼峰分析 · 矿物鉴定 / Convert · Analyze · Identify）

基本转换：
    - Excel(.xlsx)：数据 + 内嵌折线图，双击用 WPS/Excel 打开即见图；
    - PNG 图片：光谱曲线图（整数刻度、峰位标注、可自定义），文件名与源文件一致；
    - CSV：纯数据，第一列横坐标（默认拉曼位移 cm-1），之后每个通道一列；
    - 峰列表 CSV（*_peaks.csv）：峰位、强度、相对强度、半高宽 FWHM、峰突出度、来源、归属；
    - JCAMP-DX（.jdx）：等间距数据用 ##X++(Y..Y)，通用谱学软件可读。

支持读入的格式：.jws / .csv / .spc（Galactic-SPC）/ .jdx（JCAMP-DX）
    / .txt .dat .asc .xy（两列数值）/ .xlsx

格式参考开源项目 jasco_jws_reader / jasco-jws-converter 的 DataInfo 结构说明：
    DataInfo[12] channels, [20] point number, [24] x_first, [32] x_last,
    [40] x_increment, [48] x 轴类型码, [52+] 各通道类型码。
本工具的数值解析已与上述参考实现逐点比对，结果完全一致。

界面自带光谱预览：选中文件即可看图；可在图上左键补标峰、右键删除。

用法：
    双击 “启动拉曼光谱工具.bat”        打开图形界面
    python jws2csv.py                  打开图形界面
    python jws2csv.py a.jws 某文件夹    命令行批量转换（文件夹递归查找光谱文件）
    python jws2csv.py --png a.jws      导出 PNG 图片
    python jws2csv.py --xlsx a.jws     导出带图表的 Excel
    python jws2csv.py --peaks a.jws    导出峰列表 CSV
    python jws2csv.py --jcamp a.jws    导出 JCAMP-DX
    python jws2csv.py --all a.jws      导出全部（CSV + Excel + PNG + 峰列表 + 峰拟合）
    python jws2csv.py --waterfall 文件夹   多条光谱瀑布（堆叠）图

预处理（可叠加，默认只作用于出图与峰识别）：
    --despike                          尖峰 / 宇宙射线去除
    --despike-thresh 10 --despike-window 5
    --baseline iterpoly|rolling_ball|rolling|poly2|linear
    --baseline-degree 5 --baseline-iters 20 --baseline-window 60
    --smooth 9 --normalize max --calib 520.6,520.7
    --mineral Zircon                   峰位归属参考（峰表/报告多一列“归属”）

分析：
    --cluster [文件或文件夹]            层次聚类 + PCA（出树状图 / 主成分散点 / 分组表）
    --cluster-cut 0.2                  手动指定聚类分割阈值
    --map 5,5 --map-metric main_peak   二维成像热图（主峰位/强度/FWHM/峰数/总强度）
    --map-metric at:1008               指定波数附近的峰强度作为成像指标
    --report [文件夹] [--out 目录]      生成自包含 HTML 分析报告（峰表 + 图 + 参数）
    --batch 文件夹 [--out 目录]         整目录批处理（转换 + 峰表 + 出图 + 汇总统计）
    --pair 文件                        与本地参考谱库自动配对（峰位匹配 F1 排序）
    --pair-batch 文件夹 [--pair-ref 目录] 批量配对：多条实测谱逐条配参考谱集，
                                       汇总表按综合分升序（最可疑的在前），
                                       并给 CSV + HTML 报告
    --identify-batch 文件夹            批量鉴定：多条陌生谱逐条全库检索，
                                       汇总表给最佳候选 + 综合分 + 结论文本
    --db-match 文件                    与本地参考谱库比对（相关系数 / 谱角）

数据库与矿物信息：
    --db-search 关键字                 在线检索 ROD 库
    --mineral-search Zircon            内置矿物表检索（名称 / 化学式）
    --by-element "Zr Si"               按元素检索内置矿物表
    --by-formula SiO2                  按化学式检索
    --mineral-info Zircon              矿物信息卡（特征峰归属 + RRUFF 真实样品记录）
    --rruff-list                       查看全部 RRUFF 数据包（拉曼 / 红外 / XRD / 成分）
    --rruff-get 数据包key              下载并建索引
    --rruff-search Zircon              检索（也可直接输 RRUFF 编号，如 R050034）
    --rruff-export Zircon              导出到本地参考谱库
    --rruff-fetch Zircon               一键下载 + 检索 + 导出
    --import-pkg 路径.zip              导入自己下载的 RRUFF 数据包

数据目录与容量（所有工具自身数据只写这里）：
    <工具目录>/工具数据/
        参考谱库/      ROD 下载 + RRUFF 导出的拉曼参考谱 CSV
        参考谱库/IR|XRD/  其他模态参考谱（分目录存放，避免跨模态误配对）
        RRUFF数据包/   数据包 zip 与索引
        分析结果/      峰拟合、配对报告、聚类、成像、报告等输出
    --data-dir            查看数据目录位置与占用
    --data-dir <路径>     更改数据目录
    --cache-limit [MB]    查看 / 设置数据包容量上限（0 = 不限制）
    --cleanup [MB]        清理数据包，可指定要释放的容量
    下载前会自动检查磁盘剩余空间；目录不可写时自动改用系统用户目录；
    旧版“光谱数据库”会自动迁移。

导出 Excel 需 openpyxl，导出 PNG 需 Pillow；CSV 与预览功能仅用标准库。

使用说明书：
    界面里看：菜单【帮助】→【使用说明（完整手册）…】（章节跳转 + 关键词搜索 +
    字号可调）；文件形式：工具目录下的 使用说明.txt（首次运行自动生成，
    版本升级自动更新，你自己写的同名文件不会被覆盖）；
    命令行：--manual 打印全文，--manual 路径.txt 导出到文件。

界面语言：
    中英双语。菜单【设置】→【语言】切换，或命令行 --lang en|zh。
    界面、日志、命令行输出、分析报告、图注与说明书一起切换；
    磁盘上的数据文件夹名保持中文，保证两种语言下数据互通。
"""

import base64
import colorsys
import json
import math
import os
import queue
import re
import shutil
import struct
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from collections import deque

try:
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference, ScatterChart, Series
    _HAVE_XLSX = True
except Exception:
    _HAVE_XLSX = False

try:
    from PIL import Image, ImageDraw, ImageFont
    _HAVE_PIL = True
except Exception:
    _HAVE_PIL = False

END_OF_CHAIN = 0xFFFFFFFE
FREE_SECT = 0xFFFFFFFF

_DATA_TYPES = {
    268435715: "WAVELENGTH",
    4097: "CD",
    8193: "HT VOLTAGE",
    3: "ABSORBANCE",
    14: "FLUORESCENCE",
}

_DEFAULT_X_HEADER = "Raman shift (cm-1)"
_DEFAULT_Y_HEADER = "Intensity"

_PALETTE = ["#1f77b4", "#d62728", "#2ca02c", "#ff7f0e", "#9467bd",
            "#8c564b", "#e377c2", "#17becf"]

_PLOT_DEFAULT = {
    "x_step": 200.0,
    "x_start": None,
    "x_min": None,
    "x_max": None,
    "y_min": None,
    "y_max": None,
    "show_y_ticks": False,
    "show_grid": True,
    "show_title": True,
    "annotate_peaks": True,
    "peak_min_dist": 20.0,
    "peak_thresh_pct": 7.0,
    "peak_label_rel": False,
    "peak_dash_line": True,
    "peak_labels": True,
    "smooth_window": 1,
    "smooth_mode": "mean",
    "despike": False,
    "despike_window": 5,
    "despike_thresh": 8.0,
    "baseline": "none",
    "baseline_window": 60.0,
    "baseline_degree": 5,
    "baseline_iters": 20,
    "derivative": 0,
    "normalize": "none",
    "calib_pairs": None,
    "apply_to_data": False,
    "fit_shape": "voigt",
    "fit_peaks": False,
    "fig_width": 1600,
    "fig_height": 900,
    "manual_peaks": None,
    "hidden_peaks": None,
    "mineral_name": None,
    "stack_offset": 1.0,
    "stacked": False,
    "merge_peak_labels": False,
    "peak_merge_tol": None,
    "peak_marks": None,
    # 图例位置：right（绘图区右侧留白，默认，永不压谱线）/ top / bottom /
    # inside（图内自由位置，legend_xy 是比例坐标，可在叠加图预览窗口里拖）/ none
    "legend_pos": "right",
    "legend_xy": None,
}


def _plot_opts(plot=None):
    opts = dict(_PLOT_DEFAULT)
    if plot:
        for k in _PLOT_DEFAULT:
            if k in plot and plot[k] is not None:
                opts[k] = plot[k]
    return opts


# ---------------------------------------------------------------- 界面语言 --
# 中英双语：界面/命令行文案全部以中文原文为键，英文译文放在 _I18N_RAW 表里；
# 换语言时按登记表把控件文案重新套用一遍（不留残影，可来回切换）。

_UI_LANG = ["zh"]
_LANG_LABELS = {"zh": "中文", "en": "English"}
_I18N = [None]
_I18N_MISS = set()

_LOC_TEXT = []
_LOC_MENU = []
_LOC_HEAD = []
_LOC_TITLE = []
_LOC_TEXTVAR = []


def ui_lang():
    return _UI_LANG[0]


def _sp():
    """拼接两句话时用：英文需要空格，中文直接连写。"""
    return " " if _UI_LANG[0] == "en" else ""


def _unescape(text):
    """把表里的 \\n \\t \\\\ 还原成真实字符。"""
    out = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == "\\" and i + 1 < len(text):
            nxt = text[i + 1]
            if nxt == "n":
                out.append("\n")
                i += 2
                continue
            if nxt == "t":
                out.append("\t")
                i += 2
                continue
            if nxt == "\\":
                out.append("\\")
                i += 2
                continue
        out.append(ch)
        i += 1
    return "".join(out)


def _parse_i18n(raw):
    """解析 `中文||English` 表；表内用 \\n 表示换行，\\| 表示竖线。

    日志类文案带前导空格（对齐用），所以不做 strip；同时登记去空格版本，
    这样写了空格或没写空格都能命中。
    """
    table = {}
    for line in raw.splitlines():
        if not line or line.lstrip().startswith("#") or "||" not in line:
            continue
        cn, _, en = line.partition("||")
        cn = _unescape(cn.replace("\\|", "|"))
        en = _unescape(en.replace("\\|", "|"))
        if not cn.strip() or not en.strip():
            continue
        table[cn] = en
        if cn != cn.strip():
            table.setdefault(cn.strip(), en)
    return table


def i18n_table():
    if _I18N[0] is None:
        _I18N[0] = _parse_i18n(_I18N_RAW)
    return _I18N[0]


def _template_regex(cn):
    """把带占位符的中文模板编译成正则（占位符变成捕获组）。"""
    out = []
    pos = 0
    for m in _SPEC_RE.finditer(cn):
        out.append(re.escape(cn[pos:m.start()]))
        out.append("%" if m.group(0) == "%%" else "(.*?)")
        pos = m.end()
    out.append(re.escape(cn[pos:]))
    return "".join(out)


def _fill_template(en, values):
    """把捕获到的实际值填回英文模板。"""
    out = []
    pos = 0
    values = list(values)
    idx = 0
    for m in _SPEC_RE.finditer(en):
        out.append(en[pos:m.start()])
        if m.group(0) == "%%":
            out.append("%")
        else:
            out.append(values[idx] if idx < len(values) else "")
            idx += 1
        pos = m.end()
    out.append(en[pos:])
    return "".join(out)


def i18n_templates():
    """模板回退翻译表：用于已经格式化好的字符串（如先赋值模板再拼接输出）。"""
    if _I18N_TMPL[0] is None:
        pats = []
        for cn, en in i18n_table().items():
            if "%" not in cn:
                continue
            try:
                pats.append((re.compile("^" + _template_regex(cn) + "$"), cn, en))
            except re.error:
                continue
        _I18N_TMPL[0] = pats
    return _I18N_TMPL[0]


def T(text):
    """把中文界面文案翻成当前语言；未收录的文案原样返回（并记下来便于查漏）。

    先查精确表；查不到再按“模板”匹配一次，这样先格式化再输出的文案
    （例如先拼好 "已下载，索引 639 条" 再打印）也能翻译。
    """
    if not text or _UI_LANG[0] == "zh":
        return text
    if not any("\u4e00" <= ch <= "\u9fff" for ch in text):
        return text
    table = i18n_table()
    hit = table.get(text)
    if hit is None:
        stripped = text.strip()
        if stripped != text:
            hit = table.get(stripped)
            if hit is not None and text.startswith(" "):
                pad = len(text) - len(text.lstrip(" "))
                hit = " " * pad + hit
    if hit is None:
        for rx, cn, en in i18n_templates():
            m = rx.match(text)
            if m:
                return _fill_template(en, m.groups())
    if hit is None:
        _I18N_MISS.add(text)
        return text
    return hit


def i18n_missing():
    """英文模式下没找到译文的中文文案（自测用）。"""
    return sorted(_I18N_MISS)


def set_ui_lang(code, persist=True, apply_now=True):
    code = "en" if str(code).lower().startswith("en") else "zh"
    if code == _UI_LANG[0] and not apply_now:
        return code
    _UI_LANG[0] = code
    if apply_now:
        apply_language()
    if persist:
        try:
            data = _load_settings()
            data["lang"] = code
            _save_settings(data)
        except Exception:
            pass
    return code


def detect_ui_lang():
    """默认语言：中文系统用中文，其余用英文；用户选过就按用户设置。"""
    try:
        saved = (_load_settings().get("lang") or "").strip().lower()
        if saved in ("zh", "en"):
            return saved
    except Exception:
        pass
    if os.name == "nt":
        try:
            import ctypes
            langid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            if langid & 0x3FF == 0x04:
                return "zh"
            return "en"
        except Exception:
            pass
    for key in ("LANG", "LC_ALL", "LC_MESSAGES"):
        if str(os.environ.get(key) or "").lower().startswith("zh"):
            return "zh"
    return "en"


_LOC_SEEN = set()
# ===== 应用图标（内嵌 base64，冻结成 exe 后也能显示，无需外部文件）=====
_APP_ICON_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAATXUlEQVR42uWbe5BcVZ3HP79zbndP90wmM8nkTZIZkkDe"
    "TwghIUYFHwuCKA6C+GLVRReyVqFbbu1uVZaqZXcttSxFWLdqRZddFKEsNSwPeSgR5RUSEhKHCQxJJsnkNa9k3t197/nt"
    "H/fe7tuTCQRr/4ixq7ru7dv3nnO+v/fjXGGsT3Oz5aGHAoDzr2geT5C6CpErgVXAeUA1ggAgQngmiEg0gITX4/+J7hnj"
    "mowa45R7qHxOTj+uYmQQzCGMbFPso67II3sf+vrJ0ZiSHzkd+Lmrb6rVHBtV9HNiTJOIQdWhhHOdZeCjc4MYi1iLoqjT"
    "fcaYHwRu8K62++/qo/lBy0PXB6cnwIYNHlu2+I3rP/5+49m7rZea6wIfdc4hogimtIqzDXx5DEXEgYrxUsakMjjfbwuC"
    "wq1t93/zCTZs8thyhx9DtqPBN62/4Vab8n4iIhOCoOgDghGDyNkPPrwuCEbECKg6vxiIsQ3Gep+auGRtV/fDd764YcMm"
    "r719iysToLnZ8uijQdP65lttpup7GvgOVRUxNpxdyrJyNoMvzRGPKyJiDBo4VNVW5a6asPCSru0P3/kizQ9aWh5SiXX+"
    "/PU3vFdS9ml1LgA1iJHkoH964EfNAYoxzngp6wrB5a//z52/prnZCmwyF659pbqYyu0Wa2epCxxizDkFXqJzVSeptFEX"
    "HBBTWLxnZm7QwB2ukM5utOmqWc4F/jkLHsAY4/yib6tys5xLb+SOO5zM/eBNtcGw32KsN13VKSLmnARfmk+deJ6oCw67"
    "grfQuOHgKuulZ6j7MwAvgBGjga8mlZ4hWb3KqHBN6ONFKyKDcxF8/KwxihgV1WsMwipVJ4iacxG8AiqKVto1o84JyCpP"
    "lRlOXXR3YjHxub4VRc9O8MroEFlKqUs8kDqHKjM8kJyqlgYPz4Qy9j8N8FqxzgqspZ8KoBq7RBBynqKg8SRavlmTBElI"
    "zyngeUvwImViKv8/4HVMzjPq2XKKo1oS6vI1FBQ857TMnORTpTE0opyU6YRGYvX24B3CcMGS8ZSUVRx/HHiQBHtOA76U"
    "p0ZMTRAiBhyOEd8lGKdK/NXEefkbUlBL1NTwd3xNI0MTSY0ikZQJhUDwjHLp7F7qsz5DvsWMqaNlDimCxsYLKc1RIYIJ"
    "OihaWotG3C5xnDImjbA4KvF5gdPICwouDgNM5do0eYy5WyK0Uqkr4ViBQtpT7r2ui7UrJnCgvYcbf1JPR3+GtI2kSsbU"
    "ryS6hMjHC4hVVkvEjxenaFlaVMtMi24ICSIVkmKcOpwqgVMCdTh1BM7hnIbfMSRE43MUBwkOlEOJwaJh8aQB1i6txZ97"
    "N7MWLeKK2V0MFCwiWsGtSqnS0rGEOuY0lZx2Go+TWJdjDEl2IT6nOOcIouuBamgDSkbalT2eiAMjiJOoFACYiEeakBIt"
    "GxdBSoTIWEdrd449b/Zwgflbuo8d5rnD9VR5Ac4lOFfxKctaOFbMsfL1kg7LKANXsg0hEUgSi+TvkgcAwHOqiXVERZ9Y"
    "CF0M1IWuMUkMV2nDKvAoeFY52hewt/5WLpxWZGdXke0dm6mvdvgJT4TGalUGF4Mvg64kkkbqoTHgxJ1lgkgF8CRBynZQ"
    "8JK6EvOwAlCFVCQkgzB4JLIfSPlcjDBSDJhUm2HF4pVo/ThmNHVRk9lMMQglzmmZgE6TDIgNHxG4pH2J3bOWdFhVYuE9"
    "FayM4nwUFRI9owKeY5RRG02A6GgSgoiLPZFLOM9QSlCwIgznCyw6fyaT66sRHJNqM1Rns/QODJHyLKPjj9itavRbZTTw"
    "mGLl9bkIoRvlBkuSoaOMYDSmSLhaAK8iWJBKxUyoCsFYhNGES5byuVEYiQhgxDCcLzB+XDX1tTUc6+3DWhsBCoFJKTJL"
    "6Hv0uww8vGQkNLAOoSblIkMoFXZAR3uIKFKMWZ1EYsq+tywGOsq4lK1u+ehiLzBG7KCqBM5x8cK5OFWe2roLawyT6mvJ"
    "F4vh/5HnOeX5xLXS/9ERdfQXhIUTBrho8kkGiqG7LcUrkXBWxAUxYI2JEydI4dGEpj0OMeMbTo2u4m9pApUosCgTJHaP"
    "hSCgqirN6kVzOXS8m6e27gJg2sR6Cn4QjuNC96RJ0BXHMrDQUCsnC4aPze/lmVuGePpLPl+5+Dj9RYORxNoS4Cq4LSG2"
    "0fmIKYehsSyb0m8NE8TSoCWCJP2yJiaOLHO+6DN1Yj1zZkxh5xvt7GxrB2DW1IYyVyOLXY7OksdYwpJH8NXw2YVHSV34"
    "Wbjg7/jc0h6yKcXXRH5AIoo8XX1ABIwBY/DGzupkzKRHRw0kY4SyxgjDhSLzG2cgIrzcupf2Y90ANE6bhIqEFjsRP1SG"
    "mpXWPJGooOp44uA01u//KYJlc1sdg0XD+IwjOCVcHpU9ytg1A29s8GOltJUEsSaKGiRUBU1IkR84Vi2YA8Arb+znxMAQ"
    "QeBonDop9AAaWeRkAeYtKjmhSlpcYQCv8eMwexIUB3hmqI2guB2yucivVkrzqan8qXOYMwMfJwjhUUToy1uGipYTI4Yg"
    "zvAidbGeZc2ieagqew4eZaToc7i7l1lTJpKryuBHQUBsoJLGKlaj5HURwVfHuJosN7xrKTJ+JTS8i5VzZ+BwiJHyGsWc"
    "MXgQzJmBT1BMYLBouWHZME996Qjf+VAv1pgozTUUA0d9bQ0XzT+fNw8f52jPSXzn2H+0iykT6hhfU00QR0GSUDUpFS0r"
    "zkUEYwxD+SIr5s7iwpnT2PLqHobzIyxtOg9jbSQhlWn4mYBHCI3gmYIXgXxgaKor8KNPGC675g5uvWkZn1neS1/eI+UZ"
    "8kWfpulTmFBbwwt/aGNwpIAi7D/SiWdDV1h0Lgp/pcSxSimLJC0yyMYY8n7A1WuWIcB3Nm/hQOdJFjdOp6YqUyboOwQf"
    "MfTMwCeviRH6Bofw+4fw+wZw0QLECIWiz7J5s1FVXmhpK6XP+452lVyhHzWfynOUdTYGkDTEvnPU1VTTvH4lJwaHeXL7"
    "a7Qf7+G8iXVMqhtHMXCYdyD2yTqheUfgEXIp2NNp+c3Jd2G7f8FgcTqPHGgk5xVwalCESxbORUTY8cYB0qkUIob2iAAz"
    "p0wkrEGcqlq+hq7LJuoE1hgGRgpcNG82MydN4OEXX2Wg5ySth44hIjROmUjBDzDmnYMvqcCZgI+5U3TKxBrLssu+gCz8"
    "DrXLvsLCppnkiz4KZKsyXLpoLgPDI+w9fJxMOoUxwsHO3tAVTm0YVREKVau/aBmXDjDAkJ9oSotQDBzXrl0GwM9+vwNS"
    "Hq/uPwzA/POm4rsg7Om8Q/AJI/j24AEk8vELmmYyZ0oNew50IDguX7UQP3AUA8e0hjrmz5rOK6+309U3gOdZPM9ytDfc"
    "qdI0bRJiTIkIoVH1uHnpSbb+1SF+86mDLGjIMxIYjAnBT6yt4aPrVnDsRB/P79lPuqaa1o5jACxtnF6OXEe5aWveGnzC"
    "Db49+HCxhkLRZ93SC0EM3/rJoxzt6ed9q5dSnaticCTP/NnTsdbwfEsbRT9ARPCsR3ffIMOFIo1TGkinvFLdMMAwLuOz"
    "ad0xpq39Z5a86yN8ddUhhnyPlDUM5AtcMr+JqfW1/Oy5nXT1DVKTzXCw8wT5os+SxumkU145pY5w9BbS9OTTkRsdGzxR"
    "lntG4OMQ01rL+1YvRVV5+LlX2PramyxsnEHT9Mn4I3lWzw8DoK2t+7CexUVxQf/wCIe7epk1eQLV2SqCRPbnJMOINxk6"
    "HsMde4XXerIYUzaAH123HIBfPP8qKc/iWUPPwBBtRzpZMHMKddU5fOciyRICNXx63kE+P39/RORE1MpYgdAZgBcj5P2A"
    "aQ31rF1yAbv3HeLo4U5+u7MVBVYvmAN+wNol8wD4Q3sHmXQKVcUaw3DBp/1YNw1146iryeEHjpRnOTEwxG0ffi9NV3yb"
    "w539mPq1dNZdC/4AhUCZUlfLtWuWcrCrl5fbDlBdlUEkVMVd7UcYn8syfeJ4CoHDMzAYpPj7VQf5z0+O498/2cC/rdnH"
    "kO9hjYxZTjdnAj7sJ4Z5/fILGslm0jz+wk4AntvdhgDvXrGA1LhqLrqwiX1HOuno7CWTSkWVn5CT+452Y0SYVFeLAicH"
    "R3jP8gVs+sQH2HtcufZnUylOvY4v/MU6spkU/SMF1i6cw4Rx1Tz0ux2cGBwm5dmSNL7afgSAOdMaQk9gLYFTVk/owJ94"
    "Nf7E61k3qZMqLyqenNJLkIQKvE3HRkQIAsd7Vi4C4NfbW7DVOVoPHOFI9wkuW3IBS+fNpn5cNS+17qV/JI/n2Qr12n8s"
    "dIXnTaqnkC9SP66a//jyJxERNt5zP1tffoHHX9zBivPPY/mcmRSG81y/fiUAm1/cRVU6Vcr7Pc/ScvAoAItmTkWBoUKR"
    "pin1zFr113gH7sNr/x6/6l3FYMFFEnBqA8ecWa9OCFTJZjO8/5KlDAyPsPPNg9RWZ+kdGOTX21tonNrAjZdfigIvte6L"
    "8vBypcYYW84KpzRQHBzm21+8nnkzJvMvP32cX217jWx9HT966iUAPrxmGbmaHB+6eDFtR7rYsa+DXFWmVApPpzz2HusB"
    "YFnTjGjZlh9+8WqaFn6Yb71xJQOz/4k1795I2hRxMnbAZ84EvDHCSMGncdpkFjXO4Hevvs6x3pOkUx4iwtPbWxARPn/V"
    "BgTY0XYgtMwJ6fI8w6GuEwDUZDNc94F1fHzDRfx21xvc+cDj1I/LUZVO89vdb3L8RD8fuXQZH1u/kuqqNA/9/pVQoqwp"
    "NUrSnsfRE/109Q2yZNY0/MDxzc9czbr5Tdz3zMt89f6tbNnrs2bOJJY1TmcwXwyDpbGM4Fjg4yTEGMF6lkLR55JFYYT3"
    "xMu7SyWzbCbDS617w7pfdZbOk/28cfg4Vel0ydcrkPI8uvoGCJzj2rXLuXvjjZwYGOaL3/sxxoSJVMqzdA8M8eMtLzNn"
    "WgP/+umrcU55bFtLKP6J8Txr6RvO09pxnMnja/jnm67kL997Mc/vaef2+/6Xqqzlv7e8AAjNa5YwXPAxcfid8HzmdOB9"
    "pwwMjzAwnGdgOE8wkueKixYD8OzOVjzPUggCUp6l7fBxtr/RTuAcr+49xJGePqw1pQ5T4BRjDF19Axzt7WP5nJlMqavl"
    "b77/IK8dPEY2k6YQBPiBI+15PPDsdvJFn+kTxrNjXwfb93ZQlU7hB65ULlMg7we8sq+Dmqo0/3Dd5RzqPsmn7nqAvB+Q"
    "y2V5evdeDvf28YnLljO1fhwF35VD8NMVRIwJffa7VyzgzltuwKnDiMF3jovnN3G05wRHek4yuX48IoK1luLJfl5+fT/r"
    "Fs+j9cARMimP2ups6JsTiU4xCOgfGoGJcP9vXuLnz++gccpEfKdkI92syWV580gXO/d1sPqC2TyxoxURYXx1tlRHIJKA"
    "fODY39kbMixwbLz3l/QMDjN5/DgQ6Owf4uFtrdxyxWoWz5rKs63t1HoZglLMJHFJrNLaO1VqclkWNs4gcA5jBOeU7r5B"
    "fvjYs+SLAXU1WQJVjBjGVWd57KVdfPDiJTy1vYWaXBWetRhjSgQw1jAy6PPMq69zcmiYf/yvzTSMr8Eag7HJENYwUvT5"
    "wZMvUJur4pcv7aa+Joc1tpQtx3apNlfF1jcP0drRyfeffJ5nX9vHlCg7FCNMqMlx75Zt/PSFXew6eIzqqnQFeBCk7oqb"
    "tbKUFE9gSKW8ivqfGEPf8AjWmDFVJpPyyPsB1tpEbl62unHhUkxI5LA/UFlui0XUOSWV8igEAZ6xo/b4lKUqcEo6ZRkp"
    "+lHcUS62xOvynSOXSVfOEY3ljQUeCd1eMV9IFD/DATOp1CmBkyJkvLCJkUmnEpOM7VVUICUGVxpmjKqUZ8LxUqmKtvzo"
    "tXpeWEbLZTIoiimV+cNyWpWNCDqK8/F6vLHAx4vxjDkNFzlli4wmevml/h6jdn0kzl10p5b6NYkydrIzxKjub4leUjFv"
    "mfOV1WRXsdxTmeG99bYURnVmeUe59pjSxdhbaf6YYsbYW2XGTu/HAk/UxhtC3mI31tsWTN4BeDlbwIcFV4UhI2I6MLa8"
    "FeOcBx82QMVYBOkwCtvEWAVxfxbgw3OH9RQx24yxdnNog6IG8rkOPnFBjWw24tKPuKDYITYlSLxv4JwG70ilRIvFDufM"
    "I6bn8bv6xNh7JJWJd/6cu+DDEpwz6ayocE/PXV/uE9hkJl5zsNoFmd3i2Vka+NErM+ca+Kj8mM4YDfwDLl1Y3J0bGjQ0"
    "t0j35nv7ndGbwzaVjXqT5xh4UKxRjOBEb+7+xtf6aVkklpYWpbnZ5n9+397M3BVdpir3IYJAo7es5FzhPMYi2ZzVkZHb"
    "ur57+4PxW6Rh1tLSomzY5I08+a0XM3NXdhkvfZVYK2jgh/vhjPxpGjxVkEDSGSvWihZHbuv87u13s2mTxz23BZVvjrZv"
    "cWzY5I088Y0Xsxde/LyKXCqZbEO4G04dItFuSZGzOsIrvToriJc2Jltt0KBNC8UbO797+wNs2uRxR/nV2dO+PD3hgxtr"
    "XS610Vg+J2Kbwi1rrrx97GzkvBHEWLBetK3O7VMxP3DFrrt67rrjDF6eHuP1+frmr403BFcpXKkSvj4vaHWpq3n2iL2q"
    "mEExcgi8bWLk0WI6/Ujv12+JXp8/FTzA/wEX75j0sTB4cgAAAABJRU5ErkJggg=="
)

_APP_ICON = [None]


def _app_icon():
    """返回 PhotoImage 形式的窗口图标（只创建一次）。"""
    if _APP_ICON[0] is None:
        try:
            import tkinter as _tk
            _APP_ICON[0] = _tk.PhotoImage(data=_APP_ICON_B64) or False
        except Exception:
            _APP_ICON[0] = False
    return _APP_ICON[0] or None


def _apply_icon(win):
    """给任意窗口（主窗口 / 对话框）套上应用图标。"""
    img = _app_icon()
    if img is None:
        return
    try:
        win.iconphoto(True, img)
    except Exception:
        pass


_TEXT_CLASSES = ("TLabel", "TButton", "TCheckbutton", "TRadiobutton", "TLabelFrame",
                 "TLabelframe", "Label", "Button", "Checkbutton", "Radiobutton",
                 "LabelFrame")


def localize_tree(widget):
    """登记并翻译控件树上的静态文案（窗口标题/标签/按钮/勾选框/分组框/表头）。"""
    if widget is None:
        return
    try:
        cls = widget.winfo_class()
    except Exception:
        return
    if cls in ("Tk", "Toplevel"):
        try:
            title = widget.title()
        except Exception:
            title = ""
        if title:
            key = ("t", id(widget))
            if key not in _LOC_SEEN:
                _LOC_SEEN.add(key)
                _LOC_TITLE.append((widget, title))
            try:
                widget.title(T(title))
            except Exception:
                pass
    try:
        if cls in _TEXT_CLASSES:
            text = widget.cget("text")
            if text:
                key = ("w", id(widget))
                if key not in _LOC_SEEN:
                    _LOC_SEEN.add(key)
                    _LOC_TEXT.append((widget, text))
                widget.configure(text=T(text))
        elif cls == "Treeview":
            for col in str(widget.cget("columns")).split():
                head = widget.heading(col, "text")
                if not head:
                    continue
                key = ("h", id(widget), col)
                if key not in _LOC_SEEN:
                    _LOC_SEEN.add(key)
                    _LOC_HEAD.append((widget, col, head))
                widget.heading(col, text=T(head))
        # textvariable 形式的静态文案（状态行等）：登记当前值，切换语言时若没被
        # 程序改写就一起翻译；被改写过（动态内容）则不动它。
        try:
            tvar = widget.cget("textvariable")
        except Exception:
            tvar = ""
        if tvar:
            cur = widget.getvar(tvar)
            if cur and isinstance(cur, str):
                key = ("v", id(widget))
                if key not in _LOC_SEEN:
                    _LOC_SEEN.add(key)
                    _LOC_TEXTVAR.append((widget, tvar, cur, [T(cur)]))
                widget.setvar(tvar, T(cur))
    except Exception:
        pass
    try:
        for child in widget.winfo_children():
            localize_tree(child)
    except Exception:
        pass


def localize_menu(menu):
    """登记并翻译菜单项（含级联子菜单）。"""
    if menu is None:
        return
    try:
        end = menu.index("end")
    except Exception:
        return
    if end is None:
        return
    for i in range(end + 1):
        try:
            if str(menu.type(i)) == "separator":
                continue
            label = menu.entrycget(i, "label")
        except Exception:
            label = ""
        if label:
            key = ("m", id(menu), i)
            if key not in _LOC_SEEN:
                _LOC_SEEN.add(key)
                _LOC_MENU.append((menu, i, label))
            try:
                menu.entryconfigure(i, label=T(label))
            except Exception:
                pass
        try:
            sub = menu.entrycget(i, "menu")
        except Exception:
            sub = ""
        if sub:
            try:
                localize_menu(menu.nametowidget(sub))
            except Exception:
                pass


def apply_language():
    """按当前语言把登记过的窗口标题/控件/菜单/表头文案重新套用一遍。"""
    for widget, cn in _LOC_TITLE:
        try:
            if widget.winfo_exists():
                widget.title(T(cn))
        except Exception:
            pass
    for widget, cn in _LOC_TEXT:
        try:
            if widget.winfo_exists():
                widget.configure(text=T(cn))
        except Exception:
            pass
    for widget, tvar, cn, last in _LOC_TEXTVAR:
        try:
            if not widget.winfo_exists():
                continue
            cur = widget.getvar(tvar)
            if cur == last[0]:          # 还是当初那串静态文案，才跟着语言走
                new = T(cn)
                widget.setvar(tvar, new)
                last[0] = new
        except Exception:
            pass
    for widget, col, cn in _LOC_HEAD:
        try:
            if widget.winfo_exists():
                widget.heading(col, text=T(cn))
        except Exception:
            pass
    for menu, i, cn in _LOC_MENU:
        try:
            menu.entryconfigure(i, label=T(cn))
        except Exception:
            pass


def _axis_ticks(xmin, xmax, step, start=None):
    try:
        step = float(step)
    except (TypeError, ValueError):
        step = 200.0
    if not step > 0:
        step = 200.0
    if start is None:
        first = math.ceil(xmin / step) * step
    else:
        first = float(start)
        if first < xmin - 1e-9:
            first += math.ceil((xmin - first) / step) * step
    count = int(math.floor((xmax - first) / step + 1e-9)) + 1
    if count <= 0:
        return [xmin]
    count = min(count, 400)
    return [first + i * step for i in range(count)]


def _moving_average(values, window):
    n = len(values)
    if window < 3 or n < window:
        return list(values)
    half = window // 2
    prefix = [0.0]
    for v in values:
        prefix.append(prefix[-1] + v)
    out = []
    for i in range(n):
        lo = max(0, i - half)
        hi = min(n, i + half + 1)
        out.append((prefix[hi] - prefix[lo]) / (hi - lo))
    return out


def _solve(a, b):
    size = len(b)
    for col in range(size):
        piv = max(range(col, size), key=lambda r: abs(a[r][col]))
        if abs(a[piv][col]) < 1e-12:
            return None
        if piv != col:
            a[col], a[piv] = a[piv], a[col]
            b[col], b[piv] = b[piv], b[col]
        for r in range(size):
            if r != col and a[r][col]:
                factor = a[r][col] / a[col][col]
                for c in range(col, size):
                    a[r][c] -= factor * a[col][c]
                b[r] -= factor * b[col]
    return [b[i] / a[i][i] for i in range(size)]


def _baseline_poly(values, degree=2):
    n = len(values)
    size = degree + 1
    if n <= size:
        return [0.0] * n
    a = [[0.0] * size for _ in range(size)]
    b = [0.0] * size
    for i in range(n):
        pows = [1.0]
        for _k in range(2 * degree):
            pows.append(pows[-1] * i)
        for j in range(size):
            b[j] += values[i] * pows[j]
            for k in range(size):
                a[j][k] += pows[j + k]
    coef = _solve(a, b)
    if coef is None:
        return [0.0] * n
    out = []
    for i in range(n):
        v = 0.0
        for k, c in enumerate(coef):
            v += c * (i ** k)
        out.append(v)
    return out


def _baseline_linear(values):
    n = len(values)
    if n < 4:
        return [0.0] * n
    k = max(2, n // 20)
    left = sum(values[:k]) / k
    right = sum(values[-k:]) / k
    step = (right - left) / (n - 1)
    return [left + step * i for i in range(n)]


def _rolling_min(values, window):
    n = len(values)
    window = max(3, int(window))
    if window % 2 == 0:
        window += 1
    if window >= n:
        window = n if n % 2 else n - 1
    half = window // 2
    block = window
    pre = [0.0] * n
    suf = [0.0] * n
    for start in range(0, n, block):
        end = min(n, start + block)
        cur = values[start]
        for i in range(start, end):
            if values[i] < cur:
                cur = values[i]
            pre[i] = cur
        cur = values[end - 1]
        for i in range(end - 1, start - 1, -1):
            if values[i] < cur:
                cur = values[i]
            suf[i] = cur
    out = [0.0] * n
    for i in range(n):
        lo = max(0, i - half)
        hi = min(n - 1, i + half)
        if lo // block == hi // block:
            out[i] = min(values[lo:hi + 1])
        else:
            out[i] = suf[lo] if suf[lo] < pre[hi] else pre[hi]
    return out


def _rolling_max(values, window):
    return [-v for v in _rolling_min([-v for v in values], window)]


def _median_filter(values, window):
    n = len(values)
    window = max(3, int(window))
    if window % 2 == 0:
        window += 1
    half = window // 2
    if n < window:
        return list(values)
    out = [0.0] * n
    for i in range(n):
        lo = max(0, i - half)
        hi = min(n, i + half + 1)
        seg = sorted(values[lo:hi])
        out[i] = seg[len(seg) // 2]
    return out


def remove_spikes(values, window=5, thresh=8.0, max_width=3, min_rel=0.15):
    """剔除尖峰（宇宙射线 / 坏点）。

    判定用两个条件同时约束，避免误伤真实拉曼峰：
      1) 与中值滤波的偏差超过 thresh × 噪声尺度（噪声用一阶差分稳健估计）；
      2) 该偏差还要达到邻域动态范围的一定比例（尖峰是突跳，窄峰峰顶的偏差
         相对整段信号的起伏很小）；
      3) 连续命中宽度不超过 max_width（真实峰即便很窄，偏离也是连续多点的）。
    命中的点用两侧邻点线性插值补回。
    """
    n = len(values)
    window = max(3, int(window))
    if window % 2 == 0:
        window += 1
    if n < window + 2:
        return list(values), []
    diffs = sorted(abs(values[i + 1] - values[i]) for i in range(n - 1))
    sigma = diffs[len(diffs) // 2] * 1.4826 / math.sqrt(2.0)
    span_all = abs(max(values) - min(values))
    sigma = max(sigma, span_all * 1e-9, 1e-12)
    limit = sigma * float(thresh)

    med = _median_filter(values, window)
    resid = [values[i] - med[i] for i in range(n)]
    wide = max(window * 4, 21)
    half = wide // 2
    cand = []
    for i in range(n):
        if abs(resid[i]) <= limit:
            continue
        lo = max(0, i - half)
        hi = min(n, i + half + 1)
        local = max(values[lo:hi]) - min(values[lo:hi])
        if local > 0 and abs(resid[i]) < float(min_rel) * local:
            continue
        cand.append(i)
    if not cand:
        return list(values), []

    runs = []
    start = prev = cand[0]
    for i in cand[1:]:
        if i == prev + 1:
            prev = i
            continue
        runs.append((start, prev))
        start = prev = i
    runs.append((start, prev))

    accepted = [(a, b) for a, b in runs if b - a + 1 <= max(1, int(max_width))]
    if not accepted:
        return list(values), []
    out = list(values)
    for a, b in accepted:
        if a == 0 or b >= n - 1:
            for i in range(a, b + 1):
                out[i] = med[i]
            continue
        y0 = values[a - 1]
        y1 = values[b + 1]
        span = float(b - a + 2)
        for k, i in enumerate(range(a, b + 1), 1):
            out[i] = y0 + (y1 - y0) * (k / span)
    hits = [i for a, b in accepted for i in range(a, b + 1)]
    return out, hits


def _poly_fit(values, degree, weights=None):
    """加权多项式拟合（横坐标归一化到 [-1,1]，避免高阶病态）。"""
    n = len(values)
    degree = max(0, min(int(degree), 12))
    size = degree + 1
    if n <= size + 1:
        return [0.0] * n
    t = [(2.0 * i / (n - 1)) - 1.0 for i in range(n)]
    a = [[0.0] * size for _ in range(size)]
    b = [0.0] * size
    for i in range(n):
        w = 1.0 if weights is None else float(weights[i])
        if w <= 0.0:
            continue
        pows = [1.0]
        for _k in range(2 * degree):
            pows.append(pows[-1] * t[i])
        for j in range(size):
            b[j] += w * values[i] * pows[j]
            for k in range(size):
                a[j][k] += w * pows[j + k]
    coef = _solve(a, b)
    if coef is None:
        return [0.0] * n
    out = [0.0] * n
    for i in range(n):
        ti = t[i]
        acc = 0.0
        for k in range(size - 1, -1, -1):
            acc = acc * ti + coef[k]
        out[i] = acc
    return out


def _baseline_iter_poly(values, degree=5, iters=20):
    """迭代多项式基线：反复拟合并剔除基线以上的点，最终得到贴着谷底的基线。"""
    n = len(values)
    if n < 12:
        return _baseline_poly(values, 2)
    degree = max(1, min(int(degree), 12))
    weights = [1.0] * n
    base = _poly_fit(values, degree, weights)
    prev = None
    for _ in range(max(1, min(int(iters), 200))):
        base = _poly_fit(values, degree, weights)
        kept = 0
        for i in range(n):
            if values[i] <= base[i]:
                weights[i] = 1.0
                kept += 1
            else:
                weights[i] = 0.0
        if kept < degree + 2:
            for i in range(n):
                weights[i] = 1.0
        if prev is not None and all(abs(base[i] - prev[i]) <= 1e-9 for i in range(n)):
            break
        prev = base
    return base


def _baseline_rolling_ball(values, radius_pts):
    """滚动球基线：形态学开运算（先滑动最小值再滑动最大值）近似球体下切。"""
    r = max(3, int(radius_pts))
    if r % 2 == 0:
        r += 1
    opened = _rolling_max(_rolling_min(values, r), r)
    return _moving_average(opened, max(3, r // 2 * 2 + 1))


def _savgol(values, window, order=2, deriv=0, step=1.0):
    n = len(values)
    window = int(window)
    if window % 2 == 0:
        window += 1
    order = max(0, min(int(order), 6))
    if window < order + 2:
        window = order + 2 + (1 - (order + 2) % 2)
    if window <= 2 or n < window:
        return list(values)
    half = window // 2
    size = order + 1
    a = [[0.0] * size for _ in range(size)]
    for i in range(-half, half + 1):
        pows = [1.0]
        for _k in range(2 * order):
            pows.append(pows[-1] * i)
        for j in range(size):
            for k in range(size):
                a[j][k] += pows[j + k]
    b = [0.0] * size
    if deriv <= order:
        fact = 1.0
        for k in range(1, deriv + 1):
            fact *= k
        b[deriv] = fact
    sol = _solve(a, b)
    if sol is None:
        return list(values)
    coef = [0.0] * window
    for i in range(-half, half + 1):
        t = float(i)
        acc = 0.0
        for k in range(size):
            acc += sol[k] * (t ** k)
        coef[i + half] = acc
    left = [values[0] + (values[0] - values[k]) * (half - j) / k
            for j, k in enumerate(range(1, half + 1))]
    right = [values[-1] + (values[-1] - values[-1 - k]) * (j + 1) / k
             for j, k in enumerate(range(1, half + 1))]
    ext = left + list(values) + right
    scale = (step ** deriv) if step else 1.0
    out = [0.0] * n
    for i in range(n):
        acc = 0.0
        for k in range(window):
            acc += coef[k] * ext[i + k]
        out[i] = acc / scale if scale else acc
    return out


def _calibrate(xs, plot=None):
    p = _plot_opts(plot)
    pairs = p.get("calib_pairs") or []
    if not pairs:
        return xs
    if len(pairs) == 1:
        measured, standard = float(pairs[0][0]), float(pairs[0][1])
        shift = standard - measured
        return [x + shift for x in xs]
    (m1, s1), (m2, s2) = pairs[0], pairs[1]
    m1, s1, m2, s2 = float(m1), float(s1), float(m2), float(s2)
    if abs(m2 - m1) < 1e-12:
        return list(xs)
    a = (s2 - s1) / (m2 - m1)
    b = s1 - a * m1
    return [a * x + b for x in xs]


def _process_signal(ys, plot=None, xs=None):
    p = _plot_opts(plot)
    out = list(ys)
    step = None
    if xs and len(xs) > 1:
        step = abs(xs[1] - xs[0]) or None

    if p.get("despike"):
        out = remove_spikes(out, p["despike_window"], p["despike_thresh"])[0]

    mode = p["baseline"]
    if mode == "linear":
        out = [v - b for v, b in zip(out, _baseline_linear(out))]
    elif mode == "poly2":
        out = [v - b for v, b in zip(out, _baseline_poly(out, 2))]
    elif mode == "iterpoly":
        out = [v - b for v, b in zip(
            out, _baseline_iter_poly(out, p["baseline_degree"], p["baseline_iters"]))]
    elif mode == "rolling" and step:
        win = max(3, int(round(abs(float(p["baseline_window"])) / step)))
        base = _moving_average(_rolling_min(out, win), win if win % 2 else win + 1)
        out = [v - b for v, b in zip(out, base)]
    elif mode == "rolling_ball" and step:
        win = max(3, int(round(abs(float(p["baseline_window"])) / step)))
        out = [v - b for v, b in zip(out, _baseline_rolling_ball(out, win))]

    w = int(p["smooth_window"] or 1)
    deriv = int(p["derivative"] or 0)
    if deriv:
        out = _savgol(out, max(w, 11), order=2, deriv=deriv, step=step or 1.0)
    elif w >= 3:
        if p["smooth_mode"] == "sg":
            out = _savgol(out, w, order=2)
        else:
            out = _moving_average(out, w)

    norm = p["normalize"]
    if norm == "max":
        m = max(out) if out else 0.0
        if m:
            out = [v / m for v in out]
    elif norm == "minmax":
        lo = min(out) if out else 0.0
        hi = max(out) if out else 1.0
        if hi > lo:
            out = [(v - lo) / (hi - lo) for v in out]
    return out


def _peak_fwhm(xs, ys, idx, base):
    peak = ys[idx]
    if peak <= base:
        return 0.0
    half = base + (peak - base) / 2.0
    n = len(ys)
    li = idx
    while li > 0 and ys[li] > half:
        li -= 1
    ri = idx
    while ri < n - 1 and ys[ri] > half:
        ri += 1
    if len(xs) > 1:
        return abs(xs[ri] - xs[li])
    return float(ri - li)


def _nearest_index(xs, x):
    n = len(xs)
    if n == 0:
        return 0
    if n > 2:
        step = xs[1] - xs[0]
        if step:
            guess = int(round((x - xs[0]) / step))
            if 0 <= guess < n:
                return guess
    return min(range(n), key=lambda i: abs(xs[i] - x))


def _peak_at(xs, ys, idx):
    n = len(ys)
    if n == 0:
        return {"index": 0, "x": 0.0, "y": 0.0, "base": 0.0,
                "prominence": 0.0, "fwhm": 0.0}
    idx = max(0, min(n - 1, idx))
    v = ys[idx]
    left_min = v
    j = idx - 1
    while j >= 0 and ys[j] <= v:
        if ys[j] < left_min:
            left_min = ys[j]
        j -= 1
    right_min = v
    j = idx + 1
    while j < n and ys[j] <= v:
        if ys[j] < right_min:
            right_min = ys[j]
        j += 1
    base = max(left_min, right_min)
    return {"index": idx, "x": xs[idx] if idx < len(xs) else float(idx),
            "y": v, "base": base, "prominence": v - base,
            "fwhm": _peak_fwhm(xs, ys, idx, base)}


def find_peaks(xs, ys, min_distance=20.0, window=9, thresh_pct=3.0):
    n = len(ys)
    if n < window + 2:
        return []
    sm = _moving_average(ys, window)
    resid = sorted(abs(ys[i] - sm[i]) for i in range(n))
    noise = resid[n // 2] * 1.4826
    span = max(sm) - min(sm)
    thresh = max(noise * 3.0, span * float(thresh_pct) / 100.0)
    if thresh <= 0:
        return []
    step = abs(xs[1] - xs[0]) if len(xs) > 1 else 1.0
    if not step:
        step = 1.0
    min_pts = max(3, int(round(abs(min_distance) / step)))
    cand = []
    for i in range(1, n - 1):
        if sm[i] >= sm[i - 1] and sm[i] > sm[i + 1]:
            cand.append(i)
    scored = []
    for i in cand:
        left_min = sm[i]
        j = i - 1
        while j >= 0 and sm[j] <= sm[i]:
            if sm[j] < left_min:
                left_min = sm[j]
            j -= 1
        right_min = sm[i]
        j = i + 1
        while j < n and sm[j] <= sm[i]:
            if sm[j] < right_min:
                right_min = sm[j]
            j += 1
        base = max(left_min, right_min)
        prom = sm[i] - base
        if prom >= thresh:
            scored.append((i, prom, base))
    scored.sort(key=lambda t: -t[1])
    kept = []
    for i, prom, base in scored:
        if all(abs(i - k) >= min_pts for k, _p, _b in kept):
            kept.append((i, prom, base))
    kept.sort()
    peaks = []
    for i, prom, base in kept:
        peaks.append({
            "index": i,
            "x": xs[i],
            "y": ys[i],
            "base": base,
            "prominence": prom,
            "fwhm": _peak_fwhm(xs, ys, i, base),
        })
    return peaks


def analyze_peaks(xs, ys, plot=None, processed=False):
    p = _plot_opts(plot)
    sig = ys if processed else _process_signal(ys, p, xs)
    peaks = find_peaks(xs, sig, p["peak_min_dist"], thresh_pct=p["peak_thresh_pct"])
    for pk in peaks:
        pk["manual"] = False
    # 用户删掉的自动峰：按峰位抹掉。容差取“最小峰间距”的一半，与手动峰
    # 合并的判据一致 —— 该间距下自动峰之间至少隔这么远，只可能命中一个。
    hidden = p.get("hidden_peaks") or []
    if hidden:
        tol = max(1.0, abs(float(p["peak_min_dist"])) / 2.0)
        peaks = [pk for pk in peaks
                 if not any(abs(pk["x"] - float(hv)) <= tol for hv in hidden)]
    manual = p.get("manual_peaks") or []
    if manual:
        dup = max(1.0, abs(float(p["peak_min_dist"])) / 2.0)
        step = abs(xs[1] - xs[0]) if len(xs) > 1 else 0.0
        w = max(1, int(round(dup / step))) if step else 0
        for mv in manual:
            idx = _nearest_index(xs, float(mv))
            if w:
                lo = max(0, idx - w)
                hi = min(len(sig), idx + w + 1)
                if hi > lo:
                    idx = max(range(lo, hi), key=lambda i: sig[i])
            px = xs[idx] if idx < len(xs) else float(idx)
            if any(abs(px - pk["x"]) <= dup for pk in peaks):
                continue
            pk = _peak_at(xs, sig, idx)
            if pk["prominence"] <= 0 and step:
                w2 = max(3, int(round(dup / step)))
                lo2 = max(0, idx - w2)
                hi2 = min(len(sig), idx + w2 + 1)
                if hi2 > lo2:
                    base2 = min(sig[lo2:hi2])
                    pk["base"] = base2
                    pk["prominence"] = sig[idx] - base2
                    pk["fwhm"] = _peak_fwhm(xs, sig, idx, base2)
            pk["manual"] = True
            peaks.append(pk)
        peaks.sort(key=lambda pk: pk["x"])
    strongest = max((pk["prominence"] for pk in peaks), default=0.0)
    for pk in peaks:
        pk["rel"] = (100.0 * pk["prominence"] / strongest) if strongest > 0 else 0.0
    return peaks


class JwsError(Exception):
    pass


class _DirEntry:
    __slots__ = ("index", "name", "obj_type", "left", "right", "child", "start", "size")

    def __init__(self, index, name, obj_type, left, right, child, start, size):
        self.index = index
        self.name = name
        self.obj_type = obj_type
        self.left = left
        self.right = right
        self.child = child
        self.start = start
        self.size = size


class CompoundFile:
    def __init__(self, blob):
        self.data = blob
        self._read_header()
        self._build_fat()
        self._read_directory()

    def _read_header(self):
        d = self.data
        if len(d) < 512 or d[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
            raise JwsError(T("不是有效的 OLE2/复合文档，可能不是 .jws 文件"))
        self.sector_size = 1 << struct.unpack_from("<H", d, 0x1E)[0]
        self.mini_size = 1 << struct.unpack_from("<H", d, 0x20)[0]
        self.dir_start = struct.unpack_from("<I", d, 0x2C)[0]
        self.mini_cutoff = struct.unpack_from("<I", d, 0x38)[0]
        self.minifat_start = struct.unpack_from("<I", d, 0x3C)[0]
        self.difat_start = struct.unpack_from("<I", d, 0x44)[0]
        self.num_difat = struct.unpack_from("<I", d, 0x48)[0]
        self._base_difat = list(struct.unpack_from("<109I", d, 0x4C))

    def _sector(self, n):
        off = (n + 1) * self.sector_size
        return self.data[off:off + self.sector_size]

    def _fat_sectors(self):
        out = [x for x in self._base_difat if x != FREE_SECT]
        sec = self.difat_start
        per = self.sector_size // 4
        for _ in range(self.num_difat):
            if sec in (END_OF_CHAIN, FREE_SECT):
                break
            block = self._sector(sec)
            if len(block) < self.sector_size:
                break
            vals = list(struct.unpack_from("<%dI" % per, block, 0))
            out.extend(x for x in vals[:-1] if x != FREE_SECT)
            sec = vals[-1]
        return out

    def _build_fat(self):
        per = self.sector_size // 4
        fat = []
        for s in self._fat_sectors():
            block = self._sector(s)
            if len(block) < self.sector_size:
                break
            fat.extend(struct.unpack_from("<%dI" % per, block, 0))
        self.fat = fat

    def _chain(self, start):
        out = []
        s = start
        seen = set()
        while s not in (END_OF_CHAIN, FREE_SECT) and s < len(self.fat):
            if s in seen:
                break
            seen.add(s)
            out.append(s)
            s = self.fat[s]
        return out

    def _locate_directory(self):
        ss = self.sector_size
        nsec = len(self.data) // ss - 1
        for s in range(nsec + 1):
            base = (s + 1) * ss
            if base + 128 > len(self.data):
                break
            if self.data[base + 66] != 5:
                continue
            nlen = struct.unpack_from("<H", self.data, base + 64)[0]
            if nlen < 2 or nlen > 64:
                continue
            try:
                name = self.data[base:base + nlen - 2].decode("utf-16le")
            except UnicodeDecodeError:
                continue
            if name == "Root Entry":
                return s
        return self.dir_start

    def _read_directory(self):
        head = self._locate_directory()
        raw = b"".join(self._sector(s) for s in self._chain(head))
        self.entries = []
        for i in range(len(raw) // 128):
            chunk = raw[i * 128:(i + 1) * 128]
            nlen = struct.unpack_from("<H", chunk, 64)[0]
            name = ""
            if nlen >= 2:
                name = chunk[:nlen - 2].decode("utf-16le", "replace")
            self.entries.append(_DirEntry(
                i, name, chunk[66],
                struct.unpack_from("<I", chunk, 68)[0],
                struct.unpack_from("<I", chunk, 72)[0],
                struct.unpack_from("<I", chunk, 76)[0],
                struct.unpack_from("<I", chunk, 116)[0],
                struct.unpack_from("<Q", chunk, 120)[0],
            ))
        roots = [e for e in self.entries if e.obj_type == 5]
        if not roots:
            raise JwsError(T("未找到 Root Entry，文件结构异常"))
        self.root = roots[0]
        self.ministream = b"".join(
            self._sector(s) for s in self._chain(self.root.start)
        )[:self.root.size]
        self.minifat = []
        if self.minifat_start not in (END_OF_CHAIN, FREE_SECT):
            mraw = b"".join(self._sector(s) for s in self._chain(self.minifat_start))
            self.minifat = list(struct.unpack_from("<%dI" % (len(mraw) // 4), mraw, 0))

    def _mini_chain(self, start):
        out = []
        s = start
        seen = set()
        while s not in (END_OF_CHAIN, FREE_SECT) and s < len(self.minifat):
            if s in seen:
                break
            seen.add(s)
            out.append(s)
            s = self.minifat[s]
        return out

    def read_stream(self, entry):
        if entry.size < self.mini_cutoff:
            blob = b"".join(
                self.ministream[s * self.mini_size:(s + 1) * self.mini_size]
                for s in self._mini_chain(entry.start)
            )
            return blob[:entry.size]
        blob = b"".join(self._sector(s) for s in self._chain(entry.start))
        return blob[:entry.size]

    def _walk(self, idx, out):
        if idx in (FREE_SECT, END_OF_CHAIN) or idx >= len(self.entries):
            return
        e = self.entries[idx]
        self._walk(e.left, out)
        out.append(e)
        self._walk(e.right, out)

    def children(self, parent):
        out = []
        self._walk(parent.child, out)
        return out


def _decode_utf16(blob):
    if not blob:
        return ""
    return blob.decode("utf-16le", "replace").split("\x00")[0].strip()


class Spectrum:
    def __init__(self, path):
        self.path = path
        with open(path, "rb") as f:
            blob = f.read()
        self.cf = CompoundFile(blob)
        self.streams = {}
        for e in self.cf.children(self.cf.root):
            if e.obj_type == 2 and e.name not in self.streams:
                self.streams[e.name] = e
        self.sample = self._sample_name()
        self._load_data()

    def _read(self, name):
        e = self.streams.get(name)
        return None if e is None else self.cf.read_stream(e)

    def _sample_name(self):
        si = self._read("SampleInfo")
        if not si or len(si) < 8:
            return None
        blen = struct.unpack_from("<I", si, 4)[0]
        if blen <= 0 or 8 + blen > len(si):
            blen = len(si) - 8
        return _decode_utf16(si[8:8 + blen]) or None

    def _load_data(self):
        di = self._read("DataInfo")
        if di is None or len(di) < 48:
            raise JwsError(T("缺少 DataInfo 流，无法确定横坐标轴"))
        channel_number = struct.unpack_from("<I", di, 12)[0]
        npoints = struct.unpack_from("<I", di, 20)[0]
        start = struct.unpack_from("<d", di, 24)[0]
        end = struct.unpack_from("<d", di, 32)[0]
        step = struct.unpack_from("<d", di, 40)[0]
        self.x_type = struct.unpack_from("<I", di, 48)[0]
        self.y_types = []
        for i in range(min(max(channel_number, 1), 3)):
            off = 52 + 4 * i
            self.y_types.append(
                struct.unpack_from("<I", di, off)[0] if off + 4 <= len(di) else None)

        yd = self._read("Y-Data")
        if yd is None:
            raise JwsError(T("缺少 Y-Data 流，文件里没有光谱数据"))
        count = len(yd) // 4
        raw = list(struct.unpack_from("<%df" % count, yd, 0))

        channels = max(1, channel_number)
        if npoints > 0 and count % npoints == 0:
            channels = count // npoints
        else:
            npoints = count
            channels = 1

        self.npoints = npoints
        self.channel_number = channels
        self.start = start
        self.end = end
        self.step = self._resolve_step(start, end, step, npoints)
        if npoints > 0:
            self.end = start + (npoints - 1) * self.step
        self.y_data = [
            raw[i * npoints:(i + 1) * npoints] for i in range(channels)
        ]
        self.y = self.y_data[0] if self.y_data else []

    @staticmethod
    def _resolve_step(start, end, step, n):
        if n > 1:
            if step == step and step not in (0.0, float("inf"), float("-inf")):
                return step
            return (end - start) / (n - 1)
        return 1.0

    def x_values(self):
        s = self.start
        d = self.step
        return [s + i * d for i in range(self.npoints)]

    def x_value(self, i):
        return self.start + i * self.step

    def axis_name(self, default=_DEFAULT_X_HEADER):
        return _DATA_TYPES.get(self.x_type, default)

    def auto_names_supported(self):
        return self.x_type in _DATA_TYPES

    def y_names(self, default=_DEFAULT_Y_HEADER):
        if self.channel_number == 1:
            base = [_DATA_TYPES.get(self.y_types[0]) if self.y_types else None]
        else:
            base = [_DATA_TYPES.get(c) if c is not None else None for c in self.y_types]
            if len(base) < self.channel_number:
                base += [None] * (self.channel_number - len(base))
        names = []
        for i in range(self.channel_number):
            det = base[i] if i < len(base) else None
            if det:
                names.append(det)
            elif self.channel_number == 1:
                names.append(default)
            else:
                names.append("%s %d" % (default, i + 1))
        return names


_SKIP_DIR_NAMES = {"工具数据", "分析结果", "_转换结果", "参考谱库", "RRUFF数据包",
                   "__pycache__", ".git"}
_SKIP_FILE_SUFFIXES = ("_peaks.csv", "_fit.csv", "_批处理汇总.csv", "_数据库比对.csv",
                       "_配对报告.csv")


def _is_aux_dir(name):
    return name in _SKIP_DIR_NAMES or name.startswith("_报告图")


def find_input_files(paths):
    """收集光谱文件。

    规则：你显式指定的目录与文件一律算数；从它再往下递归时，
    跳过工具自身的输出目录（_转换结果 / 工具数据 / 分析结果 …）与
    衍生表（峰列表、拟合表、汇总表、配对报告），
    避免它们被当成光谱再次参与分析。
    """
    exts = (".jws", ".csv", ".spc", ".jdx")
    direct_exts = (".jws", ".csv", ".spc", ".jdx", ".txt", ".dat", ".asc", ".xy",
                   ".xlsx", ".xlsm")
    result = []
    for p in paths:
        p = os.path.abspath(p)
        if os.path.isdir(p):
            top = os.path.normcase(p)
            for root, dirs, names in os.walk(p):
                direct = os.path.normcase(os.path.abspath(root)) == top
                dirs[:] = [d for d in dirs if not _is_aux_dir(d)]
                for n in sorted(names):
                    low = n.lower()
                    if not low.endswith(exts):
                        continue
                    if not direct and any(low.endswith(sfx)
                                          for sfx in _SKIP_FILE_SUFFIXES):
                        continue
                    result.append(os.path.join(root, n))
        elif p.lower().endswith(direct_exts):
            result.append(p)
    seen = set()
    uniq = []
    for p in result:
        key = os.path.normcase(p)
        if key not in seen:
            seen.add(key)
            uniq.append(p)
    return uniq


def _column_names(spec, x_header, y_header, auto_names):
    if auto_names and spec.auto_names_supported():
        return [spec.axis_name(x_header)] + spec.y_names(y_header)
    names = [x_header]
    if spec.channel_number == 1:
        names.append(y_header)
    else:
        names.extend("%s %d" % (y_header, i + 1) for i in range(spec.channel_number))
    return names


def _export_xy(spec, plot):
    p = _plot_opts(plot)
    xs = _calibrate(spec.x_values(), p)
    if p["apply_to_data"]:
        chans = [_process_signal(spec.y_data[c], p, xs) for c in range(spec.channel_number)]
    else:
        chans = [list(spec.y_data[c]) for c in range(spec.channel_number)]
    return xs, chans


def _write_csv(dst, spec, names, write_header, plot=None):
    xs, chans = _export_xy(spec, plot)
    lines = []
    if write_header:
        lines.append(",".join(names))
    for i in range(len(xs)):
        row = ["%.6f" % xs[i]]
        for c in range(spec.channel_number):
            row.append("%.6f" % chans[c][i])
        lines.append(",".join(row))
    with open(dst, "w", encoding="utf-8", newline="") as f:
        f.write("\n".join(lines))
        f.write("\n")


def _write_xlsx(dst, spec, names, write_header, plot=None):
    if not _HAVE_XLSX:
        raise JwsError(T("导出 Excel 需要 openpyxl 组件；请改用 CSV，或先安装 openpyxl"))
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    if write_header:
        ws.append(names)
    ch = spec.channel_number
    xs, chans = _export_xy(spec, plot)
    for i in range(len(xs)):
        row = [xs[i]]
        for c in range(ch):
            row.append(chans[c][i])
        ws.append(row)
    first = 2 if write_header else 1
    last = first + len(xs) - 1
    chart = ScatterChart()
    chart.scatterStyle = "line"
    chart.title = os.path.splitext(os.path.basename(dst))[0]
    chart.style = 13
    p = _plot_opts(plot)
    step = float(p["x_step"])
    x0 = min(xs)
    x1 = max(xs)
    x_ref = Reference(ws, min_col=1, min_row=first, max_row=last)
    for i in range(ch):
        y_ref = Reference(ws, min_col=2 + i, min_row=first, max_row=last)
        label = names[1 + i] if 1 + i < len(names) else "Y%d" % (i + 1)
        chart.series.append(Series(y_ref, x_ref, title=label if write_header else None))
    chart.x_axis.title = names[0]
    chart.y_axis.title = names[-1] if ch == 1 else names[1]
    chart.x_axis.majorUnit = step
    chart.x_axis.scaling.min = math.floor(x0 / step) * step
    chart.x_axis.scaling.max = math.ceil(x1 / step) * step
    chart.x_axis.number_format = "0"
    chart.x_axis.delete = False
    chart.y_axis.delete = not p["show_y_ticks"]
    chart.width = 26
    chart.height = 15
    ws.add_chart(chart, "H2")
    wb.save(dst)


def _png_font(size):
    fonts_dir = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
    for name in ("msyh.ttc", "msyh.ttf", "simhei.ttf", "simsun.ttc", "arial.ttf"):
        try:
            return ImageFont.truetype(os.path.join(fonts_dir, name), size)
        except Exception:
            continue
    try:
        return ImageFont.load_default()
    except Exception:
        return None


def _png_text(draw, font, xy, text, anchor, fill=(70, 70, 70)):
    try:
        draw.text(xy, text, font=font, fill=fill, anchor=anchor)
    except Exception:
        draw.text(xy, text, font=font, fill=fill)


def _png_measure(draw, font):
    """拿一个“量文字宽度”的小函数（量不到就按汉语/拉丁混排估个大概）。

    中英混排时按 1 个字符 9 px 估是够用的：估宽了最多让图例条带宽一点，
    估窄了才会压到谱线，所以宁可估宽。
    """
    def measure(text):
        try:
            return float(draw.textlength(text, font=font))
        except Exception:
            return 9.0 * len(text)
    return measure


def _png_dashed_v(draw, x, y_from, y_to, color, dash=7, gap=5, width=1):
    """在 x 处画一条 y_from → y_to 的竖直虚线，用来把峰位波长引到横坐标轴上。"""
    lo, hi = (y_from, y_to) if y_from <= y_to else (y_to, y_from)
    y = lo
    while y < hi:
        ye = min(y + dash, hi)
        draw.line([(x, y), (x, ye)], fill=color, width=width)
        y = ye + gap


def _merge_tol(p):
    """峰位合并容差：peak_merge_tol 未设时沿用“最小峰间距”。"""
    tol = p.get("peak_merge_tol")
    if tol is None:
        tol = abs(float(p["peak_min_dist"]))
    return max(1e-9, abs(float(tol)))


def peak_clusters(series, plot=None):
    """把多条谱的峰位并到一起，邻近的归为同一个峰。

    返回按波数升序排列的列表，每项：
        {"x": 簇内平均波数, "xs": [各成员峰位], "n": 成员数,
         "curves": 参与的数据集序号集合, "manual": 是否全是手动峰}

    判据：与簇内已有序号的**均值**相差不超过容差就并入，否则另起一簇。
    用均值而不是“与上一个成员比较”，是为了避免链式漂移把一整段峰串成一簇。
    容差取 peak_merge_tol，未设时沿用“最小峰间距”。
    """
    p = _plot_opts(plot)
    tol = _merge_tol(p)

    found = []
    for k, item in enumerate(series):
        xs, ys = item[1], item[2]
        if not xs or not ys:
            continue
        for pk in analyze_peaks(xs, ys, p, processed=True):
            found.append((pk["x"], k, bool(pk.get("manual"))))
    found.sort(key=lambda t: t[0])

    clusters = []
    for x, k, manual in found:
        if clusters and abs(x - clusters[-1]["x"]) <= tol:
            c = clusters[-1]
            c["xs"].append(x)
            c["x"] = sum(c["xs"]) / len(c["xs"])
            c["n"] = len(c["xs"])
            c["curves"].add(k)
            c["manual"] = c["manual"] and manual
        else:
            clusters.append({"x": x, "xs": [x], "n": 1, "curves": {k},
                             "manual": manual})
    return clusters


def overlay_color(k):
    """叠加图配色：前 8 条用标准色板，之后按黄金角旋转色相生成新色，保证彼此可区分。"""
    n = len(_PALETTE)
    if k < n:
        return _PALETTE[k]
    hexs = _PALETTE[k % n].lstrip("#")
    r, g, b = (int(hexs[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    h, _l, _s = colorsys.rgb_to_hls(r, g, b)
    rounds = k // n
    hue = (h + 0.6180339887498949 * rounds) % 1.0
    sat = 0.48 + 0.22 * (rounds % 2)
    lig = (0.64, 0.42, 0.32, 0.55)[rounds % 4]
    r2, g2, b2 = colorsys.hls_to_rgb(hue, lig, sat)
    return "#%02x%02x%02x" % (int(r2 * 255), int(g2 * 255), int(b2 * 255))


# ---------------------------------------------------------------------------
# 图例摆放
#   以前图例画在绘图区里的右上角，谱线一密（尤其堆叠图，曲线从下到上铺满整幅）
#   就压在谱线上；加白底卡片也只是“盖住但读得清”，谱线本身还是被吃掉了。
#   现在按 legend_pos 摆：
#     right（默认）绘图区右侧的留白里竖排，永不压谱线；放不下就自动分列；
#     top / bottom 绘图区上/下的留白里横排；
#     inside       图内自由位置（legend_xy，比例坐标），白底卡片，可在预览窗口里拖；
#     none         不画图例。
#   条目不再截成“只列前 12 条”——谱线多少条就列多少条，空间不够才分列，
#   实在放不下才省略，并明确写出“还有 N 条没列”。
# ---------------------------------------------------------------------------
LEGEND_ROW_H = 28          # 每行高度
LEGEND_SWATCH = 32         # 色块（短线）长度
LEGEND_TEXT_GAP = 10       # 色块与文字间距
LEGEND_PAD = 10            # 图例块内边距
LEGEND_COL_GAP = 24        # 两列之间的间距
LEGEND_GUTTER_GAP = 16     # 绘图区与图例留白之间的间距
LEGEND_LABEL_MAX = 30      # 标签最多显示几个字
LEGEND_LABEL_MIN = 10      # 空间不够时最多截到几个字
LEGEND_MAX_COLUMNS = 4     # 最多分几列
LEGEND_GUTTER_SHARE = 1.0 / 3.0   # 右侧条带最多占整幅宽的 1/3
LEGEND_STRIP_SHARE = 0.34         # 上/下条带最多占绘图区高的 34%
LEGEND_POSITIONS = ("right", "top", "bottom", "inside", "none")
# legend_xy 是图例框左上角在绘图区里的比例坐标（0=左/上，1=右/下）
LEGEND_XY_DEFAULT = (0.62, 0.05)
# bottom 位置要把图例压到横坐标标题下面，这段是不含图例的最小下边距
LEGEND_BOTTOM_LABELS = 106
LEGEND_MB_BASE = 120       # 渲染器原本的下边距，算 bottom 要加多少时用


def legend_label(label, limit=LEGEND_LABEL_MAX):
    """图例标签：去换行、去过长空白，再截断。"""
    text = " ".join(str(label or "").split())
    if len(text) > limit:
        text = text[:limit - 1] + "…"
    return text or "(未命名)"


def legend_entries(series, measure, max_text_width=None, chars=LEGEND_LABEL_MAX):
    """图例条目 [(文字, 颜色, 文字宽度)]。

    几条谱就给几条图例（不再限 12 条）；chars 控制标签最多几个字，
    max_text_width 再按像素宽继续截，保证图例不把绘图区挤得太窄。
    """
    rows = []
    for item in series:
        label, color = item[0], item[3]
        text = legend_label(label, chars)
        if max_text_width:
            while len(text) > LEGEND_LABEL_MIN and measure(text) > max_text_width:
                text = text[:-2] + "…"
        rows.append((text, color, measure(text)))
    return rows


def legend_cell_width(entries):
    """一格图例（色块 + 文字 + 内边距）有多宽。"""
    if not entries:
        return 0
    widest = max(w for _t, _c, w in entries)
    return LEGEND_SWATCH + LEGEND_TEXT_GAP + widest + LEGEND_PAD * 2


def legend_reserve(entries, pos, plot_w, plot_h):
    """top / bottom 位置要从绘图区**外面**占掉多少空间。

    返回 (右侧, 上方, 下方, avail)。avail 是排图例时真正可用的 (宽, 高)，
    要原样交给 legend_place——否则两处各算一遍行数会对不上（图例跑到画布外）。
    right 的预留由 legend_fit 自己算（要兼顾标签截断与分列）。
    """
    if not entries or pos not in ("top", "bottom"):
        return 0, 0, 0, (0, 0)
    cell = legend_cell_width(entries)
    cols = max(1, min(LEGEND_MAX_COLUMNS,
                      int((plot_w + LEGEND_COL_GAP) // (cell + LEGEND_COL_GAP)) or 1))
    rows = -(-len(entries) // cols)
    # 条带最多占绘图区高的一部分，剩下的条目交给“还有 N 条”那行
    rows = min(rows, max(1, int(plot_h * LEGEND_STRIP_SHARE // LEGEND_ROW_H)))
    box_h = rows * LEGEND_ROW_H + 2 * LEGEND_PAD
    if pos == "top":
        # 图例排在标题之下、绘图区之上，整条往上撑
        return 0, box_h + LEGEND_GUTTER_GAP, 0, (plot_w, rows * LEGEND_ROW_H)
    # 图例排在横坐标标题之下：先把标题那一段让出来，再放图例
    add = max(0, LEGEND_BOTTOM_LABELS + box_h + LEGEND_GUTTER_GAP - LEGEND_MB_BASE)
    return 0, 0, add, (plot_w, rows * LEGEND_ROW_H)


def legend_place(entries, pos, ml, mt, pw, ph, xy=None, avail=None):
    """把图例条目排到具体坐标上。

    返回 (rows, hidden, box)：
      rows   [(文字, 颜色, 色块起点 x, 行中线 y)]
      hidden 没排下的条数
      box    图例外框 (x0, y0, x1, y1)；inside 模式画白底、也用它判定拖动
      box 为 None 表示不画。
    """
    if not entries or pos == "none":
        return [], len(entries), None
    cell = legend_cell_width(entries)
    n = len(entries)
    avail_w, avail_h = avail if avail else (0, 0)

    def pack(box_w, box_h, order):
        cols = max(1, min(LEGEND_MAX_COLUMNS,
                          int((box_w + LEGEND_COL_GAP) // (cell + LEGEND_COL_GAP)) or 1))
        rows_max = max(1, int(box_h // LEGEND_ROW_H))
        if order == "col":
            rows = min(rows_max, n)
            cols = min(cols, -(-n // rows))
            used = min(n, rows * cols)
            slots = [(i % rows, i // rows) for i in range(used)]
        else:
            cols = min(cols, n)
            rows = min(rows_max, -(-n // cols))
            used = min(n, rows * cols)
            slots = [(i // cols, i % cols) for i in range(used)]
        return slots, n - len(slots)

    if pos == "right":
        slots, hidden = pack(avail_w or pw, avail_h or ph, "col")
        x0 = ml + pw + LEGEND_GUTTER_GAP
        y0 = mt
    elif pos in ("top", "bottom"):
        slots, hidden = pack(avail_w or pw, avail_h or ph, "row")
        rows_n = (max(r for r, _c in slots) + 1) if slots else 1
        box_h = rows_n * LEGEND_ROW_H + 2 * LEGEND_PAD
        x0 = ml
        y0 = (mt - LEGEND_GUTTER_GAP - box_h if pos == "top"
              else mt + ph + LEGEND_BOTTOM_LABELS)   # bottom 排在横坐标标题下面
    else:                                   # inside：图内自由位置
        slots, hidden = pack(pw - 2 * LEGEND_PAD, ph - 2 * LEGEND_PAD, "col")
        fx, fy = (xy or LEGEND_XY_DEFAULT)[:2]
        cols = (max(c for _r, c in slots) + 1) if slots else 1
        rows_n = (max(r for r, _c in slots) + 1) if slots else 1
        bw = cols * cell + (cols - 1) * LEGEND_COL_GAP + 2 * LEGEND_PAD
        bh = rows_n * LEGEND_ROW_H + 2 * LEGEND_PAD
        x0 = min(max(ml + 4, ml + float(fx) * pw), max(ml + 4, ml + pw - bw - 4))
        y0 = min(max(mt + 4, mt + float(fy) * ph), max(mt + 4, mt + ph - bh - 4))

    rows_out = []
    for idx, (r, c) in enumerate(slots):
        text, color, _w = entries[idx]
        cx = x0 + LEGEND_PAD + c * (cell + LEGEND_COL_GAP)
        cy = y0 + LEGEND_PAD + r * LEGEND_ROW_H + LEGEND_ROW_H // 2
        rows_out.append((text, color, cx, cy))
    if not rows_out:
        return [], n, None
    cols = max(c for _r, c in slots) + 1
    rows_n = max(r for r, _c in slots) + 1
    box = (x0, y0,
           x0 + cols * cell + (cols - 1) * LEGEND_COL_GAP + 2 * LEGEND_PAD,
           y0 + rows_n * LEGEND_ROW_H + 2 * LEGEND_PAD)
    return rows_out, hidden, box


def legend_fit(series, measure, pos, width, plot_w, plot_h):
    """给渲染器用的“一步到位”接口：算条目、算要预留的空间、算排版可用空间。

    返回 (entries, reserve=(右,上,下), avail)，渲染器把尺寸定下来之后
    再拿 avail 调 legend_place。

    right 位置会尽量让**所有**谱线都有图例：先按最长标签试，条带超过整幅宽的
    1/3 就把标签一点点截短、必要时再加一列，直到都能放下（实在不行才省略条目）。
    """
    if len(series) < 2 or pos == "none":
        return [], (0, 0, 0), (0, 0)
    if pos != "right":
        entries = legend_entries(series, measure)
        right, top, bottom, avail = legend_reserve(entries, pos, plot_w, plot_h)
        return entries, (right, top, bottom), avail

    n = len(series)
    cap = max(LEGEND_SWATCH + LEGEND_TEXT_GAP + 2 * LEGEND_PAD + LEGEND_LABEL_MIN * 9,
              int(width * LEGEND_GUTTER_SHARE))
    rows_fit = max(1, int(plot_h // LEGEND_ROW_H))
    best = None
    for chars in range(LEGEND_LABEL_MAX, LEGEND_LABEL_MIN - 1, -1):
        entries = legend_entries(series, measure, chars=chars)
        cell = legend_cell_width(entries)
        cols = min(LEGEND_MAX_COLUMNS, -(-n // rows_fit))
        while cols > 1 and (cols * cell + (cols - 1) * LEGEND_COL_GAP
                            + LEGEND_GUTTER_GAP > cap):
            cols -= 1
        best = (entries, cols, cell)
        if cols * rows_fit >= n:       # 所有条目都排得下
            break
    entries, cols, cell = best
    box = cols * cell + (cols - 1) * LEGEND_COL_GAP
    return entries, (box + LEGEND_GUTTER_GAP, 0, 0), (box, plot_h)


def legend_defaults():
    """图例位置的全局默认（写在设置文件里，下次打开还是这个位置）。

    返回 (pos, xy)；xy 是比例坐标 [fx, fy] 或 None。
    """
    data = _load_settings()
    pos = str(data.get("legend_pos") or "").strip().lower()
    if pos not in LEGEND_POSITIONS:
        pos = "right"
    xy = None
    raw = str(data.get("legend_xy") or "").strip()
    if raw:
        try:
            parts = [float(v) for v in raw.replace(";", ",").split(",")]
            if len(parts) >= 2:
                xy = [min(1.0, max(0.0, parts[0])), min(1.0, max(0.0, parts[1]))]
        except ValueError:
            xy = None
    return pos, xy


def save_legend_defaults(pos, xy=None):
    """把图例位置写回设置文件（在预览窗口里拖完就存，下次打开还是这里）。"""
    data = _load_settings()
    data["legend_pos"] = str(pos if pos in LEGEND_POSITIONS else "right")
    if xy:
        data["legend_xy"] = "%.4f,%.4f" % (float(xy[0]), float(xy[1]))
    else:
        data.pop("legend_xy", None)
    return _save_settings(data)


def legend_hit(box, x, y):
    """鼠标 (x, y) 是否点在图例外框里（用来判断要不要开始拖动）。"""
    if not box:
        return False
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def draw_legend(d, font, pos, rows, hidden, box, text_fill=(40, 40, 40)):
    """画图例。inside 模式加一层白底卡片和淡边框，免得压在谱线上读不清。"""
    if not rows or not box:
        return
    inside = pos == "inside"
    if inside:
        d.rectangle(list(box), fill=(255, 255, 255), outline=(216, 216, 216))
    for text, color, cx, cy in rows:
        d.line([(cx, cy), (cx + LEGEND_SWATCH, cy)], fill=color, width=4)
        _png_text(d, font, (cx + LEGEND_SWATCH + LEGEND_TEXT_GAP,
                            cy - 9), text, "la", fill=text_fill)
    if hidden > 0:
        tail = T("…还有 %d 条没列（把图例位置改到右侧留白，或把图放大）") % hidden
        _png_text(d, font, (box[0] + LEGEND_PAD, box[3] - LEGEND_ROW_H + 4), tail, "la",
                  fill=(150, 60, 60))


def render_png(path, series, title, xlabel, ylabel, plot=None, width=None, height=None,
               return_geometry=False):
    if not _HAVE_PIL:
        raise JwsError(T("导出 PNG 需要 Pillow 组件（pip install pillow）"))
    series = [s for s in series if s[1] and s[2]]
    if not series:
        raise JwsError(T("没有可绘制的数据"))
    p = _plot_opts(plot)
    width = int(width or p["fig_width"] or 1600)
    height = int(height or p["fig_height"] or 900)
    series = [(lab, xs, _process_signal(ys, p, xs), col) for lab, xs, ys, col in series]
    # 堆叠排布：每条按偏移量纵向错开，而不是压在同一基线上（stacked spectra 画法）。
    # 归一化到最大值 = 1 后再错开，各条互不压线，谱型仍可横向比较。
    if p["stacked"]:
        step = max(0.0, float(p["stack_offset"] or 0.0))
        if step:
            series = [(lab, xs, [v + k * step for v in ys], col)
                      for k, (lab, xs, ys, col) in enumerate(series)]
    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f_title = _png_font(32)
    f_lab = _png_font(24)
    f_tick = _png_font(18)

    ml = 140 if p["show_y_ticks"] else 105
    mt_base, mb_base = 90, 120
    mt, mb = mt_base, mb_base
    base_pw = width - ml - 60
    base_ph = height - mt - mb

    # 图例先算好，再从绘图区**外面**留出空间；画在外面，谱线再密也压不到。
    legend_pos = p["legend_pos"] if len(series) > 1 else "none"
    entries, (l_right, l_top, l_bottom), l_avail = legend_fit(
        series, _png_measure(d, f_tick), legend_pos, width, base_pw, base_ph)

    mr = max(60, l_right)
    mt += l_top
    mb += l_bottom
    pw = width - ml - mr
    ph = height - mt - mb

    plot_xmin = min(min(s[1]) for s in series)
    plot_xmax = max(max(s[1]) for s in series)
    plot_ymin = min(min(s[2]) for s in series)
    plot_ymax = max(max(s[2]) for s in series)
    if p["x_min"] is not None:
        plot_xmin = float(p["x_min"])
    if p["x_max"] is not None:
        plot_xmax = float(p["x_max"])
    auto_y = p["y_min"] is None and p["y_max"] is None
    if p["y_min"] is not None:
        plot_ymin = float(p["y_min"])
    if p["y_max"] is not None:
        plot_ymax = float(p["y_max"])
    if plot_xmax <= plot_xmin:
        plot_xmax = plot_xmin + 1.0
    if plot_ymax <= plot_ymin:
        plot_ymax = plot_ymin + 1.0
    if auto_y:
        pad = (plot_ymax - plot_ymin) * 0.06
        plot_ymin -= pad
        plot_ymax += pad

    def sx(x):
        return ml + (x - plot_xmin) / (plot_xmax - plot_xmin) * pw

    def sy(y):
        return mt + ph - (y - plot_ymin) / (plot_ymax - plot_ymin) * ph

    def inside(a, b):
        return plot_xmin <= a <= plot_xmax and plot_ymin <= b <= plot_ymax

    ticks = _axis_ticks(plot_xmin, plot_xmax, p["x_step"], p["x_start"])

    if p["show_grid"]:
        for i in range(1, 5):
            gy = mt + ph * i / 5
            d.line([(ml, gy), (ml + pw, gy)], fill=(238, 238, 238))
        for xv in ticks:
            gx = sx(xv)
            d.line([(gx, mt), (gx, mt + ph)], fill=(238, 238, 238))

    for xv in ticks:
        gx = sx(xv)
        d.line([(gx, mt + ph), (gx, mt + ph + 6)], fill=(120, 120, 120))
        _png_text(d, f_tick, (gx, mt + ph + 12), "%.6g" % xv, "ma")

    if p["show_y_ticks"]:
        for i in range(6):
            gy = mt + ph - ph * i / 5
            yv = plot_ymin + (plot_ymax - plot_ymin) * i / 5
            d.line([(ml - 6, gy), (ml, gy)], fill=(120, 120, 120))
            _png_text(d, f_tick, (ml - 12, gy), "%.6g" % yv, "rm")

    d.line([(ml, mt), (ml, mt + ph)], fill=(90, 90, 90), width=2)
    d.line([(ml, mt + ph), (ml + pw, mt + ph)], fill=(90, 90, 90), width=2)

    for label, xs, ys, color in series:
        dx, dy = _decimate(xs, ys, int(pw))
        seg = []
        for a, b in zip(dx, dy):
            if not inside(a, b):
                if len(seg) >= 2:
                    d.line(seg, fill=color, width=2, joint="curve")
                seg = []
                continue
            seg.append((sx(a), sy(b)))
        if len(seg) >= 2:
            d.line(seg, fill=color, width=2, joint="curve")

    if p["annotate_peaks"]:
        # 跨谱合并：同一个峰在每条谱上都标一遍会糊成一片，合并后只画一条虚线、
        # 只标一个平均波数；此时各条谱自己只保留峰位标记（圆点 / 方块）。
        merged = bool(p["merge_peak_labels"]) and len(series) > 1
        marks = []
        tol = 0.0
        if merged:
            tol = _merge_tol(p)
            raw = p.get("peak_marks")
            if raw is None:
                marks = [{"x": c["x"], "manual": c["manual"]}
                         for c in peak_clusters(series, p)]
            else:
                marks = [{"x": float(m["x"]), "manual": bool(m.get("manual"))}
                         for m in raw]
        for _label, xs, ys, _color in series:
            for k, pk in enumerate(analyze_peaks(xs, ys, p, processed=True)):
                if not inside(pk["x"], pk["y"]):
                    continue
                # 这一簇被用户删掉了，连峰位标记也一起收走，免得留下没有虚线的孤点
                if merged and not any(abs(pk["x"] - m["x"]) <= tol for m in marks):
                    continue
                gx = sx(pk["x"])
                gy = sy(pk["y"])
                mark_color = (20, 80, 170) if pk.get("manual") else (170, 30, 30)
                # 峰位波长虚线：从峰顶一直引到横坐标轴（自动峰和手动峰一视同仁）
                if p["peak_dash_line"] and not merged:
                    _png_dashed_v(d, gx, gy, mt + ph, mark_color)
                if pk.get("manual"):
                    d.rectangle([gx - 4, gy - 4, gx + 4, gy + 4], fill=(30, 110, 200))
                else:
                    d.ellipse([gx - 3, gy - 3, gx + 3, gy + 3], fill=(200, 40, 40))
                text = ("%g  %.0f%%" % (round(pk["x"], 1), pk["rel"])
                        if p["peak_label_rel"] else "%g" % round(pk["x"], 1))
                ly = gy - (12, 32, 52)[k % 3]
                if p["peak_labels"] and not merged:
                    _png_text(d, f_tick, (gx, ly), text, "mb", fill=mark_color)

        if merged:
            for j, mk in enumerate(marks):
                gx = sx(mk["x"])
                if gx < ml - 2 or gx > ml + pw + 2:
                    continue
                mark_color = (20, 80, 170) if mk["manual"] else (170, 30, 30)
                # 标签压在图上缘，靠错位台阶让靠近的峰标签不叠字
                ty = mt + 6 + (j % 3) * 22
                if p["peak_dash_line"]:
                    _png_dashed_v(d, gx, ty + 24, mt + ph, mark_color)
                if p["peak_labels"]:
                    _png_text(d, f_tick, (gx, ty), "%g" % round(mk["x"], 1),
                              "mt", fill=mark_color)

    legend_rows, legend_hidden, legend_box = legend_place(
        entries, legend_pos, ml, mt, pw, ph, xy=p["legend_xy"], avail=l_avail)
    if legend_rows or legend_hidden:
        # 画在绘图区外面（right / top / bottom）；inside 时是用户自己拖的位置，带白底
        draw_legend(d, f_tick, legend_pos, legend_rows, legend_hidden, legend_box)

    if title and p["show_title"]:
        # top 位置时图例占了标题下面那一带，标题留在最上面不跟图例叠
        title_y = mt_base - 55 if (legend_pos == "top" and legend_rows) else mt - 55
        _png_text(d, f_title, (ml + pw / 2, title_y), title, "ma", fill=(30, 30, 30))
    if xlabel:
        _png_text(d, f_lab, (ml + pw / 2, mt + ph + 70), xlabel, "ma", fill=(40, 40, 40))
    if ylabel:
        tmp = Image.new("RGBA", (500, 44), (255, 255, 255, 0))
        _png_text(ImageDraw.Draw(tmp), f_lab, (250, 22), ylabel, "mm", fill=(40, 40, 40))
        tmp = tmp.rotate(90, expand=True)
        img.paste(tmp, (12, int(mt + ph / 2 - tmp.height / 2)), tmp)

    img.save(path, "PNG")
    if return_geometry:
        # 把绘图区的实际位置回报给调用方：交互式预览靠它把鼠标坐标换算成波数，
        # 也靠 legend_box / legend_pos 判断“点的是不是图例”，好让用户拖着放。
        return {"xmin": plot_xmin, "xmax": plot_xmax, "ml": ml, "pw": pw,
                "mt": mt, "ph": ph, "width": width, "height": height,
                "legend_pos": legend_pos, "legend_box": legend_box,
                "legend_rows": len(legend_rows)}
    return None


def render_pair_report(path, target_name, target_xy, results, plot=None, tol=5.0,
                       width=1700, height=1180, return_geometry=False):
    if not _HAVE_PIL:
        raise JwsError(T("导出报告图需要 Pillow 组件（pip install pillow）"))
    if not results:
        raise JwsError(T("没有配对结果可绘制"))
    p = _plot_opts(plot)
    best = results[0]
    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f_title = _png_font(38)
    f_sub = _png_font(20)
    f_h = _png_font(22)
    f_lab = _png_font(22)
    f_tick = _png_font(17)
    f_mono = _png_font(18)

    txs, tys = target_xy
    tsig = _norm_max(_process_signal(tys, p, txs))
    ref = best.get("ref_xy")
    rsig = rxs = None
    if ref:
        rxs = ref[0]
        rsig = _norm_max(_process_signal(ref[1], p, ref[0]))

    _png_text(d, f_title, (width / 2, 40), T("拉曼光谱配对对比报告"), "mm", fill=(25, 25, 25))
    _png_text(d, f_sub, (width / 2, 84),
              T("实测：%s        最佳匹配：%s        F1 %.0f%%    r=%.4f    谱角 %.1f°")
              % (target_name, best["name"], best["peak_score"], best["corr"], best["angle"]),
              "mm", fill=(80, 80, 80))

    # 图例同样挪到绘图区外面：这两条线（实测 / 参考）铺满整幅，
    # 压在图上必然被谱线穿过，读起来费劲。
    legend_source = [(T("实测 ") + target_name, "", "", (200, 40, 40))]
    if rsig is not None:
        legend_source.append((T("参考 ") + best["name"], "", "", (30, 110, 200)))
    legend_pos = p["legend_pos"]
    entries, (l_right, l_top, l_bottom), l_avail = legend_fit(
        legend_source, _png_measure(d, f_tick), legend_pos, width, width - 110 - 40, 510)
    left, top, bottom = 110, 130 + l_top, 640
    right = width - max(40, l_right)
    pw, ph = right - left, bottom - top
    xmin = min(txs) if rxs is None else min(min(txs), min(rxs))
    xmax = max(txs) if rxs is None else max(max(txs), max(rxs))
    ymin = min(tsig) if rsig is None else min(min(tsig), min(rsig))
    ymax = max(tsig) if rsig is None else max(max(tsig), max(rsig))
    if xmax <= xmin:
        xmax = xmin + 1.0
    if ymax <= ymin:
        ymax = ymin + 1.0
    pad = (ymax - ymin) * 0.06
    ymin -= pad
    ymax += pad

    def sx(x):
        return left + (x - xmin) / (xmax - xmin) * pw

    def sy(y):
        return bottom - (y - ymin) / (ymax - ymin) * ph

    ticks = _axis_ticks(xmin, xmax, p["x_step"], p["x_start"])
    for i in range(1, 5):
        gy = top + ph * i / 5
        d.line([(left, gy), (right, gy)], fill=(240, 240, 240))
    for xv in ticks:
        gx = sx(xv)
        d.line([(gx, top), (gx, bottom)], fill=(240, 240, 240))
        d.line([(gx, bottom), (gx, bottom + 6)], fill=(120, 120, 120))
        _png_text(d, f_tick, (gx, bottom + 12), "%.6g" % xv, "ma")
    d.line([(left, top), (left, bottom)], fill=(90, 90, 90), width=2)
    d.line([(left, bottom), (right, bottom)], fill=(90, 90, 90), width=2)
    _png_text(d, f_lab, ((left + right) / 2, bottom + 60), p.get("xlabel") or _DEFAULT_X_HEADER,
              "ma", fill=(40, 40, 40))

    # 峰位配对连线
    if rsig is not None:
        for pk in analyze_peaks(txs, tsig, p, processed=True):
            gx = sx(pk["x"])
            d.line([(gx, top), (gx, bottom)], fill=(255, 220, 220))
        for row in best.get("pair_rows", []):
            if row[3] and row[1] is not None:
                d.line([(sx(row[0]), top), (sx(row[1]), bottom)], fill=(190, 220, 255), width=2)

    dxs, dys = _decimate(txs, tsig, int(pw))
    d.line([(sx(a), sy(b)) for a, b in zip(dxs, dys)], fill=(200, 40, 40), width=3, joint="curve")
    if rsig is not None:
        rdx, rdy = _decimate(rxs, rsig, int(pw))
        d.line([(sx(a), sy(b)) for a, b in zip(rdx, rdy)], fill=(30, 110, 200), width=3,
               joint="curve")

    for k, pk in enumerate(analyze_peaks(txs, tsig, p, processed=True)):
        gx, gy = sx(pk["x"]), sy(pk["y"])
        d.ellipse([gx - 4, gy - 4, gx + 4, gy + 4], fill=(200, 40, 40))
        _png_text(d, f_tick, (gx, gy - (10, 28)[k % 2]), "%.6g" % round(pk["x"], 1), "mb",
                  fill=(170, 30, 30))
    if rsig is not None:
        for k, pk in enumerate(analyze_peaks(rxs, rsig, p, processed=True)):
            gx, gy = sx(pk["x"]), sy(pk["y"])
            d.rectangle([gx - 4, gy - 4, gx + 4, gy + 4], fill=(30, 110, 200))
            _png_text(d, f_tick, (gx, gy + 10), "%.6g" % round(pk["x"], 1), "mt",
                      fill=(20, 80, 170))

    legend_rows, legend_hidden, legend_box = legend_place(
        entries, legend_pos, left, top, pw, ph, xy=p["legend_xy"], avail=l_avail)
    if legend_rows or legend_hidden:
        draw_legend(d, f_tick, legend_pos, legend_rows, legend_hidden, legend_box)

    # 排名表（bottom 位置时图例插在图上和表之间，两张表一起往下让开，别只挪一张）
    table_top = bottom + 110 + l_bottom
    ty = table_top
    _png_text(d, f_h, (60, ty), T("配对排名（共 %d 条参考谱，按峰位匹配 F1 排序）") % len(results),
              "la", fill=(25, 25, 25))
    ty += 34
    head = "%-4s %-34s %7s %6s %10s %9s %8s" % (
        T("排名"), T("参考谱"), "F1(%)", T("命中"), T("实测/参考"), T("相关系数"), T("谱角"))
    _png_text(d, f_mono, (60, ty), head, "la", fill=(90, 90, 90))
    ty += 26
    for i, r in enumerate(results[:8], 1):
        line = "%-4d %-34s %7.1f %6d %5d/%-4d %9.4f %8.2f" % (
            i, r["name"][:34], r["peak_score"], r["peak_matched"],
            r["peak_left"], r["peak_right"], r["corr"], r["angle"])
        _png_text(d, f_mono, (60, ty), line, "la",
                  fill=(20, 90, 40) if i == 1 else (60, 60, 60))
        ty += 26

    # 峰位对照表
    rows = best.get("pair_rows", [])
    py = table_top + 34 + 26 * (min(len(results), 8) + 1) + 26
    _png_text(d, f_h, (60, py), T("最佳配对峰位对照（容差 %.1f cm-1）：%s") % (tol, best["name"]),
              "la", fill=(25, 25, 25))
    py += 32
    col_head = "%-10s %-10s %-9s %-8s" % (T("实测峰位"), T("参考峰位"), T("偏差"), T("是否匹配"))
    for k in (0, 1):
        _png_text(d, f_mono, (60 + k * 800, py), col_head, "la", fill=(90, 90, 90))
    py += 26
    half = (len(rows) + 1) // 2
    for idx in range(half):
        for k, base in enumerate((0, half)):
            j = idx + base
            if j >= len(rows):
                continue
            row = rows[j]
            _png_text(d, f_mono, (60 + k * 800, py),
                      "%-10s %-10s %-9s %-8s" % (
                          "%.2f" % row[0] if row[0] is not None else "-",
                          "%.2f" % row[1] if row[1] is not None else "-",
                          "%+.2f" % row[2] if row[2] is not None else "-",
                          T("匹配") if row[3] else T("单侧")),
                      "la", fill=(60, 60, 60))
        py += 26
    img.save(path, "PNG")
    if return_geometry:
        # 自测用：把图例框与两张表的位置报出去，好断言它们互不压
        return {"legend_box": legend_box, "legend_rows": len(legend_rows),
                "legend_pos": legend_pos, "chart": (left, top, right, bottom),
                "table_top": table_top, "pair_top": py,
                "width": width, "height": height}
    return None


_CMAP = [(0.00, (49, 54, 149)), (0.25, (69, 117, 180)), (0.50, (116, 196, 118)),
         (0.75, (244, 180, 66)), (1.00, (215, 48, 39))]

_CLUSTER_COLORS = [(31, 119, 180), (214, 39, 40), (44, 160, 44), (255, 127, 14),
                   (148, 103, 189), (140, 86, 75), (227, 119, 194), (23, 190, 207)]


def _cmap(t):
    t = max(0.0, min(1.0, float(t)))
    for i in range(len(_CMAP) - 1):
        x0, c0 = _CMAP[i]
        x1, c1 = _CMAP[i + 1]
        if t <= x1:
            f = (t - x0) / (x1 - x0) if x1 > x0 else 0.0
            return tuple(int(round(c0[k] + (c1[k] - c0[k]) * f)) for k in range(3))
    return _CMAP[-1][1]


def _cluster_color(cid):
    return _CLUSTER_COLORS[(max(1, int(cid)) - 1) % len(_CLUSTER_COLORS)]


def render_heatmap(path, rows, title="", cbar_label="", plot=None, width=None, height=None,
                   row_labels=None, col_labels=None, vmin=None, vmax=None, show_values=True):
    """二维成像热图：rows 为逐行数值（None 表示无数据）。"""
    if not _HAVE_PIL:
        raise JwsError(T("导出 PNG 需要 Pillow 组件（pip install pillow）"))
    rows = [list(r) for r in rows if r]
    if not rows:
        raise JwsError(T("没有可绘制的数据"))
    p = _plot_opts(plot)
    width = int(width or p["fig_width"] or 1200)
    height = int(height or p["fig_height"] or 780)
    nr = len(rows)
    nc = max(len(r) for r in rows)
    vals = [v for r in rows for v in r if v is not None]
    if not vals:
        raise JwsError(T("没有可绘制的数据"))
    lo = min(vals) if vmin is None else float(vmin)
    hi = max(vals) if vmax is None else float(vmax)
    if hi <= lo:
        hi = lo + 1.0

    ml = 170 if row_labels else 90
    mr = 250
    mt = 100
    mb = 130
    pw = max(60, width - ml - mr)
    ph = max(60, height - mt - mb)
    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f_title = _png_font(32)
    f_tick = _png_font(17)
    f_cell = _png_font(16)
    cw = pw / float(nc)
    ch = ph / float(nr)

    for r in range(nr):
        for c in range(nc):
            x0 = ml + c * cw
            y0 = mt + r * ch
            v = rows[r][c] if c < len(rows[r]) else None
            if v is None:
                d.rectangle([x0, y0, x0 + cw - 1, y0 + ch - 1], fill=(236, 236, 236))
                continue
            t = (v - lo) / (hi - lo)
            d.rectangle([x0, y0, x0 + cw - 1, y0 + ch - 1], fill=_cmap(t))
            if show_values and cw >= 52 and ch >= 24:
                _png_text(d, f_cell, (x0 + cw / 2.0, y0 + ch / 2.0), "%.4g" % v, "mm",
                          fill=(255, 255, 255) if t > 0.62 else (30, 30, 30))

    for c in range(nc + 1):
        gx = ml + c * cw
        d.line([(gx, mt), (gx, mt + ph)], fill=(255, 255, 255))
    for r in range(nr + 1):
        gy = mt + r * ch
        d.line([(ml, gy), (ml + pw, gy)], fill=(255, 255, 255))
    d.rectangle([ml, mt, ml + pw, mt + ph], outline=(120, 120, 120), width=2)

    if row_labels:
        for r in range(nr):
            label = row_labels[r] if r < len(row_labels) else str(r + 1)
            _png_text(d, f_tick, (ml - 12, mt + (r + 0.5) * ch), str(label)[:20], "rm",
                      fill=(60, 60, 60))
    if col_labels:
        for c in range(nc):
            label = col_labels[c] if c < len(col_labels) else str(c + 1)
            _png_text(d, f_tick, (ml + (c + 0.5) * cw, mt + ph + 14), str(label)[:20], "ma",
                      fill=(60, 60, 60))
    else:
        for c in range(nc):
            _png_text(d, f_tick, (ml + (c + 0.5) * cw, mt + ph + 14), str(c + 1), "ma",
                      fill=(60, 60, 60))

    cbx = ml + pw + 70
    cbw = 36
    for i in range(200):
        t = 1.0 - i / 199.0
        y0 = mt + ph * i / 200.0
        y1 = mt + ph * (i + 1) / 200.0
        d.rectangle([cbx, y0, cbx + cbw, y1], fill=_cmap(t))
    d.rectangle([cbx, mt, cbx + cbw, mt + ph], outline=(110, 110, 110), width=2)
    for i in range(5):
        gy = mt + ph - ph * i / 4.0
        yv = lo + (hi - lo) * i / 4.0
        d.line([(cbx + cbw, gy), (cbx + cbw + 8, gy)], fill=(90, 90, 90))
        _png_text(d, f_tick, (cbx + cbw + 14, gy), "%.5g" % yv, "lm", fill=(60, 60, 60))
    if cbar_label:
        _png_text(d, f_tick, (cbx + cbw / 2.0, mt - 16), cbar_label, "mb", fill=(60, 60, 60))

    if title:
        _png_text(d, f_title, (ml + pw / 2.0, mt - 58), title, "ma", fill=(30, 30, 30))
    img.save(path, "PNG")


def render_dendrogram(path, labels, merges, title="", cut=None, groups=None, plot=None,
                      width=None, height=None):
    """层次聚类树状图（叶子在左、距离在右，便于显示长文件名）。"""
    if not _HAVE_PIL:
        raise JwsError(T("导出 PNG 需要 Pillow 组件（pip install pillow）"))
    n = len(labels)
    if n < 2 or not merges:
        raise JwsError(T("至少需要 2 条光谱才能聚类"))
    width = int(width or 1200)
    height = int(height or max(560, 210 + 34 * n))
    mt, mb = 100, 170
    ml, mr = 300, 110
    ph = max(80, height - mt - mb)
    pw = max(80, width - ml - mr)
    maxh = max(m[2] for m in merges) or 1.0

    members = {}
    for k, (a, b, h, c) in enumerate(merges):
        members[n + k] = (a, b)
    order = []

    def walk(cid):
        if cid < n:
            order.append(cid)
            return
        a, b = members[cid]
        walk(a)
        walk(b)

    walk(n + len(merges) - 1)
    pos = {leaf: i for i, leaf in enumerate(order)}
    total = max(1, len(order))

    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f_title = _png_font(30)
    f_lab = _png_font(18)
    f_tick = _png_font(16)

    node_x = {i: ml for i in range(n)}
    node_y = {i: mt + (pos[i] + 0.5) * ph / total for i in range(n)}

    def hx(h):
        return ml + (h / maxh) * pw

    for k, (a, b, h, c) in enumerate(merges):
        nid = n + k
        x0, x1 = node_x[a], node_x[b]
        y0, y1 = node_y[a], node_y[b]
        xm = hx(h)
        color = (120, 120, 120)
        d.line([(x0, y0), (xm, y0)], fill=color, width=2)
        d.line([(x1, y1), (xm, y1)], fill=color, width=2)
        d.line([(xm, y0), (xm, y1)], fill=color, width=2)
        node_x[nid] = xm
        node_y[nid] = (y0 + y1) / 2.0

    for i in range(n):
        y = node_y[i]
        cid = groups[i] if groups else 1
        d.rectangle([16, y - 7, 30, y + 7], fill=_cluster_color(cid))
        _png_text(d, f_lab, (ml - 16, y), str(labels[i])[:36], "rm", fill=(50, 50, 50))
        d.line([(ml - 10, y), (ml, y)], fill=(170, 170, 170))

    for i in range(5):
        gx = hx(maxh * i / 4.0)
        d.line([(gx, mt), (gx, mt + ph)], fill=(240, 240, 240))
        d.line([(gx, mt + ph), (gx, mt + ph + 7)], fill=(110, 110, 110))
        _png_text(d, f_tick, (gx, mt + ph + 14), "%.3f" % (maxh * i / 4.0), "ma",
                  fill=(60, 60, 60))
    d.line([(ml, mt), (ml, mt + ph)], fill=(90, 90, 90), width=2)
    d.line([(ml, mt + ph), (ml + pw, mt + ph)], fill=(90, 90, 90), width=2)
    _png_text(d, f_tick, (ml + pw / 2.0, mt + ph + 48), T("聚类距离（1 − 相关系数）"), "ma",
              fill=(40, 40, 40))

    if cut is not None:
        cx = hx(min(float(cut), maxh))
        for yy in range(int(mt), int(mt + ph), 16):
            d.line([(cx, yy), (cx, min(yy + 8, mt + ph))], fill=(200, 60, 60), width=2)
        _png_text(d, f_tick, (cx + 6, mt - 12), T("自动分割阈值 %.3f") % float(cut), "lm",
                  fill=(180, 40, 40))

    if groups:
        counts = {}
        for i in range(n):
            cid = groups[i]
            counts[cid] = counts.get(cid, 0) + 1
        tx = ml
        ty = mt + ph + 82
        _png_text(d, f_tick, (tx, ty), T("簇："), "lm", fill=(60, 60, 60))
        tx += 40
        for cid in sorted(counts):
            d.rectangle([tx, ty - 7, tx + 14, ty + 7], fill=_cluster_color(cid))
            _png_text(d, f_tick, (tx + 20, ty), T("簇%d（%d 条）") % (cid, counts[cid]), "lm",
                      fill=(60, 60, 60))
            tx += 130

    if title:
        _png_text(d, f_title, (ml + pw / 2.0, mt - 60), title, "ma", fill=(30, 30, 30))
    img.save(path, "PNG")


def render_scatter(path, points, title="", xlabel="PC1", ylabel="PC2", plot=None,
                   groups=None, labels=None, width=None, height=None):
    """主成分散点图：points 为 (x, y) 列表。"""
    if not _HAVE_PIL:
        raise JwsError(T("导出 PNG 需要 Pillow 组件（pip install pillow）"))
    pts = [(float(a), float(b)) for a, b in points if a is not None and b is not None]
    if not pts:
        raise JwsError(T("没有可绘制的数据"))
    p = _plot_opts(plot)
    width = int(width or p["fig_width"] or 1100)
    height = int(height or p["fig_height"] or 820)
    ml, mr, mt, mb = 150, 330, 100, 140
    pw = max(60, width - ml - mr)
    ph = max(60, height - mt - mb)
    xs = [a for a, _b in pts]
    ys = [b for _a, b in pts]
    xmin, xmax = min(xs), max(xs)
    ymin, ymax = min(ys), max(ys)
    if xmax <= xmin:
        xmax = xmin + 1.0
    if ymax <= ymin:
        ymax = ymin + 1.0
    padx = (xmax - xmin) * 0.12
    pady = (ymax - ymin) * 0.12
    xmin -= padx
    xmax += padx
    ymin -= pady
    ymax += pady

    img = Image.new("RGB", (width, height), (255, 255, 255))
    d = ImageDraw.Draw(img)
    f_title = _png_font(32)
    f_lab = _png_font(22)
    f_tick = _png_font(17)
    f_pt = _png_font(15)

    def sx(x):
        return ml + (x - xmin) / (xmax - xmin) * pw

    def sy(y):
        return mt + ph - (y - ymin) / (ymax - ymin) * ph

    for i in range(5):
        gy = mt + ph * i / 4.0
        d.line([(ml, gy), (ml + pw, gy)], fill=(238, 238, 238))
        gx = ml + pw * i / 4.0
        d.line([(gx, mt), (gx, mt + ph)], fill=(238, 238, 238))

    for i in range(7):
        gx = ml + pw * i / 6.0
        xv = xmin + (xmax - xmin) * i / 6.0
        d.line([(gx, mt + ph), (gx, mt + ph + 7)], fill=(110, 110, 110))
        _png_text(d, f_tick, (gx, mt + ph + 14), "%.4g" % xv, "ma", fill=(60, 60, 60))
        gy = mt + ph - ph * i / 6.0
        yv = ymin + (ymax - ymin) * i / 6.0
        d.line([(ml - 7, gy), (ml, gy)], fill=(110, 110, 110))
        _png_text(d, f_tick, (ml - 14, gy), "%.4g" % yv, "rm", fill=(60, 60, 60))

    d.line([(ml, mt), (ml, mt + ph)], fill=(90, 90, 90), width=2)
    d.line([(ml, mt + ph), (ml + pw, mt + ph)], fill=(90, 90, 90), width=2)
    if xmin < 0 < xmax:
        d.line([(sx(0), mt), (sx(0), mt + ph)], fill=(200, 200, 200))
    if ymin < 0 < ymax:
        d.line([(ml, sy(0)), (ml + pw, sy(0))], fill=(200, 200, 200))

    for i, (a, b) in enumerate(pts):
        cid = groups[i] if groups else 1
        col = _cluster_color(cid)
        gx, gy = sx(a), sy(b)
        d.ellipse([gx - 6, gy - 6, gx + 6, gy + 6], fill=col, outline=(255, 255, 255))
        if labels and i < len(labels):
            _png_text(d, f_pt, (gx + 10, gy - 10), str(labels[i])[:22], "la", fill=(70, 70, 70))

    if groups:
        seen = []
        for cid in groups:
            if cid not in seen:
                seen.append(cid)
        ly = mt + 8
        for cid in seen:
            d.rectangle([ml + pw + 40, ly, ml + pw + 54, ly + 14], fill=_cluster_color(cid))
            n_in = sum(1 for c in groups if c == cid)
            _png_text(d, f_tick, (ml + pw + 62, ly + 7), T("簇%d（%d 条）") % (cid, n_in), "lm",
                      fill=(60, 60, 60))
            ly += 26
            if ly > mt + ph - 20:
                break

    if title:
        _png_text(d, f_title, (ml + pw / 2.0, mt - 58), title, "ma", fill=(30, 30, 30))
    if xlabel:
        _png_text(d, f_lab, (ml + pw / 2.0, mt + ph + 62), xlabel, "ma", fill=(40, 40, 40))
    if ylabel:
        tmp = Image.new("RGBA", (500, 44), (255, 255, 255, 0))
        _png_text(ImageDraw.Draw(tmp), f_lab, (250, 22), ylabel, "mm", fill=(40, 40, 40))
        tmp = tmp.rotate(90, expand=True)
        img.paste(tmp, (12, int(mt + ph / 2 - tmp.height / 2)), tmp)
    img.save(path, "PNG")


def render_waterfall(path, series, title="", xlabel="", ylabel="", plot=None, width=None,
                     height=None, offset=0.75, normalize=True):
    """瀑布（堆叠偏移）图：多条光谱按序纵向错开，便于对比谱型。"""
    data = [(lab, xs, ys) for lab, xs, ys in series if len(xs) > 1 and len(ys) > 1]
    if not data:
        raise JwsError(T("没有可绘制的数据"))
    step = max(0.2, min(float(offset or 1.0), 3.0))
    local = dict(plot or {})
    local.update({"baseline": "none", "smooth_window": 1, "smooth_mode": "mean",
                  "derivative": 0, "normalize": "max" if normalize else "none",
                  "despike": False, "annotate_peaks": False, "show_y_ticks": False,
                  "y_min": None, "y_max": None,
                  "stacked": True, "stack_offset": step})
    local.pop("manual_peaks", None)
    colored = [(lab, xs, ys, _PALETTE[k % len(_PALETTE)])
               for k, (lab, xs, ys) in enumerate(data)]
    render_png(path, colored, title, xlabel, ylabel, local, width, height)


def render_overlay(path, series, title="", xlabel="", ylabel="", plot=None,
                   width=None, height=None, normalize=True, common_range=True,
                   merge_peaks=True, marks=None, return_geometry=False):
    """多数据图叠加（堆叠排布）：每个数据集一种颜色，各条上下错开、谱线彼此分开。

    参照 stacked spectra 的画法：各条先归一化到最大值 = 1，再按“谱线偏移”
    纵向错开 k×偏移，所以不会压在同一条基线上，谱型仍可横向比较。

    峰位标注做了**跨谱合并**：把各条谱上邻近的峰归成同一个峰，只画一条虚线、
    只标一个平均波数（merge_peaks=False 则退回逐峰标注）。
    各条谱自己的峰位标记（圆点 / 方块）仍然保留，方便看是哪几条谱有这个峰。

    marks 指定时按这份清单标注（[{"x": 波数, "manual": 是否手动}]），
    用于交互式预览里用户手动增减峰位；None = 自动合并检测。

    common_range=True 时，横坐标取所有数据图波数范围的**交集**：
    只比较大家都有数据的波段，避免某条谱短一截时右边空出一段白。
    return_geometry=True 时返回绘图区几何，供交互式预览把鼠标坐标换算成波数。
    """
    data = [(lab, xs, ys) for lab, xs, ys in series if len(xs) > 1 and len(ys) > 1]
    if not data:
        raise JwsError(T("没有可绘制的数据"))
    local = dict(plot or {})
    if normalize:
        local["normalize"] = "max"
    local["stacked"] = True
    local.setdefault("stack_offset", 1.0)
    local["merge_peak_labels"] = bool(merge_peaks) and len(data) > 1
    if marks is not None:
        local["peak_marks"] = list(marks)
    if common_range:
        lo, hi = overlay_range(data)
        if lo is not None and hi is not None and hi > lo:
            # 高级设置里手动填过范围就以手动为准
            if local.get("x_min") is None:
                local["x_min"] = lo
            if local.get("x_max") is None:
                local["x_max"] = hi
    colored = [(lab, xs, ys, overlay_color(k)) for k, (lab, xs, ys) in enumerate(data)]
    return render_png(path, colored, title, xlabel, ylabel, local, width, height,
                      return_geometry=return_geometry)


def overlay_range(series):
    """返回叠加时共用的横坐标范围 (lo, hi)：各数据图波数范围的交集。

    各条完全不相交时交集为空，退回并集，避免画出一张空白图。
    """
    xs_list = [xs for _l, xs, ys in series if len(xs) > 1 and len(ys) > 1]
    if not xs_list:
        return None, None
    lo = max(min(xs) for xs in xs_list)
    hi = min(max(xs) for xs in xs_list)
    if hi <= lo:
        return min(min(xs) for xs in xs_list), max(max(xs) for xs in xs_list)
    return lo, hi


def _write_png(dst, spec, names, plot=None):
    if not _HAVE_PIL:
        raise JwsError(T("导出 PNG 需要 Pillow 组件（pip install pillow）"))
    xs = _calibrate(spec.x_values(), plot)
    series = []
    for c in range(spec.channel_number):
        label = names[1 + c] if 1 + c < len(names) else "Y%d" % (c + 1)
        series.append((label, xs, spec.y_data[c], _PALETTE[c % len(_PALETTE)]))
    title = os.path.splitext(os.path.basename(dst))[0]
    ylabel = names[1] if len(names) > 1 else ""
    render_png(dst, series, title, names[0], ylabel, plot=plot)


def _write_peaks_csv(dst, spec, names, plot=None):
    p = _plot_opts(plot)
    xs = _calibrate(spec.x_values(), p)
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        f.write(T("通道,峰位,强度,相对强度(%),半高宽FWHM,峰突出度,来源\n"))
        for c in range(spec.channel_number):
            label = names[1 + c] if 1 + c < len(names) else "Y%d" % (c + 1)
            for pk in analyze_peaks(xs, spec.y_data[c], p):
                f.write("%s,%.4f,%.6g,%.2f,%.4f,%.6g,%s\n" % (
                    label, pk["x"], pk["y"], pk["rel"], pk["fwhm"], pk["prominence"],
                    "手动" if pk.get("manual") else "自动"))


_FIT_SHAPE_NAMES = {"gaussian": "高斯", "lorentzian": "洛伦兹", "voigt": "伪Voigt"}


def _shape_value(dx, height, fwhm, eta, shape):
    fwhm = fwhm if fwhm > 0 else 1e-9
    t = 2.0 * dx / fwhm
    gauss = math.exp(-math.log(2.0) * t * t)
    lorentz = 1.0 / (1.0 + t * t)
    if shape == "gaussian":
        return height * gauss
    if shape == "lorentzian":
        return height * lorentz
    eta = min(1.0, max(0.0, eta))
    return height * (eta * lorentz + (1.0 - eta) * gauss)


def _shape_area(height, fwhm, eta, shape):
    gauss = height * fwhm * math.sqrt(math.pi / (4.0 * math.log(2.0)))
    lorentz = height * fwhm * math.pi / 2.0
    if shape == "gaussian":
        return gauss
    if shape == "lorentzian":
        return lorentz
    eta = min(1.0, max(0.0, eta))
    return eta * lorentz + (1.0 - eta) * gauss


def fit_peak(xs, ys, idx, base, shape="voigt", max_iter=60):
    n = len(ys)
    if n < 5:
        return None
    step = abs(xs[1] - xs[0]) if len(xs) > 1 else 1.0
    if not step:
        step = 1.0
    init_fwhm = _peak_fwhm(xs, ys, idx, base)
    if not init_fwhm or init_fwhm <= 0:
        init_fwhm = step * 10.0
    half = max(3, int(round(1.6 * init_fwhm / step)))
    lo = max(0, idx - half)
    hi = min(n, idx + half + 1)
    wx = xs[lo:hi]
    wy = ys[lo:hi]
    if len(wx) < 4:
        return None

    params = [max(ys[idx] - base, 1e-9), xs[idx], init_fwhm, 0.5 if shape == "voigt" else 0.0]

    def sse(pr):
        h, c, f, eta = pr
        if f <= 0 or h < 0:
            return float("inf")
        err = 0.0
        for x, y in zip(wx, wy):
            d = y - (base + _shape_value(x - c, h, f, eta, shape))
            err += d * d
        return err

    steps = [abs(params[0]) * 0.3 + 1e-9, max(init_fwhm * 0.3, step),
             max(init_fwhm * 0.3, step), 0.2]
    best = sse(params)
    for _it in range(max_iter):
        improved = False
        for pi in range(4):
            if shape != "voigt" and pi == 3:
                continue
            for sign in (1.0, -1.0):
                cand = list(params)
                cand[pi] += sign * steps[pi]
                if pi == 2 and cand[2] <= step:
                    continue
                if pi == 3:
                    cand[3] = min(1.0, max(0.0, cand[3]))
                val = sse(cand)
                if val < best:
                    params, best, improved = cand, val, True
        if not improved:
            steps = [s * 0.5 for s in steps]
            if max(steps[0], steps[1], steps[2]) < 1e-7:
                break
    h, c, f, eta = params
    mean = sum(wy) / len(wy)
    ss_tot = sum((y - mean) ** 2 for y in wy)
    r2 = (1.0 - best / ss_tot) if ss_tot > 0 else 0.0
    return {"center": c, "height": h, "fwhm": f, "eta": eta,
            "area": _shape_area(h, f, eta, shape), "r2": r2, "shape": shape}


def fit_spectrum_peaks(xs, ys, plot=None, processed=False):
    p = _plot_opts(plot)
    sig = ys if processed else _process_signal(ys, p, xs)
    out = []
    for pk in analyze_peaks(xs, sig, p, processed=True):
        res = fit_peak(xs, sig, pk["index"], pk["base"], p["fit_shape"])
        if res is None:
            continue
        res["x"] = pk["x"]
        res["manual"] = bool(pk.get("manual"))
        out.append(res)
    return out


def _write_fit_csv(dst, spec, names, plot=None):
    p = _plot_opts(plot)
    xs = _calibrate(spec.x_values(), p)
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        f.write(T("通道,峰位,拟合中心,拟合峰高,拟合半高宽FWHM,峰面积,峰形,混合系数,拟合R2,来源\n"))
        for c in range(spec.channel_number):
            label = names[1 + c] if 1 + c < len(names) else "Y%d" % (c + 1)
            for fit in fit_spectrum_peaks(xs, spec.y_data[c], p):
                f.write("%s,%.4f,%.4f,%.6g,%.4f,%.6g,%s,%.2f,%.4f,%s\n" % (
                    label, fit["x"], fit["center"], fit["height"], fit["fwhm"],
                    fit["area"], T(_FIT_SHAPE_NAMES.get(fit["shape"], fit["shape"])),
                    fit["eta"], fit["r2"], T("手动") if fit["manual"] else T("自动")))


def load_reference_spectrum(path):
    xs, ys = [], []
    with open(path, "r", encoding="utf-8", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.replace(",", " ").replace("\t", " ").replace(";", " ").split()
            if len(parts) < 2:
                continue
            try:
                x = float(parts[0])
                y = float(parts[1])
            except ValueError:
                continue
            xs.append(x)
            ys.append(y)
    if len(xs) < 4:
        raise JwsError(T("参考谱数据不足：%s") % os.path.basename(path))
    if xs[0] > xs[-1]:
        xs.reverse()
        ys.reverse()
    return os.path.splitext(os.path.basename(path))[0], xs, ys


def _resample(xs, ys, grid):
    out = []
    j = 0
    n = len(xs)
    for g in grid:
        if g <= xs[0]:
            out.append(ys[0])
            continue
        if g >= xs[-1]:
            out.append(ys[-1])
            continue
        while j < n - 2 and xs[j + 1] < g:
            j += 1
        x0, x1 = xs[j], xs[j + 1]
        y0, y1 = ys[j], ys[j + 1]
        out.append(y0 if x1 == x0 else y0 + (y1 - y0) * (g - x0) / (x1 - x0))
    return out


def compare_spectra(a, b, points=400):
    axs, ays = a
    bxs, bys = b
    if not axs or not bxs:
        return None
    lo = max(min(axs), min(bxs))
    hi = min(max(axs), max(bxs))
    if hi <= lo:
        return None
    grid = [lo + (hi - lo) * i / (points - 1) for i in range(points)]
    ya = _resample(axs, ays, grid)
    yb = _resample(bxs, bys, grid)
    ma = sum(ya) / len(ya)
    mb = sum(yb) / len(yb)
    da = [v - ma for v in ya]
    db = [v - mb for v in yb]
    na = math.sqrt(sum(v * v for v in da))
    nb = math.sqrt(sum(v * v for v in db))
    corr = (sum(u * v for u, v in zip(da, db)) / (na * nb)) if na and nb else 0.0
    sa = math.sqrt(sum(v * v for v in ya))
    sb = math.sqrt(sum(v * v for v in yb))
    if sa and sb:
        cosv = max(-1.0, min(1.0, sum(u * v for u, v in zip(ya, yb)) / (sa * sb)))
        angle = math.degrees(math.acos(cosv))
    else:
        angle = 90.0
    return {"corr": corr, "angle": angle, "overlap": (lo, hi)}


def match_references(target_xy, ref_paths, top=None):
    txs, tys = target_xy
    results = []
    for path in ref_paths:
        try:
            name, rxs, rys = load_reference_spectrum(path)
        except Exception:
            continue
        cmp = compare_spectra((txs, tys), (rxs, rys))
        if cmp is None:
            continue
        cmp["name"] = name
        cmp["path"] = path
        results.append(cmp)
    results.sort(key=lambda r: -r["corr"])
    return results[:top] if top else results


def search_reference_table(peak_positions, table_path, tol=5.0):
    rows = []
    with open(table_path, "r", encoding="utf-8-sig", errors="ignore") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            parts = [p.strip() for p in line.replace("\t", ",").replace(";", ",").split(",")]
            if len(parts) < 2:
                continue
            name = parts[0]
            vals = []
            for p in parts[1:]:
                try:
                    vals.append(float(p))
                except ValueError:
                    continue
            if name and vals:
                rows.append((name, vals))
    scored = []
    for name, vals in rows:
        matched = sum(1 for v in vals if any(abs(v - pos) <= tol for pos in peak_positions))
        scored.append((100.0 * matched / len(vals), name, matched, len(vals)))
    scored.sort(key=lambda r: (-r[0], r[1]))
    return scored


def similarity_matrix(spectra):
    n = len(spectra)
    mat = [[0.0] * n for _ in range(n)]
    for i in range(n):
        mat[i][i] = 1.0
        for j in range(i + 1, n):
            cmp = compare_spectra((spectra[i][1], spectra[i][2]),
                                  (spectra[j][1], spectra[j][2]))
            value = cmp["corr"] if cmp else 0.0
            mat[i][j] = mat[j][i] = value
    return mat


def _pearson(a, b):
    n = len(a)
    if n < 2:
        return 0.0
    ma = sum(a) / n
    mb = sum(b) / n
    sa = sb = sab = 0.0
    for i in range(n):
        da = a[i] - ma
        db = b[i] - mb
        sa += da * da
        sb += db * db
        sab += da * db
    den = math.sqrt(sa * sb)
    return sab / den if den > 0 else 0.0


def _pca(matrix, k=3, iters=80):
    """幂迭代 + 收缩求前 k 个主成分。返回 (载荷, 得分)。"""
    n = len(matrix)
    if n < 3:
        return [], []
    m = len(matrix[0])
    means = [sum(matrix[i][j] for i in range(n)) / n for j in range(m)]
    centered = [[matrix[i][j] - means[j] for j in range(m)] for i in range(n)]
    work = [row[:] for row in centered]
    comps = []
    for c in range(min(k, m, n - 1)):
        vec = [0.0] * m
        vec[(c * 5 + 1) % m] = 1.0
        for _ in range(max(10, int(iters))):
            new = [0.0] * m
            for i in range(n):
                row = work[i]
                s = 0.0
                for j in range(m):
                    s += row[j] * vec[j]
                if s:
                    for j in range(m):
                        new[j] += row[j] * s
            norm = math.sqrt(sum(v * v for v in new))
            if norm < 1e-12:
                break
            new = [v / norm for v in new]
            delta = sum(abs(new[j] - vec[j]) for j in range(m))
            vec = new
            if delta < 1e-10:
                break
        comps.append(vec)
        for i in range(n):
            row = work[i]
            s = sum(row[j] * vec[j] for j in range(m))
            for j in range(m):
                row[j] -= s * vec[j]
    scores = [[sum(centered[i][j] * comp[j] for j in range(m)) for comp in comps]
              for i in range(n)]
    return comps, scores


def cluster_spectra(spectra, points=0, cut=None):
    """把多条光谱重采样、归一后做聚类分析（相关系数距离 + 平均连接层次聚类）。

    返回 dict：names, grid, vecs, dist, merges, cut, labels(每条谱的簇号),
    groups(簇号 -> 索引列表), scores(前3主成分坐标), explained(方差贡献估计)。
    """
    usable = [(name, xs, ys) for name, xs, ys in spectra if len(xs) > 1]
    n = len(usable)
    if n < 2:
        return None
    span = _common_range(usable)
    if span is None:
        return None
    lo, hi = span
    densest = max(len(u[1]) for u in usable)
    npts = int(points) if (points and int(points) > 3) else densest
    npts = max(60, min(npts, 1200))
    step = (hi - lo) / (npts - 1)
    grid = [lo + i * step for i in range(npts)]
    vecs = []
    for _name, xs, ys in usable:
        v = [_interp_at(xs, ys, x) for x in grid]
        base = min(v)
        rng = max(v) - base
        if rng <= 0:
            rng = 1.0
        vecs.append([(t - base) / rng for t in v])

    dist = [[0.0] * n for _ in range(n)]
    for i in range(n):
        for j in range(i + 1, n):
            d = 1.0 - _pearson(vecs[i], vecs[j])
            dist[i][j] = dist[j][i] = d

    clusters = {i: [i] for i in range(n)}
    active = set(range(n))
    link = {(i, j): dist[i][j] for i in range(n) for j in range(i + 1, n)}
    merges = []
    next_id = n
    while len(active) > 1:
        keys = sorted(active)
        best_a, best_b, best_d = None, None, None
        for ai in range(len(keys)):
            for bi in range(ai + 1, len(keys)):
                a, b = keys[ai], keys[bi]
                d = link[(a, b) if a < b else (b, a)]
                if best_d is None or d < best_d:
                    best_a, best_b, best_d = a, b, d
        a, b = best_a, best_b
        members = clusters[a] + clusters[b]
        merges.append((a, b, best_d, len(members)))
        wa, wb = len(clusters[a]), len(clusters[b])
        active.discard(a)
        active.discard(b)
        for c in list(active):
            da = link[(a, c) if a < c else (c, a)]
            db = link[(b, c) if b < c else (c, b)]
            key = (next_id, c) if next_id < c else (c, next_id)
            link[key] = (da * wa + db * wb) / float(wa + wb)
        clusters[next_id] = members
        active.add(next_id)
        next_id += 1

    heights = [m[2] for m in merges]
    if cut is None:
        hs = sorted(heights, reverse=True)
        if len(hs) >= 3:
            gaps = [(hs[i] - hs[i + 1], i) for i in range(len(hs) - 1)]
            best_gap, k = max(gaps)
            if hs[0] < 0.05 or best_gap < 0.02:
                cut = hs[0] * 1.1 + 1e-9
            else:
                cut = (hs[k] + hs[k + 1]) / 2.0
        else:
            cut = hs[0] * 1.1 + 1e-9 if (not hs or hs[0] < 0.05) else hs[0] * 0.5
    cut = float(cut)

    total_ids = next_id
    parent = list(range(total_ids))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for k, (a, b, h, _c) in enumerate(merges):
        if h > cut:
            continue
        nid = n + k
        for child in (a, b):
            rc, rn = find(child), find(nid)
            if rc != rn:
                parent[rn] = rc
    raw = {}
    for i in range(n):
        raw.setdefault(find(i), []).append(i)
    ordered = sorted(raw.values(), key=lambda g: (-len(g), g[0]))
    labels = [0] * n
    groups = {}
    for cid, members in enumerate(ordered, 1):
        groups[cid] = members
        for i in members:
            labels[i] = cid

    comps, scores = _pca(vecs, 3)
    total_var = sum(sum((vecs[i][j] - sum(row[j] for row in vecs) / n) ** 2
                        for i in range(n)) for j in range(npts))
    explained = []
    for c, comp in enumerate(comps):
        var = sum(s[c] ** 2 for s in scores)
        explained.append(100.0 * var / total_var if total_var > 0 else 0.0)

    return {
        "names": [u[0] for u in usable],
        "grid": grid,
        "vecs": vecs,
        "dist": dist,
        "merges": merges,
        "cut": cut,
        "labels": labels,
        "groups": groups,
        "scores": scores,
        "explained": explained,
    }


_ROD_BASE = "https://solsa.crystallography.net/rod"


_MINERAL_DB = [
    {"en": "Zircon", "cn": "锆石", "formula": "ZrSiO4", "system": "四方晶系",
     "bands": [(1008.0, "ν3(SiO4) 反对称伸缩（最强峰）"), (974.0, "ν1(SiO4) 对称伸缩"),
               (438.0, "ν2(SiO4) 弯曲振动"), (393.0, "外部振动/晶格模式"),
               (356.0, "外部振动"), (264.0, "外部振动"), (224.0, "外部振动"),
               (202.0, "外部振动")]},
    {"en": "Baddeleyite", "cn": "斜锆石（单斜氧化锆）", "formula": "ZrO2", "system": "单斜晶系",
     "bands": [(178.0, "晶格振动"), (190.0, "晶格振动"), (222.0, "晶格振动"),
               (305.0, "Zr-O 振动"), (335.0, "Zr-O 振动"), (348.0, "Zr-O 振动"),
               (382.0, "Zr-O 振动"), (476.0, "Zr-O 伸缩"), (502.0, "Zr-O 伸缩"),
               (536.0, "Zr-O 伸缩"), (560.0, "Zr-O 伸缩"), (618.0, "Zr-O 伸缩")]},
    {"en": "Cubic zirconia", "cn": "立方氧化锆", "formula": "ZrO2", "system": "立方晶系",
     "bands": [(610.0, "Zr-O 伸缩（特征强峰）"), (470.0, "Zr-O 振动"),
               (260.0, "晶格振动")]},
    {"en": "Tetragonal zirconia", "cn": "四方氧化锆", "formula": "ZrO2", "system": "四方晶系",
     "bands": [(148.0, "晶格振动"), (265.0, "Zr-O 振动"), (320.0, "Zr-O 振动"),
               (460.0, "Zr-O 振动"), (610.0, "Zr-O 伸缩"), (640.0, "Zr-O 伸缩")]},
    {"en": "Quartz", "cn": "石英", "formula": "SiO2", "system": "三方晶系",
     "bands": [(464.0, "ν1 A1 Si-O 对称伸缩（最强峰）"), (206.0, "晶格振动"),
               (265.0, "晶格振动"), (355.0, "晶格振动"), (394.0, "晶格振动"),
               (795.0, "Si-O 伸缩"), (808.0, "Si-O 伸缩"), (1085.0, "Si-O 反对称伸缩"),
               (1162.0, "Si-O 反对称伸缩")]},
    {"en": "Cristobalite", "cn": "方石英", "formula": "SiO2", "system": "四方晶系",
     "bands": [(230.0, "晶格振动"), (416.0, "Si-O 弯曲"), (786.0, "Si-O 伸缩")]},
    {"en": "Coesite", "cn": "柯石英", "formula": "SiO2", "system": "单斜晶系",
     "bands": [(176.0, "晶格振动"), (270.0, "晶格振动"), (355.0, "晶格振动"),
               (521.0, "Si-O 伸缩（特征峰）"), (768.0, "Si-O 伸缩")]},
    {"en": "Rutile", "cn": "金红石", "formula": "TiO2", "system": "四方晶系",
     "bands": [(143.0, "B1g 晶格振动（弱）"), (447.0, "Eg Ti-O 伸缩"),
               (612.0, "A1g Ti-O 伸缩（最强峰）"), (826.0, "B2g Ti-O 伸缩（弱）")]},
    {"en": "Anatase", "cn": "锐钛矿", "formula": "TiO2", "system": "四方晶系",
     "bands": [(144.0, "Eg Ti-O 伸缩（很强）"), (197.0, "Eg"), (399.0, "B1g"),
               (513.0, "A1g"), (639.0, "Eg")]},
    {"en": "Hematite", "cn": "赤铁矿", "formula": "Fe2O3", "system": "三方晶系",
     "bands": [(226.0, "A1g Fe-O"), (245.0, "Eg Fe-O"), (293.0, "Eg Fe-O"),
               (299.0, "Eg Fe-O"), (412.0, "Eg Fe-O"), (498.0, "A1g Fe-O"),
               (613.0, "Eg Fe-O"), (660.0, "Eu 晶格"), (1320.0, "双磁子散射（宽峰）")]},
    {"en": "Magnetite", "cn": "磁铁矿", "formula": "Fe3O4", "system": "立方晶系",
     "bands": [(193.0, "T2g"), (306.0, "Eg"), (538.0, "T2g（特征）"), (668.0, "A1g（最强）")]},
    {"en": "Goethite", "cn": "针铁矿", "formula": "FeOOH", "system": "正交晶系",
     "bands": [(243.0, "Fe-O"), (299.0, "Fe-O"), (386.0, "Fe-O"),
               (480.0, "Fe-O"), (550.0, "Fe-O"), (685.0, "Fe-O"),
               (992.0, "Fe-OH 弯曲")]},
    {"en": "Lepidocrocite", "cn": "纤铁矿", "formula": "FeOOH", "system": "正交晶系",
     "bands": [(219.0, "Fe-O"), (252.0, "Fe-O"), (305.0, "Fe-O"), (350.0, "Fe-O"),
               (378.0, "Fe-O"), (528.0, "Fe-O"), (650.0, "Fe-OH")]},
    {"en": "Corundum", "cn": "刚玉", "formula": "Al2O3", "system": "三方晶系",
     "bands": [(378.0, "Eg Al-O"), (418.0, "A1g Al-O（最强）"), (430.0, "Eg"),
               (578.0, "Eg"), (645.0, "Eg"), (751.0, "A1g")]},
    {"en": "Gibbsite", "cn": "三水铝石", "formula": "Al(OH)3", "system": "单斜晶系",
     "bands": [(320.0, "Al-O"), (352.0, "Al-O"), (3370.0, "OH 伸缩"),
               (3430.0, "OH 伸缩"), (3625.0, "OH 伸缩")]},
    {"en": "Boehmite", "cn": "勃姆石", "formula": "AlOOH", "system": "正交晶系",
     "bands": [(310.0, "Al-O"), (350.0, "Al-O"), (480.0, "Al-O"),
               (670.0, "Al-OH"), (740.0, "Al-OH（宽）")]},
    {"en": "Cassiterite", "cn": "锡石", "formula": "SnO2", "system": "四方晶系",
     "bands": [(476.0, "Eg Sn-O"), (634.0, "A1g Sn-O（最强）"), (776.0, "B2g")]},
    {"en": "Calcite", "cn": "方解石", "formula": "CaCO3", "system": "三方晶系",
     "bands": [(156.0, "晶格振动"), (282.0, "晶格振动"), (712.0, "ν4(CO3) 弯曲"),
               (1086.0, "ν1(CO3) 对称伸缩（最强峰）"), (1435.0, "ν3(CO3) 反对称伸缩")]},
    {"en": "Aragonite", "cn": "文石", "formula": "CaCO3", "system": "正交晶系",
     "bands": [(152.0, "晶格振动"), (206.0, "晶格振动"), (703.0, "ν4(CO3)"),
               (1085.0, "ν1(CO3) 对称伸缩（最强峰）"), (1460.0, "ν3(CO3)")]},
    {"en": "Dolomite", "cn": "白云石", "formula": "CaMg(CO3)2", "system": "三方晶系",
     "bands": [(176.0, "晶格振动"), (300.0, "晶格振动"), (725.0, "ν4(CO3)"),
               (1097.0, "ν1(CO3) 对称伸缩（最强峰）"), (1440.0, "ν3(CO3)")]},
    {"en": "Magnesite", "cn": "菱镁矿", "formula": "MgCO3", "system": "三方晶系",
     "bands": [(213.0, "晶格振动"), (330.0, "晶格振动"), (738.0, "ν4(CO3)"),
               (1094.0, "ν1(CO3)（最强峰）")]},
    {"en": "Siderite", "cn": "菱铁矿", "formula": "FeCO3", "system": "三方晶系",
     "bands": [(184.0, "晶格振动"), (287.0, "晶格振动"), (730.0, "ν4(CO3)"),
               (1086.0, "ν1(CO3)（最强峰）")]},
    {"en": "Rhodochrosite", "cn": "菱锰矿", "formula": "MnCO3", "system": "三方晶系",
     "bands": [(185.0, "晶格振动"), (248.0, "晶格振动"), (290.0, "晶格振动"),
               (717.0, "ν4(CO3)"), (1085.0, "ν1(CO3)（最强峰）")]},
    {"en": "Gypsum", "cn": "石膏", "formula": "CaSO4·2H2O", "system": "单斜晶系",
     "bands": [(414.0, "ν2(SO4) 弯曲"), (493.0, "ν2(SO4)"), (620.0, "ν4(SO4)"),
               (670.0, "ν4(SO4)"), (1008.0, "ν1(SO4) 对称伸缩（最强峰）"),
               (1136.0, "ν3(SO4) 反对称伸缩")]},
    {"en": "Anhydrite", "cn": "硬石膏", "formula": "CaSO4", "system": "正交晶系",
     "bands": [(417.0, "ν2(SO4)"), (498.0, "ν2(SO4)"), (609.0, "ν4(SO4)"),
               (628.0, "ν4(SO4)"), (1017.0, "ν1(SO4)（最强峰）"), (1129.0, "ν3(SO4)")]},
    {"en": "Barite", "cn": "重晶石", "formula": "BaSO4", "system": "正交晶系",
     "bands": [(452.0, "ν2(SO4)"), (462.0, "ν2(SO4)"), (618.0, "ν4(SO4)"),
               (988.0, "ν1(SO4) 对称伸缩（最强峰）"), (1084.0, "ν3(SO4)"),
               (1136.0, "ν3(SO4)")]},
    {"en": "Celestite", "cn": "天青石", "formula": "SrSO4", "system": "正交晶系",
     "bands": [(452.0, "ν2(SO4)"), (622.0, "ν4(SO4)"),
               (1000.0, "ν1(SO4)（最强峰）"), (1085.0, "ν3(SO4)")]},
    {"en": "Apatite", "cn": "磷灰石", "formula": "Ca5(PO4)3(F,OH,Cl)", "system": "六方晶系",
     "bands": [(964.0, "ν1(PO4) 对称伸缩（最强峰）"),
               (1050.0, "ν3(PO4) 反对称伸缩")]},
    {"en": "Monazite", "cn": "独居石", "formula": "(Ce,La,Nd)PO4", "system": "单斜晶系",
     "bands": [(975.0, "ν1(PO4) 对称伸缩（最强峰）"), (1068.0, "ν3(PO4)")]},
    {"en": "Xenotime", "cn": "磷钇矿", "formula": "YPO4", "system": "四方晶系",
     "bands": [(995.0, "ν1(PO4) 对称伸缩（最强峰）"), (1050.0, "ν3(PO4)")]},
    {"en": "Olivine", "cn": "橄榄石", "formula": "(Mg,Fe)2SiO4", "system": "正交晶系",
     "bands": [(820.0, "ν1(SiO4) 对称伸缩"), (853.0, "ν1(SiO4) 对称伸缩（双峰）"),
               (918.0, "ν3(SiO4)"), (960.0, "ν3(SiO4)")]},
    {"en": "Garnet", "cn": "石榴子石", "formula": "X3Y2(SiO4)3", "system": "立方晶系",
     "bands": [(340.0, "晶格振动"), (370.0, "晶格振动"), (550.0, "ν2(SiO4)"),
               (900.0, "ν1(SiO4) 对称伸缩"), (1050.0, "ν3(SiO4)")]},
    {"en": "Pyroxene", "cn": "辉石", "formula": "(Mg,Fe,Ca)SiO3", "system": "单斜/正交",
     "bands": [(320.0, "晶格振动"), (390.0, "晶格振动"), (660.0, "ν4(SiO4)"),
               (1010.0, "ν1(SiO4) 对称伸缩")]},
    {"en": "Amphibole", "cn": "角闪石", "formula": "Ca2(Mg,Fe)5Si8O22(OH)2", "system": "单斜晶系",
     "bands": [(350.0, "晶格振动"), (670.0, "Si-O-Si"), (930.0, "ν1(SiO4)"),
               (1050.0, "ν3(SiO4)"), (3620.0, "OH 伸缩")]},
    {"en": "Orthoclase", "cn": "正长石", "formula": "KAlSi3O8", "system": "单斜晶系",
     "bands": [(454.0, "晶格振动"), (476.0, "晶格振动"), (514.0, "ν4(SiO4)"),
               (750.0, "Si-O-Si"), (1130.0, "ν3(SiO4)")]},
    {"en": "Albite", "cn": "钠长石", "formula": "NaAlSi3O8", "system": "三斜晶系",
     "bands": [(290.0, "晶格振动"), (480.0, "ν4(SiO4)"), (508.0, "ν4(SiO4)"),
               (762.0, "Si-O-Si"), (1090.0, "ν3(SiO4)")]},
    {"en": "Anorthite", "cn": "钙长石", "formula": "CaAl2Si2O8", "system": "三斜晶系",
     "bands": [(505.0, "ν4(SiO4)"), (560.0, "ν4(SiO4)"),
               (915.0, "Al-O 振动"), (1100.0, "ν3(SiO4)")]},
    {"en": "Muscovite", "cn": "白云母", "formula": "KAl2(AlSi3O10)(OH)2", "system": "单斜晶系",
     "bands": [(265.0, "晶格振动"), (410.0, "Si-O"), (700.0, "Si-O-Si"),
               (900.0, "Al-OH"), (3620.0, "OH 伸缩")]},
    {"en": "Kaolinite", "cn": "高岭石", "formula": "Al2Si2O5(OH)4", "system": "三斜晶系",
     "bands": [(450.0, "Si-O"), (790.0, "Si-O"), (915.0, "Al-OH 弯曲"),
               (3620.0, "内层 OH 伸缩"), (3652.0, "内层 OH 伸缩"),
               (3695.0, "表面 OH 伸缩")]},
    {"en": "Talc", "cn": "滑石", "formula": "Mg3Si4O10(OH)2", "system": "单斜晶系",
     "bands": [(195.0, "晶格振动"), (360.0, "晶格振动"), (670.0, "Si-O-Si"),
               (1040.0, "ν3(SiO4)"), (3675.0, "OH 伸缩")]},
    {"en": "Pyrite", "cn": "黄铁矿", "formula": "FeS2", "system": "立方晶系",
     "bands": [(343.0, "Ag S-S 伸缩（最强峰）"), (379.0, "Eg S-S 伸缩"),
               (430.0, "Tg")]},
    {"en": "Chalcopyrite", "cn": "黄铜矿", "formula": "CuFeS2", "system": "四方晶系",
     "bands": [(267.0, "晶格振动"), (290.0, "S-S 伸缩"),
               (318.0, "晶格振动"), (352.0, "S-S 伸缩")]},
    {"en": "Sphalerite", "cn": "闪锌矿", "formula": "ZnS", "system": "立方晶系",
     "bands": [(352.0, "TO/LO Zn-S 伸缩")]},
    {"en": "Molybdenite", "cn": "辉钼矿", "formula": "MoS2", "system": "六方晶系",
     "bands": [(383.0, "E2g Mo-S 面内伸缩（最强峰）"), (408.0, "A1g Mo-S 面外伸缩")]},
    {"en": "Graphite", "cn": "石墨", "formula": "C", "system": "六方晶系",
     "bands": [(1350.0, "D 带（缺陷诱导）"), (1580.0, "G 带（sp2 骨架）"),
               (2700.0, "2D 带（二阶）")]},
    {"en": "Fluorite", "cn": "萤石", "formula": "CaF2", "system": "立方晶系",
     "bands": [(322.0, "Ca-F 伸缩")]},
    {"en": "Arsenopyrite", "cn": "毒砂", "formula": "FeAsS", "system": "单斜晶系",
     "bands": [(196.0, "晶格振动"), (300.0, "As-S"), (350.0, "As-S"), (380.0, "As-S")]},
    {"en": "Diamond", "cn": "金刚石", "formula": "C", "system": "立方晶系",
     "bands": [(1332.0, "sp3 C-C 伸缩（特征单峰）")]},
]


def _formula_elements(formula):
    out = []
    i = 0
    n = len(formula)
    while i < n:
        ch = formula[i]
        if ch.isupper():
            sym = ch
            i += 1
            while i < n and formula[i].islower():
                sym += formula[i]
                i += 1
            if sym not in out:
                out.append(sym)
        else:
            i += 1
    return out


def lookup_mineral(name):
    """按英文名 / 中文名 / 编号前缀查找内置矿物记录。"""
    if not name:
        return None
    key = str(name).strip().lower()
    if not key:
        return None
    for rec in _MINERAL_DB:
        if rec["en"].lower() == key or rec["cn"] == str(name).strip():
            return rec
    for rec in _MINERAL_DB:
        if key in rec["en"].lower() or key in rec["cn"]:
            return rec
    simple = key.split("_")[0].replace(" ", "")
    for rec in _MINERAL_DB:
        if rec["en"].lower().replace(" ", "").startswith(simple):
            return rec
    return None


def assign_bands(peak_positions, mineral=None, tol=8.0):
    """把峰位逐个归属到内置矿物特征峰；返回 [(峰位, 归属文本), ...]。"""
    rec = lookup_mineral(mineral)
    if rec is None:
        return [(p, "") for p in peak_positions]
    out = []
    for pos in peak_positions:
        best = None
        for band, text in rec["bands"]:
            d = abs(float(pos) - band)
            if d <= tol and (best is None or d < best[0]):
                best = (d, band, text)
        if best is None:
            out.append((pos, T("未归属")))
        else:
            out.append((pos, T("%s（参考 %.1f cm-1）") % (T(best[2]), best[1])))
    return out


def mineral_search(keyword, mode="name"):
    """按名称 / 化学式 / 元素检索内置矿物表。"""
    text = str(keyword or "").strip()
    if not text:
        return []
    low = text.lower()
    hits = []
    for rec in _MINERAL_DB:
        if mode == "element":
            syms = [s.upper() for s in
                    text.replace(",", " ").replace(";", " ").split() if s.strip()]
            elems = [e.upper() for e in _formula_elements(rec["formula"])]
            if syms and all(s in elems for s in syms):
                hits.append(rec)
            continue
        if mode == "formula":
            if low in rec["formula"].lower():
                hits.append(rec)
            continue
        if (low in rec["en"].lower() or low in rec["cn"]
                or low in rec["formula"].lower()):
            hits.append(rec)
    return hits


def mineral_info_text(mineral):
    rec = lookup_mineral(mineral)
    if rec is None:
        return ""
    if ui_lang() == "en":
        lines = ["%s (%s)" % (rec["en"], rec["cn"]),
                 T("化学式：%s") % rec["formula"],
                 T("晶系：%s") % T(rec["system"]),
                 T("元素组成：%s") % ", ".join(_formula_elements(rec["formula"])),
                 "",
                 T("特征拉曼峰（cm-1）：")]
    else:
        lines = ["%s（%s）" % (rec["en"], rec["cn"]),
                 "化学式：%s" % rec["formula"],
                 "晶系：%s" % rec["system"],
                 "元素组成：%s" % "、".join(_formula_elements(rec["formula"])),
                 "",
                 "特征拉曼峰（cm-1）："]
    for band, text in rec["bands"]:
        lines.append("   %8.1f   %s" % (band, T(text)))
    return "\n".join(lines)


def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


_DATA_ROOT_NAME = "工具数据"
_SETTINGS_NAME = "工具设置.ini"
_DATA_ROOT_CACHE = [None]


def _settings_path():
    return os.path.join(_app_dir(), _SETTINGS_NAME)


def _load_settings():
    data = {}
    try:
        with open(_settings_path(), "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                key, _, value = line.partition("=")
                data[key.strip()] = value.strip()
    except OSError:
        pass
    return data


def _save_settings(data):
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            f.write(T("# 拉曼光谱工具 设置文件（请勿删除本文件）\n"))
            for key, value in data.items():
                f.write("%s=%s\n" % (key, value))
        return True
    except OSError:
        return False


def _can_write(path):
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, ".writetest")
        with open(probe, "w", encoding="utf-8") as f:
            f.write("1")
        os.remove(probe)
        return True
    except OSError:
        return False


def data_root(create=True):
    if _DATA_ROOT_CACHE[0]:
        return _DATA_ROOT_CACHE[0]
    custom = _load_settings().get("data_dir")
    if custom and os.path.isdir(custom):
        _DATA_ROOT_CACHE[0] = custom
        return custom
    default = os.path.join(_app_dir(), _DATA_ROOT_NAME)
    if _can_write(default) or not create:
        _DATA_ROOT_CACHE[0] = default
        return default
    local = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    fallback = os.path.join(local, "拉曼光谱工具", _DATA_ROOT_NAME)
    legacy = os.path.join(local, "光谱转换工具", _DATA_ROOT_NAME)
    if not os.path.isdir(fallback) and os.path.isdir(legacy):
        fallback = legacy          # 沿用旧目录名，避免升级后找不到原有数据
    _can_write(fallback)
    _DATA_ROOT_CACHE[0] = fallback
    return fallback


def set_data_root(path):
    path = os.path.abspath(path)
    os.makedirs(path, exist_ok=True)
    _DATA_ROOT_CACHE[0] = path
    data = _load_settings()
    data["data_dir"] = path
    return _save_settings(data)


def ref_dir():
    folder = os.path.join(data_root(), "参考谱库")
    os.makedirs(folder, exist_ok=True)
    return folder


def rruff_dir():
    folder = os.path.join(data_root(), "RRUFF数据包")
    os.makedirs(folder, exist_ok=True)
    return folder


def results_dir():
    folder = os.path.join(data_root(), "分析结果")
    os.makedirs(folder, exist_ok=True)
    return folder


def data_folders():
    return [("参考谱库", ref_dir()), ("RRUFF数据包", rruff_dir()), ("分析结果", results_dir())]


def pkg_kind(pkg):
    """数据包类型（只用于显示，随语言变化；内部判断仍用原值）。"""
    return T((pkg or {}).get("kind") or "")


def pkg_label(pkg, fallback=""):
    """数据包说明（只用于显示，随语言变化）。"""
    return T((pkg or {}).get("label") or fallback)


def dir_stats(path):
    count = 0
    size = 0
    for root, _dirs, files in os.walk(path):
        for name in files:
            try:
                size += os.path.getsize(os.path.join(root, name))
                count += 1
            except OSError:
                pass
    return count, size


def db_dir():
    return ref_dir()


def migrate_legacy_layout():
    legacy = os.path.join(_app_dir(), "光谱数据库")
    if not os.path.isdir(legacy):
        return []
    moved = []
    try:
        names = os.listdir(legacy)
    except OSError:
        return []
    for name in names:
        src = os.path.join(legacy, name)
        if os.path.isfile(src) and name.lower().endswith((".csv", ".txt", ".dat")):
            dst = os.path.join(ref_dir(), name)
            if not os.path.exists(dst):
                try:
                    shutil.move(src, dst)
                    moved.append(name)
                except OSError:
                    pass
    old_pkg = os.path.join(legacy, "_RRUFF包")
    if os.path.isdir(old_pkg):
        try:
            pkg_names = os.listdir(old_pkg)
        except OSError:
            pkg_names = []
        for name in pkg_names:
            src = os.path.join(old_pkg, name)
            dst = os.path.join(rruff_dir(), name)
            if os.path.isfile(src) and not os.path.exists(dst):
                try:
                    shutil.move(src, dst)
                    moved.append(name)
                except OSError:
                    pass
        try:
            if not os.listdir(old_pkg):
                os.rmdir(old_pkg)
        except OSError:
            pass
    try:
        if not os.listdir(legacy):
            os.rmdir(legacy)
    except OSError:
        pass
    return moved


def db_list():
    folder = ref_dir()
    out = []
    for name in sorted(os.listdir(folder)):
        if name.lower().endswith((".csv", ".txt", ".dat")):
            out.append(os.path.join(folder, name))
    return out


DEFAULT_CACHE_LIMIT_MB = 2048


def cache_limit_mb():
    """数据包目录的容量上限（MB），0 表示不限制。"""
    try:
        value = float(_load_settings().get("max_cache_mb", DEFAULT_CACHE_LIMIT_MB))
    except (TypeError, ValueError):
        value = float(DEFAULT_CACHE_LIMIT_MB)
    return value if value > 0 else 0.0


def set_cache_limit_mb(value):
    data = _load_settings()
    data["max_cache_mb"] = "%.0f" % max(0.0, float(value))
    return _save_settings(data)


def _size_text_to_bytes(text):
    s = str(text or "").strip().upper().replace(" ", "")
    try:
        if s.endswith("GB"):
            return float(s[:-2]) * 1073741824.0
        if s.endswith("MB"):
            return float(s[:-2]) * 1048576.0
        if s.endswith("KB"):
            return float(s[:-2]) * 1024.0
        if s.endswith("G"):
            return float(s[:-1]) * 1073741824.0
        if s.endswith("M"):
            return float(s[:-1]) * 1048576.0
        return float(s)
    except ValueError:
        return 0.0


def disk_free(path=None):
    """返回 (可用字节, 总字节)；无法判断时返回 (None, None)。"""
    target = path or data_root(create=False)
    while target and not os.path.isdir(target):
        parent = os.path.dirname(target)
        if parent == target:
            break
        target = parent
    try:
        usage = shutil.disk_usage(target or _app_dir())
        return usage.free, usage.total
    except OSError:
        return None, None


def check_space(need_bytes, path=None, margin_mb=20.0):
    """下载前预检：返回 (是否可继续, 提示文本)。"""
    free, _total = disk_free(path)
    if free is None:
        return True, ""
    margin = max(0.0, float(margin_mb)) * 1048576.0
    if need_bytes and free < need_bytes * 1.1 + margin:
        return False, ("磁盘剩余空间不足：本次约需 %.0f MB，可用仅 %.0f MB。\n"
                       "建议先清理旧数据包（设置 → 数据文件夹… → 清空），"
                       "或把数据文件夹改到空间更大的分区。"
                       % (need_bytes / 1048576.0, free / 1048576.0))
    return True, ""


def cache_usage():
    return dir_stats(rruff_dir())


def over_cache_limit():
    """返回 (是否超限, 上限MB, 当前字节数)。"""
    limit = cache_limit_mb()
    _n, size = cache_usage()
    return (bool(limit) and size > limit * 1048576.0), limit, size


def package_sizes():
    out = []
    for pkg in all_packages():
        path = rruff_zip_path(pkg["key"])
        if os.path.exists(path):
            try:
                out.append((pkg["key"], path, os.path.getsize(path)))
            except OSError:
                pass
    return out


def cleanup_packages(keep_keys=None, target_mb=0.0):
    """删除数据包 zip 及其索引释放空间；返回 (删除个数, 释放字节数)。"""
    removed = 0
    freed = 0
    need = max(0.0, float(target_mb)) * 1048576.0
    for key, path, size in package_sizes():
        if keep_keys and key in keep_keys:
            continue
        try:
            os.remove(path)
        except OSError:
            continue
        removed += 1
        freed += size
        for extra in (rruff_index_path(key), _peak_index_path(key)):
            try:
                if os.path.exists(extra):
                    os.remove(extra)
            except OSError:
                pass
        if need and freed >= need:
            break
    return removed, freed


def cache_limit_warning():
    over, limit, size = over_cache_limit()
    if not over:
        return None
    return ("RRUFF 数据包已占用 %.0f MB，超过设定的上限 %.0f MB。\n"
            "可在“设置 → 数据文件夹…”中清理不再需要的数据包，或调高上限。"
            % (size / 1048576.0, limit))


def write_peaks_table(dst, xs, ys, plot=None):
    """把峰列表（含相对强度、FWHM、突出度、可选矿物归属）写成 CSV。"""
    p = _plot_opts(plot)
    peaks = analyze_peaks(xs, ys, p)
    mineral = p.get("mineral_name")
    assign = {}
    if mineral and peaks:
        for pos, text in assign_bands([pk["x"] for pk in peaks], mineral):
            assign[round(pos, 4)] = text
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        head = T("峰位(cm-1),强度,相对强度(%),半高宽FWHM,峰突出度,来源")
        if mineral:
            head += "," + T("归属")
        f.write(head + "\n")
        for pk in peaks:
            row = "%.4f,%.6g,%.2f,%.4f,%.6g,%s" % (
                pk["x"], pk["y"], pk["rel"], pk["fwhm"], pk["prominence"],
                T("手动") if pk.get("manual") else T("自动"))
            if mineral:
                row += "," + assign.get(round(pk["x"], 4), "")
            f.write(row + "\n")
    return peaks


def export_generic(src, out_dir=None, formats=("csv", "png"), plot=None,
                   skip_existing=False):
    """把任意受支持的光谱文件导出为 CSV / PNG / JCAMP-DX / 峰列表。

    返回 (首个输出路径, [(label, xs, ys)], 是否跳过)。
    """
    xlabel, series = read_any_series(src)
    if not series:
        raise JwsError(T("没有读到任何光谱数据"))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        base_dir = out_dir
    else:
        base_dir = os.path.dirname(os.path.abspath(src))
    base = os.path.splitext(os.path.basename(src))[0]
    wanted = set(formats) or {"csv"}
    targets = []
    if "csv" in wanted:
        targets.append(os.path.join(base_dir, base + ".csv"))
    if "png" in wanted:
        targets.append(os.path.join(base_dir, base + ".png"))
    if "jcamp" in wanted:
        targets.append(os.path.join(base_dir, base + ".jdx"))
    if "peaks" in wanted:
        targets.append(os.path.join(base_dir, base + "_peaks.csv"))
    if skip_existing and targets and all(os.path.exists(p) for p in targets):
        return targets[0], series, True

    p = _plot_opts(plot)
    if "csv" in wanted:
        dst = os.path.join(base_dir, base + ".csv")
        with open(dst, "w", encoding="utf-8-sig", newline="") as f:
            f.write(",".join([xlabel] + [lab for lab, _x, _y in series]) + "\n")
            xs0 = series[0][1]
            for i in range(len(xs0)):
                cells = ["%.6f" % xs0[i]]
                for _lab, _xs, ys in series:
                    cells.append("%.6f" % ys[i] if i < len(ys) else "")
                f.write(",".join(cells) + "\n")
    if "png" in wanted:
        dst = os.path.join(base_dir, base + ".png")
        render_png(dst, [(lab, xs, ys, _PALETTE[i % len(_PALETTE)])
                         for i, (lab, xs, ys) in enumerate(series)],
                   os.path.basename(src) if p["show_title"] else "", xlabel, "Intensity", p)
    if "jcamp" in wanted:
        dst = os.path.join(base_dir, base + ".jdx")
        lab, xs, ys = series[0]
        write_jcamp(dst, xs, ys, {"SAMPLE": lab}, title=os.path.basename(src),
                    xunits="1/CM", yunits="ARBITRARY UNITS")
    if "peaks" in wanted:
        dst = os.path.join(base_dir, base + "_peaks.csv")
        write_peaks_table(dst, series[0][1], series[0][2], p)
    first = targets[0] if targets else os.path.join(base_dir, base + ".csv")
    return first, series, False


def batch_convert(folder, out_dir=None, formats=("xlsx", "png", "peaks"), plot=None,
                  skip_existing=False, progress=None):
    """整目录批处理：逐文件转换 + 峰识别 + 汇总统计表。

    返回 dict：ok, skip, fail, rows, errors, summary, out_dir
    """
    files = find_input_files([folder])
    ok = skip = fail = 0
    rows = []
    errors = []
    target_dir = out_dir or results_dir()
    os.makedirs(target_dir, exist_ok=True)
    for i, src in enumerate(files, 1):
        if progress:
            progress(i, len(files), src)
        try:
            if src.lower().endswith(".jws"):
                dst, spec, skipped = convert_file(src, out_dir=out_dir,
                                                  skip_existing=skip_existing,
                                                  formats=tuple(formats), plot=plot)
                if skipped:
                    skip += 1
                    continue
                xs = _calibrate(spec.x_values(), plot)
                ys = list(spec.y_data[0])
            else:
                _dst, series, skipped = export_generic(src, out_dir=out_dir,
                                                       formats=tuple(formats),
                                                       plot=plot,
                                                       skip_existing=skip_existing)
                if skipped:
                    skip += 1
                    continue
                xs, ys = series[0][1], series[0][2]
            peaks = analyze_peaks(xs, ys, plot)
            main = max(peaks, key=lambda pk: pk["prominence"]) if peaks else None
            rows.append({
                "file": os.path.basename(src),
                "path": src,
                "points": len(xs),
                "peaks": len(peaks),
                "x": main["x"] if main else None,
                "y": main["y"] if main else None,
                "rel": main["rel"] if main else None,
                "fwhm": main["fwhm"] if main else None,
            })
            ok += 1
        except Exception as exc:
            fail += 1
            errors.append((os.path.basename(src), str(exc)))
    summary = os.path.join(target_dir, "批处理汇总.csv")
    if rows:
        base = os.path.splitext(os.path.basename(os.path.abspath(folder)))[0] or "批量"
        summary = os.path.join(target_dir, "%s_批处理汇总.csv" % base)
        with open(summary, "w", encoding="utf-8-sig", newline="") as f:
            f.write(T("文件,数据点数,识别峰数,主峰位(cm-1),主峰强度,主峰相对强度(%),主峰半高宽FWHM\n"))
            for r in rows:
                f.write("%s,%d,%d,%s,%s,%s,%s\n" % (
                    r["file"], r["points"], r["peaks"],
                    "" if r["x"] is None else "%.2f" % r["x"],
                    "" if r["y"] is None else "%.6g" % r["y"],
                    "" if r["rel"] is None else "%.1f" % r["rel"],
                    "" if r["fwhm"] is None else "%.2f" % r["fwhm"]))
    return {"ok": ok, "skip": skip, "fail": fail, "rows": rows, "errors": errors,
            "summary": summary if rows else "", "out_dir": target_dir, "total": len(files)}


def _b64_png(path):
    try:
        with open(path, "rb") as f:
            return "data:image/png;base64," + base64.b64encode(f.read()).decode("ascii")
    except OSError:
        return ""


def _html_escape(text):
    return (str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def write_report_html(path, title, subtitle="", info=None, images=None, tables=None,
                      notes=None):
    """生成自包含 HTML 分析报告（图片内嵌 base64，可浏览器打印为 PDF 或 Word 打开）。"""
    info = info or []
    images = images or []
    tables = tables or []
    notes = notes or []
    parts = [
        "<!DOCTYPE html>",
        "<html lang=\"%s\"><head><meta charset=\"utf-8\">"
        % ("zh-CN" if ui_lang() == "zh" else "en"),
        "<title>%s</title>" % _html_escape(title),
        "<style>",
        "body{font-family:'Microsoft YaHei',SimSun,Arial,sans-serif;margin:32px 40px;color:#222;}",
        "h1{font-size:24px;border-bottom:2px solid #1a4a8a;padding-bottom:8px;margin-bottom:6px;}",
        ".sub{color:#666;font-size:13px;margin-bottom:18px;}",
        "h2{font-size:17px;color:#1a4a8a;margin-top:26px;margin-bottom:8px;}",
        "table{border-collapse:collapse;margin:8px 0 16px 0;font-size:13px;}",
        "th,td{border:1px solid #bbb;padding:5px 10px;text-align:left;}",
        "th{background:#eef3fa;}",
        "tr:nth-child(even) td{background:#fafafa;}",
        "img{max-width:100%;border:1px solid #ddd;margin:6px 0 18px 0;}",
        ".kv{font-size:14px;line-height:1.9;}",
        ".note{color:#666;font-size:12px;margin-top:6px;}",
        "ul{font-size:13px;line-height:1.9;}",
        "</style></head><body>",
        "<h1>%s</h1>" % _html_escape(title),
    ]
    if subtitle:
        parts.append("<div class=\"sub\">%s</div>" % _html_escape(subtitle))
    if info:
        parts.append("<div class=\"kv\">")
        kv_sep = "：" if ui_lang() == "zh" else ": "
        for key, value in info:
            parts.append("<div><b>%s%s</b>%s</div>"
                         % (_html_escape(key), kv_sep, _html_escape(value)))
        parts.append("</div>")
    for heading, rows in tables:
        parts.append("<h2>%s</h2>" % _html_escape(heading))
        if not rows:
            parts.append("<div class=\"note\">%s</div>" % _html_escape(T("（无数据）")))
            continue
        parts.append("<table>")
        parts.append("<tr>" + "".join("<th>%s</th>" % _html_escape(c) for c in rows[0]) + "</tr>")
        for row in rows[1:]:
            parts.append("<tr>" + "".join("<td>%s</td>" % _html_escape(c) for c in row) + "</tr>")
        parts.append("</table>")
    for heading, img_path in images:
        parts.append("<h2>%s</h2>" % _html_escape(heading))
        data = _b64_png(img_path)
        if data:
            parts.append("<img src=\"%s\" alt=\"%s\">" % (data, _html_escape(heading)))
        else:
            parts.append("<div class=\"note\">%s</div>"
                         % _html_escape(T("图片缺失：%s") % img_path))
    if notes:
        parts.append("<h2>%s</h2><ul>" % T("说明"))
        for line in notes:
            parts.append("<li>%s</li>" % _html_escape(line))
        parts.append("</ul>")
    parts.append("</body></html>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(parts))
    return path


def collect_spectra(paths, plot=None, verbose=False):
    """把路径/文件夹读成 [(名称, xs, ys)]，自动处理多列与多通道。"""
    out = []
    for path in find_input_files(paths):
        try:
            xlabel, series = read_any_series(path)
        except Exception as exc:
            if verbose:
                print(T("  ! 跳过 %s：%s") % (os.path.basename(path), exc))
            continue
        if not series:
            continue
        xs = _calibrate(series[0][1], plot)
        for lab, _sxs, ys in series:
            name = os.path.basename(path)
            if len(series) > 1:
                name += "|" + lab
            out.append((name, xs, list(ys)))
    return out


def _safe_name(text):
    clean = "".join(c for c in str(text) if c not in '\\/:*?"<>|()')
    clean = clean.replace(" ", "_").strip("_")
    return clean[:60] or "item"


def build_report(paths, out_path=None, title=None, plot=None, verbose=False):
    """把若干光谱整理成一份自包含 HTML 报告（总览 + 逐条峰表 + 光谱图）。"""
    spectra = collect_spectra(paths, plot, verbose=verbose)
    if not spectra:
        raise JwsError(T("没有读到任何光谱，无法生成报告"))
    p = _plot_opts(plot)
    if out_path is None:
        out_path = os.path.join(results_dir(), "光谱分析报告.html")
    out_dir = os.path.dirname(os.path.abspath(out_path)) or os.getcwd()
    os.makedirs(out_dir, exist_ok=True)
    mineral = p.get("mineral_name")
    images = []
    tables = []
    summary = [[T("文件"), T("数据点数"), T("识别峰数"), T("主峰位(cm-1)"),
                T("主峰强度"), T("主峰FWHM")]]
    for name, xs, ys in spectra:
        peaks = analyze_peaks(xs, ys, p)
        main = max(peaks, key=lambda pk: pk["prominence"]) if peaks else None
        summary.append([
            name, str(len(xs)), str(len(peaks)),
            "" if main is None else "%.2f" % main["x"],
            "" if main is None else "%.5g" % main["y"],
            "" if main is None else "%.2f" % main["fwhm"],
        ])
        assign = {}
        if mineral and peaks:
            for pos, text in assign_bands([pk["x"] for pk in peaks], mineral):
                assign[round(pos, 4)] = text
        rows = [[T("峰位(cm-1)"), T("强度"), T("相对强度(%)"), T("半高宽FWHM"),
                 T("突出度"), T("来源")]
                + ([T("归属")] if mineral else [])]
        for pk in peaks:
            row = ["%.2f" % pk["x"], "%.5g" % pk["y"], "%.1f" % pk["rel"],
                   "%.2f" % pk["fwhm"], "%.5g" % pk["prominence"],
                   T("手动") if pk.get("manual") else T("自动")]
            if mineral:
                row.append(assign.get(round(pk["x"], 4), ""))
            rows.append(row)
        tables.append((T("%s · 峰表") % name, rows))
        try:
            png = os.path.join(out_dir, "_报告图_%s.png" % _safe_name(name))
            render_png(png, [(name, xs, ys, _PALETTE[0])], name,
                       _DEFAULT_X_HEADER, "Intensity", p)
            images.append((T("光谱图 · %s") % name, png))
        except Exception as exc:
            if verbose:
                print(T("  ! 出图失败 %s：%s") % (name, exc))
    tables.insert(0, (T("总览"), summary))

    params = []
    base_cn = {"none": T("无"), "linear": T("线性"), "poly2": T("多项式2阶"),
               "rolling": T("滚动最小值"), "iterpoly": T("迭代多项式"),
               "rolling_ball": T("滚动球")}
    if p.get("despike"):
        params.append(T("尖峰去除(窗口%d, 阈值%.1f)")
                      % (p["despike_window"], p["despike_thresh"]))
    if p["baseline"] != "none":
        params.append(T("基线校正:") + base_cn.get(p["baseline"], p["baseline"]))
    if int(p["smooth_window"]) >= 3:
        params.append(T("平滑:%s 窗口%d点") % (
            "SG" if p["smooth_mode"] == "sg" else T("移动平均"), p["smooth_window"]))
    if int(p["derivative"]):
        params.append(T("%d 阶导数") % int(p["derivative"]))
    if p["normalize"] != "none":
        params.append(T("归一化:") + str(p["normalize"]))
    if p.get("calib_pairs"):
        params.append(T("位移校准:%s") % p["calib_pairs"])
    params.append(T("峰识别阈值 %.1f%% 强度范围，最小峰间距 %.1f cm-1")
                  % (p["peak_thresh_pct"], p["peak_min_dist"]))
    if mineral:
        params.append(T("峰位归属参考：%s") % mineral)

    info = [(T("生成时间"), time.strftime("%Y-%m-%d %H:%M:%S")),
            (T("光谱条数"), str(len(spectra))),
            (T("数据文件夹"), data_root())]
    notes = [T("峰位、FWHM 由本工具自动识别，请在发表前人工核对。"),
             T("处理参数：") + ("；" if ui_lang() == "zh" else "; ").join(params)
             + ("。" if ui_lang() == "zh" else ".")]
    if mineral:
        notes.append(T("归属文本参考内置矿物特征峰表（容差 8 cm-1），仅供参考。"))
    notes.append(T("报告为单个 HTML 文件，图片已内嵌；可用浏览器“打印 → 另存为 PDF”，")
                 + _sp() + T("也可直接用 Word / WPS 打开。"))
    write_report_html(out_path, title or T("光谱分析报告"),
                      subtitle=T("由 拉曼光谱工具 生成 · 共 %d 条光谱") % len(spectra),
                      info=info, images=images, tables=tables, notes=notes)
    return out_path, len(spectra)


def _http_get(url, timeout=30, retries=3, max_bytes=32 * 1048576):
    """取一个文本响应（ROD 检索 / JDX 谱线）。

    带重试（网络抖动很常见）与体积上限（原来是无上限整体读进内存）。
    """
    last = None
    tries = max(1, int(retries))
    opener = make_opener()
    for attempt in range(1, tries + 1):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "SpectraTool/1.0"})
            with opener.open(req, timeout=timeout) as resp:
                raw = resp.read(max_bytes + 1)
            if len(raw) > max_bytes:
                raise JwsError(T("响应过大（超过 %d MB），已放弃")
                               % (max_bytes // 1048576))
            for enc in ("utf-8", "latin-1"):
                try:
                    return raw.decode(enc)
                except UnicodeDecodeError:
                    continue
            return raw.decode("utf-8", "ignore")
        except JwsError:
            raise
        except Exception as exc:
            last = exc
            if attempt < tries:
                time.sleep(1.5 * attempt)
    raise JwsError(T("网络请求失败（重试 %d 次）：%s") % (tries, last))


def _num(text):
    try:
        return float(str(text).strip())
    except (TypeError, ValueError):
        return None


def parse_jcamp(text):
    meta = {}
    xs, ys = [], []
    mode = None
    cur_x = None
    first_x = None
    for raw in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = raw.strip()
        if not line:
            continue
        if line.startswith("##"):
            head = line.upper()
            if head.startswith(("##XYPOINTS", "##PEAK TABLE", "##X++", "##Y++", "##XY..XY")):
                body = line.split("=", 1)[1].upper() if "=" in line else ""
                if "X++" in head or "X++" in body:
                    mode = "xpp"
                elif "Y++" in head or "Y++" in body:
                    mode = "ypp"
                else:
                    mode = "xy"
                continue
            key, _, value = line[2:].partition("=")
            key = key.strip().upper()
            meta[key] = value.strip()
            if key == "FIRSTX":
                first_x = _num(value)
                cur_x = first_x
            continue
        if mode is None:
            continue
        nums = []
        for token in line.replace(",", " ").split():
            value = _num(token)
            if value is not None:
                nums.append(value)
        if not nums:
            continue
        if mode == "xy":
            for i in range(0, len(nums) - 1, 2):
                xs.append(nums[i])
                ys.append(nums[i + 1])
            continue
        inc = _num(meta.get("DELTAX")) or _num(meta.get("XINCREMENT")) or 0.0
        if not inc:
            last = _num(meta.get("LASTX"))
            cnt = _num(meta.get("NPOINTS"))
            if last is not None and cnt and cnt > 1 and first_x is not None:
                inc = (last - first_x) / (cnt - 1)
        if cur_x is None:
            cur_x = first_x if first_x is not None else 0.0
        for value in nums:
            xs.append(cur_x)
            ys.append(value)
            cur_x += inc if inc else 1.0
    xf = _num(meta.get("XFACTOR"))
    yf = _num(meta.get("YFACTOR"))
    if mode == "xy" and xf and xf != 1.0:
        xs = [v * xf for v in xs]
    if yf and yf != 1.0:
        ys = [v * yf for v in ys]
    return xs, ys, meta


def rod_search(query, timeout=30):
    url = "%s/result.php?text=%s&format=json" % (_ROD_BASE, urllib.parse.quote(str(query)))
    text = _http_get(url, timeout).strip()
    if not text:
        return []
    try:
        data = json.loads(text)
    except ValueError:
        raise JwsError(T("数据库返回内容无法解析（可能网络被拦截）"))
    if isinstance(data, dict):
        data = [data]
    return data


def rod_fetch(rod_id, timeout=40):
    url = "%s/%s.jdx" % (_ROD_BASE, urllib.parse.quote(str(rod_id)))
    xs, ys, meta = parse_jcamp(_http_get(url, timeout))
    if len(xs) < 4:
        raise JwsError(T("未取得有效光谱数据：%s") % rod_id)
    # 传输被截断时 JDX 的 NPOINTS 会对不上；只判“收少了”，宁可报错也别把半截谱存进库
    expect = _num(meta.get("NPOINTS"))
    if expect and len(xs) < int(expect) - max(2, int(expect) * 0.02):
        raise JwsError(T("光谱数据不完整：应为 %d 点，实际收到 %d 点（可稍后重试）")
                       % (int(expect), len(xs)))
    return xs, ys, meta


def db_save_spectrum(rod_id, mineral, xs, ys, meta=None, xlabel=None, tag=None):
    meta = meta or {}
    label = "".join(c for c in str(mineral or "unknown") if c not in '\\/:*?"<>|').strip()
    label = (label or "unknown").replace(" ", "_")
    wave = _num(meta.get("WAVELENGTH"))
    suffix = "_%gnm" % wave if wave else ""
    plain = (not tag) or tag in ("拉曼", "Raman")
    prefix = "" if plain else ("%s_" % tag)
    folder = db_dir()
    if not plain:
        folder = os.path.join(folder, str(tag))
        os.makedirs(folder, exist_ok=True)
    dst = os.path.join(folder, "%s%s_%s%s.csv" % (prefix, label, rod_id, suffix))
    tmp = dst + ".part"
    try:
        with open(tmp, "w", encoding="utf-8-sig", newline="") as f:
            f.write("%s,Intensity\n" % (xlabel or "Raman shift (cm-1)"))
            for x, y in zip(xs, ys):
                f.write("%.4f,%.6g\n" % (x, y))
        # 写完才转正：否则写盘中途出错会留下半截 CSV，还会被 db_list 当成正常谱进库
        os.replace(tmp, dst)
    except OSError:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise
    return dst


def rod_meta(rod_id, timeout=30):
    url = "%s/result.php?id=%s&format=json" % (_ROD_BASE, urllib.parse.quote(str(rod_id)))
    text = _http_get(url, timeout).strip()
    if not text:
        return {}
    try:
        data = json.loads(text)
    except ValueError:
        return {}
    if isinstance(data, list):
        return data[0] if data else {}
    return data if isinstance(data, dict) else {}


def db_download(rod_id, mineral=None, wavelength=None, timeout=40):
    if not mineral:
        try:
            info = rod_meta(rod_id, timeout)
            mineral = info.get("mineral") or info.get("chemname")
            wavelength = wavelength or info.get("wavelength")
        except Exception:
            pass
    xs, ys, meta = rod_fetch(rod_id, timeout)
    if wavelength:
        meta["WAVELENGTH"] = str(wavelength)
    return db_save_spectrum(rod_id, mineral or meta.get("TITLE") or rod_id,
                            xs, ys, meta), meta


def db_match(target_xy, top=None):
    return match_references(target_xy, db_list(), top=top)


RRUFF_PACKAGES = [
    {"key": "unrated_unoriented", "file": "unrated_unoriented.zip",
     "label": "未评级 · 非定向", "size": "12 MB"},
    {"key": "poor_unoriented", "file": "poor_unoriented.zip",
     "label": "poor · 非定向", "size": "33 MB"},
    {"key": "unrated_oriented", "file": "unrated_oriented.zip",
     "label": "未评级 · 定向", "size": "39 MB"},
    {"key": "fair_oriented", "file": "fair_oriented.zip",
     "label": "fair · 定向", "size": "271 KB"},
    {"key": "fair_unoriented", "file": "fair_unoriented.zip",
     "label": "fair · 非定向", "size": "59 MB"},
    {"key": "excellent_oriented", "file": "excellent_oriented.zip",
     "label": "excellent · 定向", "size": "77 MB"},
    {"key": "excellent_unoriented", "file": "excellent_unoriented.zip",
     "label": "excellent · 非定向", "size": "229 MB"},
    {"key": "LR-Raman", "file": "LR-Raman.zip",
     "label": "低分辨率 LR", "size": "227 MB"},
]

_RRUFF_BASE = "https://www.rruff.net/zipped_data_files/raman/"


RRUFF_EXTRA = [
    {"key": "infrared_RAW", "label": "红外光谱 IR（RAW）", "size": "14 MB",
     "url": "https://www.rruff.net/zipped_data_files/infrared/RAW.zip",
     "kind": "IR", "xlabel": "Wavenumber (cm-1)"},
    {"key": "powder_DIF", "label": "XRD 峰位列表（DIF）", "size": "7.6 MB",
     "url": "https://www.rruff.net/zipped_data_files/powder/DIF.zip",
     "kind": "XRD", "xlabel": "d spacing (A)"},
    {"key": "powder_XY_Processed", "label": "XRD 粉末衍射（XY 处理后）", "size": "67 MB",
     "url": "https://www.rruff.net/zipped_data_files/powder/XY_Processed.zip",
     "kind": "XRD", "xlabel": "2-Theta (deg)"},
    {"key": "chemistry_Microprobe", "label": "化学成分（电子探针）", "size": "19 MB",
     "url": "https://www.rruff.net/zipped_data_files/chemistry/Microprobe_Data.zip",
     "kind": "成分", "xlabel": ""},
]

_ALL_PACKAGES_CACHE = [None]


def all_packages():
    """拉曼 + 红外 + XRD + 化学成分 的全部数据包。"""
    if _ALL_PACKAGES_CACHE[0] is None:
        merged = []
        for pkg in RRUFF_PACKAGES:
            item = dict(pkg)
            item.setdefault("kind", "拉曼")
            item.setdefault("xlabel", "Raman shift (cm-1)")
            merged.append(item)
        for pkg in RRUFF_EXTRA:
            item = dict(pkg)
            item.setdefault("kind", item.get("kind", "其他"))
            merged.append(item)
        _ALL_PACKAGES_CACHE[0] = merged
    return _ALL_PACKAGES_CACHE[0]


def _package_url(pkg):
    if pkg.get("url"):
        return pkg["url"]
    return _RRUFF_BASE + pkg.get("file", "")


def package_bytes(pkg):
    if pkg.get("bytes"):
        return float(pkg["bytes"])
    size = _size_text_to_bytes(pkg.get("size"))
    if size:
        return size
    try:
        req = urllib.request.Request(_package_url(pkg),
                                     headers={"User-Agent": "SpectraTool/1.0"})
        req.get_method = lambda: "HEAD"
        with make_opener().open(req, timeout=20) as resp:
            return float(resp.headers.get("Content-Length") or 0)
    except Exception:
        return 0.0


# ---------------------------------------------------------------------------
# 下载引擎：断点续传 + 并发分片 + 自动重试 + 完整性校验
#
# 为什么并发要开这么大（实网实测，2026-09）：
#   到 rruff.net（OVH 加拿大）ping 丢包 25%、RTT 236 ms。这种“长肥丢包链路”下
#   单个 TCP 连接会被拥塞控制压死，实测 1 路只有 0.027 MB/s；并发数几乎就是吞吐量：
#       1 路 0.027 / 8 路 0.32 / 16 路 0.41 / 32 路 0.89 / 64 路 1.28 / 96 路 1.45 MB/s
#   227 MB 的包：1 路要 140 分钟，32 路约 4 分钟，64 路约 3 分钟（服务器全程无拒绝）。
#   所以默认给 32 路，可用 --dl-conns / 高级设置里调（1~128）。
# 另一条更快的路：如果你有代理 / VPN，把地址填进设置（或打开系统代理），
#   走代理能绕开这条高丢包链路，往往比堆并发快得多。
# ---------------------------------------------------------------------------

DOWNLOAD_CHUNK = 1 << 16
DOWNLOAD_CONNECTIONS = 32
DOWNLOAD_CONNECTIONS_MAX = 128
DOWNLOAD_RETRIES = 5
DOWNLOAD_TIMEOUT = 60
# 工作单元（work unit）大小：文件切成一个个小段放进队列，
# 谁先下完谁再领下一段，免得总耗时被最慢的那一条连接拖住。
# 单元数必须 ≫ 连接数（否则没得可领），但也不能太碎（每个单元要新开一条连接，
# 跨境链路上一次握手约 0.5 s）。实测吞吐几乎正比于同时在跑的连接数，
# 所以「连接数」这一项不能因为文件小就被砍掉。
DOWNLOAD_WORK_UNIT = 2 << 20              # 单个工作单元上限
DOWNLOAD_MIN_UNIT = 64 * 1024             # 单个工作单元下限
DOWNLOAD_UNITS_PER_LANE = 4               # 每条连接分几个单元，留出领活余地
DOWNLOAD_MIN_LANE_BYTES = 256 * 1024      # 每条连接至少分到这么多活，否则少开几路
DOWNLOAD_MAX_UNITS = 512
_UA = {"User-Agent": "SpectraTool/1.0"}
# 默认下行速率（字节/秒），仅用于“预计耗时”提示；每次真实下载完会写回实测值。
DEFAULT_DOWNLOAD_BPS = 1.0 * 1048576.0


class DownloadCancelled(Exception):
    """用户在下载过程中点了取消。"""


def download_connections():
    """并发连接数：设置里可调，默认 DOWNLOAD_CONNECTIONS。"""
    try:
        value = int(float(_load_settings().get("download_conns") or 0))
    except (TypeError, ValueError):
        value = 0
    if value <= 0:
        return DOWNLOAD_CONNECTIONS
    return max(1, min(value, DOWNLOAD_CONNECTIONS_MAX))


def download_lanes(connections, total):
    """真正开几条连接。

    实测吞吐几乎正比于同时在跑的连接数（32 路 0.89、64 路 1.28、96 路 1.45 MB/s），
    所以只要文件装得下就一路给到 connections；文件太小才收敛，
    否则 271 KB 的包也会去开 32 条连接，全耗在握手上。
    """
    if total <= 0:
        return max(1, int(connections))
    fit = max(1, int(total // DOWNLOAD_MIN_LANE_BYTES))
    return max(1, min(int(connections), fit))


def download_units(total, lanes):
    """切成多少个工作单元。

    目标：每条连接分到 DOWNLOAD_UNITS_PER_LANE 个单元，这样才有“谁先下完谁再领”
    的余地；单元大小被夹在 [DOWNLOAD_MIN_UNIT, DOWNLOAD_WORK_UNIT] 之间。
    """
    if total <= 0:
        return 1
    want = max(1, int(lanes) * DOWNLOAD_UNITS_PER_LANE)
    if total // want < DOWNLOAD_MIN_UNIT:
        want = max(1, int(total // DOWNLOAD_MIN_UNIT))
    unit = min(DOWNLOAD_WORK_UNIT,
               max(DOWNLOAD_MIN_UNIT, total // max(1, want)))
    return max(1, min(int(math.ceil(total / float(unit))), DOWNLOAD_MAX_UNITS))


def download_spans(total, units):
    """工作单元 → (起始字节, 结束字节) 列表；total 未知时返回单段读到 EOF。"""
    if total <= 0:
        return [(0, None)]
    return [(total * i // units, total * (i + 1) // units - 1)
            for i in range(units)]


def system_proxy():
    """读 Windows 系统代理（“设置 → 网络 → 代理”）。没开或读不到就返回 ""。"""
    if os.name != "nt":
        return ""
    try:
        import winreg
    except ImportError:
        return ""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Internet Settings")
    except OSError:
        return ""
    try:
        try:
            enable, _ = winreg.QueryValueEx(key, "ProxyEnable")
            server, _ = winreg.QueryValueEx(key, "ProxyServer")
        except OSError:
            return ""
    finally:
        try:
            winreg.CloseKey(key)
        except OSError:
            pass
    if not enable:
        return ""
    server = str(server or "").strip()
    if "=" in server:                 # 形如 http=host:port;https=host:port
        parts = dict(p.split("=", 1) for p in server.split(";") if "=" in p)
        server = (parts.get("https") or parts.get("http") or "").strip()
    if not server:
        return ""
    return server if "://" in server else "http://" + server


def proxy_url():
    """实际用的代理：设置里填了用它，否则用系统代理；都没有返回 ""（直连）。

    设置里填 none / off / direct 可以强制直连（忽略系统代理）。
    """
    explicit = str(_load_settings().get("proxy") or "").strip()
    if explicit:
        if explicit.lower() in ("none", "off", "direct", "0", "关闭", "直连"):
            return ""
        return explicit if "://" in explicit else "http://" + explicit
    return system_proxy()


def make_opener():
    """按当前代理设置生成 opener；直连时显式绕开环境变量里的代理。

    高丢包链路上并发是唯一能提吞吐的手段，所以这里一次生成、多个线程复用。
    """
    proxy = proxy_url()
    if proxy:
        handler = urllib.request.ProxyHandler({"http": proxy, "https": proxy})
    else:
        handler = urllib.request.ProxyHandler({})
    return urllib.request.build_opener(handler)


def download_speed_bps():
    """上次实测到的下行速率（字节/秒），没有记录时用默认值。"""
    try:
        value = float(_load_settings().get("download_bps") or 0)
    except (TypeError, ValueError):
        value = 0.0
    return value if value > 1024 else DEFAULT_DOWNLOAD_BPS


def remember_download_speed(bps):
    """把本次实测速率记进设置，让下次的“预计耗时”更准。"""
    if not bps or bps <= 1024:
        return
    data = _load_settings()
    try:
        old = float(data.get("download_bps") or 0)
    except (TypeError, ValueError):
        old = 0.0
    if old and abs(bps - old) / max(old, 1.0) < 0.25:
        return                      # 变化不大就别反复写盘
    data["download_bps"] = "%d" % int(bps)
    _save_settings(data)


def estimate_download_seconds(size_bytes, bps=None):
    """按实测速率估算下载耗时（秒）；大小未知时返回 None。"""
    if not size_bytes:
        return None
    rate = bps or download_speed_bps()
    if rate <= 0:
        return None
    return float(size_bytes) / rate


def format_eta(seconds):
    """把秒数写成“不到 1 分钟 / 约 6 分钟 / 约 1 小时 5 分钟”。"""
    if seconds is None:
        return ""
    sec = int(max(0, round(seconds)))
    if sec < 60:
        return T("不到 1 分钟")
    minutes = sec / 60.0
    if minutes < 60:
        return T("约 %.0f 分钟") % max(1.0, round(minutes))
    hours = int(minutes // 60)
    rest = int(round(minutes - hours * 60))
    if rest <= 0:
        return T("约 %d 小时") % hours
    return T("约 %d 小时 %d 分钟") % (hours, rest)


def format_mb(nbytes):
    return "%.1f MB" % (float(nbytes) / 1048576.0)


class DownloadMeter:
    """把“已收 / 总数”换算成瞬时速率与剩余时间，供界面显示。"""

    def __init__(self, window=8.0):
        self.window = float(window)
        self.samples = []
        self.t0 = None

    def update(self, got, total=0):
        """返回 (速率 字节/秒, 剩余秒数或 None)。"""
        now = time.time()
        if self.t0 is None:
            self.t0 = now
        self.samples.append((now, got))
        cutoff = now - self.window
        while len(self.samples) > 2 and self.samples[0][0] < cutoff:
            self.samples.pop(0)
        t_a, n_a = self.samples[0]
        span = now - t_a
        speed = (got - n_a) / span if span > 0.35 else 0.0
        eta = ((total - got) / speed) if (total and speed > 0 and got < total) else None
        return speed, eta

    def average(self, got):
        if self.t0 is None or got <= 0:
            return 0.0
        return got / max(0.001, time.time() - self.t0)


def _probe_download(url, timeout=30, opener=None):
    """探测远端大小与是否支持 Range，返回 (总字节, 是否支持断点续传)。"""
    total = 0
    ranged = False
    op = opener or make_opener()
    try:
        req = urllib.request.Request(url, headers=dict(_UA))
        req.get_method = lambda: "HEAD"
        with op.open(req, timeout=timeout) as r:
            try:
                total = int(r.headers.get("Content-Length") or 0)
            except (TypeError, ValueError):
                total = 0
            ranged = str(r.headers.get("Accept-Ranges") or "").lower() == "bytes"
    except Exception:
        return 0, False
    if not total:
        return 0, False
    if ranged:
        return total, True
    # 有的服务器不报 Accept-Ranges 但其实支持，试一次 bytes=0-0
    try:
        req = urllib.request.Request(url, headers=dict(_UA, **{"Range": "bytes=0-0"}))
        with op.open(req, timeout=timeout) as r:
            if r.status != 206:
                return total, False
            content_range = str(r.headers.get("Content-Range") or "")
            r.read(1)
        if "/" in content_range:
            try:
                total = int(content_range.rsplit("/", 1)[1])
            except ValueError:
                pass
        return total, True
    except Exception:
        return total, False


def download_file(url, dst, progress=None, cancel=None,
                  connections=None, timeout=60, retries=DOWNLOAD_RETRIES):
    """下载 url 到 dst：断点续传 + 并发分片 + 自动重试 + 完整性校验。

    · 续传：分片写到 dst.part0 / .part1 …，下次接着下；中途取消或断网都不白费。
    · 并发：默认 download_connections()（32 路）。高丢包跨境链路下并发数几乎就是吞吐量，
      实测 1 路 0.027 MB/s、32 路 0.89 MB/s、96 路 1.45 MB/s；文件太小（<256 KB/路）才少开。
      分片按固定大小的“工作单元”（download_units）切，谁下完谁再领下一段，
      不是把文件平均分成 N 份 —— 后者总耗时等于最慢的那一份。
      单元数记在 .part.meta 里，中途改并发数也能接着续传。
    · 代理：按 proxy_url() 走（设置里的代理 > Windows 系统代理 > 直连）。
    · 重试：分片内失败按 1.5×n 秒退避重试，且从该分片已收到的位置继续。
    · 校验：探到 Content-Length 就必须收满，收不满不转正、抛错并保留分片。
    · 原子：全部收满才拼成 dst.part.join 再 os.replace 到 dst。
    progress(got, total, speed, eta) 可能被多个线程调用（界面侧只存值、由主线程渲染）；
    cancel 传 threading.Event，置位后抛 DownloadCancelled。
    """
    dst = os.path.abspath(dst)
    parent = os.path.dirname(dst)
    if parent:
        os.makedirs(parent, exist_ok=True)

    if connections is None:
        connections = download_connections()

    def cancelled():
        return cancel is not None and cancel.is_set()

    if cancelled():
        raise DownloadCancelled()

    opener = make_opener()
    total, ranged = _probe_download(url, timeout=timeout, opener=opener)

    # 旧版留下的是单个 .part（从 0 开始的连续前缀），迁移成第 0 个分片，别浪费
    legacy = dst + ".part"
    if os.path.exists(legacy) and not os.path.exists(dst + ".part0"):
        try:
            os.replace(legacy, dst + ".part0")
        except OSError:
            pass

    def part_path(i):
        return "%s.part%d" % (dst, i)

    # 切法（单元数）记在 .part.meta 里：并发数改了、文件大小没变时沿用上次的切法，
    # 已有的 .partN 才对得上号，断点续传才有意义。
    meta_path = dst + ".part.meta"
    old_units = 0
    if os.path.exists(meta_path):
        try:
            with open(meta_path, encoding="utf-8") as f:
                old = f.read().splitlines()
        except OSError:
            old = []
        # 元信息里记着 url / 总长 / 单元数；三者对得上才认旧分片
        if bool(old) and len(old) > 2 and old[0] == url and old[1] == str(total):
            try:
                old_units = int(old[2])
            except ValueError:
                old_units = 0
        if old_units <= 0:
            # 远端文件换了（大小变了）或元信息不全 → 丢弃旧分片，免得拼出个坏文件
            for i in range(0, DOWNLOAD_MAX_UNITS + 1):
                try:
                    os.remove(part_path(i))
                except OSError:
                    pass

    if total > 0 and ranged and old_units > 0:
        # 续传：无条件沿用上次的切法（哪怕这次并发数不一样），否则分片编号会错位、
        # 拼出来的文件长度就不对了。并发数只影响同时开几条连接。
        units = min(old_units, DOWNLOAD_MAX_UNITS)
        lanes = max(1, min(int(connections), units))
    elif ranged and total > 0 and connections > 1:
        lanes = download_lanes(connections, total)
        units = download_units(total, lanes)
        lanes = max(1, min(lanes, units))
    else:
        lanes, units = 1, 1

    try:
        with open(meta_path, "w", encoding="utf-8") as f:
            f.write("%s\n%d\n%d\n" % (url, total, units))
    except OSError:
        pass

    spans = download_spans(total, units)

    counts = [0] * units
    for i, (start, end) in enumerate(spans):
        path = part_path(i)
        if not os.path.exists(path):
            continue
        have = os.path.getsize(path)
        limit = None if end is None else (end - start + 1)
        if limit is not None and have > limit:    # 分片比应有长度还长：截断重来
            try:
                os.truncate(path, limit)
            except OSError:
                try:
                    os.remove(path)
                except OSError:
                    pass
            have = limit
        counts[i] = have                          # 已收到的计入总进度

    lock = threading.Lock()
    meter = DownloadMeter()
    failures = []
    throttled = []          # 被服务器限流（429/503）的次数，用来给出“调低并发”的提示

    def report():
        got = sum(counts)
        speed, eta = meter.update(got, total)
        if progress:
            progress(got, total, speed, eta)

    def fetch(i, start, end):
        path = part_path(i)
        attempt = 0
        while True:
            if cancelled():
                raise DownloadCancelled()
            have = os.path.getsize(path) if os.path.exists(path) else 0
            limit = None if end is None else (end - start + 1)
            if limit is not None and have >= limit:
                return                            # 这一片已经齐了
            if have and not ranged:
                # 服务器不支持续传，只能从 0 重来
                try:
                    os.truncate(path, 0)
                except OSError:
                    try:
                        os.remove(path)
                    except OSError:
                        pass
                with lock:
                    counts[i] = 0
                have = 0
            headers = dict(_UA)
            need_range = (units > 1) or (start + have) > 0
            if need_range:
                headers["Range"] = "bytes=%d-%s" % (
                    start + have, "" if end is None else str(end))
            try:
                req = urllib.request.Request(url, headers=headers)
                with opener.open(req, timeout=timeout) as r:
                    if need_range and r.status != 206:
                        raise JwsError(T("服务器不支持断点续传（HTTP %s）") % r.status)
                    with open(path, "ab") as f:
                        while True:
                            if cancelled():
                                raise DownloadCancelled()
                            chunk = r.read(DOWNLOAD_CHUNK)
                            if not chunk:
                                break
                            f.write(chunk)
                            with lock:
                                counts[i] += len(chunk)
                                report()
                # 服务端可能提前断流：长度没到就不能当成下完了
                if limit is not None:
                    now_have = os.path.getsize(path) if os.path.exists(path) else 0
                    if now_have < limit:
                        raise JwsError(T("连接提前中断（本段还差 %d 字节）")
                                       % (limit - now_have))
                return
            except DownloadCancelled:
                raise
            except urllib.error.HTTPError as exc:
                if exc.code in (429, 503):
                    with lock:
                        throttled.append(exc.code)
                attempt += 1
                if cancelled():
                    raise DownloadCancelled()
                if attempt >= retries:
                    with lock:
                        failures.append((i, exc))
                    return
                time.sleep(min(8.0, 1.5 * attempt))
            except Exception as exc:
                attempt += 1
                if cancelled():
                    raise DownloadCancelled()
                if attempt >= retries:
                    with lock:
                        failures.append((i, exc))
                    return
                time.sleep(min(8.0, 1.5 * attempt))

    def lane():
        """一条连接：下完手上的工作单元就回队列再领一个。

        静态均分（每连接固定一段）时总耗时等于最慢那一段，实测 33 MB 的包
        最快段 4.3 s、最慢段 79 s，快连接干完只能干等；改成领活之后，
        快连接会把后面的段一起吃掉，慢尾被摊平。
        取消异常在这里就地吞掉（join 之后主流程统一判断），
        否则线程里抛出会打到 stderr，看着像崩溃。
        """
        while True:
            if cancelled():
                return
            try:
                i = job_queue.get_nowait()
            except queue.Empty:
                return
            try:
                fetch(i, spans[i][0], spans[i][1])
            except DownloadCancelled:
                return

    job_queue = queue.Queue()
    for i in range(units):
        job_queue.put(i)

    if lanes <= 1 and units == 1:
        fetch(0, spans[0][0], spans[0][1])
    else:
        threads = [threading.Thread(target=lane, daemon=True)
                   for _ in range(lanes)]
        for t in threads:
            t.start()
        for t in threads:
            t.join()

    if cancelled():
        raise DownloadCancelled()
    if failures:
        hint = ""
        if throttled:
            hint = (T("　服务器多次限流（HTTP %s），请把并发连接数调小些再试。")
                    % throttled[0])
        raise JwsError(T("下载失败（重试 %d 次仍未完成）：%s")
                       % (retries, failures[0][1]) + hint)

    def cleanup():
        for i in range(units):
            try:
                os.remove(part_path(i))
            except OSError:
                pass
        try:
            os.remove(meta_path)
        except OSError:
            pass

    joined = dst + ".part.join"
    try:
        with open(joined, "wb") as out:
            for i in range(units):
                path = part_path(i)
                if not os.path.exists(path):
                    continue
                with open(path, "rb") as f:
                    shutil.copyfileobj(f, out, DOWNLOAD_CHUNK)
    except OSError as exc:
        raise JwsError(T("合并下载分片失败：%s") % exc)
    size = os.path.getsize(joined)
    if total and size != total:
        try:
            os.remove(joined)
        except OSError:
            pass
        raise JwsError(T("下载不完整：应为 %d 字节，实际收到 %d 字节（已保留分片，可重试续传）")
                       % (total, size))
    if ranged:                                     # 单连接 + 支持 Range 时也可续传
        remember_download_speed(meter.average(size) or 0.0)
    os.replace(joined, dst)
    cleanup()
    return dst


def rruff_package(key):
    for pkg in all_packages():
        if pkg["key"] == key:
            return pkg
    return None


def rruff_zip_path(key):
    return os.path.join(rruff_dir(), key + ".zip")


def rruff_index_path(key):
    return os.path.join(rruff_dir(), key + ".index.csv")


def rruff_download(key, progress=None, timeout=None, check_disk=True, cancel=None,
                   connections=None):
    """下载 RRUFF 数据包 zip。

    走 download_file：断点续传 + 并发分片 + 自动重试 + 长度校验；
    下完再验一次 zip 能不能打开（长度对得上不等于压缩包没坏）。
    """
    pkg = rruff_package(key)
    if pkg is None:
        raise JwsError(T("未知的 RRUFF 数据包：%s") % key)
    dst = rruff_zip_path(key)
    if check_disk:
        need = package_bytes(pkg)
        ok, msg = check_space(need, rruff_dir())
        if not ok:
            raise JwsError(msg)
    download_file(_package_url(pkg), dst, progress=progress, cancel=cancel,
                  connections=connections,
                  timeout=timeout or DOWNLOAD_TIMEOUT)
    if not zipfile.is_zipfile(dst):
        try:
            os.remove(dst)
        except OSError:
            pass
        raise JwsError(T("下载的压缩包打不开（传输中损坏），已删除，请重新下载。"))
    return dst


def rruff_import_zip(src, key=None, progress=None):
    """导入用户自己下载的 RRUFF 数据包 zip，并建立索引（返回 (key, 索引条数)）。"""
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        raise JwsError(T("文件不存在：%s") % src)
    if key is None:
        key = os.path.splitext(os.path.basename(src))[0]
    safe = "".join(c for c in key if c not in '\\/:*?"<>|').strip() or "imported"
    dst = rruff_zip_path(safe)
    if os.path.abspath(src) != os.path.abspath(dst):
        shutil.copy2(src, dst)
    try:
        rows = rruff_index(safe, rebuild=True, progress=progress)
    except Exception as exc:
        raise JwsError(T("导入成功但建立索引失败：%s") % exc)
    return safe, len(rows)


def _read_index(key):
    path = rruff_index_path(key)
    if not os.path.exists(path):
        return None
    rows = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        for line in f:
            line = line.rstrip("\n")
            if not line:
                continue
            parts = line.split(",")
            if len(parts) >= 5:
                rows.append([parts[0], parts[1], parts[2], parts[3], ",".join(parts[4:])])
    return rows


def _parse_rruff_headers(text):
    meta = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.startswith("##"):
            key, _, value = line[2:].partition("=")
            meta[key.strip().upper()] = value.strip()
            continue
        if line[0].isdigit() or line[0] == "-":
            break
    return meta


def parse_rruff_text(text):
    xs, ys = [], []
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.replace(",", " ").replace("\t", " ").replace(";", " ").split()
        if len(parts) < 2:
            continue
        try:
            x = float(parts[0])
            y = float(parts[1])
        except ValueError:
            continue
        xs.append(x)
        ys.append(y)
    if len(xs) > 1 and xs[0] > xs[-1]:
        xs.reverse()
        ys.reverse()
    return xs, ys


def _parse_rruff_name(name):
    """从 RRUFF 包内文件名解析 矿物 / 编号 / 类型（用于非 XY 文本条目）。"""
    stem = os.path.basename(name).split(".")[0]
    parts = [p for p in stem.split("__")]
    mineral = parts[0] if parts else ""
    rid = parts[1] if len(parts) > 1 else ""
    dtype = parts[2] if len(parts) > 2 else ""
    extra = parts[3] if len(parts) > 3 else ""
    return mineral, rid, dtype, extra


def _rruff_wavelength(meta, kind=None):
    for key in ("RAMAN WAVELENGTH", "INFRARED WAVELENGTH", "XRAY WAVELENGTH",
                "WAVELENGTH", "WAVELENGTH (NM)"):
        if meta.get(key):
            return meta[key]
    return ""


def rruff_index(key, progress=None, rebuild=False):
    rows = None if rebuild else _read_index(key)
    if rows:
        return rows
    path = rruff_zip_path(key)
    if not os.path.exists(path):
        raise JwsError(T("尚未下载数据包：%s") % key)
    rows = []
    with zipfile.ZipFile(path) as z:
        names = sorted(n for n in z.namelist() if not n.endswith("/"))
        for i, name in enumerate(names):
            if progress:
                progress(i, len(names))
            low = name.lower()
            if low.endswith((".txt", ".jdx", ".dx", ".csv", ".xy", ".dat")):
                try:
                    with z.open(name) as fh:
                        head = fh.read(2000).decode("utf-8", "ignore")
                except Exception:
                    continue
                meta = _parse_rruff_headers(head)
                mineral = meta.get("NAMES", "")
                rid = meta.get("RRUFFID", "")
                ftype = meta.get("FILETYPE", "")
                wave = _rruff_wavelength(meta)
                if not mineral or not rid:
                    pm, pr, pd_, pe = _parse_rruff_name(name)
                    mineral = mineral or pm
                    rid = rid or pr
                    ftype = ftype or (pd_ + "/" + pe).strip("/")
                rows.append([mineral, rid, wave, ftype, name])
            else:
                mineral, rid, dtype, extra = _parse_rruff_name(name)
                if not mineral:
                    continue
                rows.append([mineral, rid, "", ("%s/%s" % (dtype, extra)).strip("/"), name])
    with open(rruff_index_path(key), "w", encoding="utf-8-sig", newline="") as f:
        for row in rows:
            f.write(",".join(row) + "\n")
    return rows


def rruff_available_keys():
    return [p["key"] for p in all_packages() if os.path.exists(rruff_zip_path(p["key"]))]


def rruff_search(mineral, keys=None):
    query = str(mineral).strip().lower()
    if not query:
        return []
    out = []
    for key in (keys if keys is not None else rruff_available_keys()):
        rows = _read_index(key)
        if not rows:
            continue
        for row in rows:
            if query in row[0].lower() or query in row[1].lower():
                out.append([key] + row)
    return out


def rruff_export_matches(mineral, keys=None, only_processed=False):
    rows = rruff_search(mineral, keys)
    if not rows:
        return [], 0
    best = {}
    for key, _name, rid, wave, ftype, entry in rows:
        pkg = rruff_package(key) or {}
        kind = pkg.get("kind") or "拉曼"
        rank = 2 if "Processed" in ftype else 1
        if only_processed and rank < 2:
            continue
        signature = (kind, rid, wave)
        if signature not in best or rank > best[signature][0]:
            best[signature] = (rank, key, entry)
    grouped = {}
    for _rank, key, entry in best.values():
        grouped.setdefault(key, []).append(entry)
    saved = []
    for key, entries in grouped.items():
        saved.extend(rruff_export(key, entries))
    return saved, len(rows)


def rruff_packages_sorted():
    order = ["unrated_unoriented", "unrated_oriented", "fair_unoriented", "fair_oriented",
             "poor_unoriented", "excellent_unoriented", "excellent_oriented", "LR-Raman",
             "infrared_RAW", "powder_DIF", "powder_XY_Processed", "chemistry_Microprobe"]
    out = []
    for key in order:
        pkg = rruff_package(key)
        if pkg:
            out.append(pkg)
    return out


def _clean_rruff_value(text):
    return str(text or "").replace("_", "").strip()


_RRUFF_INFO_FIELDS = (("RRUFFID", "编号"), ("NAMES", "矿物名"),
                      ("IDEAL CHEMISTRY", "理想化学式"),
                      ("MEASURED CHEMISTRY", "实测化学式"),
                      ("CELL PARAMETERS", "晶胞参数 / 晶系"),
                      ("LOCALITY", "产地"), ("STATUS", "鉴定状态"),
                      ("DESCRIPTION", "样品描述"), ("URL", "数据链接"))


def rruff_entry_info(key, entry_name):
    """读取 RRUFF 数据包内某条目的表头信息。"""
    with zipfile.ZipFile(rruff_zip_path(key)) as z:
        with z.open(entry_name) as fh:
            head = fh.read(4000).decode("utf-8", "ignore")
    return _parse_rruff_headers(head)


def rruff_sample_info(mineral, keys=None):
    """在本地已下载的数据包里找一条同矿物记录，返回其 RRUFF 表头信息。

    优先取名字完全相同的样品，避免 “Zircon” 命中 “Cubic zirconia”。
    """
    rows = rruff_search(mineral, keys)
    target = str(mineral or "").strip().lower()

    def rank(row):
        name = (row[1] or "").strip().lower()
        if name == target:
            return 0
        if name.startswith(target):
            return 1
        if target and target.startswith(name):
            return 2
        return 3

    for row in sorted(rows, key=rank):
        key, name = row[0], row[5]
        if not name.lower().endswith((".txt", ".jdx", ".dx")):
            continue
        try:
            info = rruff_entry_info(key, name)
        except Exception:
            continue
        if info:
            info["_PACKAGE"] = key
            return info
    return None


def rruff_info_text(info):
    """把 RRUFF 表头信息排版成信息卡文本。"""
    if not info:
        return ""
    lines = [T("── RRUFF 样品记录（来源数据包：%s）──") % info.get("_PACKAGE", "")]
    for key, label in _RRUFF_INFO_FIELDS:
        value = info.get(key)
        if value:
            lines.append("%s：%s" % (T(label), _clean_rruff_value(value)))
    return "\n".join(lines)


def mineral_full_info(mineral):
    """内置矿物表 + （若本地有数据包）RRUFF 真实样品记录。"""
    parts = []
    builtin = mineral_info_text(mineral)
    if builtin:
        parts.append(builtin)
    sample = None
    try:
        sample = rruff_sample_info(mineral)
    except Exception:
        sample = None
    if sample:
        parts.append(rruff_info_text(sample))
    return "\n\n".join(parts)


def rruff_export(key, entry_names, only_processed=False, xlabel=None, tag=None):
    """导出选定条目：XY 文本转成本地库 CSV，其他类型（如电子探针表）原样提取。"""
    pkg = rruff_package(key) or {}
    if xlabel is None:
        xlabel = pkg.get("xlabel") or "Raman shift (cm-1)"
    if tag is None:
        tag = pkg.get("kind") or ""
    saved = []
    with zipfile.ZipFile(rruff_zip_path(key)) as z:
        for name in entry_names:
            low = name.lower()
            head = b""
            try:
                with z.open(name) as fh:
                    head = fh.read(2000)
            except Exception:
                continue
            if low.endswith((".txt", ".jdx", ".dx", ".csv", ".xy", ".dat")):
                raw = head.decode("utf-8", "ignore")
                meta = _parse_rruff_headers(raw)
                if only_processed and "PROCESSED" not in (meta.get("FILETYPE", "").upper()):
                    continue
                try:
                    with z.open(name) as fh:
                        raw = fh.read().decode("utf-8", "ignore")
                except Exception:
                    continue
                xs, ys = parse_rruff_text(raw)
                if len(xs) < 4:
                    continue
                info = {"WAVELENGTH": _rruff_wavelength(meta)}
                saved.append(db_save_spectrum(meta.get("RRUFFID", name),
                                              meta.get("NAMES", ""), xs, ys, info,
                                              xlabel=xlabel, tag=tag))
                continue
            mineral, rid, dtype, extra = _parse_rruff_name(name)
            dst = os.path.join(results_dir(), "%s_%s_%s" % (
                "".join(c for c in (mineral or "unknown") if c not in '\\/:*?"<>|'),
                rid or "na", os.path.basename(name).split("__")[-1]))
            try:
                with z.open(name) as fh, open(dst, "wb") as out:
                    shutil.copyfileobj(fh, out)
                saved.append(dst)
            except Exception:
                continue
    return saved


def _norm_max(ys):
    top = max(ys) if ys else 0.0
    return [v / top for v in ys] if top else list(ys)


def _interp_at(xs, ys, x):
    n = len(xs)
    if n == 0:
        return 0.0
    if n == 1:
        return ys[0]
    if x <= xs[0]:
        return ys[0]
    if x >= xs[-1]:
        return ys[-1]
    lo, hi = 0, n - 1
    while hi - lo > 1:
        mid = (lo + hi) // 2
        if xs[mid] <= x:
            lo = mid
        else:
            hi = mid
    span = xs[hi] - xs[lo]
    if not span:
        return ys[lo]
    t = (x - xs[lo]) / span
    return ys[lo] * (1.0 - t) + ys[hi] * t


def _common_range(spectra):
    lo = max(s[1][0] for s in spectra if len(s[1]))
    hi = min(s[1][-1] for s in spectra if len(s[1]))
    if hi <= lo:
        return None
    return lo, hi


def average_spectra(spectra, points=0):
    """把多条光谱插值到公共波数区间后逐点平均，返回 (xs, ys, 条数)。"""
    usable = [s for s in spectra if len(s[1]) > 1]
    if len(usable) < 2:
        return None
    span = _common_range(usable)
    if span is None:
        return None
    lo, hi = span
    if points and points > 3:
        step = (hi - lo) / (points - 1)
        xs = [lo + i * step for i in range(points)]
    else:
        base = max(usable, key=lambda s: len(s[1]))
        xs = [x for x in base[1] if lo <= x <= hi]
    if len(xs) < 4:
        return None
    total = [0.0] * len(xs)
    for _name, sxs, sys_ in usable:
        for i, x in enumerate(xs):
            total[i] += _interp_at(sxs, sys_, x)
    n = float(len(usable))
    return xs, [v / n for v in total], len(usable)


def subtract_scaled(a_xy, b_xy, k=1.0):
    """A − k·B：在公共波数区间上做加权相减，返回 (xs, ys)。"""
    axs, ays = a_xy
    bxs, bys = b_xy
    if len(axs) < 2 or len(bxs) < 2:
        return None
    lo = max(axs[0], bxs[0])
    hi = min(axs[-1], bxs[-1])
    if hi <= lo:
        return None
    xs = [x for x in axs if lo <= x <= hi]
    if len(xs) < 4:
        step = (hi - lo) / 399.0
        xs = [lo + i * step for i in range(400)]
    ys = [_interp_at(axs, ays, x) - float(k) * _interp_at(bxs, bys, x) for x in xs]
    return xs, ys


def replace_segment(xs, ys, other_xs, other_ys, lo, hi, blend=0.0):
    """把 [lo, hi] 区间替换为另一条谱的对应区间，边缘可做线性过渡避免台阶。"""
    if hi < lo:
        lo, hi = hi, lo
    lo = float(lo)
    hi = float(hi)
    blend = max(0.0, float(blend))
    if blend * 2.0 > (hi - lo):
        blend = (hi - lo) / 2.0
    out = list(ys)
    replaced = 0
    for i, x in enumerate(xs):
        if x < lo or x > hi:
            continue
        src = _interp_at(other_xs, other_ys, x)
        w = 1.0
        if blend > 0:
            if x < lo + blend:
                w = (x - lo) / blend
            elif x > hi - blend:
                w = (hi - x) / blend
            w = max(0.0, min(1.0, w))
        out[i] = ys[i] * (1.0 - w) + src * w
        replaced += 1
    return out, replaced


def pair_peaks(target_xy, ref_xy, plot=None, tol=5.0):
    txs, tys = target_xy
    rxs, rys = ref_xy
    left = analyze_peaks(txs, tys, plot)
    right = analyze_peaks(rxs, rys, plot)
    used = set()
    rows = []
    for pk in left:
        best = None
        for j, q in enumerate(right):
            if j in used:
                continue
            dist = abs(q["x"] - pk["x"])
            if dist <= tol and (best is None or dist < best[0]):
                best = (dist, j, q)
        if best:
            used.add(best[1])
            rows.append((pk["x"], best[2]["x"], best[2]["x"] - pk["x"], True))
        else:
            rows.append((pk["x"], None, None, False))
    for j, q in enumerate(right):
        if j not in used:
            rows.append((None, q["x"], None, False))
    return rows


def peak_match_score(target_xy, ref_xy, plot=None, tol=5.0):
    rows = pair_peaks(target_xy, ref_xy, plot, tol)
    left = [row for row in rows if row[0] is not None]
    right = [row for row in rows if row[1] is not None]
    matched = sum(1 for row in rows if row[3])
    n_left, n_right = len(left), len(right)
    if not n_left or not n_right:
        return 0.0, matched, n_left, n_right
    precision = matched / n_right
    recall = matched / n_left
    if precision + recall <= 0:
        f1 = 0.0
    else:
        f1 = 2 * precision * recall / (precision + recall)
    if matched < 2:
        f1 = 0.0
    return 100.0 * f1, matched, n_left, n_right


def pair_spectra(target_xy, ref_paths, plot=None, tol=5.0, top=None):
    results = match_references(target_xy, ref_paths, top=None)
    for r in results:
        r["peak_score"] = 0.0
        r["peak_matched"] = 0
        r["peak_left"] = 0
        r["peak_right"] = 0
        r["pair_rows"] = []
        r["ref_xy"] = None
        try:
            _name, rxs, rys = load_reference_spectrum(r["path"])
            r["ref_xy"] = (rxs, rys)
            score, matched, n_left, n_right = peak_match_score(
                target_xy, (rxs, rys), plot, tol)
            r["peak_score"] = score
            r["peak_matched"] = matched
            r["peak_left"] = n_left
            r["peak_right"] = n_right
        except Exception:
            continue
    results.sort(key=lambda r: (-r["peak_score"], -r["corr"]))
    if results and results[0]["ref_xy"]:
        results[0]["pair_rows"] = pair_peaks(target_xy, results[0]["ref_xy"], plot, tol)
    return results[:top] if top else results


IDENTIFY_TOL = 5.0
IDENTIFY_POOL = 20
IDENTIFY_STRONG = 3
_UNNAMED_NAMES = {"unknown", "unidentified", "unnamed", "n/a", "-", "", "未知", "未命名"}
_SPEC_RE = re.compile(r"%[-#0 +]?[\d*]*(?:\.\d+)?[hlL]?[diouxXeEfFgGcrsa%]")
_I18N_TMPL = [None]


def _peak_index_path(key):
    return os.path.join(rruff_dir(), key + ".peaks.csv")


_PEAK_INDEX_HEADER = "矿物,RRUFFID,类型,波长,化学式,晶系,峰位,相对强度,条目"
_PEAK_INDEX_MAX = 25


def _csv_field(text):
    return str(text or "").replace(",", " ").replace(";", " ").strip()


def _crystal_system(text):
    """从 ##CELL PARAMETERS 文本里取出晶系，如 tetragonal。"""
    raw = str(text or "")
    low = raw.lower()
    pos = low.find("crystal system:")
    if pos < 0:
        return ""
    tail = raw[pos + len("crystal system:"):].strip()
    return tail.split()[0].strip(",;") if tail else ""


def _peak_list_text(values):
    return ";".join(("%.2f" % v) for v in values)


def _peak_list_parse(text):
    out = []
    for token in str(text or "").replace("|", ";").split(";"):
        token = token.strip()
        if not token:
            continue
        try:
            out.append(float(token))
        except ValueError:
            pass
    return out


def rruff_peak_index(key, progress=None, rebuild=False):
    """为数据包建立“特征索引”：每条谱的主要峰位/相对强度 + 化学式 + 晶系。

    数据包内每条谱都要解析一次（600 条约 2 秒），结果缓存到 <key>.peaks.csv。
    返回行列表，每行 9 列：矿物, 编号, 类型, 波长, 化学式, 晶系, 峰位, 相对强度, 条目。
    """
    path = _peak_index_path(key)
    if not rebuild and os.path.exists(path):
        rows = []
        try:
            with open(path, "r", encoding="utf-8-sig", newline="") as f:
                next(f, None)
                for line in f:
                    parts = line.rstrip("\n").split(",")
                    if len(parts) >= 9:
                        rows.append(parts[:9])
        except OSError:
            rows = []
        if rows:
            return rows
    zpath = rruff_zip_path(key)
    if not os.path.exists(zpath):
        raise JwsError(T("尚未下载数据包：%s") % key)
    rows = []
    with zipfile.ZipFile(zpath) as z:
        names = [n for n in z.namelist()
                 if n.lower().endswith((".txt", ".jdx", ".dx"))]
        for i, name in enumerate(names):
            if progress:
                progress(i, len(names))
            try:
                with z.open(name) as fh:
                    raw = fh.read().decode("utf-8", "ignore")
            except Exception:
                continue
            meta = _parse_rruff_headers(raw)
            xs, ys = parse_rruff_text(raw)
            if len(xs) < 10:
                continue
            peaks = find_peaks(xs, ys, 20.0, thresh_pct=7.0)
            if not peaks:
                continue
            peaks.sort(key=lambda pk: -pk["prominence"])
            peaks = peaks[:_PEAK_INDEX_MAX]
            peaks.sort(key=lambda pk: pk["x"])
            top = max(pk["prominence"] for pk in peaks) or 1.0
            rows.append([
                _csv_field(meta.get("NAMES", "")), _csv_field(meta.get("RRUFFID", "")),
                _csv_field(meta.get("FILETYPE", "")), _csv_field(_rruff_wavelength(meta)),
                _csv_field(_clean_rruff_value(meta.get("IDEAL CHEMISTRY", ""))),
                _csv_field(_crystal_system(meta.get("CELL PARAMETERS", ""))),
                _peak_list_text([pk["x"] for pk in peaks]),
                _peak_list_text([100.0 * pk["prominence"] / top for pk in peaks]),
                name,
            ])
    try:
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(_PEAK_INDEX_HEADER + "\n")
            for row in rows:
                f.write(",".join(row) + "\n")
    except OSError:
        pass
    return rows


def peak_index_keys(keys=None):
    """已建好特征索引的数据包。"""
    out = []
    for key in (keys if keys is not None else rruff_available_keys()):
        if os.path.exists(_peak_index_path(key)):
            out.append(key)
    return out


def _binom_tail(n, k, p):
    """P(X >= k)，X ~ B(n, p)；n 很小（峰的个数），直接累加。"""
    if k <= 0:
        return 1.0
    if p <= 0.0:
        return 0.0
    if p >= 1.0:
        return 1.0
    k = min(int(k), int(n))
    total = 0.0
    for i in range(k, int(n) + 1):
        total += math.comb(int(n), i) * (p ** i) * ((1.0 - p) ** (int(n) - i))
    return min(1.0, max(0.0, total))


def score_peak_sets(target, ref_pos, ref_rel, tol=5.0, strong=IDENTIFY_STRONG):
    """峰位集合打分（未知谱 vs 候选参考谱）。

    target: [(峰位, 相对强度)]，未知谱识别到的峰；
    ref_pos / ref_rel: 参考谱的峰位与相对强度。
    按“参考谱最强峰优先”做容差内一对一匹配，返回各项指标：
      f1       峰位匹配 F1（命中<2 记 0）
      strong_hit 未知谱最强 strong 个峰里命中几个
      score    综合分 = 0.5×F1 + 0.5×强峰命中率
      pvalue   仅凭偶然也至少命中这么多的概率（用参考谱峰密度估算）
    """
    tpos = [float(p[0]) for p in target]
    trel = [float(p[1]) if len(p) > 1 else 0.0 for p in target]
    n_t = len(tpos)
    n_r = len(ref_pos)
    if not n_t or not n_r:
        return {"f1": 0.0, "matched": 0, "n_target": n_t, "n_ref": n_r,
                "strong_hit": 0, "strong_n": 0, "dev": 0.0, "score": 0.0,
                "pvalue": 1.0}
    order = sorted(range(n_r), key=lambda j: -(ref_rel[j] if j < len(ref_rel) else 0.0))
    used = set()
    devs = []
    for j in order:
        best = None
        for i, x in enumerate(tpos):
            if i in used:
                continue
            d = abs(x - ref_pos[j])
            if d <= tol and (best is None or d < best[0]):
                best = (d, i)
        if best:
            used.add(best[1])
            devs.append(best[0])
    matched = len(used)
    precision = matched / float(n_r)
    recall = matched / float(n_t)
    f1 = 0.0
    if precision + recall > 0:
        f1 = 2 * precision * recall / (precision + recall)
    if matched < 2:
        f1 = 0.0
    strong_n = min(int(strong), n_t)
    tops = sorted(range(n_t), key=lambda i: -trel[i])[:strong_n]
    strong_hit = sum(1 for i in tops if i in used)
    ratio = (strong_hit / float(strong_n)) if strong_n else 0.0
    score = 0.5 * (100.0 * f1) + 0.5 * (100.0 * ratio)
    span = 0.0
    if len(ref_pos) > 1:
        span = float(max(ref_pos) - min(ref_pos))
    span = max(300.0, span)
    p_single = min(1.0, n_r * 2.0 * tol / span)
    pvalue = _binom_tail(n_t, matched, p_single)
    return {"f1": 100.0 * f1, "matched": matched, "n_target": n_t, "n_ref": n_r,
            "strong_hit": strong_hit, "strong_n": strong_n,
            "dev": (sum(devs) / len(devs)) if devs else 0.0, "score": score,
            "pvalue": pvalue}


def _element_filter(chem, must, not_):
    if not must and not not_:
        return True
    elems = [e.upper() for e in _formula_elements(chem)]
    if must and not all(e in elems for e in must):
        return False
    if not_ and any(e in elems for e in not_):
        return False
    return True


def _split_elements(text):
    return [s.upper() for s in str(text or "").replace(",", " ").replace("，", " ").split()
            if s.strip()]


def identify_unknown(target_xy, plot=None, keys=None, kinds=None, must="", not_="",
                     top=15, tol=IDENTIFY_TOL, exact_pool=IDENTIFY_POOL, group=True,
                     progress=None):
    """未知光谱鉴定：拿未知谱的峰去全库检索，返回按可信度排序的候选。

    · 先按“特征索引”做峰位匹配粗筛（毫秒级）；
    · 再对前 exact_pool 个候选读原始谱做精算（精确 F1 + 相关系数 + 谱角）；
    · group=True 时同一矿物只保留最佳的一条，并给出库内该矿物的谱数。
    返回 (候选列表, 参与检索的条目数, 结论文本)。
    """
    txs, tys = target_xy
    peaks = analyze_peaks(txs, tys, plot)
    if not peaks:
        raise JwsError(T("这条光谱没有识别到峰，无法检索（可调低峰灵敏阈值后重试）"))
    tlist = [(pk["x"], pk["rel"]) for pk in peaks]
    must_e = _split_elements(must)
    not_e = _split_elements(not_)
    scan = keys if keys is not None else rruff_available_keys()
    pool = []
    total = 0
    name_count = {}
    for key in scan:
        pkg = rruff_package(key) or {}
        kind = pkg.get("kind") or "拉曼"
        if kinds and kind not in kinds:
            continue
        try:
            rows = rruff_peak_index(key)
        except Exception:
            continue
        for row in rows:
            name, rid, ftype, wave, chem, system, pos_s, rel_s, entry = row[:9]
            total += 1
            if not _element_filter(chem, must_e, not_e):
                continue
            low = (name or "").strip().lower()
            name_count[low] = name_count.get(low, 0) + 1
            pos = _peak_list_parse(pos_s)
            rel = _peak_list_parse(rel_s)
            if not pos:
                continue
            sc = score_peak_sets(tlist, pos, rel, tol)
            if sc["matched"] <= 0:
                continue
            named = bool(low) and low not in _UNNAMED_NAMES
            item = {"key": key, "kind": kind,
                    "name": name if named else (name or "（未命名样品）"),
                    "named": named,
                    "rruffid": rid, "filetype": ftype, "wavelength": wave,
                    "chemistry": chem, "system": system, "entry": entry,
                    "ref_peaks": pos, "variants": name_count.get(low, 1),
                    "exact_f1": None, "corr": None, "angle": None, "ref_xy": None}
            item.update(sc)
            pool.append(item)
    pool.sort(key=lambda r: (-r["score"], -r["f1"], -r["matched"]))
    pool = pool[:max(int(exact_pool or 0), int(top))]
    exact_n = max(0, min(int(exact_pool or 0), len(pool)))
    for i, r in enumerate(pool[:exact_n]):
        if progress:
            progress(i, exact_n)
        try:
            with zipfile.ZipFile(rruff_zip_path(r["key"])) as z:
                raw = z.read(r["entry"]).decode("utf-8", "ignore")
        except Exception:
            continue
        rxs, rys = parse_rruff_text(raw)
        if len(rxs) < 10:
            continue
        f1, matched, nl, nr = peak_match_score((txs, tys), (rxs, rys), plot, tol)
        cmp = compare_spectra((txs, tys), (rxs, rys))
        r["exact_f1"] = f1
        r["exact_matched"] = matched
        r["corr"] = cmp["corr"] if cmp else None
        r["angle"] = cmp["angle"] if cmp else None
        r["ref_xy"] = (rxs, rys)
    exacted = [r for r in pool if r["exact_f1"] is not None]
    if exacted:
        exacted.sort(key=lambda r: (-r["exact_f1"], -(r["corr"] or -9), r["dev"]))
        rest = [r for r in pool if r["exact_f1"] is None]
        rest.sort(key=lambda r: (-r["score"], -r["f1"]))
        pool = exacted + rest
    if group:
        seen = {}
        merged = []
        for r in pool:
            if r.get("named", True):
                key = ("n", (r["name"] or "").strip().lower())
            else:
                key = ("u", r["rruffid"], r["wavelength"], r["entry"])
            if key in seen:
                seen[key]["variants"] = max(seen[key]["variants"], r["variants"])
                continue
            seen[key] = r
            merged.append(r)
        pool = merged
    pool = ([r for r in pool if r.get("named", True)]
            + [r for r in pool if not r.get("named", True)])
    result = pool[:max(1, int(top))]
    named = [r for r in result if r.get("named", True)]
    best = named[0] if named else None
    # 精算后排序键换成 exact_f1，第 2 名的综合分有可能反而更高，
    # 相减会出现负数；文案里写成“仅领先 -14 分”很别扭，这里夹到 0。
    margin = max(0.0, (best["score"] - named[1]["score"]) if (best and len(named) > 1)
                 else (best["score"] if best else 0.0))
    unnamed_top = next((r for r in result if not r.get("named", True)), None)
    if unnamed_top is None:
        extra = next((r for r in pool if not r.get("named", True)), None)
        if extra is not None and (best is None or extra["score"] > best["score"]):
            unnamed_top = extra
            result = result + [extra]
    hint = ""
    if unnamed_top is not None and (best is None or unnamed_top["score"] > best["score"]):
        peaks_txt = ("、" if ui_lang() == "zh" else ", ").join(
            "%.0f" % v for v in unnamed_top["ref_peaks"][:6])
        hint = (T("另有「未命名样品 %s」（分数 %.0f）更接近，其峰位 %s；") + _sp()
                + T("可作价态/物相线索，但该记录没有矿物名，无法据此定名。")
                ) % (unnamed_top["rruffid"], unnamed_top["score"], peaks_txt)
    if best is None:
        if unnamed_top is not None:
            verdict = T("库里没有命名的矿物峰位接近（%s）") % hint
        else:
            verdict = (T("库里没有找到任何峰位接近的参考谱：可能该矿物不在已下载的数据包里，")
                       + _sp() + T("或这不是拉曼谱（峰位单位/谱区不同）。"))
    else:
        strong_ok = best["strong_hit"] >= 2
        lead = margin >= 15.0
        if best["score"] >= 60 and strong_ok and lead:
            verdict = (T("可信候选：%s（%s，库内 %d 条同名谱；领先第二名 %.0f 分）。")
                       + _sp() + T("建议再用配对报告核对峰位对照表。")) % (
                           best["name"], best["rruffid"], best["variants"], margin)
        elif best["score"] >= 35:
            verdict = T("可能的候选：%s（%s），置信度一般。") % (
                best["name"], best["rruffid"])
            if not lead:
                verdict += (_sp() + T("注意前几名分数很接近（仅领先 %.0f 分），")
                            + _sp() + T("这种情况也可能是库中没有对应矿物，请结合峰位人工判断。")
                            ) % margin
            else:
                verdict += _sp() + T("可下载更大 / 质量更高的数据包后再检索以提高把握。")
        else:
            verdict = (T("没有可靠匹配（有名字的最高分 %.0f）。可能该矿物不在已下载数据包里，")
                       + _sp() + T("或激光波长/谱区与库中不同，或这不是拉曼谱。")) % best["score"]
    if hint:
        verdict += ("　" if ui_lang() == "zh" else "  ") + hint
    return result, total, verdict


def format_identify_rows(results, limit=None):
    """把候选列表格式化成表格行（供界面/命令行/CSV 共用）。"""
    rows = []
    for i, r in enumerate(results if limit is None else results[:limit], 1):
        f1 = r["exact_f1"] if r["exact_f1"] is not None else r["f1"]
        name = r["name"]
        if not r.get("named", True):
            name = T("%s（未命名样品）") % name
        rows.append([
            str(i), name, r["rruffid"], T(r["kind"]), r["wavelength"] or "-",
            "%.1f" % r["score"], "%.1f" % f1,
            "%d/%d" % (r["strong_hit"], r["strong_n"]),
            "%d/%d" % (r["matched"], r["n_ref"]),
            "%.2f" % r["dev"],
            "-" if r["corr"] is None else "%.4f" % r["corr"],
            str(r.get("variants", 1)),
            "%.3f" % r.get("pvalue", 1.0),
            r["chemistry"] or "-", T(r["system"]) if r["system"] else "-",
            T("精算") if r["exact_f1"] is not None else T("索引"),
        ])
    return rows


IDENTIFY_COLUMNS = ("排名", "矿物", "RRUFF编号", "类型", "波长(nm)", "综合分", "F1",
                    "强峰命中", "命中/参考峰数", "平均偏差", "相关系数", "库内同名谱数",
                    "偶然概率", "化学式", "晶系", "评分方式")

# 命令行表格：列索引与列宽
_IDENTIFY_CLI_COLS = [(0, 4), (1, 26), (2, 9), (3, 5), (5, 7), (6, 6), (7, 9),
                      (9, 7), (11, 5), (12, 8), (13, 22)]


def print_identify_table(results):
    """命令行下打印候选表。"""
    head = " ".join("%-*s" % (w, T(IDENTIFY_COLUMNS[i])) for i, w in _IDENTIFY_CLI_COLS)
    print(" " + head)
    print("-" * (len(head) + 2))
    for row in format_identify_rows(results):
        line = []
        for i, w in _IDENTIFY_CLI_COLS:
            text = row[i]
            if len(text) > w:
                text = text[:w]
            if i in (5, 6, 9, 11):
                line.append("%*s" % (w, text))
            else:
                line.append("%-*s" % (w, text))
        print(" " + " ".join(line))


def write_identify_csv(dst, target_name, results, verdict="", total=0):
    """把鉴定候选写成 CSV。"""
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        f.write(T("== 未知光谱检索：%s ==\n") % target_name)
        f.write(T("参与检索条目数：%d\n") % total)
        f.write(T("结论：%s\n") % verdict.replace("\n", " "))
        f.write((T("评分说明：综合分 = 0.5×F1 + 0.5×强峰命中率；")
                 + T("F1 为峰位匹配 F1（容差 %.0f cm-1，命中<2 记 0）；")
                 + T("强峰命中 = 未知谱最强 %d 个峰里命中几个。")
                 + T("同一矿物只列最佳一条（库内同名谱数为该矿物在库中的条数）。\n"))
                % (IDENTIFY_TOL, IDENTIFY_STRONG))
        f.write(",".join(T(c) for c in IDENTIFY_COLUMNS) + "\n")
        for row in format_identify_rows(results):
            f.write(",".join(c.replace(",", " ") for c in row) + "\n")
    return dst


# ---------------------------------------------------------------------------
# 批量配对 / 批量鉴定
#
# 单个谱的「配对」与「鉴定」各自都只处理一条谱，这里加上外层批量循环：
#   · batch_pair     多条实测谱 × 一组参考谱（如一批锆石标准谱），逐条给最佳参考；
#   · batch_identify 多条陌生谱，逐条在全库里找最像的矿物。
# 两者共用 score_peaks_vs_ref 这一份打分口径，所以汇总表里的数字互相对得上。
# 汇总表只列「辅助指标」，最终是哪个物相由使用者自己判断。
# ---------------------------------------------------------------------------

def _target_peak_list(target_xy, plot=None):
    """目标谱的峰位表 [(峰位, 相对强度)]；没有识别到峰则返回空表。"""
    xs, ys = target_xy
    return [(pk["x"], pk["rel"]) for pk in analyze_peaks(xs, ys, plot)]


def score_peaks_vs_ref(tlist, target_xy, ref_xy, plot=None, tol=IDENTIFY_TOL):
    """目标峰的峰位表 vs 一条参考谱：综合分 / F1 / 强峰命中 / 相关系数 / 谱角。

    口径与「未知谱鉴定」完全一致（score_peak_sets），
    免得配对报告和鉴定结果两处给出不同的数字。
    """
    rpeaks = analyze_peaks(ref_xy[0], ref_xy[1], plot)
    if not tlist or not rpeaks:
        return None
    sc = score_peak_sets(tlist, [p["x"] for p in rpeaks],
                         [p["rel"] for p in rpeaks], tol)
    cmp = compare_spectra(target_xy, ref_xy)
    sc["corr"] = cmp["corr"] if cmp else None
    sc["angle"] = cmp["angle"] if cmp else None
    return sc


def load_reference_set(paths):
    """把参考谱文件 / 文件夹读成 [(名称, xs, ys)]。"""
    refs = []
    for path in find_input_files(paths):
        if not path.lower().endswith((".csv", ".txt", ".dat")):
            continue
        try:
            refs.append(load_reference_spectrum(path))
        except Exception:
            continue
    return refs


def pair_reading(best):
    """配对结果的「参考判读」。只是提示，最终判断留给人。"""
    if not best:
        return T("没有可比对的参考谱")
    if best["score"] >= 60 and best["strong_hit"] >= 2:
        return T("匹配良好，可作同一物相（参考判读）")
    if best["score"] >= 35:
        return T("部分匹配，建议核对峰位对照（参考判读）")
    return T("匹配很差，很可能不是同一物相（参考判读）")


def batch_pair(target_spectra, refs, plot=None, tol=IDENTIFY_TOL, progress=None):
    """多条实测谱 × 一组参考谱：逐条给出最佳参考与全部辅助指标。

    target_spectra: [(名称, xs, ys)] 实测谱；
    refs:           [(名称, xs, ys)] 参考谱（load_reference_set 读入）。
    返回 [{name, n_points, n_peaks, best, margin, candidates, error}]，
    candidates 按（综合分, F1）降序。
    """
    rows = []
    total = len(target_spectra)
    for i, (name, xs, ys) in enumerate(target_spectra):
        if progress is not None:
            progress(i, total, name)
        row = {"name": name, "n_points": len(xs), "n_peaks": 0,
               "best": None, "margin": None, "candidates": [], "error": ""}
        try:
            tlist = _target_peak_list((xs, ys), plot)
            row["n_peaks"] = len(tlist)
            if not tlist:
                row["error"] = T("这条光谱没有识别到峰")
                rows.append(row)
                continue
            cands = []
            for rname, rxs, rys in refs:
                sc = score_peaks_vs_ref(tlist, (xs, ys), (rxs, rys), plot, tol)
                if sc is None:
                    continue
                item = dict(sc)
                item["name"] = rname
                cands.append(item)
            if not cands:
                row["error"] = T("与参考谱没有可比对的峰（波数区间不重叠或参考谱没有峰）")
                rows.append(row)
                continue
            cands.sort(key=lambda r: (-r["score"], -r["f1"], -(r["corr"] or -9.0)))
            row["candidates"] = cands
            row["best"] = cands[0]
            if len(cands) > 1:
                row["margin"] = cands[0]["score"] - cands[1]["score"]
        except Exception as exc:
            row["error"] = str(exc)
        rows.append(row)
    if progress is not None:
        progress(total, total, "")
    return rows


def batch_identify(target_spectra, plot=None, keys=None, kinds=None, must="", not_="",
                   top=5, progress=None):
    """多条未知谱逐条鉴定：每条在全库里找最像的矿物。

    返回 [{name, n_points, n_peaks, candidates, verdict, total}]，
    每条独立调用 identify_unknown，互不影响。
    """
    rows = []
    total = len(target_spectra)
    for i, (name, xs, ys) in enumerate(target_spectra):
        if progress is not None:
            progress(i, total, name)
        row = {"name": name, "n_points": len(xs), "n_peaks": 0,
               "candidates": [], "verdict": "", "total": 0}
        try:
            row["n_peaks"] = len(_target_peak_list((xs, ys), plot))
            cands, n_total, verdict = identify_unknown(
                (xs, ys), plot, keys=keys, kinds=kinds, must=must, not_=not_,
                top=max(1, int(top)))
            row["candidates"] = cands
            row["total"] = n_total
            row["verdict"] = verdict
        except Exception as exc:
            row["verdict"] = str(exc)
        rows.append(row)
    if progress is not None:
        progress(total, total, "")
    return rows


def _best_named(cands):
    """候选里最靠前的「有矿物名」的一条（未命名样品不参与定名）。"""
    for c in cands or []:
        if c.get("named", True):
            return c
    return cands[0] if cands else None


BATCH_PAIR_COLUMNS = (
    "文件名", "数据点数", "识别峰数", "最佳参考谱", "综合分", "F1(%)",
    "强峰命中", "强峰总数", "命中峰数", "实测峰数", "参考峰数",
    "平均偏差(cm-1)", "偶然概率", "相关系数", "谱角(度)", "领先第二名(分)", "参考判读")

BATCH_IDENTIFY_COLUMNS = (
    "文件名", "数据点数", "识别峰数", "最佳候选", "RRUFF编号", "类型", "波长(nm)",
    "综合分", "F1(%)", "强峰命中", "强峰总数", "命中峰数", "参考峰数",
    "平均偏差(cm-1)", "相关系数", "库内同名谱数", "偶然概率", "参考判读")


def batch_pair_row(row):
    """一条实测谱 → 汇总表的一行。"""
    best = row.get("best")
    if not best:
        out = [row["name"], str(row["n_points"]), str(row["n_peaks"])] + ["-"] * 13
        out.append(row.get("error") or T("没有可比对的参考谱"))
        return out
    margin = row.get("margin")
    return [
        row["name"], str(row["n_points"]), str(row["n_peaks"]),
        best["name"], "%.1f" % best["score"], "%.1f" % best["f1"],
        str(best["strong_hit"]), str(best["strong_n"]),
        str(best["matched"]), str(best["n_target"]), str(best["n_ref"]),
        "%.2f" % best["dev"], "%.3f" % best["pvalue"],
        "-" if best["corr"] is None else "%.4f" % best["corr"],
        "-" if best["angle"] is None else "%.2f" % best["angle"],
        "-" if margin is None else "%+.1f" % margin,
        pair_reading(best),
    ]


def batch_identify_row(row):
    """一条陌生谱 → 汇总表的一行（只取最佳候选，判读用鉴定那套结论）。"""
    best = _best_named(row.get("candidates"))
    if not best:
        out = [row["name"], str(row["n_points"]), str(row.get("n_peaks", 0))] + ["-"] * 14
        out.append(row.get("verdict") or T("没有找到候选"))
        return out
    f1 = best["exact_f1"] if best.get("exact_f1") is not None else best["f1"]
    return [
        row["name"], str(row["n_points"]), str(row.get("n_peaks", 0)),
        best["name"], best["rruffid"], T(best["kind"]), best["wavelength"] or "-",
        "%.1f" % best["score"], "%.1f" % f1,
        str(best["strong_hit"]), str(best["strong_n"]),
        str(best["matched"]), str(best["n_ref"]),
        "%.2f" % best["dev"],
        "-" if best["corr"] is None else "%.4f" % best["corr"],
        str(best.get("variants", 1)), "%.3f" % best.get("pvalue", 1.0),
        row.get("verdict", ""),
    ]


def _pair_sort_key(row):
    """综合分升序：对不上的排最前，一眼就能看到可疑的。"""
    best = row.get("best")
    if not best:
        return (0, 0.0, row["name"])
    return (1, best["score"], row["name"])


def _identify_sort_key(row):
    best = _best_named(row.get("candidates"))
    if not best:
        return (0, 0.0, row["name"])
    return (1, best["score"], row["name"])


_BATCH_SCORE_NOTE = ("综合分 = 0.5×F1 + 0.5×强峰命中率（与「未知谱鉴定」同一口径）；"
                     "这些只是辅助指标，最终是哪个物相请自行核对峰位与谱型。")


def write_batch_pair_csv(dst, rows, ref_label="", tol=IDENTIFY_TOL):
    """批量配对汇总 CSV：一条实测谱一行。"""
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        f.write(T("== 批量配对汇总 ==\n"))
        f.write(T("参考谱集：%s\n") % ref_label)
        f.write(T("实测谱：%d 条；峰位匹配容差 %.0f cm-1\n") % (len(rows), tol))
        f.write(T(_BATCH_SCORE_NOTE) + "\n")
        f.write(T("（按综合分升序排列：最可疑的排在最前面）\n"))
        f.write(",".join(T(c) for c in BATCH_PAIR_COLUMNS) + "\n")
        for row in sorted(rows, key=_pair_sort_key):
            f.write(",".join(c.replace(",", " ") for c in batch_pair_row(row)) + "\n")
    return dst


def write_batch_identify_csv(dst, rows, total_entries=0, top=1):
    """批量鉴定汇总 CSV：一条陌生谱一行（最佳候选）。"""
    with open(dst, "w", encoding="utf-8-sig", newline="") as f:
        f.write(T("== 批量未知谱鉴定汇总 ==\n"))
        f.write(T("参与检索条目数：%d\n") % total_entries)
        f.write(T("实测谱：%d 条\n") % len(rows))
        f.write((T("评分说明：综合分 = 0.5×F1 + 0.5×强峰命中率；")
                 + T("F1 为峰位匹配 F1（容差 %.0f cm-1，命中<2 记 0）。")
                 + T("这些只是辅助指标，最终是哪个物相请自行核对峰位与谱型。\n"))
                % IDENTIFY_TOL)
        f.write(T("（按最佳候选的综合分升序排列：最可疑的排在最前面）\n"))
        f.write(",".join(T(c) for c in BATCH_IDENTIFY_COLUMNS) + "\n")
        for row in sorted(rows, key=_identify_sort_key):
            f.write(",".join(c.replace(",", " ") for c in batch_identify_row(row)) + "\n")
    return dst


def write_batch_pair_html(dst, rows, ref_label="", tol=IDENTIFY_TOL):
    """批量配对汇总的可打印 HTML 报告。"""
    ordered = sorted(rows, key=_pair_sort_key)
    table = [list(BATCH_PAIR_COLUMNS)] + [batch_pair_row(r) for r in ordered]
    return write_report_html(
        dst,
        T("批量配对汇总"),
        T("参考谱集：%s") % ref_label,
        info=[(T("实测谱"), str(len(rows))),
              (T("峰位匹配容差"), "%.0f cm-1" % tol),
              (T("排序"), T("按综合分升序（最可疑的在前）"))],
        tables=[(T("逐条配对结果"), table)],
        notes=[T(_BATCH_SCORE_NOTE),
               T("要对某条谱细看峰位对照，可单独对它跑一次配对，"
                 "会生成 _配对报告.csv 与 _配对报告.png。")])


def write_batch_identify_html(dst, rows, total_entries=0):
    """批量鉴定汇总的可打印 HTML 报告。"""
    ordered = sorted(rows, key=_identify_sort_key)
    table = [list(BATCH_IDENTIFY_COLUMNS)] + [batch_identify_row(r) for r in ordered]
    return write_report_html(
        dst,
        T("批量未知谱鉴定汇总"),
        T("参与检索条目数：%d") % total_entries,
        info=[(T("实测谱"), str(len(rows))),
              (T("排序"), T("按最佳候选综合分升序（最可疑的在前）"))],
        tables=[(T("逐条鉴定结果"), table)],
        notes=[T("评分说明：综合分 = 0.5×F1 + 0.5×强峰命中率；"
                 "F1 为峰位匹配 F1（容差 %.0f cm-1，命中<2 记 0）。") % IDENTIFY_TOL,
               T("这些只是辅助指标，最终是哪个物相请自行核对峰位与谱型。")])


def convert_file(src, out_dir=None, x_header=_DEFAULT_X_HEADER,
                 y_header=_DEFAULT_Y_HEADER, write_header=True,
                 skip_existing=False, auto_names=False, formats=("csv",), plot=None):
    spec = Spectrum(src)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        base_dir = out_dir
    else:
        base_dir = os.path.dirname(os.path.abspath(src))
    base = os.path.splitext(os.path.basename(src))[0]
    wanted = set(formats) or {"csv"}

    order = [("xlsx", os.path.join(base_dir, base + ".xlsx")),
             ("csv", os.path.join(base_dir, base + ".csv")),
             ("png", os.path.join(base_dir, base + ".png")),
             ("peaks", os.path.join(base_dir, base + "_peaks.csv")),
             ("fit", os.path.join(base_dir, base + "_fit.csv"))]
    targets = [(k, p) for k, p in order if k in wanted]
    if not targets:
        targets = [("csv", os.path.join(base_dir, base + ".csv"))]

    if skip_existing and all(os.path.exists(p) for _k, p in targets):
        return targets[0][1], spec, True

    names = _column_names(spec, x_header, y_header, auto_names)
    for kind, path in targets:
        if kind == "xlsx":
            _write_xlsx(path, spec, names, write_header, plot)
        elif kind == "png":
            _write_png(path, spec, names, plot)
        elif kind == "peaks":
            _write_peaks_csv(path, spec, names, plot)
        elif kind == "fit":
            _write_fit_csv(path, spec, names, plot)
        else:
            _write_csv(path, spec, names, write_header, plot)
    return targets[0][1], spec, False


def convert_csv_to_png(src, out_dir=None, plot=None):
    if not _HAVE_PIL:
        raise JwsError(T("导出 PNG 需要 Pillow 组件（pip install pillow）"))
    xlabel, series = load_csv_series(src)
    colored = [(lab, xs, ys, _PALETTE[i % len(_PALETTE)])
               for i, (lab, xs, ys) in enumerate(series)]
    base = os.path.splitext(os.path.basename(src))[0]
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
        dst = os.path.join(out_dir, base + ".png")
    else:
        dst = os.path.splitext(src)[0] + ".png"
    ylabel = series[0][0] if series else "Intensity"
    render_png(dst, colored, base, xlabel, ylabel, plot=plot)
    return dst, len(colored)


def _decimate(xs, ys, width):
    n = len(ys)
    if width <= 2 or n <= width * 2:
        return xs, ys
    out_x = []
    out_y = []
    for c in range(width):
        i0 = int(c * n / width)
        i1 = int((c + 1) * n / width)
        if i1 <= i0:
            i1 = i0 + 1
        seg = ys[i0:i1]
        mn = min(seg)
        mx = max(seg)
        imn = i0 + seg.index(mn)
        imx = i0 + seg.index(mx)
        if imn <= imx:
            out_x.append(xs[imn])
            out_y.append(mn)
            out_x.append(xs[imx])
            out_y.append(mx)
        else:
            out_x.append(xs[imx])
            out_y.append(mx)
            out_x.append(xs[imn])
            out_y.append(mn)
    return out_x, out_y


def _table_to_series(xs, cols, names, default_x):
    """把 (第一列数值, 其余列) 组装成 (xlabel, series)。

    容忍末尾的文本列（如峰列表的“来源/归属”）：全空列直接丢弃，
    个别空值用前一个有效值补齐，保证绘图与处理不出现 NaN。
    """
    if not xs:
        raise JwsError(T("文件里没有找到可用的数值数据（需第一列为数字）"))
    kept = [ys for ys in cols if any(v == v for v in ys)]
    if not kept:
        raise JwsError(T("文件里没有找到可用的数值数据"))
    fixed = []
    for ys in kept:
        out = []
        last = 0.0
        for v in ys:
            if v == v:
                last = v
            out.append(last)
        fixed.append(out)
    series = []
    for i, ys in enumerate(fixed):
        label = names[i + 1] if i + 1 < len(names) else "Y%d" % (i + 1)
        series.append((label, list(xs), ys))
    xlabel = names[0] if names else default_x
    return xlabel, series


def load_csv_series(path):
    xs = []
    cols = []
    names = []
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            parts = line.replace(";", ",").split(",")
            if not parts:
                continue
            try:
                x = float(parts[0])
            except ValueError:
                if not xs and not names and any(p.strip() for p in parts):
                    names = [p.strip() for p in parts]
                continue
            vals = [x]
            for p in parts[1:]:
                try:
                    vals.append(float(p))
                except ValueError:
                    vals.append(float("nan"))
            if not xs:
                cols = [[] for _ in range(max(1, len(vals) - 1))]
            xs.append(x)
            for i in range(len(cols)):
                cols[i].append(vals[i + 1] if i + 1 < len(vals) else float("nan"))
    return _table_to_series(xs, cols, names, _DEFAULT_X_HEADER)


def _rows_to_series(rows, default_x="X"):
    """把二维单元格表转成 (xlabel, [(label, xs, ys)])，自动跳过表头行。

    与 CSV 读取一致：只要求第一列是数字，其余列允许出现文本（按空值处理）。
    """
    names = []
    xs = []
    cols = []
    for row in rows:
        if row is None:
            continue
        cells = list(row)
        if not cells:
            continue
        try:
            x = float(cells[0])
        except (TypeError, ValueError):
            if not xs and not names:
                names = [str(c).strip() for c in cells
                         if c is not None and str(c).strip()]
            continue
        vals = []
        for cell in cells[1:]:
            if cell is None or (isinstance(cell, str) and not cell.strip()):
                vals.append(float("nan"))
                continue
            if isinstance(cell, (int, float)):
                vals.append(float(cell))
                continue
            text = str(cell).strip().replace(",", "")
            try:
                vals.append(float(text))
            except ValueError:
                vals.append(float("nan"))
        if not xs:
            cols = [[] for _ in range(max(1, len(vals)))]
        xs.append(x)
        for i in range(len(cols)):
            cols[i].append(vals[i] if i < len(vals) else float("nan"))
    return _table_to_series(xs, cols, names, default_x)


def load_txt_series(path):
    """通用两列/多列数值文本：自动识别制表符 / 逗号 / 分号 / 空白分隔。"""
    with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
        text = f.read()
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    rows = []
    for raw in lines:
        line = raw.strip()
        if not line or line.startswith("#") or line.startswith("//") or line.startswith(";"):
            continue
        if "\t" in line:
            parts = line.split("\t")
        elif "," in line:
            parts = line.split(",")
        elif ";" in line:
            parts = line.split(";")
        else:
            parts = line.split()
        if parts:
            rows.append([p.strip() for p in parts])
    return _rows_to_series(rows, default_x=_DEFAULT_X_HEADER)


def load_xlsx_series(path, sheet=None, max_rows=200000):
    """读取 Excel 第一个工作表（或指定表）里的两列以上数值。"""
    if not _HAVE_XLSX:
        raise JwsError(T("读取 Excel 需要 openpyxl 组件（pip install openpyxl）"))
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        name = sheet if sheet and sheet in wb.sheetnames else wb.sheetnames[0]
        ws = wb[name]
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(row)
            if len(rows) >= max_rows:
                break
    finally:
        wb.close()
    return _rows_to_series(rows, default_x=_DEFAULT_X_HEADER)


_SPC_XNAMES = {0: "X", 1: "Raman shift (cm-1)", 2: "Wavelength (um)",
               3: "Wavelength (nm)", 4: "Wavenumber (cm-1)", 5: "Mass (m/z)",
               6: "Time (s)", 7: "Temperature", 8: "Pressure"}


def load_spc_series(path):
    """Galactic / Thermo GRAMS 的 SPC 光谱（新版 0x4B，小端）。

    按 spc.h 规范解析：主头(512) → [X 数组(可选)] → 子文件头(32) → Y 数组。
    Y 值若 fexp/subexp = 0x80 则为 IEEE float32；否则是定点小数，
    真实值 = 2^exponent × 整数 / 2^(8×字节数)。
    """
    with open(path, "rb") as f:
        blob = f.read()
    if len(blob) < 544:
        raise JwsError(T("SPC 文件过小或已损坏"))
    ftflgs, fversn, _fexper, fexp = struct.unpack_from("<4B", blob, 0)
    if fversn == 0x4C:
        raise JwsError(T("SPC 大端格式（fversn=0x4C）暂不支持"))
    if fversn != 0x4B:
        raise JwsError(T("暂不支持该 SPC 版本（fversn=0x%02X，仅支持新版 0x4B）") % fversn)
    fnpts = struct.unpack_from("<I", blob, 4)[0]
    ffirst, flast = struct.unpack_from("<dd", blob, 8)
    fnsub = struct.unpack_from("<I", blob, 24)[0]
    fxtype = struct.unpack_from("<I", blob, 28)[0]
    if fnpts <= 0:
        raise JwsError(T("SPC 未声明数据点数，无法解析"))

    def _signed(byte):
        return byte - 256 if byte > 127 else byte

    pos = 512
    xarr = None
    if ftflgs & 0x80:
        if len(blob) < pos + fnpts * 4:
            raise JwsError(T("SPC 声明了 X 数组但数据段不足"))
        xarr = list(struct.unpack_from("<%df" % fnpts, blob, pos))
        pos += fnpts * 4
    if len(blob) < pos + 32:
        raise JwsError(T("SPC 缺少子文件头，文件可能被截断"))
    subflgs, subexp = struct.unpack_from("<2B", blob, pos)
    subnpts = struct.unpack_from("<i", blob, pos + 16)[0]
    pos += 32
    npts = subnpts if 0 < subnpts <= fnpts else fnpts
    multi = bool(subflgs & 0x01)
    nsub = fnsub if (fnsub > 0 and multi) else 1
    if (subflgs & 0x10) and xarr is None:
        if len(blob) < pos + npts * 4:
            raise JwsError(T("SPC 子文件 X 数组不完整"))
        xarr = list(struct.unpack_from("<%df" % npts, blob, pos))
        pos += npts * 4

    ieee = (fexp == 0x80) or (subexp == 0x80)
    if ieee:
        if len(blob) < pos + npts * 4:
            raise JwsError(T("SPC 数据段长度不足（浮点）"))
        ys = list(struct.unpack_from("<%df" % npts, blob, pos))
        enc = "float32"
    elif ftflgs & 0x01:
        if len(blob) < pos + npts * 2:
            raise JwsError(T("SPC 数据段长度不足（16 位）"))
        raw = struct.unpack_from("<%dh" % npts, blob, pos)
        scale = (2.0 ** _signed(fexp)) / 65536.0
        ys = [v * scale for v in raw]
        enc = "16 位定点"
    else:
        if len(blob) < pos + npts * 4:
            raise JwsError(T("SPC 数据段长度不足（32 位）"))
        raw = struct.unpack_from("<%di" % npts, blob, pos)
        scale = (2.0 ** _signed(fexp)) / 4294967296.0
        ys = [v * scale for v in raw]
        enc = "32 位定点"
    if ftflgs & 0x40:
        ys = [10.0 ** v if -300 < v < 300 else v for v in ys]
        enc += "+log"

    if xarr and len(xarr) >= npts:
        xs = [float(v) for v in xarr[:npts]]
    else:
        step = (flast - ffirst) / (npts - 1) if npts > 1 else 0.0
        xs = [ffirst + i * step for i in range(npts)]
    if len(xs) > 1 and xs[0] > xs[-1]:
        xs.reverse()
        ys.reverse()

    note = ""
    if nsub > 1:
        note = "（含 %d 个子谱，取第 1 个）" % nsub
    xlabel = _SPC_XNAMES.get(int(fxtype), "X")
    return xlabel, [("SPC[%s]%s" % (enc, note), xs, ys)]


def read_any_series(path):
    """按扩展名分派读取任意受支持的光谱文件，返回 (xlabel, [(label, xs, ys)])。"""
    low = path.lower()
    if low.endswith(".jws"):
        spec = Spectrum(path)
        xs = spec.x_values()
        series = []
        for c in range(spec.channel_number):
            label = spec.y_names(_DEFAULT_Y_HEADER)[c] if spec.channel_number <= 1 \
                else "通道%d" % (c + 1)
            series.append((label, xs, list(spec.y_data[c])))
        return spec.axis_name(_DEFAULT_X_HEADER), series
    if low.endswith(".csv"):
        return load_csv_series(path)
    if low.endswith((".xlsx", ".xlsm")):
        return load_xlsx_series(path)
    if low.endswith(".spc"):
        return load_spc_series(path)
    if low.endswith((".jdx", ".dx")):
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            xs, ys, meta = parse_jcamp(f.read())
        if len(xs) < 2:
            raise JwsError(T("JCAMP-DX 文件里没有可用的数据点"))
        return "Raman shift (cm-1)", [("JCAMP-DX", xs, ys)]
    if low.endswith(".txt"):
        with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
            head = f.read(400)
        if head.lstrip().startswith("##"):
            with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                xs, ys, _meta = parse_jcamp(f.read())
            if len(xs) >= 2:
                return "Raman shift (cm-1)", [("JCAMP-DX", xs, ys)]
        return load_txt_series(path)
    return load_txt_series(path)


def _jcamp_factor(values, limit=6):
    """选择能把数值写成整数的 XFACTOR / YFACTOR（10^-n）。

    采样必须覆盖极值点与全区间，否则会误判（例如只采到接近 0 的区段）。
    """
    n = len(values)
    if n == 0:
        return 10.0 ** (-limit)
    step = max(1, n // 240)
    idxs = set(range(0, n, step))
    idxs.update((0, n - 1))
    idxs.add(values.index(max(values)))
    idxs.add(values.index(min(values)))
    sample = [values[i] for i in sorted(idxs)]
    for digits in range(0, limit + 1):
        f = 10.0 ** (-digits)
        ok = True
        for v in sample:
            q = v / f
            if abs(q - round(q)) > 1e-6 * max(1.0, abs(q)):
                ok = False
                break
        if ok:
            return f
    return 10.0 ** (-limit)


def write_jcamp(path, xs, ys, meta=None, title=None, dtype="RAMAN SPECTRUM",
                xunits="1/CM", yunits="ARBITRARY UNITS", per_line=5):
    """导出 JCAMP-DX (.jdx) 文件。

    等间距数据用 ##X++(Y..Y) + ##DELTAX（横坐标无编码误差，与 RRUFF/ROD 一致）；
    非等间距回退为 ##XYPOINTS=(XY..XY)。纵坐标按 ##YFACTOR 整数化。
    """
    n = min(len(xs), len(ys))
    if n < 2:
        raise JwsError(T("数据点太少，无法导出 JCAMP-DX"))
    xs = [float(v) for v in xs[:n]]
    ys = [float(v) for v in ys[:n]]
    meta = meta or {}
    yf = _jcamp_factor(ys)
    dx = xs[1] - xs[0]
    even = abs(dx) > 0
    if even and n > 3:
        tol = abs(dx) * 1e-6 + 1e-9
        for i in range(1, n - 1):
            if abs((xs[i + 1] - xs[i]) - dx) > tol:
                even = False
                break
    lines = []
    lines.append("##TITLE=%s" % (title or os.path.splitext(os.path.basename(path))[0]))
    lines.append("##JCAMP-DX=5.01")
    lines.append("##DATA TYPE=%s" % dtype)
    lines.append("##ORIGIN=%s" % T("拉曼光谱工具"))
    for key in ("RRUFFID", "MINERAL", "WAVELENGTH", "SAMPLE"):
        if meta.get(key):
            lines.append("##%s=%s" % (key, meta[key]))
    lines.append("##XUNITS=%s" % xunits)
    lines.append("##YUNITS=%s" % yunits)
    lines.append("##FIRSTX=%.10g" % xs[0])
    lines.append("##LASTX=%.10g" % xs[-1])
    lines.append("##NPOINTS=%d" % n)
    lines.append("##FIRSTY=%.10g" % ys[0])
    if even:
        lines.append("##DELTAX=%.10g" % dx)
        lines.append("##XFACTOR=1")
        lines.append("##YFACTOR=%.10g" % yf)
        lines.append("##X++(Y..Y)")
        buf = []
        for v in ys:
            buf.append("%d" % round(v / yf))
            if len(buf) >= max(1, int(per_line)):
                lines.append(" ".join(buf))
                buf = []
        if buf:
            lines.append(" ".join(buf))
    else:
        xf = _jcamp_factor(xs)
        lines.append("##XFACTOR=%.10g" % xf)
        lines.append("##YFACTOR=%.10g" % yf)
        lines.append("##XYPOINTS=(XY..XY)")
        buf = []
        for i in range(n):
            buf.append("%d,%d" % (round(xs[i] / xf), round(ys[i] / yf)))
            if len(buf) >= max(1, int(per_line)):
                lines.append(" ".join(buf))
                buf = []
        if buf:
            lines.append(" ".join(buf))
    lines.append("##END=")
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines) + "\n")
    return path


def _cli_batch(folder, out_arg, formats, plot):
    if not os.path.isdir(folder):
        print(T("文件夹不存在：%s") % folder)
        return 1
    wanted = tuple(f for f in (formats or ("xlsx", "png", "peaks"))
                   if f in ("csv", "xlsx", "png", "peaks", "fit", "jcamp"))
    if not wanted:
        wanted = ("xlsx", "png", "peaks")
    target = out_arg or os.path.join(os.path.abspath(folder), "_转换结果")
    os.makedirs(target, exist_ok=True)
    print(T("批处理目录：%s") % folder)
    print(T("输出目录：  %s") % target)
    print(T("输出内容：  %s") % "+".join(wanted))
    print("-" * 48)

    def prog(i, n, src):
        print("  (%d/%d) %s" % (i, n, os.path.basename(src)))

    try:
        res = batch_convert(folder, out_dir=target, formats=wanted, plot=plot or None,
                            progress=prog)
    except Exception as exc:
        print(T("批处理失败：%s") % exc)
        return 1
    print("-" * 48)
    print(T("完成：共 %d 个文件，成功 %d，跳过 %d，失败 %d")
          % (res["total"], res["ok"], res["skip"], res["fail"]))
    for name, msg in res["errors"][:10]:
        print("  ✘ %s：%s" % (name, msg))
    if res["summary"]:
        print(T("汇总统计：%s") % res["summary"])
        sample = ["%s→%.1f" % (r["file"], r["x"]) for r in res["rows"][:5]
                  if r["x"] is not None]
        if sample:
            print(T("  主峰位示例：") + "  ".join(sample))
    return 0 if res["fail"] == 0 else 1


def _print_usage():
    """命令行用法（双语）。"""
    for line in (
            "用法：",
            "    双击“启动拉曼光谱工具.bat”或直接运行（不带参数）= 打开图形界面",
            "    jws2csv.py 文件/文件夹 [--xlsx --png --csv --peaks]    批量转换",
            "    jws2csv.py --identify 文件           未知光谱全库鉴定",
            "    jws2csv.py --pair 文件               与本地库自动配对",
            "    jws2csv.py --pair-batch 文件夹       批量配对（找对不上的）",
            "    jws2csv.py --identify-batch 文件夹   批量鉴定（逐条给最佳候选）",
            "    jws2csv.py --rruff-list              查看 RRUFF 数据包",
            "    jws2csv.py --mineral-search 名称     查内置矿物特征峰表",
            "    jws2csv.py --manual [路径]           打印 / 导出说明书",
            "    jws2csv.py --lang en|zh              切换界面语言",
            "    更多参数见说明书第 16 章（命令行速查）"):
        print(T(line))
    return 0


def _cli(args):
    # 界面语言：--lang en|zh 优先；否则按用户设置 / 系统界面语言
    _lang_pick = None
    for _i, _a in enumerate(args):
        if _a in ("--lang", "-lang", "--语言") and _i + 1 < len(args):
            _lang_pick = args[_i + 1]
            break
    if _lang_pick is not None:
        set_ui_lang(_lang_pick, persist=True, apply_now=False)
    else:
        set_ui_lang(detect_ui_lang(), persist=False, apply_now=False)
    if any(a in ("-h", "--help", "/?") for a in args):
        return _print_usage()
    formats = set()
    plot = {}
    rest = []
    db_query = None
    db_ids = None
    db_target = None
    pair_target = None
    pair_batch_arg = None
    pair_ref_arg = None
    db_ls = False
    rruff_list = False
    rruff_get = None
    rruff_query = None
    rruff_export_q = None
    rruff_fetch_q = None
    data_dir_arg = None
    cluster_sel = False
    cluster_cut = None
    map_arg = None
    map_metric = "main_peak"
    batch_folder = None
    out_arg = None
    report_flag = False
    cache_arg = None
    dl_conns_arg = None
    proxy_arg = None
    cleanup_arg = None
    import_pkgs = []
    mineral_q = None
    mineral_mode = "name"
    manual_arg = None
    identify_arg = None
    identify_batch_arg = None
    identify_top = 15
    identify_must = ""
    identify_not = ""
    identify_kind = "拉曼"
    identify_index = None
    identify_no_exact = False
    identify_all = False
    i = 0
    while i < len(args):
        a = args[i]
        low = a.lower()
        if low == "--csv":
            formats.add("csv")
        elif low in ("--xlsx", "--excel"):
            formats.add("xlsx")
        elif low == "--png":
            formats.add("png")
        elif low in ("--jcamp", "--jdx"):
            formats.add("jcamp")
        elif low == "--waterfall":
            formats.add("waterfall")
        elif low in ("--overlay", "--stack-overlay"):
            formats.add("overlay")
        elif low == "--both":
            formats.update(("csv", "xlsx"))
        elif low == "--all":
            formats.update(("csv", "xlsx", "png", "peaks", "fit"))
        elif low == "--peaks":
            formats.add("peaks")
        elif low in ("--fit", "--fitting"):
            formats.add("fit")
        elif low == "--despike":
            plot["despike"] = True
        elif low in ("--legend-pos", "--legend"):
            if i + 1 < len(args):
                value = args[i + 1].strip().lower()
                if value in LEGEND_POSITIONS:
                    plot["legend_pos"] = value
                    i += 1
        elif low == "--cluster":
            cluster_sel = True
        elif low == "--report":
            report_flag = True
        elif low == "--map":
            if i + 1 < len(args) and "," in args[i + 1]:
                parts = args[i + 1].replace("x", ",").replace("*", ",").split(",")
                try:
                    map_arg = (int(parts[0]), int(parts[1]))
                except (ValueError, IndexError):
                    map_arg = None
                i += 1
        elif low == "--map-metric":
            if i + 1 < len(args):
                map_metric = args[i + 1].strip()
                i += 1
        elif low in ("--batch", "--folder"):
            if i + 1 < len(args):
                batch_folder = args[i + 1]
                i += 1
        elif low in ("--out", "--out-dir"):
            if i + 1 < len(args):
                out_arg = args[i + 1]
                i += 1
        elif low == "--cache-limit":
            if i + 1 < len(args):
                cache_arg = args[i + 1]
                i += 1
            else:
                cache_arg = ""
        elif low == "--cleanup":
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                cleanup_arg = args[i + 1]
                i += 1
            else:
                cleanup_arg = "0"
        elif low in ("--dl-conns", "--connections"):
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                dl_conns_arg = args[i + 1]
                i += 1
            else:
                dl_conns_arg = ""
        elif low == "--proxy":
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                proxy_arg = args[i + 1]
                i += 1
            else:
                proxy_arg = ""
        elif low in ("--import-pkg", "--import-zip"):
            if i + 1 < len(args):
                import_pkgs.append(args[i + 1])
                i += 1
        elif low in ("--mineral-search", "--by-element", "--by-formula", "--mineral-info",
                     "--rruff-info"):
            if low == "--by-element":
                mineral_mode = "element"
            elif low == "--by-formula":
                mineral_mode = "formula"
            elif low == "--mineral-info":
                mineral_mode = "info"
            elif low == "--rruff-info":
                mineral_mode = "rruffinfo"
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                mineral_q = args[i + 1]
                i += 1
        elif low in ("--manual", "--help-text", "--说明书"):
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                manual_arg = args[i + 1]
                i += 1
            else:
                manual_arg = ""
        elif low in ("--identify", "--search-unknown", "--鉴定"):
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                identify_arg = args[i + 1]
                i += 1
        elif low == "--identify-batch":
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                identify_batch_arg = args[i + 1]
                i += 1
        elif low == "--identify-top":
            if i + 1 < len(args):
                try:
                    identify_top = max(1, int(float(args[i + 1])))
                except ValueError:
                    pass
                i += 1
        elif low == "--identify-must":
            if i + 1 < len(args):
                identify_must = args[i + 1]
                i += 1
        elif low == "--identify-not":
            if i + 1 < len(args):
                identify_not = args[i + 1]
                i += 1
        elif low == "--identify-kind":
            if i + 1 < len(args):
                identify_kind = args[i + 1].strip()
                i += 1
        elif low == "--identify-index":
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                identify_index = args[i + 1]
                i += 1
            else:
                identify_index = ""
        elif low == "--identify-no-exact":
            identify_no_exact = True
        elif low == "--identify-all":
            identify_all = True
        elif low == "--baseline":
            if i + 1 < len(args):
                plot["baseline"] = args[i + 1].strip().lower()
                i += 1
        elif low == "--mineral":
            if i + 1 < len(args):
                plot["mineral_name"] = args[i + 1]
                i += 1
        elif low == "--y-ticks":
            plot["show_y_ticks"] = True
        elif low == "--peak-label-rel":
            plot["peak_label_rel"] = True
        elif low == "--no-grid":
            plot["show_grid"] = False
        elif low == "--no-title":
            plot["show_title"] = False
        elif low == "--no-peaks":
            plot["annotate_peaks"] = False
        elif low == "--no-peak-dash":
            plot["peak_dash_line"] = False
        elif low in ("--no-peak-labels", "--no-peak-numbers"):
            plot["peak_labels"] = False
        elif low == "--peak-merge":
            if i + 1 < len(args):
                try:
                    plot["peak_merge_tol"] = max(0.001, float(args[i + 1]))
                except ValueError:
                    pass
                i += 1
        elif low == "--no-peak-merge":
            plot["merge_peak_labels"] = False
        elif low == "--normalize":
            if i + 1 < len(args):
                plot["normalize"] = args[i + 1].strip().lower()
                i += 1
        elif low in ("--calib", "--shift"):
            if i + 1 < len(args):
                parts = args[i + 1].replace(";", ",").split(",")
                try:
                    if len(parts) >= 4:
                        plot["calib_pairs"] = [[float(parts[0]), float(parts[1])],
                                               [float(parts[2]), float(parts[3])]]
                    elif len(parts) >= 2:
                        plot["calib_pairs"] = [[float(parts[0]), float(parts[1])]]
                except ValueError:
                    pass
                i += 1
        elif low == "--cluster-cut":
            if i + 1 < len(args):
                try:
                    cluster_cut = float(args[i + 1])
                except ValueError:
                    cluster_cut = None
                i += 1
        elif low == "--manual":
            if i + 1 < len(args):
                manual = []
                for token in args[i + 1].replace(";", ",").split(","):
                    token = token.strip()
                    if token:
                        try:
                            manual.append(float(token))
                        except ValueError:
                            pass
                if manual:
                    plot["manual_peaks"] = manual
                i += 1
        elif low == "--db-search":
            if i + 1 < len(args):
                db_query = args[i + 1]
                i += 1
        elif low == "--db-get":
            if i + 1 < len(args):
                db_ids = [t for t in args[i + 1].replace(";", ",").split(",") if t.strip()]
                i += 1
        elif low == "--db-match":
            if i + 1 < len(args):
                db_target = args[i + 1]
                i += 1
        elif low == "--pair":
            if i + 1 < len(args):
                pair_target = args[i + 1]
                i += 1
        elif low == "--pair-batch":
            if i + 1 < len(args):
                pair_batch_arg = args[i + 1]
                i += 1
        elif low == "--pair-ref":
            if i + 1 < len(args):
                pair_ref_arg = args[i + 1]
                i += 1
        elif low == "--db-ls":
            db_ls = True
        elif low == "--rruff-list":
            rruff_list = True
        elif low == "--rruff-get":
            if i + 1 < len(args):
                rruff_get = args[i + 1]
                i += 1
        elif low == "--rruff-search":
            if i + 1 < len(args):
                rruff_query = args[i + 1]
                i += 1
        elif low == "--rruff-export":
            if i + 1 < len(args):
                rruff_export_q = args[i + 1]
                i += 1
        elif low == "--rruff-fetch":
            if i + 1 < len(args):
                rruff_fetch_q = args[i + 1]
                i += 1
        elif low in ("--data-dir", "--data"):
            if i + 1 < len(args) and not args[i + 1].startswith("-"):
                data_dir_arg = args[i + 1]
                i += 1
            else:
                data_dir_arg = ""
        elif low in ("--x-step", "--x-start", "--x-min", "--x-max", "--y-min", "--y-max",
                     "--peak-dist", "--peak-thresh", "--smooth", "--fig-width", "--fig-height",
                     "--despike-thresh", "--baseline-degree", "--baseline-iters",
                     "--baseline-window", "--stack-offset"):
            if i + 1 < len(args):
                try:
                    val = float(args[i + 1])
                except ValueError:
                    val = None
                if val is not None:
                    key = {
                        "--x-step": "x_step", "--x-start": "x_start",
                        "--x-min": "x_min", "--x-max": "x_max",
                        "--y-min": "y_min", "--y-max": "y_max",
                        "--peak-dist": "peak_min_dist", "--peak-thresh": "peak_thresh_pct",
                        "--smooth": "smooth_window", "--fig-width": "fig_width",
                        "--fig-height": "fig_height",
                        "--despike-thresh": "despike_thresh",
                        "--baseline-degree": "baseline_degree",
                        "--baseline-iters": "baseline_iters",
                        "--baseline-window": "baseline_window",
                        "--stack-offset": "stack_offset",
                    }[low]
                    if key in ("smooth_window", "fig_width", "fig_height",
                               "baseline_degree", "baseline_iters"):
                        plot[key] = int(val)
                    else:
                        plot[key] = val
                i += 1
        elif low == "--despike-window":
            if i + 1 < len(args):
                try:
                    plot["despike_window"] = max(3, int(float(args[i + 1])))
                except ValueError:
                    pass
                i += 1
        elif low in ("--lang", "-lang", "--语言"):
            if i + 1 < len(args):
                i += 1
        else:
            rest.append(a)
        i += 1
    if db_query is not None:
        rows = rod_search(db_query)
        if not rows:
            print(T("数据库中没有找到：%s") % db_query)
            return 1
        print(T("共 %d 条结果（数据来源：ROD / RRUFF 拉曼库）：") % len(rows))
        for row in rows:
            print("  %-10s %-18s %-24s %6s nm   %s" % (
                row.get("file") or "", (row.get("mineral") or "")[:18],
                (row.get("chemname") or "")[:24], row.get("wavelength") or "",
                row.get("devicecompany") or ""))
        print("\n" + T("下载参考谱：--db-get 编号[,编号...]"))
        return 0
    if db_ls:
        files = db_list()
        print(T("本地数据库：%s（%d 条）") % (db_dir(), len(files)))
        for path in files:
            print("  " + os.path.basename(path))
        return 0
    if db_ids:
        ok = 0
        for rid in db_ids:
            try:
                dst, _meta = db_download(rid.strip())
                print("OK  %s  ->  %s" % (rid.strip(), os.path.basename(dst)))
                ok += 1
            except Exception as exc:
                print("ERR %s  ->  %s" % (rid.strip(), exc))
        print("\n" + T("完成：成功 %d / %d，保存在 %s") % (ok, len(db_ids), db_dir()))
        return 0 if ok else 1
    if db_target is not None:
        local = db_list()
        if not local:
            print(T("本地数据库为空，请先用 --db-get 下载参考谱。"))
            return 1
        try:
            if db_target.lower().endswith(".jws"):
                spec = Spectrum(db_target)
                txs = _calibrate(spec.x_values(), plot)
                tys = list(spec.y_data[0])
            else:
                _xl, rows = load_csv_series(db_target)
                txs, tys = rows[0][1], rows[0][2]
        except Exception as exc:
            print(T("读取待比对文件失败：%s") % exc)
            return 1
        results = match_references((txs, tys), local, top=20)
        if not results:
            print(T("与本地数据库没有可比较的重叠区间。"))
            return 1
        print(T("与本地数据库比对（库内 %d 条）：") % len(local))
        for rank, r in enumerate(results, 1):
            print(T("  %2d. %-40s 相关系数 %.4f  谱角 %.2f°") % (
                rank, r["name"][:40], r["corr"], r["angle"]))
        return 0
    if pair_target is not None:
        local = db_list()
        if not local:
            print(T("本地数据库为空，请先用 --db-get 下载参考谱。"))
            return 1
        try:
            if pair_target.lower().endswith(".jws"):
                spec = Spectrum(pair_target)
                txs = _calibrate(spec.x_values(), plot)
                tys = list(spec.y_data[0])
            else:
                _xl, rows = load_csv_series(pair_target)
                txs, tys = rows[0][1], rows[0][2]
        except Exception as exc:
            print(T("读取待配对文件失败：%s") % exc)
            return 1
        results = pair_spectra((txs, tys), local, plot, 5.0, top=20)
        if not results:
            print(T("与本地数据库没有可比较的重叠区间。"))
            return 1
        best = results[0]
        base = os.path.splitext(os.path.basename(pair_target))[0]
        dst = os.path.join(os.path.dirname(os.path.abspath(pair_target)),
                           base + "_配对报告.csv")
        with open(dst, "w", encoding="utf-8-sig", newline="") as f:
            f.write(T("== 配对排名（本地库 %d 条；按峰位匹配 F1 优先排序）==\n") % len(local))
            f.write(T("排名,参考谱,峰位匹配F1(%),命中峰数,实测峰数,参考峰数,相关系数,谱角(度)\n"))
            for rank, r in enumerate(results, 1):
                f.write("%d,%s,%.1f,%d,%d,%d,%.4f,%.2f\n" % (
                    rank, r["name"], r["peak_score"], r["peak_matched"],
                    r["peak_left"], r["peak_right"], r["corr"], r["angle"]))
            f.write("\n== " + T("最佳配对峰位对照（容差 5 cm-1）：%s") % best["name"] + " ==\n")
            f.write(T("实测峰位,参考峰位,偏差(cm-1),是否匹配\n"))
            for row in best.get("pair_rows", []):
                f.write("%s,%s,%s,%s\n" % (
                    "" if row[0] is None else "%.2f" % row[0],
                    "" if row[1] is None else "%.2f" % row[1],
                    "" if row[2] is None else "%.2f" % row[2],
                    T("匹配") if row[3] else T("单侧")))
        print(T("配对排名（本地库 %d 条，按峰位匹配 F1 优先）：") % len(local))
        for rank, r in enumerate(results, 1):
            print(T("  %2d. %-32s F1 %5.1f%% (命中%d/%d 参考%d)  r=%.4f  谱角 %.2f°") % (
                rank, r["name"][:32], r["peak_score"], r["peak_matched"],
                r["peak_left"], r["peak_right"], r["corr"], r["angle"]))
        print("\n" + T("最佳配对：%s") % best["name"])
        print(T("  实测峰位    参考峰位    偏差     匹配"))
        for row in best.get("pair_rows", []):
            print("  %-11s %-11s %-8s %s" % (
                "%.2f" % row[0] if row[0] is not None else "-",
                "%.2f" % row[1] if row[1] is not None else "-",
                "%+.2f" % row[2] if row[2] is not None else "-",
                T("匹配") if row[3] else T("单侧")))
        try:
            png = os.path.join(os.path.dirname(os.path.abspath(pair_target)),
                               base + "_配对报告.png")
            render_pair_report(png, os.path.basename(pair_target), (txs, tys), results, plot, 5.0)
            print("\n" + T("报告图已保存：%s") % png)
        except Exception as exc:
            print("\n" + T("报告图生成失败：%s") % exc)
        print(T("报告已保存：%s") % dst)
        return 0

    if pair_batch_arg is not None:
        if pair_ref_arg:
            ref_paths = find_input_files([pair_ref_arg])
            ref_label = os.path.basename(os.path.abspath(pair_ref_arg))
        else:
            ref_paths = db_list()
            ref_label = T("本地参考谱库")
        if not ref_paths:
            print(T("没有找到参考谱：用 --pair-ref 指定参考谱文件或文件夹，"))
            print(T("或先用 --db-get 把参考谱下载到本地库。"))
            return 1
        refs = load_reference_set(ref_paths)
        if not refs:
            print(T("参考谱一条也读不出来（需要两列文本 / CSV）：%s") % ref_label)
            return 1
        files = [p for p in find_input_files([pair_batch_arg])
                 if not any(os.path.basename(p).lower().endswith(sfx)
                            for sfx in _SKIP_FILE_SUFFIXES)]
        spectra = collect_spectra(files, plot, verbose=True)
        if not spectra:
            print(T("没有读到待配对的实测谱：%s") % pair_batch_arg)
            return 1
        print(T("批量配对：%d 条实测谱 × %d 条参考谱（%s）")
              % (len(spectra), len(refs), ref_label))
        rows = batch_pair(
            spectra, refs, plot,
            progress=lambda i, n, nm: print("  [%d/%d] %s" % (i + 1, n, nm)) if nm else None)
        out = results_dir()
        base = _safe_name(os.path.splitext(os.path.basename(
            os.path.abspath(pair_batch_arg)))[0]) or "批量配对"
        csv_dst = os.path.join(out, "%s_批量配对汇总.csv" % base)
        html_dst = os.path.join(out, "%s_批量配对汇总.html" % base)
        write_batch_pair_csv(csv_dst, rows, ref_label)
        try:
            write_batch_pair_html(html_dst, rows, ref_label)
        except Exception as exc:
            print(T("HTML 汇总生成失败：%s") % exc)
        print()
        print(T("批量配对结果（按综合分升序，最可疑的排在最前）："))
        print("  %-28s %-28s %7s %7s %9s  %s" % (
            T("实测谱"), T("最佳参考谱"), T("综合分"), T("F1(%)"),
            T("强峰命中"), T("参考判读")))
        for row in sorted(rows, key=_pair_sort_key):
            best = row.get("best")
            if not best:
                print("  %-28s %-28s %7s %7s %9s  %s" % (
                    row["name"][:28], "-", "-", "-", "-",
                    row.get("error") or T("没有可比对的参考谱")))
                continue
            print("  %-28s %-28s %7.1f %7.1f %9s  %s" % (
                row["name"][:28], best["name"][:28], best["score"], best["f1"],
                "%d/%d" % (best["strong_hit"], best["strong_n"]),
                pair_reading(best)))
        print()
        print(T(_BATCH_SCORE_NOTE))
        print(T("汇总表：%s") % csv_dst)
        print(T("汇总报告：%s") % html_dst)
        return 0

    if data_dir_arg is not None:
        if data_dir_arg:
            set_data_root(data_dir_arg)
            print(T("数据文件夹已设置为：%s") % data_root())
        else:
            migrate_legacy_layout()
            print(T("数据文件夹：%s") % data_root())
        print(T("（下载的参考谱、RRUFF 数据包、分析结果都保存在此目录）"))
        total_n = 0
        total_s = 0
        for label, path in data_folders():
            n, s = dir_stats(path)
            total_n += n
            total_s += s
            print(T("   %-10s %5d 个文件  %9.2f MB   %s") % (T(label), n, s / 1048576.0, path))
        print(T("   合计：%d 个文件，%.2f MB") % (total_n, total_s / 1048576.0))
        return 0
    if identify_index is not None:
        avail = rruff_available_keys()
        if not avail:
            print(T("本地还没有数据包。请先 --rruff-get（或 --import-pkg）下载。"))
            return 1
        targets = [identify_index] if identify_index else avail
        for key in targets:
            pkg = rruff_package(key)
            if pkg is None:
                print(T("未知数据包：%s") % key)
                continue
            print(T("正在为 %s（%s）建立特征索引…") % (key, pkg.get("label", "")))

            def prog(i, n, _k=key):
                if n and (i % 100 == 0 or i + 1 == n):
                    print("   %s %d/%d" % (_k, i + 1, n))

            rows = rruff_peak_index(key, progress=prog, rebuild=True)
            print(T("   完成：%d 条特征 -> %s") % (len(rows), os.path.basename(
                _peak_index_path(key))))
        print("\n" + T("现在可以鉴定未知光谱了：--identify 你的谱.csv"))
        return 0
    if identify_arg is not None:
        spectra = collect_spectra([identify_arg], plot, verbose=True)
        if not spectra:
            print(T("没有读到光谱：%s") % identify_arg)
            return 1
        name, txs, tys = spectra[0]
        keys = peak_index_keys()
        if not keys:
            print(T("还没有可用于检索的特征索引。"))
            print(T("请先：--rruff-get <数据包key> 然后 --identify-index（为已下载包建索引）"))
            return 1
        kind_low = identify_kind.strip().lower()
        kinds = None if kind_low in ("all", "全部", "*", "") else {identify_kind.strip()}
        print(T("未知光谱：%s（%d 点，%d 个峰）")
              % (name, len(txs), len(analyze_peaks(txs, tys, plot))))
        print(T("检索范围：%s；已建索引的数据包 %d 个：%s")
              % (T(identify_kind) if identify_kind else T("全部"), len(keys),
                 ("、" if ui_lang() == "zh" else ", ").join(keys)))
        if identify_must or identify_not:
            print(T("元素筛选：必须含 [%s]；必须不含 [%s]")
                  % (identify_must or "-", identify_not or "-"))
        print(T("正在检索…"))
        try:
            results, total, verdict = identify_unknown(
                (txs, tys), plot, keys=keys, kinds=kinds, must=identify_must,
                not_=identify_not, top=identify_top, group=not identify_all,
                exact_pool=0 if identify_no_exact else max(IDENTIFY_POOL, identify_top))
        except JwsError as exc:
            print(T("检索失败：%s") % exc)
            return 1
        print(T("共比对 %d 条参考记录，以下按可信度排序：") % total)
        print()
        print_identify_table(results)
        print(T("结　论：%s") % verdict)
        base = os.path.splitext(os.path.basename(identify_arg))[0]
        dst = os.path.join(results_dir(), "%s_未知光谱检索.csv" % _safe_name(base))
        write_identify_csv(dst, name, results, verdict, total)
        print(T("候选清单：%s") % dst)
        best = results[0] if results else None
        if best and best.get("ref_xy"):
            try:
                png = os.path.join(results_dir(), "%s_未知光谱检索_对比.png" % _safe_name(base))
                render_png(png, [(name, txs, tys, _PALETTE[0]),
                                 ("%s %s" % (best["name"], best["rruffid"]),
                                  best["ref_xy"][0], best["ref_xy"][1], _PALETTE[1])],
                           T("未知光谱 vs 最佳候选 %s") % best["name"],
                           _DEFAULT_X_HEADER, "Intensity", plot)
                print(T("对比图：  %s") % png)
            except Exception as exc:
                print(T("对比图生成失败：%s") % exc)
        return 0
    if identify_batch_arg is not None:
        files = [p for p in find_input_files([identify_batch_arg])
                 if not any(os.path.basename(p).lower().endswith(sfx)
                            for sfx in _SKIP_FILE_SUFFIXES)]
        spectra = collect_spectra(files, plot, verbose=True)
        if not spectra:
            print(T("没有读到光谱：%s") % identify_batch_arg)
            return 1
        keys = peak_index_keys()
        if not keys:
            print(T("还没有可用于检索的特征索引。"))
            print(T("请先：--rruff-get <数据包key> 然后 --identify-index（为已下载包建索引）"))
            return 1
        kind_low = identify_kind.strip().lower()
        kinds = None if kind_low in ("all", "全部", "*", "") else {identify_kind.strip()}
        print(T("批量鉴定：%d 条未知谱；检索范围 %s；已建索引的数据包 %d 个：%s")
              % (len(spectra), T(identify_kind) if identify_kind else T("全部"),
                 len(keys), ("、" if ui_lang() == "zh" else ", ").join(keys)))
        if identify_must or identify_not:
            print(T("元素筛选：必须含 [%s]；必须不含 [%s]")
                  % (identify_must or "-", identify_not or "-"))
        rows = batch_identify(
            spectra, plot, keys=keys, kinds=kinds, must=identify_must,
            not_=identify_not, top=identify_top,
            progress=lambda i, n, nm: print("  [%d/%d] %s" % (i + 1, n, nm)) if nm else None)
        total_entries = max([r.get("total", 0) for r in rows] or [0])
        out = results_dir()
        base = _safe_name(os.path.splitext(os.path.basename(
            os.path.abspath(identify_batch_arg)))[0]) or "批量鉴定"
        csv_dst = os.path.join(out, "%s_批量鉴定汇总.csv" % base)
        html_dst = os.path.join(out, "%s_批量鉴定汇总.html" % base)
        write_batch_identify_csv(csv_dst, rows, total_entries, identify_top)
        try:
            write_batch_identify_html(html_dst, rows, total_entries)
        except Exception as exc:
            print(T("HTML 汇总生成失败：%s") % exc)
        print()
        print(T("批量鉴定结果（按最佳候选综合分升序，最可疑的排在最前）："))
        print("  %-28s %-26s %7s %7s  %s" % (
            T("实测谱"), T("最佳候选"), T("综合分"), T("F1(%)"), T("参考判读")))
        for row in sorted(rows, key=_identify_sort_key):
            best = _best_named(row.get("candidates"))
            if not best:
                print("  %-28s %-26s %7s %7s  %s" % (
                    row["name"][:28], "-", "-", "-",
                    (row.get("verdict") or T("没有找到候选"))[:60]))
                continue
            f1 = best["exact_f1"] if best.get("exact_f1") is not None else best["f1"]
            print("  %-28s %-26s %7.1f %7.1f  %s" % (
                row["name"][:28], best["name"][:26], best["score"], f1,
                row.get("verdict", "")[:60]))
        print()
        print(T("汇总表：%s") % csv_dst)
        print(T("汇总报告：%s") % html_dst)
        return 0
    if manual_arg is not None:
        if manual_arg:
            target = os.path.abspath(manual_arg)
            if os.path.isdir(target):
                target = os.path.join(target, "使用说明.txt")
            try:
                with open(target, "w", encoding="utf-8-sig", newline="") as f:
                    f.write(manual_text())
            except OSError as exc:
                print(T("导出说明书失败：%s") % exc)
                return 1
            print(T("说明书已导出：%s") % target)
            return 0
        print(manual_text())
        return 0
    if mineral_q:
        if mineral_mode in ("info", "rruffinfo"):
            rec = lookup_mineral(mineral_q)
            name = rec["en"] if rec else mineral_q
            if mineral_mode == "info":
                text = mineral_full_info(name)
            else:
                sample = None
                try:
                    sample = rruff_sample_info(mineral_q)
                except Exception:
                    sample = None
                text = rruff_info_text(sample)
                if not text:
                    print((T("本地数据包里没有 \"%s\" 的记录，或尚未下载数据包。")
                           + T("可先用 --rruff-get 或 --import-pkg 准备数据。")) % mineral_q)
                    return 1
            if not text:
                print(T("内置矿物表里没有「%s」。可用 --mineral-search <名称> 查找。") % mineral_q)
                return 1
            print(text)
            return 0
        mode_cn = {"element": "元素", "formula": "化学式", "name": "名称"}[mineral_mode]
        recs = mineral_search(mineral_q, mineral_mode)
        if not recs:
            print(T("内置矿物表里没有匹配（检索方式：%s）：%s") % (T(mode_cn), mineral_q))
            print(T("内置表共 %d 种常见矿物，可用 --mineral-search 名称 模糊查找。")
                  % len(_MINERAL_DB))
            return 1
        print(T("按%s检索 \"%s\"：命中 %d 种矿物") % (T(mode_cn), mineral_q, len(recs)))
        for rec in recs:
            print("  %-22s %-24s %-14s %s" % (rec["en"], rec["cn"], rec["formula"],
                                              T(rec["system"])))
            print(T("      特征峰：%s") % "、".join("%.0f" % b for b, _t in rec["bands"]))
        if len(recs) == 1:
            print("\n" + mineral_info_text(recs[0]["en"]))
        return 0
    if dl_conns_arg is not None or proxy_arg is not None:
        data = _load_settings()
        if dl_conns_arg is not None:
            if dl_conns_arg:
                try:
                    value = int(float(dl_conns_arg))
                except ValueError:
                    value = -1
                if value < 1 or value > DOWNLOAD_CONNECTIONS_MAX:
                    print(T("并发连接数要在 1 ~ %d 之间。") % DOWNLOAD_CONNECTIONS_MAX)
                    return 1
                data["download_conns"] = str(value)
            else:
                data.pop("download_conns", None)
        if proxy_arg is not None:
            if proxy_arg:
                data["proxy"] = proxy_arg
            else:
                data.pop("proxy", None)
        _save_settings(data)
        print(T("下载并发连接数：%d（可用 --dl-conns 1~%d 调整）")
              % (download_connections(), DOWNLOAD_CONNECTIONS_MAX))
        print(T("下载代理：%s") % (proxy_url() or T("直连（不用代理）")))
        sys_px = system_proxy()
        if sys_px and not str(data.get("proxy") or "").strip():
            print(T("（检测到 Windows 系统代理：%s，已自动使用）") % sys_px)
        print(T("提示：跨境高丢包链路下并发数几乎决定速度；有代理 / VPN 时走代理通常更快。"))
        return 0
    if cache_arg is not None:
        if cache_arg:
            try:
                set_cache_limit_mb(float(cache_arg))
            except ValueError:
                print(T("容量上限必须是数字（单位 MB，0 表示不限制）。"))
                return 1
        n, s = cache_usage()
        over, limit, size = over_cache_limit()
        free, total = disk_free()
        print(T("数据文件夹：%s") % data_root())
        print(T("数据包目录：%s") % rruff_dir())
        print(T("  数据包：%d 个文件，占用 %.1f MB") % (n, s / 1048576.0))
        if limit:
            print(T("  容量上限：%.0f MB%s") % (limit, T("（已超限，建议清理）") if over else ""))
        else:
            print(T("  容量上限：不限制"))
        if free is not None:
            print(T("  所在磁盘可用：%.1f GB / 共 %.1f GB")
                  % (free / 1073741824.0, total / 1073741824.0))
        for key, path, sz in package_sizes():
            pkg = rruff_package(key)
            print("   %-22s %8.1f MB   %s" % (key, sz / 1048576.0,
                                              pkg_label(pkg)))
        print("\n" + T("清理全部数据包：--cleanup ；只释放指定空间：--cleanup 500（MB）"))
        return 0
    if cleanup_arg is not None:
        try:
            want = float(cleanup_arg or 0)
        except ValueError:
            want = 0.0
        removed, freed = cleanup_packages(target_mb=want)
        if removed:
            print(T("已删除 %d 个数据包，释放 %.1f MB。") % (removed, freed / 1048576.0))
        else:
            print(T("没有可清理的数据包。"))
        return 0 if removed else 1
    if import_pkgs:
        bad = 0
        for src in import_pkgs:
            try:
                key, count = rruff_import_zip(src)
                print(T("导入成功：%s → %s（索引 %d 条）") % (os.path.basename(src), key, count))
            except Exception as exc:
                print(T("导入失败：%s → %s") % (os.path.basename(src), exc))
                bad += 1
        print(T("数据包目录：%s") % rruff_dir())
        return 1 if bad else 0
    if cluster_sel:
        spectra = collect_spectra(rest, plot, verbose=True)
        if len(spectra) < 2:
            print(T("聚类至少需要 2 条光谱（请给出文件或文件夹）。"))
            return 1
        res = cluster_spectra(spectra, cut=cluster_cut)
        if res is None:
            print(T("所选光谱没有共同的波数区间，无法聚类。"))
            return 1
        out = results_dir()
        group_csv = os.path.join(out, "聚类分析_分组.csv")
        with open(group_csv, "w", encoding="utf-8-sig", newline="") as f:
            f.write(T("文件,簇号,PC1,PC2,PC3\n"))
            for i, name in enumerate(res["names"]):
                sc = (res["scores"][i] if i < len(res["scores"]) else []) + [0.0, 0.0, 0.0]
                f.write("%s,%d,%.4f,%.4f,%.4f\n" % (name, res["labels"][i],
                                                    sc[0], sc[1], sc[2]))
        tree_png = os.path.join(out, "聚类分析_树状图.png")
        render_dendrogram(tree_png, res["names"], res["merges"], T("层次聚类树状图"),
                          cut=res["cut"], groups=res["labels"])
        print(T("聚类完成：%d 条光谱 → %d 个簇（自动分割阈值 %.3f）")
              % (len(res["names"]), len(res["groups"]), res["cut"]))
        for cid in sorted(res["groups"]):
            members = res["groups"][cid]
            shown = ("、" if ui_lang() == "zh" else ", ").join(
                res["names"][i] for i in members[:5])
            print(T("  簇%d（%d 条）：%s%s") % (cid, len(members), shown,
                                          " …" if len(members) > 5 else ""))
        print(T("分组表：%s") % group_csv)
        print(T("树状图：%s") % tree_png)
        if res["scores"] and len(res["scores"][0]) >= 2:
            xlab = "PC1"
            ylab = "PC2"
            if res["explained"]:
                pcm = "PC1（%.1f%%）" if ui_lang() == "zh" else "PC1 (%.1f%%)"
                xlab = pcm % res["explained"][0]
                if len(res["explained"]) > 1:
                    pcm2 = "PC2（%.1f%%）" if ui_lang() == "zh" else "PC2 (%.1f%%)"
                    ylab = pcm2 % res["explained"][1]
            png = os.path.join(out, "聚类分析_主成分散点.png")
            render_scatter(png, [(s[0], s[1]) for s in res["scores"]],
                           T("主成分散点（PC1-PC2）"), xlab, ylab,
                           groups=res["labels"], labels=res["names"])
            print(T("主成分散点：%s") % png)
        return 0
    if map_arg:
        rows_n, cols_n = map_arg
        if rows_n < 1 or cols_n < 1 or rows_n * cols_n > 10000:
            print(T("二维成像网格无效：--map 行,列（例如 --map 5,5）。"))
            return 1
        spectra = collect_spectra(rest, plot, verbose=True)
        need = rows_n * cols_n
        if len(spectra) < need:
            print(T("二维成像需要 %d×%d = %d 个点位，当前只读到 %d 条。")
                  % (rows_n, cols_n, need, len(spectra)))
            return 1
        spectra = spectra[:need]
        metric = map_metric
        metric_cn = {"main_peak": "主峰位(cm-1)", "intensity": "主峰强度",
                     "fwhm": "主峰半高宽FWHM", "peaks": "识别峰数",
                     "total": "总强度"}.get(metric, metric)
        target = None
        if metric.startswith("at:"):
            try:
                target = float(metric[3:])
            except ValueError:
                target = None
            metric_cn = "%.1f cm-1 附近峰强度" % (target or 0.0)
        metric_shown = T(metric_cn)
        values = []
        for name, xs, ys in spectra:
            peaks = analyze_peaks(xs, ys, plot)
            if not peaks:
                values.append(None)
                continue
            main = max(peaks, key=lambda pk: pk["prominence"])
            if target is not None:
                near = min(peaks, key=lambda pk: abs(pk["x"] - target))
                values.append(near["y"] if abs(near["x"] - target) <= 20 else None)
            elif metric == "intensity":
                values.append(main["y"])
            elif metric == "fwhm":
                values.append(main["fwhm"])
            elif metric == "peaks":
                values.append(float(len(peaks)))
            elif metric == "total":
                values.append(sum(ys))
            else:
                values.append(main["x"])
        grid = [values[r * cols_n:(r + 1) * cols_n] for r in range(rows_n)]
        out = results_dir()
        png = os.path.join(out, "二维成像_%s.png" % _safe_name(metric_cn))
        render_heatmap(png, grid, T("二维成像 · %s") % metric_shown, metric_shown, plot,
                       row_labels=[T("第%d行") % (r + 1) for r in range(rows_n)])
        csv_path = os.path.join(out, "二维成像_%s.csv" % _safe_name(metric_cn))
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(T("行\\列,") + ",".join(str(c + 1) for c in range(cols_n)) + "\n")
            for r in range(rows_n):
                cells = ["" if v is None else "%.6g" % v for v in grid[r]]
                f.write(T("第%d行,") % (r + 1) + ",".join(cells) + "\n")
        valid = [v for v in values if v is not None]
        print(T("二维成像完成：%d 个点位，%s 范围 %.4g ~ %.4g")
              % (len(spectra), metric_shown, min(valid), max(valid)) if valid else
              T("二维成像完成：%d 个点位（无有效数值）") % len(spectra))
        print(T("热图：%s") % png)
        print(T("数值表：%s") % csv_path)
        return 0
    if batch_folder:
        return _cli_batch(batch_folder, out_arg, formats, plot)
    if report_flag:
        out_path = out_arg
        if out_path and os.path.isdir(out_path):
            out_path = os.path.join(out_path, "光谱分析报告.html")
        try:
            path, count = build_report(rest, out_path, plot=plot, verbose=True)
        except Exception as exc:
            print(T("生成报告失败：%s") % exc)
            return 1
        print(T("报告已生成（%d 条光谱）：%s") % (count, path))
        return 0
    if rruff_list:
        print(T("RRUFF 数据包（来源 www.rruff.net/zipped_data_files）："))
        print("  %-22s %-6s %-30s %-8s %s" % (T("数据包 key"), T("类型"), T("说明"), T("大小"), T("状态")))
        for pkg in all_packages():
            have = os.path.exists(rruff_zip_path(pkg["key"]))
            idx = _read_index(pkg["key"])
            if have and idx:
                state = T("已下载，索引 %d 条") % len(idx)
            elif have:
                state = T("已下载，未建索引")
            else:
                state = T("未下载")
            print("  %-22s %-6s %-30s %-8s %s" % (
                pkg["key"], T(pkg.get("kind", "")), T(pkg["label"]), pkg["size"], state))
        print("\n" + T("下载并建索引：--rruff-get 数据包key"))
        print(T("按矿物检索：  --rruff-search Zircon（也可直接输 RRUFF 编号，如 R050034）"))
        print(T("导出到本地库：--rruff-export Zircon"))
        print(T("导入自己的包：--import-pkg 路径.zip"))
        print(T("查看占用/清理：--cache-limit ；--cleanup 500"))
        return 0
    if rruff_get:
        pkg = rruff_package(rruff_get)
        if pkg is None:
            print(T("未知的数据包 key：%s（用 --rruff-list 查看可选值）") % rruff_get)
            return 1
        print(T("正在下载 %s（%s）…") % (pkg["label"], pkg["size"]))
        try:
            rruff_download(rruff_get, lambda got, total, speed, eta: None)
        except Exception as exc:
            print(T("下载失败：%s") % exc)
            return 1
        print(T("正在建立索引 …"))
        rows = rruff_index(rruff_get, rebuild=True)
        print(T("完成：%d 条，保存在 %s") % (len(rows), rruff_dir()))
        warn = cache_limit_warning()
        if warn:
            print("\n" + warn)
        return 0
    if rruff_query:
        rows = rruff_search(rruff_query)
        print(T("检索 \"%s\" -> %d 条") % (rruff_query, len(rows)))
        for row in rows[:40]:
            print("  %-18s %-9s %-6s %-20s %s" % (
                row[1][:18], row[2], row[3], row[4][:20], row[0]))
        return 0
    if rruff_export_q:
        rows = rruff_search(rruff_export_q)
        if not rows:
            print(T("未找到 \"%s\"。请先用 --rruff-get 下载数据包（或 --rruff-list 查看）。")
                  % rruff_export_q)
            return 1
        saved, hits = rruff_export_matches(rruff_export_q)
        print(T("命中 %d 条，已导出 %d 条（同类优先 Processed）到本地库：%s")
              % (hits, len(saved), db_dir()))
        for path in saved:
            print("   " + os.path.basename(path))
        return 0
    if rruff_fetch_q:
        if not rruff_available_keys():
            print(T("本地还没有 RRUFF 数据包，先下载最小的（未评级·非定向，12 MB）…"))
            rruff_download("unrated_unoriented", lambda got, total, speed, eta: None)
            rows = rruff_index("unrated_unoriented", rebuild=True)
            print(T("索引完成：%d 条") % len(rows))
        saved, hits = rruff_export_matches(rruff_fetch_q)
        if not hits:
            print(T("已下载的数据包里没有找到 \"%s\"。") % rruff_fetch_q)
            print(T("可下载更大的数据包后重试（每个约 1~2 分钟）："))
            for pkg in rruff_packages_sorted():
                if not os.path.exists(rruff_zip_path(pkg["key"])):
                    print("   --rruff-get %-22s %s" % (pkg["key"], pkg["size"]))
            return 1
        print(T("按矿物批量抓取 \"%s\"：命中 %d 条，导出 %d 条到 %s")
              % (rruff_fetch_q, hits, len(saved), db_dir()))
        for path in saved:
            print("   " + os.path.basename(path))
        return 0

    if not formats:
        formats = {"csv"}
    want_waterfall = "waterfall" in formats
    want_overlay = "overlay" in formats
    formats = set(f for f in formats if f not in ("waterfall", "overlay"))
    out_dir_arg = out_arg
    if out_dir_arg:
        os.makedirs(out_dir_arg, exist_ok=True)
    files = find_input_files(rest)
    if not files:
        print(T("没有找到任何光谱文件（支持 .jws / .csv / .spc / .jdx / .txt / .xlsx）"))
        return 1
    ok, skip, fail = 0, 0, 0
    for src in files:
        try:
            if src.lower().endswith(".jws"):
                jws_formats = tuple(f for f in formats
                                    if f in ("csv", "xlsx", "png", "peaks", "fit"))
                if not jws_formats:
                    jws_formats = ("csv",)
                dst, spec, skipped = convert_file(src, out_dir=out_dir_arg,
                                                  skip_existing=False,
                                                  formats=jws_formats, plot=plot or None)
                if skipped:
                    skip += 1
                    print("SKIP %s" % os.path.basename(dst))
                    continue
                if "jcamp" in formats:
                    base_dir = out_dir_arg or os.path.dirname(os.path.abspath(src))
                    jdst = os.path.join(base_dir,
                                        os.path.splitext(os.path.basename(src))[0] + ".jdx")
                    write_jcamp(jdst, _calibrate(spec.x_values(), plot),
                                list(spec.y_data[0]), {"SAMPLE": spec.sample or ""},
                                title=os.path.basename(src))
                    print("JCA %s  ->  %s" % (os.path.basename(src),
                                              os.path.basename(jdst)))
                print(T("OK  %s  ->  %s  (%d 通道, %d 点, %.3f ~ %.3f)") % (
                    os.path.basename(src), os.path.basename(dst),
                    spec.channel_number, spec.npoints, spec.start, spec.end))
                ok += 1
                continue
            gen_formats = tuple(f for f in formats if f in ("csv", "png", "jcamp", "peaks"))
            if not gen_formats:
                gen_formats = ("csv",)
            dst, series, skipped = export_generic(src, out_dir=out_dir_arg,
                                                  formats=gen_formats, plot=plot or None)
            if skipped:
                skip += 1
                print("SKIP %s" % os.path.basename(dst))
                continue
            print(T("OK  %s  ->  %s  (%d 条曲线, %d 点)") % (
                os.path.basename(src), os.path.basename(dst),
                len(series), len(series[0][1])))
            ok += 1
        except Exception as exc:
            print("ERR %s  ->  %s" % (os.path.basename(src), exc))
            fail += 1
    if want_waterfall:
        spectra = collect_spectra(files, plot, verbose=False)
        if len(spectra) >= 2:
            try:
                p = _plot_opts(plot)
                dst = os.path.join(out_dir_arg or results_dir(),
                                   "瀑布图_%d条.png" % len(spectra))
                render_waterfall(dst, spectra, "瀑布图（%d 条）" % len(spectra),
                                 _DEFAULT_X_HEADER, T("归一化强度"), plot,
                                 offset=p["stack_offset"])
                print(T("瀑布图：%s") % dst)
            except Exception as exc:
                print(T("ERR 瀑布图 -> %s") % exc)
                fail += 1
        else:
            print(T("瀑布图至少需要 2 条光谱，已跳过。"))
    if want_overlay:
        spectra = collect_spectra(files, plot, verbose=False)
        if len(spectra) >= 2:
            try:
                dst = os.path.join(out_dir_arg or results_dir(),
                                   "叠加图_%d条.png" % len(spectra))
                render_overlay(dst, spectra, "叠加图（%d 条）" % len(spectra),
                               _DEFAULT_X_HEADER, "归一化强度", plot)
                print(T("叠加图：%s") % dst)
            except Exception as exc:
                print(T("ERR 叠加图 -> %s") % exc)
                fail += 1
        else:
            print(T("叠加图至少需要 2 条光谱，已跳过。"))
    print("\n" + T("完成：成功 %d 个，跳过 %d 个，失败 %d 个") % (ok, skip, fail))
    return 0 if fail == 0 else 1


def _run_gui():
    import queue
    import threading
    import traceback
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog

    set_ui_lang(detect_ui_lang(), persist=False, apply_now=False)

    # 双语：所有对话框在布局完成后自动套用当前语言；弹窗与文本同步翻译。
    _orig_toplevel = tk.Toplevel

    class _LocToplevel(_orig_toplevel):
        def __init__(self, *args, **kwargs):
            _orig_toplevel.__init__(self, *args, **kwargs)
            _apply_icon(self)
            self.after_idle(self._localize_self)

        def _localize_self(self):
            try:
                localize_tree(self)
            except Exception:
                pass

    tk.Toplevel = _LocToplevel
    for _name in ("showinfo", "showwarning", "showerror", "askyesno",
                  "askokcancel", "askretrycancel"):
        _orig_mb = getattr(messagebox, _name)

        def _make(orig):
            def _wrapper(title=None, message=None, *args, **kwargs):
                try:
                    if message is not None:
                        message = T(message)
                    if title is not None:
                        title = T(title)
                except Exception:
                    pass
                return orig(title, message, *args, **kwargs)
            return _wrapper
        setattr(messagebox, _name, _make(_orig_mb))
    _orig_askstring = simpledialog.askstring

    def _askstring_wrapper(title=None, prompt=None, *args, **kwargs):
        try:
            return _orig_askstring(T(title), T(prompt), *args, **kwargs)
        except Exception:
            return _orig_askstring(title, prompt, *args, **kwargs)

    simpledialog.askstring = _askstring_wrapper

    moved = []
    try:
        data_root()
        moved = migrate_legacy_layout()
    except Exception:
        moved = []
    try:
        manual_file = ensure_manual_file()
    except Exception:
        manual_file = ""

    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("拉曼光谱工具 · 转换 / 分析 / 鉴定")
            _apply_icon(self)
            self.geometry("1220x780")
            self.minsize(1020, 680)
            self.files = []
            self.q = queue.Queue()
            self.worker = None
            self.cache = {}
            self._last_draw = None
            self._resize_job = None
            self._refresh_job = None
            self.adv_values = {
                "x_min": None, "x_max": None, "y_min": None, "y_max": None,
                "smooth_window": 1, "smooth_mode": "mean",
                "despike": False, "despike_window": 5, "despike_thresh": 8.0,
                "baseline": "none", "baseline_window": 60.0,
                "baseline_degree": 5, "baseline_iters": 20,
                "derivative": 0,
                "normalize": "none", "apply_to_data": False,
                "calib_a": None, "calib_b": None, "calib_a2": None, "calib_b2": None,
                "peak_thresh_pct": 7.0, "peak_label_rel": False,
                "peak_dash_line": True,
                "fit_shape": "voigt",
                "fig_width": 1600, "fig_height": 900,
                "mineral_name": None, "stack_offset": 1.0,
            }
            # 图例位置：全局默认存在设置文件里，下次打开还是上次摆的那个位置
            self.adv_values["legend_pos"], self.adv_values["legend_xy"] = legend_defaults()
            self.manual_peaks = {}
            self.hidden_peaks = {}
            self._view = None
            self._view_series = None
            self._build_style()
            self._build_ui()
            self.after(100, self._poll)
            self._center()

        def _build_style(self):
            style = ttk.Style(self)
            try:
                style.theme_use("vista")
            except tk.TclError:
                style.theme_use("clam")
            style.configure("Big.TButton", font=("Microsoft YaHei UI", 11, "bold"), padding=8)
            style.configure("TButton", font=("Microsoft YaHei UI", 9), padding=5)
            style.configure("TLabel", font=("Microsoft YaHei UI", 9))
            style.configure("TCheckbutton", font=("Microsoft YaHei UI", 9))
            style.configure("TLabelframe.Label", font=("Microsoft YaHei UI", 9, "bold"))

        def _center(self):
            self.update_idletasks()
            w, h = self.winfo_width(), self.winfo_height()
            x = (self.winfo_screenwidth() - w) // 2
            y = max(0, (self.winfo_screenheight() - h) // 3)
            self.geometry("+%d+%d" % (x, y))

        def _build_ui(self):
            self._build_menu()
            pane = ttk.PanedWindow(self, orient="horizontal")
            pane.pack(fill="both", expand=True)
            left = ttk.Frame(pane)
            right = ttk.Frame(pane)
            pane.add(left, weight=3)
            pane.add(right, weight=2)

            self._build_left(left)
            self._build_right(right)
            self._toggle_dir()
            self._toggle_header()
            self._log(T("① 可读：.jws / .csv / .spc / .jdx / .txt / .xlsx；可出 Excel(图表) / PNG / CSV / 峰列表 / 峰拟合 / JCAMP-DX。"))
            self._log(T("② 漏标的峰：在右侧图上左键点击即可手动补标（蓝色方块）；标错的峰（含自动峰）对准它右键即可删掉。"))
            self._log(T("③ 菜单“分析工具”：未知光谱检索（全库鉴定）/ 峰拟合 / 比对 / 峰位检索 /"))
            self._log(T("   相似度矩阵 / 平均 / 相减 / 谱段替换 / A−k·B / 瀑布图 / 聚类 / 二维成像。"))
            self._log(T("④ 菜单“文件”：文件夹批处理、导出分析报告；菜单“设置”：高级设置（含尖峰去除、"))
            self._log(T("   迭代多项式与滚动球基线）、矿物信息与峰归属库、数据文件夹与容量上限。"))
            self._log(T("⑤ 数据文件夹：%s") % data_root())
            self._log(T("   下载的参考谱/数据包与分析结果都只写到这里；数据包超上限或磁盘不足会提醒。"))
            if manual_file:
                self._log(T("⑥ 说明书：%s") % manual_file)
                self._log(T("   新手建议先读一遍：菜单【帮助】→【使用说明（完整手册）】。"))
            if moved:
                self._log(T("   已把旧的“光谱数据库”内容迁移进数据文件夹：%d 项。") % len(moved))

        def _build_menu(self):
            bar = tk.Menu(self)
            fm = tk.Menu(bar, tearoff=0)
            fm.add_command(label="添加文件…", command=self.add_files)
            fm.add_command(label="添加文件夹…", command=self.add_folder)
            fm.add_separator()
            fm.add_command(label="批处理文件夹…（预处理 + 峰表 + 出图 + 汇总）",
                           command=self.open_batch)
            fm.add_command(label="导出分析报告（选中文件）…", command=self.tool_report)
            fm.add_separator()
            fm.add_command(label="退出", command=self.destroy)
            bar.add_cascade(label="文件", menu=fm)

            tm = tk.Menu(bar, tearoff=0)
            tm.add_command(label="未知光谱检索（全库鉴定）…", command=self.open_identify)
            tm.add_command(label="批量鉴定（文件夹逐条鉴定）…",
                           command=self.tool_batch_identify)
            tm.add_separator()
            pm = tk.Menu(tm, tearoff=0)
            pm.add_command(label="手动配对（选参考谱文件 / 文件夹）…",
                           command=self.tool_pair_manual)
            pm.add_command(label="从数据库配对（在线检索选择）…", command=self.open_db_search)
            pm.add_separator()
            pm.add_command(label="自动配对（本地数据库）", command=self.tool_auto_pair)
            pm.add_command(label="自动配对（在线检索并下载）…",
                           command=self.tool_auto_pair_online)
            pm.add_separator()
            pm.add_command(label="批量配对（文件夹 × 参考谱）…",
                           command=self.tool_batch_pair)
            tm.add_cascade(label="配对比较（手动 / 自动）", menu=pm)
            tm.add_separator()
            tm.add_command(label="峰拟合（选中文件）", command=self.tool_fit)
            tm.add_command(label="峰位检索（参考峰表）…", command=self.tool_search)
            tm.add_command(label="相似度矩阵（所选文件）", command=self.tool_similarity)
            tm.add_separator()
            tm.add_command(label="聚类分析 + 主成分（所选文件）…", command=self.tool_cluster)
            tm.add_command(label="二维成像热图（所选点位）…", command=self.tool_map)
            tm.add_separator()
            tm.add_command(label="平均所选光谱", command=self.tool_average)
            tm.add_command(label="光谱相减（A−B）", command=self.tool_subtract)
            tm.add_command(label="交互式相减 A−k·B（拖动找平）…",
                           command=self.tool_subtract_interactive)
            tm.add_separator()
            tm.add_command(label="谱段替换 / 拼接…", command=self.tool_replace)
            tm.add_command(label="瀑布图（所选光谱）", command=self.tool_waterfall)
            tm.add_command(label="多数据图叠加（所选光谱）…", command=self.tool_overlay)
            tm.add_separator()
            tm.add_command(label="光谱比对（参考谱文件）…", command=self.tool_match)
            bar.add_cascade(label="分析工具", menu=tm)

            sm = tk.Menu(bar, tearoff=0)
            sm.add_command(label="高级设置…（处理 / 校准 / 坐标轴 / 峰 / 图幅）",
                           command=self.open_advanced)
            sm.add_command(label="矿物信息与拉曼峰归属库…", command=self.open_mineral_info)
            sm.add_separator()
            sm.add_command(label="数据文件夹…（位置 / 占用 / 容量上限 / 清理）",
                           command=self.open_data_manager)
            sm.add_command(label="清空手动峰标注", command=self.clear_manual_peaks)
            sm.add_command(label="恢复被删的自动峰", command=self.restore_hidden_peaks)
            bar.add_cascade(label="设置", menu=sm)

            dm = tk.Menu(bar, tearoff=0)
            dm.add_command(label="在线检索数据库（ROD / RRUFF 数据）…", command=self.open_db_search)
            dm.add_command(label="RRUFF 数据源：拉曼 / 红外 / XRD / 成分…", command=self.open_rruff)
            dm.add_separator()
            dm.add_command(label="用本地数据库比对（选中文件）", command=self.db_compare)
            dm.add_separator()
            dm.add_command(label="打开数据文件夹", command=lambda: self._open_folder(data_root()))
            dm.add_command(label="打开参考谱库", command=lambda: self._open_folder(ref_dir()))
            bar.add_cascade(label="数据库", menu=dm)

            hm = tk.Menu(bar, tearoff=0)
            hm.add_command(label="使用说明（完整手册）…", command=self.open_manual)
            hm.add_command(label="导出说明书为 TXT…", command=self._export_manual)
            hm.add_separator()
            hm.add_command(label="关于", command=self._about)
            bar.add_cascade(label="帮助", menu=hm)

            lm = tk.Menu(bar, tearoff=0)
            self._lang_var = tk.StringVar(value=_LANG_LABELS[ui_lang()])
            for _code in ("zh", "en"):
                lm.add_radiobutton(label=_LANG_LABELS[_code],
                                   value=_LANG_LABELS[_code],
                                   variable=self._lang_var,
                                   command=lambda c=_code: self._switch_lang(c))
            sm.add_cascade(label="语言 / Language", menu=lm)
            self.config(menu=bar)
            localize_menu(bar)
            localize_tree(self)

        def _switch_lang(self, code):
            """切换界面语言：已登记的控件/菜单当场重刷。"""
            set_ui_lang(code)
            self._lang_var.set(_LANG_LABELS[ui_lang()])
            self._log(T("界面语言已切换为 %s") % _LANG_LABELS[ui_lang()])

        def open_manual(self):
            win = tk.Toplevel(self)
            win.title("使用说明书 · 拉曼光谱工具")
            win.transient(self)
            win.geometry("1060x760")
            win.minsize(860, 580)

            top = ttk.Frame(win)
            top.pack(fill="x", padx=12, pady=(10, 4))
            ttk.Label(top, text="使用说明书",
                      font=("Microsoft YaHei UI", 13, "bold")).pack(side="left")
            ttk.Label(top, text="点左侧章节跳转", foreground="#777").pack(side="left", padx=8)
            ttk.Label(top, text="查找：").pack(side="left", padx=(16, 2))
            q = tk.StringVar(value="")
            ent = ttk.Entry(top, textvariable=q, width=16)
            ent.pack(side="left")
            ttk.Button(top, text="下一个", width=8,
                       command=lambda: find_next()).pack(side="left", padx=4)
            ttk.Label(top, text="字号：").pack(side="left", padx=(14, 2))
            size_var = tk.StringVar(value="11")
            size_box = ttk.Combobox(top, textvariable=size_var, state="readonly", width=4,
                                    values=("10", "11", "12", "13", "14", "16"))
            size_box.pack(side="left")
            lang_btn = ttk.Button(top, text="English", width=8,
                                  command=lambda: switch_manual_lang())
            lang_btn.pack(side="right")

            body = ttk.Frame(win)
            body.pack(fill="both", expand=True, padx=12)

            left = ttk.Frame(body)
            left.pack(side="left", fill="y")
            ttk.Label(left, text="章节目录", font=("Microsoft YaHei UI", 9, "bold")).pack(anchor="w")
            toc = tk.Listbox(left, width=24, activestyle="none", exportselection=False,
                             font=("Microsoft YaHei UI", 10), height=26)
            tsb = ttk.Scrollbar(left, orient="vertical", command=toc.yview)
            toc.configure(yscrollcommand=tsb.set)
            toc.pack(side="left", fill="both", expand=True, pady=(2, 0))
            tsb.pack(side="left", fill="y", pady=(2, 0))

            right = ttk.Frame(body)
            right.pack(side="left", fill="both", expand=True, padx=(12, 0))
            txt = tk.Text(right, wrap="word", font=("Microsoft YaHei UI", 11),
                          padx=16, pady=12, spacing1=2, spacing3=3,
                          background="#fdfdfd", highlightthickness=1,
                          highlightbackground="#dcdcdc")
            ysb = ttk.Scrollbar(right, orient="vertical", command=txt.yview)
            txt.configure(yscrollcommand=ysb.set)
            txt.pack(side="left", fill="both", expand=True)
            ysb.pack(side="left", fill="y")

            txt.tag_configure("hit", background="#ffe08a")
            txt.tag_configure("sec", font=("Microsoft YaHei UI", 12, "bold"),
                              foreground="#1a4a8a")
            sep = "=" * 60
            cur_lang = [ui_lang()]

            def load(code):
                """按语言重新装入说明书正文与目录（中英各一份）。"""
                sections = manual_sections(code)
                txt.configure(state="normal")
                txt.delete("1.0", "end")
                txt.insert("1.0", manual_text(code))
                for title, _b in sections:
                    pos = txt.search(sep + "\n" + title, "1.0", stopindex="end")
                    if pos:
                        start = "%s+%dc" % (pos, len(sep) + 1)
                        txt.tag_add("sec", start, "%s lineend" % start)
                txt.configure(state="disabled")
                toc.delete(0, "end")
                for title, _b in sections:
                    toc.insert("end", title)
                cur_lang[0] = code
                lang_btn.configure(text="中文" if code == "en" else "English")
                win.title(T("使用说明书 · 拉曼光谱工具"))

            def switch_manual_lang():
                load("en" if cur_lang[0] == "zh" else "zh")

            def goto(index_str):
                line = int(txt.index(index_str).split(".")[0])
                total = max(1, int(txt.index("end-1c").split(".")[0]))
                txt.yview_moveto(max(0.0, (line - 4) / float(total)))

            def jump(_e=None):
                sel = toc.curselection()
                if not sel:
                    return
                pos = txt.search(sep + "\n" + toc.get(sel[0]), "1.0", stopindex="end")
                if pos:
                    goto("%s+%dc" % (pos, len(sep) + 1))

            def find_next(_e=None):
                key = q.get().strip()
                if not key:
                    return
                pos = txt.search(key, txt.index("insert"), stopindex="end", nocase=True)
                if not pos:
                    pos = txt.search(key, "1.0", stopindex="end", nocase=True)
                txt.tag_remove("hit", "1.0", "end")
                if not pos:
                    self._log(T("说明书中没有找到“%s”。") % key)
                    return
                txt.tag_add("hit", pos, "%s+%dc" % (pos, len(key)))
                txt.mark_set("insert", "%s+%dc" % (pos, len(key)))
                txt.tag_raise("hit")
                goto(pos)

            def set_font(*_a):
                try:
                    size = int(size_var.get())
                except ValueError:
                    return
                txt.configure(font=("Microsoft YaHei UI", size))

            toc.bind("<<ListboxSelect>>", jump)
            ent.bind("<Return>", find_next)
            size_var.trace_add("write", set_font)

            bar = ttk.Frame(win)
            bar.pack(fill="x", padx=12, pady=(6, 12))
            ttk.Button(bar, text="导出为 TXT…", command=self._export_manual).pack(side="left")
            ttk.Button(bar, text="打开说明书所在文件夹",
                       command=lambda: self._open_manual_folder()).pack(side="left", padx=6)
            ttk.Label(bar, text="提示：Ctrl+A 全选后 Ctrl+C 可复制任意段落",
                      foreground="#888").pack(side="left", padx=10)
            load(cur_lang[0])
            ttk.Button(bar, text="关闭", command=win.destroy).pack(side="right")

        def _export_manual(self):
            path = filedialog.asksaveasfilename(
                title=T("导出说明书"), defaultextension=".txt",
                initialfile="拉曼光谱工具_使用说明.txt",
                filetypes=[(T("文本文件"), "*.txt"), (T("所有文件"), "*.*")])
            if not path:
                return
            try:
                with open(path, "w", encoding="utf-8-sig", newline="") as f:
                    f.write(manual_text())
            except OSError as exc:
                messagebox.showerror("导出失败", str(exc))
                return
            self._log(T("说明书已导出：%s") % path)
            messagebox.showinfo("导出完成", "说明书已保存到：\n%s" % path)

        def _open_manual_folder(self):
            path = ensure_manual_file()
            folder = os.path.dirname(path) if path else _app_dir()
            self._open_folder(folder)

        def _about(self):
            messagebox.showinfo(
                "关于",
                T("拉曼光谱工具 · Raman Spectrum Toolkit") + "\n"
                + T("（JASCO .jws 光谱转换 · 拉曼峰分析 · 矿物鉴定）") + "\n\n"
                + T("读入：.jws / .csv / .spc / .jdx / .txt / .xlsx") + "\n"
                + T("输出：Excel（含图表）/ PNG / CSV / 峰列表 / 峰拟合 / JCAMP-DX") + "\n"
                + T("预处理：尖峰去除 · 基线校正 · 平滑 · 导数 · 归一化 · 位移校准") + "\n"
                + T("分析：未知谱检索 · 峰拟合 · 峰位检索 · 相似度矩阵 · 谱运算 · 瀑布图 · 聚类 · 二维成像 · 报告") + "\n"
                + T("数据库：ROD 在线检索 · RRUFF 拉曼 / 红外 / XRD / 化学成分数据包") + "\n"
                + T("矿物：内置特征峰归属库 + RRUFF 真实样品记录") + "\n\n"
                + T("能力参考 RRUFF 项目所列工具（RamanCrystalHunter、RamanLab 等）") + "\n"
                + T("版本：%s") % _MANUAL_VERSION)

        def _build_left(self, root):
            top = ttk.Frame(root)
            top.pack(fill="x", padx=10, pady=(10, 4))
            ttk.Label(top, text="① 选择文件：.jws 转数据/图表，.csv 转图片（可多选或整个文件夹）",
                      font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w")

            mid = ttk.Frame(root)
            mid.pack(fill="both", expand=True, padx=10)
            left = ttk.Frame(mid)
            left.pack(side="left", fill="both", expand=True)
            self.listbox = tk.Listbox(left, selectmode="extended", activestyle="none",
                                      font=("Microsoft YaHei UI", 9), height=8,
                                      exportselection=False)
            sb = ttk.Scrollbar(left, orient="vertical", command=self.listbox.yview)
            self.listbox.configure(yscrollcommand=sb.set)
            self.listbox.bind("<<ListboxSelect>>", self._on_select)
            self.listbox.pack(side="left", fill="both", expand=True)
            sb.pack(side="left", fill="y")

            btns = ttk.Frame(mid)
            btns.pack(side="left", fill="y", padx=(8, 0))
            ttk.Button(btns, text="添加文件", command=self.add_files).pack(fill="x", pady=2)
            ttk.Button(btns, text="添加文件夹", command=self.add_folder).pack(fill="x", pady=2)
            ttk.Button(btns, text="移除选中", command=self.remove_selected).pack(fill="x", pady=2)
            ttk.Button(btns, text="清空列表", command=self.clear_files).pack(fill="x", pady=2)
            ttk.Separator(btns, orient="horizontal").pack(fill="x", pady=6)
            ttk.Button(btns, text="使用说明", command=self.open_manual).pack(fill="x", pady=2)

            out = ttk.LabelFrame(root, text="② 输出位置")
            out.pack(fill="x", padx=10, pady=6)
            self.same_dir = tk.BooleanVar(value=True)
            self.skip_existing = tk.BooleanVar(value=False)
            ttk.Checkbutton(out, text="保存到源文件所在目录（推荐）", variable=self.same_dir,
                            command=self._toggle_dir).grid(row=0, column=0, columnspan=2,
                                                           sticky="w", padx=8, pady=(6, 2))
            ttk.Checkbutton(out, text="跳过已存在的 CSV（不覆盖）", variable=self.skip_existing
                            ).grid(row=0, column=2, sticky="w", padx=8, pady=(6, 2))
            ttk.Label(out, text="输出文件夹：").grid(row=1, column=0, sticky="w", padx=8)
            self.dir_var = tk.StringVar()
            self.dir_entry = ttk.Entry(out, textvariable=self.dir_var)
            self.dir_entry.grid(row=1, column=1, sticky="we", padx=4, pady=(0, 6))
            self.dir_btn = ttk.Button(out, text="浏览…", command=self.choose_dir)
            self.dir_btn.grid(row=1, column=2, padx=8, pady=(0, 6))
            fmtrow = ttk.Frame(out)
            fmtrow.grid(row=2, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 6))
            ttk.Label(fmtrow, text="输出内容：").pack(side="left")
            self.fmt_excel = tk.BooleanVar(value=True)
            self.fmt_png = tk.BooleanVar(value=True)
            self.fmt_csv = tk.BooleanVar(value=False)
            ttk.Checkbutton(fmtrow, text="Excel（含图表）", variable=self.fmt_excel).pack(side="left")
            ttk.Checkbutton(fmtrow, text="PNG 图片", variable=self.fmt_png).pack(side="left", padx=8)
            ttk.Checkbutton(fmtrow, text="CSV（纯数据）", variable=self.fmt_csv).pack(side="left")
            self.fmt_peaks = tk.BooleanVar(value=False)
            ttk.Checkbutton(fmtrow, text="峰列表", variable=self.fmt_peaks).pack(side="left", padx=8)
            self.fmt_fit = tk.BooleanVar(value=False)
            ttk.Checkbutton(fmtrow, text="峰拟合", variable=self.fmt_fit).pack(side="left")
            ttk.Label(fmtrow, text="PNG 文件名与源文件一致", foreground="#777").pack(
                side="left", padx=8)
            out.columnconfigure(1, weight=1)

            head = ttk.LabelFrame(root, text="③ 表头与列名")
            head.pack(fill="x", padx=10, pady=2)
            self.write_header = tk.BooleanVar(value=True)
            self.auto_names = tk.BooleanVar(value=False)
            ttk.Checkbutton(head, text="写入表头行", variable=self.write_header,
                            command=self._toggle_header).grid(row=0, column=0, sticky="w",
                                                              padx=8, pady=(6, 2))
            ttk.Checkbutton(head, text="自动识别光谱列名（波长 / 吸光度 / CD 等）",
                            variable=self.auto_names, command=self._toggle_header
                            ).grid(row=0, column=1, sticky="w", padx=8, pady=(6, 2))
            ttk.Label(head, text="横坐标列名：").grid(row=1, column=0, sticky="w", padx=8)
            self.x_var = tk.StringVar(value=_DEFAULT_X_HEADER)
            self.x_entry = ttk.Entry(head, textvariable=self.x_var, width=22)
            self.x_entry.grid(row=1, column=1, sticky="w", padx=4, pady=2)
            ttk.Label(head, text="纵坐标列名：").grid(row=2, column=0, sticky="w", padx=8, pady=(0, 6))
            self.y_var = tk.StringVar(value=_DEFAULT_Y_HEADER)
            self.y_entry = ttk.Entry(head, textvariable=self.y_var, width=22)
            self.y_entry.grid(row=2, column=1, sticky="w", padx=4, pady=(0, 6))

            run = ttk.Frame(root)
            run.pack(fill="x", padx=10, pady=6)
            self.run_btn = ttk.Button(run, text="开始转换", style="Big.TButton",
                                      command=self.start_convert)
            self.run_btn.pack(side="left")
            self.progress = ttk.Progressbar(run, mode="determinate")
            self.progress.pack(side="left", fill="x", expand=True, padx=10)
            ttk.Button(run, text="打开输出目录", command=self.open_out_dir).pack(side="left")

            logf = ttk.LabelFrame(root, text="④ 运行日志")
            logf.pack(fill="both", expand=True, padx=10, pady=(2, 10))
            self.log = tk.Text(logf, height=7, font=("Consolas", 9), wrap="word",
                               state="disabled", background="#fbfbfb")
            lsb = ttk.Scrollbar(logf, orient="vertical", command=self.log.yview)
            self.log.configure(yscrollcommand=lsb.set)
            self.log.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
            lsb.pack(side="left", fill="y", pady=6, padx=(0, 6))

        def _build_right(self, root):
            box = ttk.LabelFrame(root, text="光谱预览")
            box.pack(fill="both", expand=True, padx=10, pady=10)
            bar = ttk.Frame(box)
            bar.pack(fill="x", padx=6, pady=(6, 2))
            ttk.Button(bar, text="预览选中文件", command=self.preview_selected).pack(side="left")
            ttk.Button(bar, text="打开 CSV 看图", command=self.open_csv_preview).pack(side="left",
                                                                                    padx=6)
            ttk.Button(bar, text="CSV→图片", command=self.csv_to_png).pack(side="left")
            ttk.Button(bar, text="保存为 PNG", command=self.save_preview_png).pack(side="left", padx=6)
            ttk.Button(bar, text="清空视图", command=self.clear_preview).pack(side="left")
            self.info_var = tk.StringVar(value=T("尚未选择文件"))
            ttk.Label(box, textvariable=self.info_var, foreground="#555",
                      font=("Microsoft YaHei UI", 9)).pack(anchor="w", padx=10)

            setf = ttk.LabelFrame(box, text="图表设置")
            setf.pack(fill="x", padx=8, pady=(2, 0))
            ttk.Label(setf, text="横坐标刻度间隔：").grid(row=0, column=0, sticky="w", padx=(6, 0),
                                                        pady=(6, 2))
            self.xstep_var = tk.StringVar(value="200")
            ttk.Entry(setf, textvariable=self.xstep_var, width=6).grid(row=0, column=1, sticky="w",
                                                                      padx=2, pady=(6, 2))
            ttk.Label(setf, text="起始刻度：").grid(row=0, column=2, sticky="w", padx=(8, 0),
                                                  pady=(6, 2))
            self.xstart_var = tk.StringVar(value="")
            ttk.Entry(setf, textvariable=self.xstart_var, width=6).grid(row=0, column=3, sticky="w",
                                                                       padx=2, pady=(6, 2))
            ttk.Label(setf, text="(空=自动)", foreground="#888").grid(row=0, column=4, sticky="w",
                                                                    padx=4, pady=(6, 2))
            self.show_y_ticks = tk.BooleanVar(value=False)
            self.show_grid = tk.BooleanVar(value=True)
            self.show_title = tk.BooleanVar(value=True)
            self.annotate_peaks = tk.BooleanVar(value=True)
            self.peak_labels = tk.BooleanVar(value=True)
            self.peakdist_var = tk.StringVar(value="20")
            ttk.Checkbutton(setf, text="显示纵坐标数值", variable=self.show_y_ticks,
                            command=self._refresh_preview).grid(row=1, column=0, columnspan=2,
                                                                sticky="w", padx=6)
            ttk.Checkbutton(setf, text="网格线", variable=self.show_grid,
                            command=self._refresh_preview).grid(row=1, column=2, sticky="w")
            ttk.Checkbutton(setf, text="标题", variable=self.show_title,
                            command=self._refresh_preview).grid(row=1, column=3, columnspan=2,
                                                                sticky="w")
            ttk.Checkbutton(setf, text="标注峰位", variable=self.annotate_peaks,
                            command=self._refresh_preview).grid(row=2, column=0, columnspan=2,
                                                                sticky="w", padx=6, pady=(0, 6))
            ttk.Label(setf, text="最小峰间距(cm-1)：").grid(row=2, column=2, columnspan=2,
                                                        sticky="w", pady=(0, 6))
            ttk.Entry(setf, textvariable=self.peakdist_var, width=6).grid(row=2, column=4,
                                                                        sticky="w", padx=4,
                                                                        pady=(0, 6))
            ttk.Button(setf, text="高级设置…", command=self.open_advanced).grid(
                row=3, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 4))
            ttk.Checkbutton(setf, text="显示峰位数值", variable=self.peak_labels,
                            command=self._refresh_preview).grid(
                row=3, column=2, columnspan=3, sticky="w", padx=6, pady=(0, 4))
            ttk.Button(setf, text="撤销手动峰", command=lambda: self.remove_manual_peak(None)
                       ).grid(row=4, column=0, sticky="w", padx=6, pady=(0, 6))
            ttk.Button(setf, text="清空手动峰", command=self.clear_manual_peaks).grid(
                row=4, column=1, sticky="w", padx=4, pady=(0, 6))
            ttk.Button(setf, text="恢复自动峰", command=self.restore_hidden_peaks).grid(
                row=4, column=2, sticky="w", padx=4, pady=(0, 6))
            ttk.Label(setf, text="左键点图=补标峰，右键=删掉最近的峰（自动 / 手动都可）",
                      foreground="#888").grid(row=5, column=0, columnspan=5, sticky="w",
                                              padx=6, pady=(0, 6))
            for var in (self.xstep_var, self.xstart_var, self.peakdist_var):
                var.trace_add("write", lambda *a: self._refresh_preview())

            self.canvas = tk.Canvas(box, background="white", highlightthickness=1,
                                    highlightbackground="#c9c9c9", cursor="crosshair")
            self.canvas.pack(fill="both", expand=True, padx=8, pady=(4, 8))
            self.canvas.bind("<Configure>", self._on_canvas_resize)
            self.canvas.bind("<Button-1>", self.add_manual_peak)
            self.canvas.bind("<Button-3>", self.remove_annotated_peak)
            self._draw(None, "", "")

        def _toggle_dir(self):
            state = "disabled" if self.same_dir.get() else "normal"
            self.dir_entry.configure(state=state)
            self.dir_btn.configure(state=state)

        def _toggle_header(self):
            editable = self.write_header.get() and not self.auto_names.get()
            state = "normal" if editable else "disabled"
            self.x_entry.configure(state=state)
            self.y_entry.configure(state=state)

        def plot_options(self):
            def num(text, default):
                try:
                    return float(str(text).strip())
                except (TypeError, ValueError):
                    return default

            st = self.xstart_var.get().strip()
            start = None
            if st:
                start = num(st, None)
            opts = {
                "x_step": num(self.xstep_var.get(), 200.0),
                "x_start": start,
                "show_y_ticks": self.show_y_ticks.get(),
                "show_grid": self.show_grid.get(),
                "show_title": self.show_title.get(),
                "annotate_peaks": self.annotate_peaks.get(),
                "peak_labels": self.peak_labels.get(),
                "peak_min_dist": num(self.peakdist_var.get(), 20.0),
            }
            opts.update(self.adv_values)
            pairs = []
            if opts.get("calib_a") is not None and opts.get("calib_b") is not None:
                pairs.append((opts["calib_a"], opts["calib_b"]))
                if opts.get("calib_a2") is not None and opts.get("calib_b2") is not None:
                    pairs.append((opts["calib_a2"], opts["calib_b2"]))
            opts["calib_pairs"] = pairs or None
            return opts

        def current_plot_options(self):
            opts = self.plot_options()
            sel = self.listbox.curselection()
            if len(sel) == 1:
                opts.update(self._peak_overrides(self.files[sel[0]]))
            return opts

        def _peak_key(self, path):
            """手动峰 / 被删自动峰按文件记录，键统一成绝对路径，避免同一个文件对不上。"""
            return os.path.abspath(path)

        def _peak_overrides(self, path):
            """某个文件上“手动补的峰 + 被删掉的自动峰”，打包成 plot 覆盖项。"""
            out = {}
            manual = self.manual_peaks.get(self._peak_key(path))
            if manual:
                out["manual_peaks"] = list(manual)
            hidden = self.hidden_peaks.get(self._peak_key(path))
            if hidden:
                out["hidden_peaks"] = list(hidden)
            return out

        def _manual_list(self):
            sel = self.listbox.curselection()
            if len(sel) != 1:
                return None, None
            path = self._peak_key(self.files[sel[0]])
            return path, self.manual_peaks.setdefault(path, [])

        def _hidden_list(self):
            sel = self.listbox.curselection()
            if len(sel) != 1:
                return None, None
            path = self._peak_key(self.files[sel[0]])
            return path, self.hidden_peaks.setdefault(path, [])

        def add_manual_peak(self, event=None):
            path, lst = self._manual_list()
            if lst is None:
                self.info_var.set(T("手动标注请只选择 1 个文件"))
                return
            if not self._view or not self._view_series:
                return
            v = self._view
            if not (v["ml"] <= event.x <= v["ml"] + v["pw"]) or not (
                    v["mt"] <= event.y <= v["mt"] + v["ph"]):
                return
            xs = self._view_series[0][1]
            ys = self._view_series[0][2]
            if not xs:
                return
            data_x = v["xmin"] + (event.x - v["ml"]) / v["pw"] * (v["xmax"] - v["xmin"])
            idx = _nearest_index(xs, data_x)
            if len(xs) > 1:
                step = abs(xs[1] - xs[0]) or 1.0
                half = abs(float(self.plot_options()["peak_min_dist"])) / 2.0
                w = max(1, int(round(half / step)))
                lo = max(0, idx - w)
                hi = min(len(ys), idx + w + 1)
                if hi > lo:
                    idx = max(range(lo, hi), key=lambda i: ys[i])
            px = xs[idx]
            cur = self.plot_options()
            dup = max(1.0, abs(float(cur["peak_min_dist"])) / 2.0)
            auto = find_peaks(xs, ys, cur["peak_min_dist"], thresh_pct=cur["peak_thresh_pct"])
            if any(abs(px - a["x"]) <= dup for a in auto):
                self._log(T("该位置已有自动峰（%.1f），无需重复标注。") % px)
                return
            if all(abs(px - q) > 1e-9 for q in lst):
                lst.append(px)
                lst.sort()
            self.preview_selected()
            self._log(T("手动标注峰：%.2f（共 %d 个）") % (px, len(lst)))

        def remove_manual_peak(self, event=None):
            """撤销最后补的一个手动峰（工具栏按钮用）。"""
            path, lst = self._manual_list()
            if lst is None or not lst:
                return
            lst.pop()
            self.preview_selected()
            self._log(T("已移除手动峰，剩余 %d 个。") % len(lst))

        def remove_annotated_peak(self, event=None):
            """右键图上的峰：手动峰直接删掉，自动峰记进“已删除”名单不再画。

            要求点在峰附近（横向 ± 图宽的 1/40），否则不动 —— 自动峰往往很多，
            不加这道距离判断，在空白处随手一右键就会误删一个峰。
            """
            path, mlst = self._manual_list()
            hpath, hlst = self._hidden_list()
            if mlst is None or hlst is None or path != hpath:
                self.info_var.set(T("删除峰标注请只选择 1 个文件"))
                return
            if not self._view or not self._view_series or event is None:
                return
            v = self._view
            xs = self._view_series[0][1]
            ys = self._view_series[0][2]
            if not xs:
                return
            plot = self.current_plot_options()
            peaks = analyze_peaks(xs, ys, plot, processed=True)
            if not peaks:
                self.info_var.set(T("图上没有可删除的标注峰"))
                return
            data_x = v["xmin"] + (event.x - v["ml"]) / v["pw"] * (v["xmax"] - v["xmin"])
            span = v["xmax"] - v["xmin"]
            near = min(peaks, key=lambda pk: abs(pk["x"] - data_x))
            if abs(near["x"] - data_x) > max(1e-9, span / 40.0):
                self.info_var.set(T("附近没有峰，请点在峰上再右键"))
                return
            px = near["x"]
            tol = max(1.0, abs(float(plot["peak_min_dist"])) / 2.0)
            if near.get("manual"):
                for q in list(mlst):
                    if abs(q - px) <= tol:
                        mlst.remove(q)
                self._log(T("已删除手动峰：%.2f（剩余 %d 个）") % (px, len(mlst)))
            else:
                if all(abs(px - q) > tol for q in hlst):
                    hlst.append(px)
                    hlst.sort()
                self._log(T("已删除自动峰：%.2f（该文件共隐藏 %d 个）") % (px, len(hlst)))
            self.preview_selected()

        def restore_hidden_peaks(self):
            """把当前文件里被右键删掉的自动峰全部恢复回来。"""
            path, hlst = self._hidden_list()
            if hlst is None:
                return
            if not hlst:
                self.info_var.set(T("当前文件没有被删除的自动峰"))
                return
            n = len(hlst)
            hlst.clear()
            self.preview_selected()
            self._log(T("已恢复 %d 个被删除的自动峰。") % n)

        def clear_manual_peaks(self):
            path, lst = self._manual_list()
            if lst is None:
                return
            lst.clear()
            self.preview_selected()
            self._log(T("已清空该文件的手动标注。"))

        def open_advanced(self):
            win = tk.Toplevel(self)
            win.title("高级设置 · 处理 / 校准 / 坐标轴 / 峰 / 图幅")
            win.transient(self)
            adv = self.adv_values
            vars_ = {}

            # 底部按钮条先占位：小屏 / 高 DPI 下内容再长，按钮也不会被顶出屏幕。
            btns = ttk.Frame(win)
            btns.pack(side="bottom", fill="x", padx=10, pady=(6, 10))

            # 其余内容放进可滚动区，一屏放不下就滚，而不是把窗口撑到屏幕外。
            body = ttk.Frame(win)
            body.pack(side="top", fill="both", expand=True)
            canvas = tk.Canvas(body, borderwidth=0, highlightthickness=0)
            vbar = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=vbar.set)
            vbar.pack(side="right", fill="y")
            canvas.pack(side="left", fill="both", expand=True)
            inner = ttk.Frame(canvas)
            inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

            def _sync_scrollregion(_event=None):
                canvas.configure(scrollregion=canvas.bbox("all"))

            def _fit_width(event):
                canvas.itemconfigure(inner_id, width=event.width)

            def _on_wheel(event):
                canvas.yview_scroll(-1 if getattr(event, "delta", 0) > 0 else 1, "units")

            inner.bind("<Configure>", _sync_scrollregion)
            canvas.bind("<Configure>", _fit_width)
            # 滚轮：子控件不处理本事件时会向上冒泡到 inner / canvas，各挂一次即可
            inner.bind("<MouseWheel>", _on_wheel)
            canvas.bind("<MouseWheel>", _on_wheel)

            def add_entry(parent, label, key, row, hint="", width=9):
                ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=3)
                value = adv.get(key)
                var = tk.StringVar(value="" if value is None else str(value))
                ttk.Entry(parent, textvariable=var, width=width).grid(
                    row=row, column=1, sticky="w", padx=4, pady=3)
                if hint:
                    ttk.Label(parent, text=hint, foreground="#888").grid(
                        row=row, column=2, sticky="w", padx=6)
                vars_[key] = var

            def add_combo(parent, label, key, row, labels, default):
                ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=8, pady=3)
                # 显示值随语言变，但内部键永远是 labels 的 key（用译文反查）
                shown = {k: T(v) for k, v in labels.items()}
                var = tk.StringVar(value=shown.get(adv.get(key, default),
                                                   list(shown.values())[0]))
                ttk.Combobox(parent, textvariable=var, state="readonly",
                             values=list(shown.values()), width=14).grid(
                    row=row, column=1, sticky="w", padx=4, pady=3)
                return var, {v: k for k, v in shown.items()}

            axes = ttk.LabelFrame(inner, text="坐标轴范围（留空 = 自动）")
            axes.grid(row=0, column=0, sticky="we", padx=10, pady=(10, 4))
            add_entry(axes, "横坐标最小值：", "x_min", 0)
            add_entry(axes, "横坐标最大值：", "x_max", 1)
            add_entry(axes, "纵坐标最小值：", "y_min", 2)
            add_entry(axes, "纵坐标最大值：", "y_max", 3)

            proc = ttk.LabelFrame(inner, text="光谱处理（默认仅用于出图与峰识别）")
            proc.grid(row=1, column=0, sticky="we", padx=10, pady=4)
            dsp = tk.BooleanVar(value=bool(adv.get("despike")))
            ttk.Checkbutton(proc, text="尖峰去除（宇宙射线 / 坏点）", variable=dsp).grid(
                row=0, column=0, columnspan=2, sticky="w", padx=8, pady=3)
            add_entry(proc, "尖峰窗口(点)：", "despike_window", 1, hint="3~7 即可")
            add_entry(proc, "尖峰阈值(噪声倍数)：", "despike_thresh", 2,
                      hint="越小越激进，建议 8~15")
            base_labels = {"none": "无", "linear": "线性（端点连线）",
                           "poly2": "多项式（2阶）", "iterpoly": "迭代多项式（推荐）",
                           "rolling": "滚动最小值", "rolling_ball": "滚动球"}
            bvar, binv = add_combo(proc, "基线校正：", "baseline", 3, base_labels, "none")
            add_entry(proc, "基线窗口(cm-1)：", "baseline_window", 4,
                      hint="滚动最小值 / 滚动球用")
            add_entry(proc, "迭代阶数：", "baseline_degree", 5, hint="迭代多项式用")
            add_entry(proc, "迭代次数：", "baseline_iters", 6, hint="迭代多项式用")
            sm_labels = {"mean": "移动平均", "sg": "Savitzky-Golay"}
            svar, sinv = add_combo(proc, "平滑方式：", "smooth_mode", 7, sm_labels, "mean")
            add_entry(proc, "平滑窗口(点)：", "smooth_window", 8, hint="1=不平滑")
            dv_labels = {0: "无", 1: "一阶导数", 2: "二阶导数"}
            dvar, dinv = add_combo(proc, "导数：", "derivative", 9, dv_labels, 0)
            n_labels = {"none": "无", "max": "最大值=1", "minmax": "最小-最大"}
            nvar, ninv = add_combo(proc, "归一化：", "normalize", 10, n_labels, "none")
            add_entry(proc, "峰位归属矿物：", "mineral_name", 11,
                      hint="如 Zircon / 锆石，留空 = 不归属", width=18)
            add_entry(proc, "堆叠偏移（瀑布图 / 叠加图）：", "stack_offset", 12,
                      hint="0.2 ~ 2.0，1.0 = 谱线刚好不压线")
            avar = tk.BooleanVar(value=bool(adv.get("apply_to_data")))
            ttk.Checkbutton(proc, text="同时应用到导出的数据（CSV/Excel）", variable=avar).grid(
                row=13, column=0, columnspan=3, sticky="w", padx=8, pady=(0, 4))

            cal = ttk.LabelFrame(inner, text="拉曼位移校准（实测峰位 → 标准峰位，留空 = 不校准）")
            cal.grid(row=2, column=0, sticky="we", padx=10, pady=4)
            add_entry(cal, "第1对 实测/标准：", "calib_a", 0, hint="例 520.6 → 520.7")
            add_entry(cal, "对应标准值：", "calib_b", 1)
            add_entry(cal, "第2对 实测：", "calib_a2", 2, hint="两点校准才需填")
            add_entry(cal, "对应标准值：", "calib_b2", 3)

            peak = ttk.LabelFrame(inner, text="峰识别与峰拟合")
            peak.grid(row=3, column=0, sticky="we", padx=10, pady=4)
            add_entry(peak, "峰灵敏阈值(%)：", "peak_thresh_pct", 0, hint="占强度范围，越小越灵敏")
            sh_labels = {"gaussian": "高斯", "lorentzian": "洛伦兹", "voigt": "伪Voigt"}
            shvar, shinv = add_combo(peak, "拟合峰形：", "fit_shape", 1, sh_labels, "voigt")
            rvar = tk.BooleanVar(value=bool(adv.get("peak_label_rel")))
            ttk.Checkbutton(peak, text="峰标签同时显示相对强度(%)", variable=rvar).grid(
                row=2, column=0, columnspan=3, sticky="w", padx=8, pady=3)
            dashvar = tk.BooleanVar(value=bool(adv.get("peak_dash_line", True)))
            ttk.Checkbutton(peak, text="峰位虚线引到横坐标轴（自动峰 + 手动峰）",
                            variable=dashvar).grid(
                row=3, column=0, columnspan=3, sticky="w", padx=8, pady=3)

            fig = ttk.LabelFrame(inner, text="图幅（PNG 输出像素）")
            fig.grid(row=4, column=0, sticky="we", padx=10, pady=4)
            add_entry(fig, "宽：", "fig_width", 0)
            add_entry(fig, "高：", "fig_height", 1)
            legend_labels = {
                "right": "绘图区右侧留白（不压谱线，推荐）",
                "top": "绘图区上方",
                "bottom": "绘图区下方",
                "inside": "图内自由位置（可拖动）",
                "none": "不显示图例",
            }
            lgvar, lginv = add_combo(fig, "图例位置：", "legend_pos", 2,
                                     legend_labels, "right")
            ttk.Label(fig, text="几条谱线就列几条图例；选“图内自由位置”后可以在"
                                "叠加图预览窗口里拖着放，位置会记住",
                      foreground="#888").grid(row=3, column=0, columnspan=3,
                                              sticky="w", padx=8, pady=(0, 3))

            def apply(close=False):
                def to_float(var, default=None):
                    text = str(var.get()).strip()
                    if not text:
                        return default
                    try:
                        return float(text)
                    except ValueError:
                        return default

                def to_int(var, default, minimum):
                    value = to_float(var, None)
                    if value is None:
                        return default
                    return max(minimum, int(round(value)))

                adv["x_min"] = to_float(vars_["x_min"])
                adv["x_max"] = to_float(vars_["x_max"])
                adv["y_min"] = to_float(vars_["y_min"])
                adv["y_max"] = to_float(vars_["y_max"])
                adv["baseline"] = binv.get(bvar.get(), "none")
                adv["baseline_window"] = to_float(vars_["baseline_window"], 60.0)
                adv["baseline_degree"] = to_int(vars_["baseline_degree"], 5, 1)
                adv["baseline_iters"] = to_int(vars_["baseline_iters"], 20, 1)
                adv["despike"] = bool(dsp.get())
                adv["despike_window"] = to_int(vars_["despike_window"], 5, 3)
                adv["despike_thresh"] = to_float(vars_["despike_thresh"], 6.0)
                adv["stack_offset"] = to_float(vars_["stack_offset"], 0.75)
                name = str(vars_["mineral_name"].get()).strip()
                adv["mineral_name"] = name or None
                adv["smooth_mode"] = sinv.get(svar.get(), "mean")
                adv["smooth_window"] = to_int(vars_["smooth_window"], 1, 1)
                adv["derivative"] = dinv.get(dvar.get(), 0)
                adv["normalize"] = ninv.get(nvar.get(), "none")
                adv["apply_to_data"] = bool(avar.get())
                adv["calib_a"] = to_float(vars_["calib_a"])
                adv["calib_b"] = to_float(vars_["calib_b"])
                adv["calib_a2"] = to_float(vars_["calib_a2"])
                adv["calib_b2"] = to_float(vars_["calib_b2"])
                adv["peak_thresh_pct"] = to_float(vars_["peak_thresh_pct"], 7.0)
                adv["fit_shape"] = shinv.get(shvar.get(), "voigt")
                adv["peak_label_rel"] = bool(rvar.get())
                adv["peak_dash_line"] = bool(dashvar.get())
                adv["fig_width"] = to_int(vars_["fig_width"], 1600, 200)
                adv["fig_height"] = to_int(vars_["fig_height"], 900, 200)
                adv["legend_pos"] = lginv.get(lgvar.get(), "right")
                save_legend_defaults(adv["legend_pos"], adv.get("legend_xy"))
                if close:
                    win.destroy()
                self.preview_selected()

            ttk.Button(btns, text="取消", command=win.destroy).pack(side="right")
            ttk.Button(btns, text="应用", command=lambda: apply(False)).pack(side="right", padx=6)
            ttk.Button(btns, text="确定", command=lambda: apply(True)).pack(side="right")

            # 尺寸：内容有多高就给多高，但绝不超出屏幕可用高度；
            # 超出部分靠滚动，按钮条固定在底部始终可点。
            win.update_idletasks()
            need_w = max(inner.winfo_reqwidth(), 560)
            need_h = inner.winfo_reqheight() + btns.winfo_reqheight() + 30
            sw = win.winfo_screenwidth()
            sh = win.winfo_screenheight()
            w = min(need_w + 30, int(sw * 0.92))
            h = min(need_h, int(sh * 0.85))
            _sync_scrollregion()
            win.geometry("%dx%d+%d+%d" % (w, h, max(0, (sw - w) // 2),
                                          max(0, (sh - h) // 2 - 20)))
            win.minsize(min(w, 460), min(h, 300))
            win.resizable(True, True)

        def _refresh_preview(self):
            if self._refresh_job:
                self.after_cancel(self._refresh_job)
            self._refresh_job = self.after(250, self.preview_selected)

        def choose_dir(self):
            d = filedialog.askdirectory(title=T("选择输出文件夹"))
            if d:
                self.dir_var.set(d)

        def _add_paths(self, paths):
            added = 0
            for p in find_input_files(paths):
                if p not in self.files:
                    self.files.append(p)
                    self.listbox.insert("end", p)
                    added += 1
            return added

        def add_files(self):
            paths = filedialog.askopenfilenames(
                title=T("选择光谱文件（.jws / .csv / .spc / .jdx / .txt / .xlsx）"),
                filetypes=[("光谱文件",
                            "*.jws *.csv *.spc *.jdx *.dx *.txt *.dat *.asc *.xy *.xlsx"),
                           (T("JASCO 光谱(*.jws)"), "*.jws"),
                           ("CSV(*.csv)", "*.csv"),
                           (T("SPC 光谱(*.spc)"), "*.spc"),
                           ("JCAMP-DX(*.jdx)", "*.jdx"),
                           (T("文本(*.txt)"), "*.txt"),
                           ("Excel(*.xlsx)", "*.xlsx"),
                           (T("所有文件"), "*.*")])
            if paths:
                added = self._add_paths(paths)
                self._log(T("已添加 %d 个文件。") % added)
                if added:
                    self._select_first_new()

        def add_folder(self):
            d = filedialog.askdirectory(
                title=T("选择文件夹（会递归查找 .jws / .csv / .spc / .jdx）"))
            if d:
                added = self._add_paths([d])
                self._log(T("从文件夹添加了 %d 个光谱文件。") % added)

        def _select_first_new(self):
            if self.files:
                self.listbox.selection_clear(0, "end")
                self.listbox.selection_set(0)
                self.listbox.activate(0)
                self.preview_selected()

        def remove_selected(self):
            for idx in reversed(self.listbox.curselection()):
                self.listbox.delete(idx)
                del self.files[idx]
            self.clear_preview()

        def clear_files(self):
            self.listbox.delete(0, "end")
            self.files.clear()
            self.clear_preview()

        def _log(self, msg):
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        def _get_spectrum(self, path):
            spec = self.cache.get(path)
            if spec is None:
                spec = Spectrum(path)
                self.cache[path] = spec
            return spec

        def _on_select(self, _event=None):
            self.preview_selected()

        def preview_selected(self):
            sel = self.listbox.curselection()
            if not sel:
                self._draw(None, "", "")
                self.info_var.set(T("尚未选择文件"))
                return
            paths = [self.files[i] for i in sel]
            show = paths[:8]
            series = []
            xlabel = _DEFAULT_X_HEADER
            ylabel = _DEFAULT_Y_HEADER
            note = ""
            try:
                for k, p in enumerate(show):
                    if p.lower().endswith(".jws"):
                        spec = self._get_spectrum(p)
                        xs = spec.x_values()
                        for c in range(spec.channel_number):
                            label = os.path.basename(p)
                            if spec.channel_number > 1:
                                label += " | 通道%d" % (c + 1)
                            series.append((label, xs, spec.y_data[c],
                                           _PALETTE[(k + c) % len(_PALETTE)]))
                        if len(paths) == 1:
                            note = "%d 通道  ·  %d 点  ·  %.3f ~ %.3f  ·  强度 %.3f ~ %.3f" % (
                                spec.channel_number, spec.npoints, spec.start, spec.end,
                                min(spec.y_data[0]), max(spec.y_data[0]))
                            if spec.sample:
                                note = "样品：%s   |   " % spec.sample + note
                    else:
                        xl, rows = read_any_series(p)
                        xlabel = xl
                        for c, (lab, xs, ys) in enumerate(rows):
                            series.append((lab, xs, ys, _PALETTE[(k + c) % len(_PALETTE)]))
                            ylabel = lab
                        note = "%s   |   %d 列数据  ·  %d 点" % (
                            os.path.basename(p), len(rows), len(rows[0][1]) if rows else 0)
            except Exception as exc:
                self._draw(None, "", "")
                self.info_var.set(T("无法读取：%s") % exc)
                return
            if len(paths) > 1:
                note = "叠加显示前 %d 个文件" % len(show)
            title = os.path.basename(paths[0]) if len(paths) == 1 else "%d 个文件" % len(paths)
            self.info_var.set(note)
            self._draw(series, title, xlabel, ylabel)

        def open_csv_preview(self):
            p = filedialog.askopenfilename(
                title=T("选择要查看的 CSV 文件"),
                filetypes=[(T("CSV 文件"), "*.csv"), (T("所有文件"), "*.*")])
            if not p:
                return
            try:
                xlabel, series = load_csv_series(p)
            except Exception as exc:
                messagebox.showerror("无法读取 CSV", str(exc))
                return
            colored = [(lab, xs, ys, _PALETTE[i % len(_PALETTE)])
                       for i, (lab, xs, ys) in enumerate(series)]
            self.info_var.set(T("%s   |   %d 列数据  ·  %d 点") % (
                os.path.basename(p), len(series), len(series[0][1]) if series else 0))
            self._draw(colored, os.path.basename(p), xlabel, "Intensity")
            self._log(T("已在预览中打开 CSV：%s") % os.path.basename(p))

        def clear_preview(self):
            self._last_draw = None
            self._draw(None, "", "")
            self.info_var.set(T("尚未选择文件"))

        def save_preview_png(self):
            if not self._last_draw:
                messagebox.showinfo("提示", "当前没有可导出的预览图，请先选择文件或打开 CSV。")
                return
            if not _HAVE_PIL:
                messagebox.showerror("缺少组件", "导出 PNG 需要 Pillow（pip install pillow）。")
                return
            series, title, xlabel, ylabel = self._last_draw
            default = (os.path.splitext(os.path.basename(title))[0] or "preview") + ".png"
            p = filedialog.asksaveasfilename(
                title=T("保存预览为 PNG"), defaultextension=".png", initialfile=default,
                filetypes=[(T("PNG 图片"), "*.png")])
            if not p:
                return
            try:
                render_png(p, series, title, xlabel, ylabel, plot=self.current_plot_options())
            except Exception as exc:
                messagebox.showerror("导出失败", str(exc))
                return
            self._log(T("已保存预览图片：%s") % os.path.basename(p))

        def csv_to_png(self):
            paths = filedialog.askopenfilenames(
                title=T("选择要转成图片的 CSV 文件（可多选）"),
                filetypes=[(T("CSV 文件"), "*.csv"), (T("所有文件"), "*.*")])
            if not paths:
                return
            if not _HAVE_PIL:
                messagebox.showerror("缺少组件", "导出 PNG 需要 Pillow（pip install pillow）。")
                return
            out_dir = self._out_dir()
            base_plot = self.plot_options()
            ok = 0
            for src in paths:
                try:
                    plot = base_plot
                    ov = self._peak_overrides(src)
                    if ov:
                        plot = dict(base_plot)
                        plot.update(ov)
                    dst, nser = convert_csv_to_png(src, out_dir=out_dir, plot=plot)
                    self._log(T("CSV → PNG：%s（%d 列）") % (os.path.basename(dst), nser))
                    ok += 1
                except Exception as exc:
                    self._log(T("CSV → PNG 失败：%s（%s）") % (os.path.basename(src), exc))
            self._log(T("CSV 转图片完成，共 %d 个。") % ok)

        def _tool_out_dir(self, src):
            out = self._out_dir()
            if out:
                os.makedirs(out, exist_ok=True)
                return out
            return results_dir()

        def _selected_spectra(self):
            out = []
            opts = self.plot_options()
            for i in self.listbox.curselection():
                path = self.files[i]
                try:
                    _xl, series = read_any_series(path)
                    if not series:
                        raise JwsError(T("没有读到数据"))
                    xs = _calibrate(series[0][1], opts)
                    for lab, _sxs, ys in series:
                        name = os.path.basename(path)
                        if len(series) > 1:
                            name += "|" + lab
                        out.append((name, xs, list(ys)))
                except Exception as exc:
                    self._log(T("读取失败：%s（%s）") % (os.path.basename(path), exc))
            return out

        def _save_series_csv(self, dst, label, xs, ys):
            with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                f.write("%s,%s\n" % (_DEFAULT_X_HEADER, label))
                for x, y in zip(xs, ys):
                    f.write("%.6f,%.6f\n" % (x, y))

        def tool_fit(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择文件。")
                return
            plot = self.plot_options()
            shape_cn = T(_FIT_SHAPE_NAMES.get(plot["fit_shape"], plot["fit_shape"]))
            for name, xs, ys in specs:
                try:
                    fits = fit_spectrum_peaks(xs, ys, plot)
                except Exception as exc:
                    self._log(T("峰拟合失败：%s（%s）") % (name, exc))
                    continue
                src = self.files[self.listbox.curselection()[0]]
                base = os.path.splitext(name.split("|")[0])[0]
                dst = os.path.join(self._tool_out_dir(src), base + "_fit.csv")
                with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                    f.write(T("峰位,拟合中心,拟合峰高,拟合半高宽FWHM,峰面积,峰形,混合系数,拟合R2,来源\n"))
                    for ft in fits:
                        f.write("%.4f,%.4f,%.6g,%.4f,%.6g,%s,%.2f,%.4f,%s\n" % (
                            ft["x"], ft["center"], ft["height"], ft["fwhm"], ft["area"],
                            T(_FIT_SHAPE_NAMES.get(ft["shape"], ft["shape"])), ft["eta"],
                            ft["r2"], T("手动") if ft["manual"] else T("自动")))
                self._log(T("峰拟合（%s，%d 个峰）→ %s") % (shape_cn, len(fits), os.path.basename(dst)))
                for ft in fits[:8]:
                    self._log(T("   %8.1f  中心 %7.2f  FWHM %6.2f  面积 %9.1f  R2 %.3f") % (
                        ft["x"], ft["center"], ft["fwhm"], ft["area"], ft["r2"]))

        def _overlay_pair(self, target_name, target_xy, best):
            txs, tys = target_xy
            series = [(target_name, txs, _norm_max(tys), _PALETTE[0])]
            ref = best.get("ref_xy")
            if ref:
                series.append((best["name"], ref[0], _norm_max(ref[1]), _PALETTE[1]))
            self._draw(series,
                       "配对：%s  ↔  %s   (r=%.3f)" % (target_name, best["name"], best["corr"]),
                       _DEFAULT_X_HEADER, "归一化强度")

        def _do_pair(self, target_name, target_xy, refs, mode="配对", tol=5.0):
            mode = T(mode)
            if not refs:
                messagebox.showinfo("提示", "没有可用的参考谱。")
                return
            results = pair_spectra(target_xy, refs, self.plot_options(), tol, top=20)
            if not results:
                messagebox.showwarning("提示", "与参考谱没有重叠的波数区间，无法配对。")
                return
            sel = self.listbox.curselection()
            src = self.files[sel[0]] if sel else self.files[0]
            base = os.path.splitext(os.path.basename(src))[0]
            dst = os.path.join(self._tool_out_dir(src), base + "_配对报告.csv")
            best = results[0]
            with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                f.write(T("== 配对排名（%s，共 %d 条参考谱；按峰位匹配 F1 优先排序）==\n")
                        % (mode, len(results)))
                f.write(T("排名,参考谱,峰位匹配F1(%),命中峰数,实测峰数,参考峰数,相关系数,谱角(度)\n"))
                for i, r in enumerate(results, 1):
                    f.write("%d,%s,%.1f,%d,%d,%d,%.4f,%.2f\n" % (
                        i, r["name"], r["peak_score"], r["peak_matched"],
                        r["peak_left"], r["peak_right"], r["corr"], r["angle"]))
                f.write("\n== " + T("最佳配对峰位对照（容差 %.1f cm-1）：%s") % (tol, best["name"]) + " ==\n")
                f.write(T("实测峰位,参考峰位,偏差(cm-1),是否匹配\n"))
                for row in best.get("pair_rows", []):
                    f.write("%s,%s,%s,%s\n" % (
                        "" if row[0] is None else "%.2f" % row[0],
                        "" if row[1] is None else "%.2f" % row[1],
                        "" if row[2] is None else "%.2f" % row[2],
                        "匹配" if row[3] else "单侧"))
            self._log(T("配对完成（%s）：最佳 %s，峰位匹配 F1 %.0f%%（命中 %d/%d 实测峰，参考 %d 峰），相关系数 %.4f") % (
                mode, best["name"], best["peak_score"], best["peak_matched"],
                best["peak_left"], best["peak_right"], best["corr"]))
            for r in results[:5]:
                self._log(T("   %-30s F1 %5.1f%% (命中%d/%d 参考%d)  r=%.4f  谱角 %.2f°") % (
                    r["name"][:30], r["peak_score"], r["peak_matched"],
                    r["peak_left"], r["peak_right"], r["corr"], r["angle"]))
            matched = sum(1 for row in best.get("pair_rows", []) if row[3])
            total = len(best.get("pair_rows", []))
            if total:
                self._log(T("   峰位匹配：%d/%d 已在容差 %.1f cm-1 内") % (matched, total, tol))
            self._log(T("   报告 → %s") % os.path.basename(dst))
            try:
                png = os.path.join(self._tool_out_dir(src), base + "_配对报告.png")
                render_pair_report(png, target_name, target_xy, results, self.plot_options(), tol)
                self._log(T("   报告图 → %s") % os.path.basename(png))
            except Exception as exc:
                self._log(T("   报告图生成失败：%s") % exc)
            self._overlay_pair(target_name, target_xy, best)

        def open_identify(self):
            """未知光谱检索：不依赖矿物名，直接拿峰去全库鉴定。"""
            win = tk.Toplevel(self)
            win.title("未知光谱检索 · 全库鉴定")
            win.transient(self)
            win.geometry("1080x760")
            win.minsize(900, 620)

            state = {"name": None, "xy": None, "results": [], "total": 0, "verdict": ""}

            ttk.Label(win, text=T("不知道样品是什么？选中一条光谱，工具拿它的峰去参考库里比对，")
                                 + _sp() + T("按可信度给出候选矿物。"),
                      font=("Microsoft YaHei UI", 10, "bold")).pack(
                anchor="w", padx=12, pady=(10, 2))

            tgt = ttk.LabelFrame(win, text="① 待鉴定的未知光谱")
            tgt.pack(fill="x", padx=12, pady=4)
            tgt_var = tk.StringVar(value=T("（尚未选择）"))
            ttk.Label(tgt, textvariable=tgt_var, foreground="#1a4a8a").pack(
                side="left", padx=10, pady=6)

            def take_selection(quiet=False):
                specs = self._selected_spectra()
                if not specs:
                    if not quiet:
                        messagebox.showinfo("提示", "请先在主界面左侧选中 1 条光谱，"
                                                    "或点“选择文件…”直接指定。")
                    return
                name, xs, ys = specs[0]
                state["name"] = name
                state["xy"] = (xs, ys)
                tgt_var.set(T("%s（%d 点，%d 个峰）") % (
                    name, len(xs), len(analyze_peaks(xs, ys, self.plot_options()))))

            def pick_file():
                path = filedialog.askopenfilename(
                    title=T("选择要鉴定的光谱文件"),
                    filetypes=[("光谱文件",
                                "*.jws *.csv *.spc *.jdx *.dx *.txt *.dat *.asc *.xy"),
                               (T("所有文件"), "*.*")])
                if not path:
                    return
                try:
                    _xl, series = read_any_series(path)
                except Exception as exc:
                    messagebox.showerror("读取失败", str(exc))
                    return
                if not series:
                    messagebox.showerror("读取失败", "文件里没有可用的光谱")
                    return
                xs = _calibrate(series[0][1], self.plot_options())
                state["name"] = os.path.basename(path)
                state["xy"] = (xs, series[0][2])
                tgt_var.set(T("%s（%d 点，%d 个峰）") % (
                    state["name"], len(xs),
                    len(analyze_peaks(xs, series[0][2], self.plot_options()))))

            ttk.Button(tgt, text="用左侧选中的光谱", command=take_selection).pack(
                side="right", padx=4, pady=6)
            ttk.Button(tgt, text="选择文件…", command=pick_file).pack(
                side="right", padx=4, pady=6)

            lib = ttk.LabelFrame(win, text="② 参考库来源（勾选参与检索的数据包）")
            lib.pack(fill="x", padx=12, pady=4)
            avail = rruff_available_keys()
            pkg_vars = {}
            pkg_box = ttk.Frame(lib)
            pkg_box.pack(fill="x", padx=8, pady=(6, 2))
            if not avail:
                ttk.Label(pkg_box, text=T("本地还没有数据包。请先到【数据库 → RRUFF 数据源…】")
                                        + T("下载（建议至少下载“fair / excellent”等级，谱质更好）。"),
                          foreground="#a33").pack(anchor="w")
            for key in avail:
                pkg = rruff_package(key) or {}
                rows = None
                if os.path.exists(_peak_index_path(key)):
                    try:
                        rows = rruff_peak_index(key)
                    except Exception:
                        rows = None
                var = tk.BooleanVar(value=True)
                pkg_vars[key] = var
                text = "%s · %s（%s）" % (pkg_kind(pkg), pkg_label(pkg, key),
                                        T("索引 %d 条") % len(rows) if rows else T("未建索引"))
                ttk.Checkbutton(pkg_box, text=text, variable=var).pack(anchor="w")
            ctrl = ttk.Frame(lib)
            ctrl.pack(fill="x", padx=8, pady=(2, 6))
            ttk.Label(ctrl, text="检索类型：").pack(side="left")
            kind_list = ("拉曼", "红外", "XRD", "成分", "全部")
            kind_var = tk.StringVar(value=T("拉曼"))
            ttk.Combobox(ctrl, textvariable=kind_var, state="readonly", width=8,
                         values=[T(v) for v in kind_list]).pack(side="left")

            def kind_internal():
                """下拉框显示的是译文，这里换回内部值。"""
                cur = kind_var.get()
                for v in kind_list:
                    if T(v) == cur:
                        return v
                return "拉曼"
            ttk.Label(ctrl, text="必须含元素：").pack(side="left", padx=(12, 2))
            must_var = tk.StringVar(value="")
            ttk.Entry(ctrl, textvariable=must_var, width=10).pack(side="left")
            ttk.Label(ctrl, text="必须不含：").pack(side="left", padx=(8, 2))
            not_var = tk.StringVar(value="")
            ttk.Entry(ctrl, textvariable=not_var, width=10).pack(side="left")
            ttk.Label(ctrl, text="（如 Zr Si，空格分隔；不知道就留空）",
                      foreground="#888").pack(side="left", padx=6)
            row2 = ttk.Frame(lib)
            row2.pack(fill="x", padx=8, pady=(0, 6))
            ttk.Label(row2, text="候选个数：").pack(side="left")
            top_var = tk.StringVar(value="12")
            ttk.Entry(row2, textvariable=top_var, width=5).pack(side="left")
            exact_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(row2, text="对前列候选读原始谱精算（更准，稍慢）",
                            variable=exact_var).pack(side="left", padx=10)
            rebuild_var = tk.BooleanVar(value=False)
            ttk.Checkbutton(row2, text="强制重建特征索引", variable=rebuild_var).pack(
                side="left", padx=6)

            status = tk.StringVar(value=T("选择好未知光谱后，点【开始检索】。"))
            ttk.Label(win, textvariable=status, foreground="#444", wraplength=1020,
                      justify="left").pack(anchor="w", padx=12, pady=(4, 2))

            res_box = ttk.LabelFrame(win, text="③ 候选矿物（按可信度排序）")
            res_box.pack(fill="both", expand=True, padx=12, pady=4)
            cols = ("rank", "name", "rid", "kind", "wave", "score", "f1", "strong",
                    "hit", "dev", "corr", "variants", "chem", "system", "src")
            heads = ("排名", "矿物", "编号", "类型", "波长", "综合分", "F1", "强峰命中",
                     "命中/参考", "平均偏差", "相关系数", "库内谱数", "化学式", "晶系", "方式")
            widths = (46, 150, 84, 50, 54, 62, 56, 74, 74, 72, 74, 70, 150, 90, 50)
            tree = ttk.Treeview(res_box, columns=cols, show="headings", selectmode="browse")
            for c, h, w in zip(cols, heads, widths):
                tree.heading(c, text=h)
                tree.column(c, width=w, anchor="w")
            ysb = ttk.Scrollbar(res_box, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=ysb.set)
            tree.pack(side="left", fill="both", expand=True, padx=(6, 0), pady=6)
            ysb.pack(side="left", fill="y", pady=6, padx=(0, 6))

            verdict_var = tk.StringVar(value="")
            ttk.Label(win, textvariable=verdict_var, foreground="#1a4a8a",
                      wraplength=1020, justify="left").pack(anchor="w", padx=12)

            def selected_keys():
                return [k for k, v in pkg_vars.items() if v.get()]

            def build_index():
                keys = selected_keys()
                if not keys:
                    status.set(T("请先勾选至少一个数据包。"))
                    return
                for key in keys:
                    need = rebuild_var.get() or not os.path.exists(_peak_index_path(key))
                    if not need:
                        continue
                    pkg = rruff_package(key) or {}
                    status.set(T("正在为 %s 建立特征索引（每条谱都要解析一次）…")
                               % pkg.get("label", key))
                    win.update_idletasks()
                    try:
                        rows = rruff_peak_index(
                            key, progress=lambda i, n: win.update_idletasks(),
                            rebuild=True)
                    except Exception as exc:
                        status.set(T("建立索引失败（%s）：%s") % (key, exc))
                        return
                    self._log(T("特征索引：%s -> %d 条") % (key, len(rows)))
                status.set(T("索引已就绪，可以开始检索。"))

            def run():
                if state["xy"] is None:
                    take_selection()
                    if state["xy"] is None:
                        return
                keys = selected_keys()
                keys = [k for k in keys if os.path.exists(_peak_index_path(k))]
                if not keys:
                    status.set(T("所选的包还没有特征索引，正在先建立…"))
                    win.update_idletasks()
                    build_index()
                    keys = [k for k in selected_keys()
                            if os.path.exists(_peak_index_path(k))]
                    if not keys:
                        status.set(T("特征索引建立失败或不完整，无法检索。"))
                        return
                _kind = kind_internal()
                kinds = None if _kind == "全部" else {_kind}
                try:
                    top = max(1, int(float(top_var.get())))
                except ValueError:
                    top = 12
                status.set(T("正在检索…"))
                win.update_idletasks()
                try:
                    results, total, verdict = identify_unknown(
                        state["xy"], self.plot_options(), keys=keys, kinds=kinds,
                        must=must_var.get(), not_=not_var.get(), top=top,
                        exact_pool=(max(IDENTIFY_POOL, top) if exact_var.get() else 0),
                        progress=lambda i, n: win.update_idletasks())
                except Exception as exc:
                    status.set(T("检索失败：%s") % exc)
                    return
                state["results"] = results
                state["total"] = total
                state["verdict"] = verdict
                tree.delete(*tree.get_children())
                for i, row in enumerate(format_identify_rows(results)):
                    tree.insert("", "end", iid=str(i), values=tuple(row))
                verdict_var.set(T("结论：") + verdict)
                status.set(T("检索完成：共比对 %d 条参考记录，列出 %d 个候选。")
                           % (total, len(results)))
                self._log(T("未知光谱检索（%s）：比对 %d 条，最佳候选 %s")
                          % (state["name"], total,
                             results[0]["name"] if results else "无"))

            def picked():
                sel = tree.selection()
                if not sel:
                    return None
                try:
                    return state["results"][int(sel[0])]
                except (ValueError, IndexError):
                    return None

            def export_csv():
                if not state["results"]:
                    status.set(T("还没有检索结果。"))
                    return
                base = os.path.splitext(str(state["name"]).split("|")[0])[0]
                dst = os.path.join(results_dir(),
                                   "%s_未知光谱检索.csv" % _safe_name(base))
                write_identify_csv(dst, state["name"], state["results"],
                                   state["verdict"], state["total"])
                status.set(T("候选清单已保存：%s") % dst)
                self._log(T("未知光谱检索结果：%s") % dst)

            def show_overlay():
                if not state["results"]:
                    status.set(T("还没有检索结果。"))
                    return
                best = next((r for r in state["results"] if r.get("ref_xy")), None)
                if best is None:
                    status.set(T("最佳候选没有读到原始谱（可能未精算），先勾选“精算”重跑一次。"))
                    return
                base = os.path.splitext(str(state["name"]).split("|")[0])[0]
                dst = os.path.join(results_dir(),
                                   "%s_未知光谱检索_对比.png" % _safe_name(base))
                txs, tys = state["xy"]
                try:
                    render_png(dst, [(state["name"], txs, tys, _PALETTE[0]),
                                     ("%s %s" % (best["name"], best["rruffid"]),
                                      best["ref_xy"][0], best["ref_xy"][1], _PALETTE[1])],
                               T("未知光谱 vs 最佳候选 %s") % best["name"],
                               _DEFAULT_X_HEADER, "Intensity", self.plot_options())
                except Exception as exc:
                    status.set(T("对比图生成失败：%s") % exc)
                    return
                status.set(T("对比图已保存：%s") % dst)
                self._show_image(dst, T("未知光谱 vs 最佳候选"))

            def export_candidate():
                item = picked()
                if item is None:
                    status.set(T("请先在候选表里选中一条。"))
                    return
                try:
                    saved = rruff_export(item["key"], [item["entry"]])
                except Exception as exc:
                    status.set(T("导出失败：%s") % exc)
                    return
                status.set(T("已导出到参考谱库：%s（可在【分析工具 → 配对比较】里继续比对）")
                           % "、".join(os.path.basename(p) for p in saved))
                self._log(T("候选参考谱已导出：%s") % "、".join(saved))

            def pair_best():
                item = picked() or next(
                    (r for r in state["results"] if r.get("named", True)), None)
                if item is None:
                    status.set(T("还没有可用的候选。"))
                    return
                try:
                    paths = rruff_export(item["key"], [item["entry"]])
                except Exception as exc:
                    status.set(T("导出候选失败：%s") % exc)
                    return
                if not paths:
                    status.set(T("候选导出失败，无法配对。"))
                    return
                txs, tys = state["xy"]
                self._do_pair(state["name"], (txs, tys), list(paths),
                              "未知光谱 × %s" % item["name"])

            row = ttk.Frame(win)
            row.pack(fill="x", padx=12, pady=(4, 12))
            ttk.Button(row, text="建立 / 更新特征索引", command=build_index).pack(side="left")
            ttk.Button(row, text="开始检索", style="Big.TButton",
                       command=run).pack(side="left", padx=8)
            ttk.Button(row, text="导出候选 CSV", command=export_csv).pack(side="left", padx=4)
            ttk.Button(row, text="最佳候选对比图", command=show_overlay).pack(side="left", padx=4)
            ttk.Button(row, text="导出选中到参考谱库", command=export_candidate).pack(
                side="left", padx=4)
            ttk.Button(row, text="用候选做配对报告", command=pair_best).pack(side="left", padx=4)
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right")

            take_selection(quiet=True)
            if state["xy"] is None:
                status.set(T("请先在主界面左侧选中 1 条未知光谱，或点【选择文件…】。")
                           + T("第一次使用请先【建立 / 更新特征索引】。"))

        def _batch_window(self, mode):
            """批量配对 / 批量鉴定的共用窗口。

            mode="pair"     多条实测谱 × 一组参考谱，逐条给最佳参考（找对不上的）
            mode="identify" 多条陌生谱逐条全库鉴定，给最像的矿物
            两者都只列辅助指标与「参考判读」，最终判断由使用者自己做。
            """
            is_pair = (mode == "pair")
            plot = self.plot_options()

            win = tk.Toplevel(self)
            win.title(T("批量配对（文件夹 × 参考谱）") if is_pair
                      else T("批量鉴定（文件夹逐条鉴定）"))
            win.transient(self)
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            win.geometry("%dx%d" % (min(1180, max(760, sw - 100)),
                                    min(720, max(420, sh - 200))))

            ttk.Label(
                win,
                text=(T("把一整个文件夹的实测谱逐条与参考谱打分，汇总成一张表；")
                      + T("表里只有辅助指标，最后是哪个物相由你自己判断。")
                      if is_pair else
                      T("把一整个文件夹的谱逐条在全库里检索，汇总成一张表；")
                      + T("每条给出最佳候选与参考判读，最终判断由你自己做。")),
                wraplength=1100, justify="left",
                foreground="#444").pack(anchor="w", padx=12, pady=(10, 4))

            # ---- 实测谱来源：文件夹，或左侧选中 ----
            trow = ttk.Frame(win)
            trow.pack(fill="x", padx=12, pady=2)
            ttk.Label(trow, text=T("实测谱文件夹：")).pack(side="left")
            dir_var = tk.StringVar(value="")

            def pick_dir():
                d = filedialog.askdirectory(title=T("选择放实测谱的文件夹"))
                if d:
                    dir_var.set(d)

            ttk.Entry(trow, textvariable=dir_var).pack(
                side="left", fill="x", expand=True, padx=4)
            ttk.Button(trow, text=T("浏览…"), command=pick_dir).pack(side="left")
            ttk.Label(trow, text=T("（留空 = 用主界面左侧选中的光谱）"),
                      foreground="#888").pack(side="left", padx=6)

            # ---- 参考谱（只配对需要）----
            ref_var = tk.StringVar(value="lib")
            ref_dir_var = tk.StringVar(value="")
            if is_pair:
                rrow = ttk.Frame(win)
                rrow.pack(fill="x", padx=12, pady=2)
                ttk.Label(rrow, text=T("参考谱：")).pack(side="left")
                ttk.Radiobutton(rrow, text=T("本地参考谱库（%d 条）") % len(db_list()),
                                variable=ref_var, value="lib").pack(side="left")
                ttk.Radiobutton(rrow, text=T("指定文件夹 / 文件："),
                                variable=ref_var, value="dir").pack(side="left", padx=(10, 2))

                def pick_ref():
                    d = filedialog.askdirectory(title=T("选择参考谱文件夹"))
                    if d:
                        ref_dir_var.set(d)
                        ref_var.set("dir")

                ttk.Entry(rrow, textvariable=ref_dir_var).pack(
                    side="left", fill="x", expand=True, padx=4)
                ttk.Button(rrow, text=T("浏览…"), command=pick_ref).pack(side="left")

            # ---- 鉴定选项 ----
            top_var = tk.StringVar(value="5")
            must_var = tk.StringVar(value="")
            not_var = tk.StringVar(value="")
            if not is_pair:
                orow = ttk.Frame(win)
                orow.pack(fill="x", padx=12, pady=2)
                ttk.Label(orow, text=T("候选个数：")).pack(side="left")
                ttk.Entry(orow, textvariable=top_var, width=5).pack(side="left")
                ttk.Label(orow, text=T("必须含元素：")).pack(side="left", padx=(12, 2))
                ttk.Entry(orow, textvariable=must_var, width=10).pack(side="left")
                ttk.Label(orow, text=T("必须不含：")).pack(side="left", padx=(8, 2))
                ttk.Entry(orow, textvariable=not_var, width=10).pack(side="left")
                ttk.Label(orow, text=T("（如 Zr Si，空格分隔；不知道就留空）"),
                          foreground="#888").pack(side="left", padx=6)

            brow = ttk.Frame(win)
            brow.pack(fill="x", padx=12, pady=(6, 2))
            run_btn = ttk.Button(brow, text=T("开始批量处理"))
            run_btn.pack(side="left")
            csv_btn = ttk.Button(brow, text=T("打开汇总表 CSV"), state="disabled")
            csv_btn.pack(side="right")
            html_btn = ttk.Button(brow, text=T("打开汇总报告 HTML"), state="disabled")
            html_btn.pack(side="right", padx=6)
            ttk.Button(brow, text=T("关闭"), command=win.destroy).pack(side="right")

            status = tk.StringVar(value=T("选好文件夹后点【开始批量处理】。"))
            ttk.Label(win, textvariable=status, foreground="#1a4a8a",
                      wraplength=1100, justify="left").pack(anchor="w", padx=12)

            heads = BATCH_PAIR_COLUMNS if is_pair else BATCH_IDENTIFY_COLUMNS
            if is_pair:
                widths = (250, 64, 64, 230, 60, 56, 64, 64, 64, 64, 64, 88, 68, 68,
                          64, 92, 260)
            else:
                widths = (250, 64, 64, 170, 84, 50, 54, 60, 56, 64, 64, 64, 64, 88,
                          68, 74, 68, 300)
            box = ttk.LabelFrame(win, text=T("结果（按综合分升序，最可疑的在前）"))
            box.pack(fill="both", expand=True, padx=12, pady=(6, 10))
            inner = ttk.Frame(box)
            inner.pack(fill="both", expand=True, padx=6, pady=6)
            cols = tuple("c%d" % i for i in range(len(heads)))
            tree = ttk.Treeview(inner, columns=cols, show="headings", selectmode="browse")
            for c, h, w in zip(cols, heads, widths):
                tree.heading(c, text=T(h))
                tree.column(c, width=w, anchor="w")
            ysb = ttk.Scrollbar(inner, orient="vertical", command=tree.yview)
            xsb = ttk.Scrollbar(inner, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
            tree.grid(row=0, column=0, sticky="nsew")
            ysb.grid(row=0, column=1, sticky="ns")
            xsb.grid(row=1, column=0, sticky="ew")
            inner.rowconfigure(0, weight=1)
            inner.columnconfigure(0, weight=1)

            state = {"rows": [], "csv": "", "html": ""}

            def target_spectra():
                """返回 (光谱列表, 用来命名输出文件的标签)。"""
                folder = dir_var.get().strip()
                if folder:
                    files = [p for p in find_input_files([folder])
                             if not any(os.path.basename(p).lower().endswith(sfx)
                                        for sfx in _SKIP_FILE_SUFFIXES)]
                    return (collect_spectra(files, plot, verbose=False),
                            os.path.basename(os.path.abspath(folder)))
                return self._selected_spectra(), T("主界面选中")

            def fill(rows, order_key, row_fn):
                tree.delete(*tree.get_children())
                for i, r in enumerate(sorted(rows, key=order_key)):
                    tree.insert("", "end", iid=str(i), values=tuple(row_fn(r)))

            def open_file(path):
                if path and os.path.isfile(path):
                    try:
                        os.startfile(path)
                        return
                    except OSError:
                        pass
                status.set(T("文件不存在：%s") % path)

            def run():
                run_btn.configure(state="disabled")
                win.update_idletasks()

                def tick(i, n, name):
                    if name:
                        status.set(T("正在处理 %d/%d：%s") % (i + 1, n, name))
                        win.update_idletasks()

                try:
                    spectra, label = target_spectra()
                    if not spectra:
                        status.set(T("没有读到实测谱：请选一个文件夹，或先在主界面左侧选中光谱。"))
                        return
                    out = results_dir()
                    base = _safe_name(label) or T("批量结果")
                    if is_pair:
                        if ref_var.get() == "lib":
                            ref_paths = db_list()
                            ref_label = T("本地参考谱库")
                        else:
                            rp = ref_dir_var.get().strip()
                            if not rp:
                                status.set(T("请先指定参考谱文件夹或文件。"))
                                return
                            ref_paths = find_input_files([rp])
                            ref_label = os.path.basename(os.path.abspath(rp))
                        refs = load_reference_set(ref_paths)
                        if not refs:
                            status.set(T("参考谱一条也读不出来（需要两列文本 / CSV）：%s")
                                       % ref_label)
                            return
                        status.set(T("正在批量配对：%d 条实测谱 × %d 条参考谱…")
                                   % (len(spectra), len(refs)))
                        win.update_idletasks()
                        rows = batch_pair(spectra, refs, plot, progress=tick)
                        csv_path = os.path.join(out, "%s_批量配对汇总.csv" % base)
                        html_path = os.path.join(out, "%s_批量配对汇总.html" % base)
                        write_batch_pair_csv(csv_path, rows, ref_label)
                        fill(rows, _pair_sort_key, batch_pair_row)
                        bad = sum(1 for r in rows
                                  if not r.get("best") or r["best"]["score"] < 35)
                        self._log(T("批量配对完成：%d 条实测谱 × %d 条参考谱，"
                                    "其中 %d 条匹配很差（<35 分）")
                                  % (len(rows), len(refs), bad))
                        status.set(T("完成：%d 条；其中 %d 条匹配很差（<35 分），"
                                     "已排在最前。汇总表：%s")
                                   % (len(rows), bad, csv_path))
                    else:
                        keys = peak_index_keys()
                        if not keys:
                            status.set(T("还没有可用于检索的特征索引："
                                         "请先在【未知光谱检索】窗口里建立索引。"))
                            return
                        try:
                            top = max(1, int(float(top_var.get())))
                        except ValueError:
                            top = 5
                        status.set(T("正在批量鉴定：%d 条谱 × %d 个数据包…")
                                   % (len(spectra), len(keys)))
                        win.update_idletasks()
                        rows = batch_identify(spectra, plot, keys=keys, kinds=None,
                                              must=must_var.get(), not_=not_var.get(),
                                              top=top, progress=tick)
                        total_entries = max([r.get("total", 0) for r in rows] or [0])
                        csv_path = os.path.join(out, "%s_批量鉴定汇总.csv" % base)
                        html_path = os.path.join(out, "%s_批量鉴定汇总.html" % base)
                        write_batch_identify_csv(csv_path, rows, total_entries, top)
                        fill(rows, _identify_sort_key, batch_identify_row)
                        bad = 0
                        for r in rows:
                            b = _best_named(r.get("candidates"))
                            if b is None or b["score"] < 35:
                                bad += 1
                        self._log(T("批量鉴定完成：%d 条谱，比对 %d 条参考记录，"
                                    "其中 %d 条没有可靠匹配（<35 分）")
                                  % (len(rows), total_entries, bad))
                        status.set(T("完成：%d 条；其中 %d 条没有可靠匹配（<35 分），"
                                     "已排在最前。汇总表：%s")
                                   % (len(rows), bad, csv_path))
                    state["rows"] = rows
                    state["csv"] = csv_path
                    state["html"] = html_path
                    try:
                        if is_pair:
                            write_batch_pair_html(html_path, rows, ref_label)
                        else:
                            write_batch_identify_html(html_path, rows, total_entries)
                        self._log(T("汇总报告：%s") % html_path)
                    except Exception as exc:
                        self._log(T("HTML 汇总生成失败：%s") % exc)
                    csv_btn.configure(state="normal",
                                      command=lambda: open_file(csv_path))
                    html_btn.configure(state="normal",
                                       command=lambda: open_file(html_path))
                except Exception as exc:
                    status.set(T("批量处理失败：%s") % exc)
                finally:
                    run_btn.configure(state="normal")

            run_btn.configure(command=run)

        def tool_batch_pair(self):
            """批量配对：一整个文件夹的实测谱 × 一组参考谱，找对不上的。"""
            self._batch_window("pair")

        def tool_batch_identify(self):
            """批量鉴定：一整个文件夹的未知谱逐条全库检索。"""
            self._batch_window("identify")

        def tool_pair_manual(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择 1 个待配对的实测谱。")
                return
            refs = filedialog.askopenfilenames(
                title=T("选择参考谱文件（可多选；取消则改为选文件夹）"),
                filetypes=[(T("谱文件"), "*.txt *.csv *.dat"), (T("所有文件"), "*.*")])
            if not refs:
                folder = filedialog.askdirectory(title=T("选择参考谱所在文件夹"))
                if not folder:
                    return
                refs = find_input_files([folder])
            name, xs, ys = specs[0]
            self._do_pair(name, (xs, ys), list(refs), "手动配对")

        def tool_auto_pair(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择 1 个待配对的实测谱。")
                return
            refs = db_list()
            if not refs:
                if messagebox.askyesno("本地数据库为空",
                                       "本地数据库还没有参考谱线。\n是否现在在线检索并下载？"):
                    self.tool_auto_pair_online()
                return
            name, xs, ys = specs[0]
            self._do_pair(name, (xs, ys), refs, "自动配对（本地库 %d 条）" % len(refs))

        def tool_auto_pair_online(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择 1 个待配对的实测谱。")
                return
            name, xs, ys = specs[0]
            guess = ""
            local = db_list()
            if local:
                first = match_references((xs, ys), local, top=1)
                if first:
                    guess = os.path.basename(first[0]["name"]).split("_")[0].replace("_", " ")
            key = simpledialog.askstring(
                "自动配对 · 在线检索",
                "将在 ROD 数据库（RRUFF 拉曼数据）检索以下关键词，\n"
                "下载候选谱线后与实测谱自动配对（可修改）：",
                initialvalue=guess, parent=self)
            if not key:
                return
            self._log(T("正在检索数据库：%s …（在后台进行，界面不会卡住）") % key)

            def work(report, cancel):
                report(text=T("正在检索数据库：%s …") % key, got=0, total=0)
                rows = rod_search(key)
                if not rows:
                    return [], 0
                todo = rows[:8]
                got = 0
                for k, row in enumerate(todo):
                    if cancel.is_set():
                        raise DownloadCancelled()
                    report(text=T("正在下载候选谱线 %d/%d：%s")
                           % (k + 1, len(todo), row["file"]), got=k, total=len(todo))
                    try:
                        dst, _meta = db_download(row["file"], row.get("mineral"),
                                                 row.get("wavelength"))
                        got += 1
                        self._log(T("   下载 %s → %s")
                                  % (row["file"], os.path.basename(dst)))
                    except Exception as exc:
                        self._log(T("   下载失败 %s：%s") % (row["file"], exc))
                return rows, got

            def done(result):
                rows, got = result
                if not rows:
                    messagebox.showinfo("提示", "数据库中没有找到“%s”。" % key)
                    return
                self._log(T("检索到 %d 条，下载了 %d 条用于配对。") % (len(rows), got))
                if not got and not db_list():
                    messagebox.showwarning("提示", "没能下载到参考谱。")
                    return
                self._do_pair(name, (xs, ys), db_list(),
                              "自动配对（在线，新下载 %d 条）" % got)

            def failed(exc):
                if isinstance(exc, DownloadCancelled):
                    self._log(T("在线检索 / 下载已取消。"))
                    return
                messagebox.showerror("检索失败", "无法访问数据库：%s" % exc)

            self._bg_run(self, work, tk.StringVar(), on_done=done, on_error=failed,
                         busy_text=T("正在连接数据库 …"))

        def tool_match(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择 1 个待识别文件。")
                return
            paths = filedialog.askopenfilenames(
                title=T("选择参考谱文件（可多选，支持 RRUFF 导出的 .txt / .csv）"),
                filetypes=[(T("谱文件"), "*.txt *.csv *.dat"), (T("所有文件"), "*.*")])
            if not paths:
                return
            name, xs, ys = specs[0]
            results = match_references((xs, ys), list(paths), top=15)
            if not results:
                messagebox.showwarning("提示", "参考谱与待测谱没有重叠的波数区间，无法比对。")
                return
            src = self.files[self.listbox.curselection()[0]]
            base = os.path.splitext(os.path.basename(src))[0]
            dst = os.path.join(self._tool_out_dir(src), base + "_match.csv")
            with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                f.write(T("排名,参考谱,相关系数,谱角(度),重叠下限,重叠上限\n"))
                for i, r in enumerate(results, 1):
                    f.write("%d,%s,%.4f,%.2f,%.2f,%.2f\n" % (
                        i, r["name"], r["corr"], r["angle"],
                        r["overlap"][0], r["overlap"][1]))
            self._log(T("光谱比对：%s 与 %d 个参考谱 → %s") % (
                os.path.basename(src), len(results), os.path.basename(dst)))
            for r in results[:6]:
                self._log(T("   %-26s 相关系数 %.4f  谱角 %.2f°") % (
                    r["name"][:26], r["corr"], r["angle"]))

        def tool_search(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择文件。")
                return
            path = filedialog.askopenfilename(
                title=T("选择参考峰位表（CSV：名称,峰位1,峰位2,…）"),
                filetypes=[(T("CSV/文本"), "*.csv *.txt"), (T("所有文件"), "*.*")])
            if not path:
                return
            name, xs, ys = specs[0]
            positions = [pk["x"] for pk in analyze_peaks(xs, ys, self.plot_options())]
            if not positions:
                messagebox.showwarning("提示", "未识别到峰，无法检索。可降低“峰灵敏阈值”。")
                return
            rows = search_reference_table(positions, path, 5.0)
            src = self.files[self.listbox.curselection()[0]]
            base = os.path.splitext(os.path.basename(src))[0]
            dst = os.path.join(self._tool_out_dir(src), base + "_search.csv")
            with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                f.write(T("排名,候选,匹配率(%),匹配峰数,参考峰数\n"))
                for i, (score, rname, matched, total) in enumerate(rows, 1):
                    f.write("%d,%s,%.1f,%d,%d\n" % (i, rname, score, matched, total))
            self._log(T("峰位检索（用 %d 个峰，容差 5 cm-1）→ %s") % (
                len(positions), os.path.basename(dst)))
            for i, (score, rname, matched, total) in enumerate(rows[:6], 1):
                self._log(T("   %d. %-24s 匹配 %.0f%% (%d/%d)") % (
                    i, rname[:24], score, matched, total))

        def tool_similarity(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "请先选择至少 2 个文件。")
                return
            mat = similarity_matrix(specs)
            src = self.files[self.listbox.curselection()[0]]
            dst = os.path.join(self._tool_out_dir(src), "相似度矩阵.csv")
            with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                f.write(T("相似度(相关系数),") + ",".join(s[0] for s in specs) + "\n")
                for i, s in enumerate(specs):
                    f.write(s[0] + "," + ",".join("%.4f" % v for v in mat[i]) + "\n")
            self._log(T("相似度矩阵（%d×%d）→ %s") % (len(specs), len(specs), os.path.basename(dst)))
            for i in range(min(len(specs), 6)):
                best = max(range(len(specs)), key=lambda j: mat[i][j] if j != i else -2)
                if best != i:
                    self._log(T("   %-22s 最相似：%s（r=%.4f）") % (
                        specs[i][0][:22], specs[best][0][:22], mat[i][best]))

        def tool_average(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "请先选择至少 2 个文件。")
                return
            lo = max(min(s[1]) for s in specs)
            hi = min(max(s[1]) for s in specs)
            if hi <= lo:
                messagebox.showwarning("提示", "所选光谱没有共同的重叠区间，无法平均。")
                return
            grid = [x for x in specs[0][1] if lo <= x <= hi]
            if len(grid) < 4:
                messagebox.showwarning("提示", "重叠区间采样点太少，无法平均。")
                return
            acc = [0.0] * len(grid)
            for _n, xs, ys in specs:
                for i, v in enumerate(_resample(xs, ys, grid)):
                    acc[i] += v
            avg = [v / len(specs) for v in acc]
            src = self.files[self.listbox.curselection()[0]]
            dst = os.path.join(self._tool_out_dir(src), "平均光谱_%d条.csv" % len(specs))
            self._save_series_csv(dst, "平均强度(%d条)" % len(specs), grid, avg)
            self._log(T("平均完成（%d 条，%d 点）→ %s") % (len(specs), len(grid), os.path.basename(dst)))
            self._draw([("平均光谱(%d条)" % len(specs), grid, avg, _PALETTE[0])],
                       "平均光谱（%d 条）" % len(specs), _DEFAULT_X_HEADER, "Intensity")

        def tool_subtract(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "请先选择 2 个文件（第一个作为 A，第二个作为 B）。")
                return
            name_a, axs, ays = specs[0]
            name_b, bxs, bys = specs[1]
            lo = max(min(axs), min(bxs))
            hi = min(max(axs), max(bxs))
            if hi <= lo:
                messagebox.showwarning("提示", "两条光谱没有重叠区间，无法相减。")
                return
            grid = [x for x in axs if lo <= x <= hi]
            if len(grid) < 4:
                messagebox.showwarning("提示", "重叠区间采样点太少。")
                return
            va = _resample(axs, ays, grid)
            vb = _resample(bxs, bys, grid)
            diff = [u - v for u, v in zip(va, vb)]
            src = self.files[self.listbox.curselection()[0]]
            dst = os.path.join(self._tool_out_dir(src), "相减_A减B.csv")
            self._save_series_csv(dst, "A-B", grid, diff)
            self._log(T("相减完成：A=%s  B=%s（%d 点）→ %s") % (
                name_a, name_b, len(grid), os.path.basename(dst)))
            self._draw([("A−B", grid, diff, _PALETTE[1])],
                       "A−B：%s − %s" % (name_a, name_b), _DEFAULT_X_HEADER, "强度差")

        def _on_canvas_resize(self, _event=None):
            if self._resize_job:
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(120, self._redraw)

        def _redraw(self):
            if self._last_draw:
                series, title, xlabel, ylabel = self._last_draw
                self._paint(series, title, xlabel, ylabel)
            else:
                self._draw(None, "", "")

        def _draw(self, series, title, xlabel, ylabel="Intensity"):
            self._last_draw = (series, title, xlabel, ylabel) if series else None
            self._paint(series, title, xlabel, ylabel)

        def _legend_measure(self):
            """画布预览里量文字宽度的函数（字号比 PNG 导出的那个小一号）。"""
            if getattr(self, "_leg_font", None) is None:
                try:
                    from tkinter import font as tkfont
                    self._leg_font = tkfont.Font(font=("Microsoft YaHei UI", 8))
                except Exception:
                    self._leg_font = False
            font = self._leg_font

            def measure(text):
                if not font:
                    return 8.0 * len(text)
                try:
                    return float(font.measure(text))
                except Exception:
                    return 8.0 * len(text)
            return measure

        def _paint(self, series, title, xlabel, ylabel):
            cv = self.canvas
            cv.delete("all")
            w = cv.winfo_width()
            h = cv.winfo_height()
            if w < 60 or h < 60:
                return
            if not series:
                self._view = None
                self._view_series = None
                cv.create_text(w // 2, h // 2, text="暂无预览\n选中左侧文件，或点“打开 CSV 看图”",
                               fill="#999", font=("Microsoft YaHei UI", 11), justify="center")
                return

            p = self.current_plot_options()
            series = [(lab, xs, _process_signal(ys, p, xs), col) for lab, xs, ys, col in series]
            ml = 76 if p["show_y_ticks"] else 46
            mt_base, mb_base = 40, 48
            mt, mb = mt_base, mb_base
            base_pw = w - ml - 22
            base_ph = h - mt - mb

            # 图例跟导出图一个口径：画在绘图区外面（或用户拖到的图内位置），不压谱线
            legend_pos = p["legend_pos"] if len(series) > 1 else "none"
            entries, (l_right, l_top, l_bottom), l_avail = legend_fit(
                series, self._legend_measure(), legend_pos, w, base_pw, base_ph)
            mr = max(22, l_right)
            mt += l_top
            mb += l_bottom
            pw = w - ml - mr
            ph = h - mt - mb
            if pw < 20 or ph < 20:
                return

            plot_xmin = min(min(s[1]) for s in series)
            plot_xmax = max(max(s[1]) for s in series)
            plot_ymin = min(min(s[2]) for s in series)
            plot_ymax = max(max(s[2]) for s in series)
            if p["x_min"] is not None:
                plot_xmin = float(p["x_min"])
            if p["x_max"] is not None:
                plot_xmax = float(p["x_max"])
            auto_y = p["y_min"] is None and p["y_max"] is None
            if p["y_min"] is not None:
                plot_ymin = float(p["y_min"])
            if p["y_max"] is not None:
                plot_ymax = float(p["y_max"])
            if plot_xmax <= plot_xmin:
                plot_xmax = plot_xmin + 1.0
            if plot_ymax <= plot_ymin:
                plot_ymax = plot_ymin + 1.0
            if auto_y:
                pad = (plot_ymax - plot_ymin) * 0.06
                plot_ymin -= pad
                plot_ymax += pad

            def sx(x):
                return ml + (x - plot_xmin) / (plot_xmax - plot_xmin) * pw

            def sy(y):
                return mt + ph - (y - plot_ymin) / (plot_ymax - plot_ymin) * ph

            def inside(a, b):
                return plot_xmin <= a <= plot_xmax and plot_ymin <= b <= plot_ymax

            self._view = {"xmin": plot_xmin, "xmax": plot_xmax, "ymin": plot_ymin,
                          "ymax": plot_ymax, "ml": ml, "pw": pw, "mt": mt, "ph": ph}
            self._view_series = series

            ticks = _axis_ticks(plot_xmin, plot_xmax, p["x_step"], p["x_start"])

            cv.create_rectangle(ml, mt, ml + pw, mt + ph, outline="#dddddd", fill="#fdfdfd")

            if p["show_grid"]:
                for i in range(1, 5):
                    gy = mt + ph * i / 5
                    cv.create_line(ml, gy, ml + pw, gy, fill="#eeeeee")
                for xv in ticks:
                    gx = sx(xv)
                    cv.create_line(gx, mt, gx, mt + ph, fill="#eeeeee")

            for xv in ticks:
                gx = sx(xv)
                cv.create_line(gx, mt + ph, gx, mt + ph + 4, fill="#888888")
                cv.create_text(gx, mt + ph + 13, text="%.6g" % xv, fill="#555",
                               font=("Consolas", 8))
            if p["show_y_ticks"]:
                for i in range(6):
                    gy = mt + ph - ph * i / 5
                    yv = plot_ymin + (plot_ymax - plot_ymin) * i / 5
                    cv.create_line(ml - 4, gy, ml, gy, fill="#888888")
                    cv.create_text(ml - 7, gy, text="%.6g" % yv, fill="#555",
                                   font=("Consolas", 8), anchor="e")

            cv.create_line(ml, mt, ml, mt + ph, fill="#666666")
            cv.create_line(ml, mt + ph, ml + pw, mt + ph, fill="#666666")

            for label, xs, ys, color in series:
                dx, dy = _decimate(xs, ys, int(pw))
                seg = []
                for xv, yv in zip(dx, dy):
                    if not inside(xv, yv):
                        if len(seg) >= 4:
                            cv.create_line(*seg, fill=color, width=1)
                        seg = []
                        continue
                    seg.append(sx(xv))
                    seg.append(sy(yv))
                if len(seg) >= 4:
                    cv.create_line(*seg, fill=color, width=1)

            if p["annotate_peaks"]:
                for _label, xs, ys, _color in series:
                    for k, pk in enumerate(analyze_peaks(xs, ys, p, processed=True)):
                        if not inside(pk["x"], pk["y"]):
                            continue
                        gx = sx(pk["x"])
                        gy = sy(pk["y"])
                        mark_color = "#1450aa" if pk.get("manual") else "#aa1e1e"
                        # 峰位波长虚线：从峰顶引到横坐标轴（自动峰和手动峰一视同仁）
                        if p["peak_dash_line"]:
                            cv.create_line(gx, gy, gx, mt + ph, fill=mark_color,
                                           dash=(3, 3))
                        if pk.get("manual"):
                            cv.create_rectangle(gx - 4, gy - 4, gx + 4, gy + 4,
                                                fill="#1e6ec8", outline="")
                        else:
                            cv.create_oval(gx - 3, gy - 3, gx + 3, gy + 3,
                                           fill="#c83232", outline="")
                        text = ("%g  %.0f%%" % (round(pk["x"], 1), pk["rel"])
                                if p["peak_label_rel"] else "%g" % round(pk["x"], 1))
                        ly = gy - (9, 25, 41)[k % 3]
                        if p["peak_labels"]:
                            cv.create_text(gx, ly, text=text, fill=mark_color,
                                           font=("Microsoft YaHei UI", 8), anchor="s")

            legend_rows, legend_hidden, legend_box = legend_place(
                entries, legend_pos, ml, mt, pw, ph, xy=p.get("legend_xy"),
                avail=l_avail)
            if legend_rows:
                if legend_pos == "inside" and legend_box:
                    cv.create_rectangle(*legend_box, fill="#ffffff", outline="#d8d8d8")
                for text, color, cx, cy in legend_rows:
                    cv.create_line(cx, cy, cx + LEGEND_SWATCH, cy, fill=color, width=2)
                    cv.create_text(cx + LEGEND_SWATCH + LEGEND_TEXT_GAP, cy, text=text,
                                   fill="#333", anchor="w",
                                   font=("Microsoft YaHei UI", 8))
                if legend_hidden > 0 and legend_box:
                    cv.create_text(legend_box[0] + LEGEND_PAD,
                                   legend_box[3] - LEGEND_ROW_H // 2,
                                   text=T("…还有 %d 条") % legend_hidden,
                                   fill="#963c3c", anchor="w",
                                   font=("Microsoft YaHei UI", 8))

            if title and p["show_title"]:
                # top 位置时图例占了标题下面那一带，标题留在最上面不跟图例叠
                ty = mt_base - 20 if (legend_pos == "top" and legend_rows) else mt - 20
                cv.create_text(ml + pw / 2, ty, text=title, fill="#222",
                               font=("Microsoft YaHei UI", 10, "bold"))
            if xlabel:
                cv.create_text(ml + pw / 2, mt + ph + 34, text=xlabel, fill="#333",
                               font=("Microsoft YaHei UI", 9))
            if ylabel:
                cv.create_text(14, mt + ph / 2, text=ylabel, fill="#333", angle=90,
                               font=("Microsoft YaHei UI", 9))

        def _out_dir(self):
            if self.same_dir.get():
                return None
            return self.dir_var.get().strip() or None

        def start_convert(self):
            if self.worker and self.worker.is_alive():
                return
            if not self.files:
                messagebox.showwarning("没有文件", "请先添加要转换的 .jws 文件或文件夹。")
                return
            out_dir = self._out_dir()
            if not self.same_dir.get() and not out_dir:
                messagebox.showwarning("未设置输出目录",
                                       "请选择输出文件夹，或勾选“保存到源文件所在目录”。")
                return
            formats = []
            if self.fmt_excel.get():
                formats.append("xlsx")
            if self.fmt_png.get():
                formats.append("png")
            if self.fmt_csv.get():
                formats.append("csv")
            if self.fmt_peaks.get():
                formats.append("peaks")
            if self.fmt_fit.get():
                formats.append("fit")
            if not formats:
                messagebox.showwarning("未选择输出内容",
                                       "请至少勾选一种输出：Excel / PNG / CSV / 峰列表。")
                return
            opts = dict(
                out_dir=out_dir,
                x_header=self.x_var.get().strip() or _DEFAULT_X_HEADER,
                y_header=self.y_var.get().strip() or _DEFAULT_Y_HEADER,
                write_header=self.write_header.get(),
                skip_existing=self.skip_existing.get(),
                auto_names=self.auto_names.get(),
                formats=tuple(formats),
                plot=self.plot_options(),
            )
            self.run_btn.configure(state="disabled")
            self.progress.configure(maximum=len(self.files), value=0)
            self._log("=" * 48)
            peak_map = {}
            for p in self.files:
                ov = self._peak_overrides(p)
                if ov:
                    peak_map[self._peak_key(p)] = ov
            self.worker = threading.Thread(
                target=self._work, args=(list(self.files), opts, peak_map),
                daemon=True)
            self.worker.start()

        def _work(self, files, opts, peak_map):
            ok, skip, fail = 0, 0, 0
            for i, src in enumerate(files, 1):
                try:
                    ov = peak_map.get(self._peak_key(src)) or {}
                    if src.lower().endswith(".csv"):
                        plot = opts.get("plot")
                        if ov:
                            plot = dict(plot or {})
                            plot.update(ov)
                        dst, nser = convert_csv_to_png(src, out_dir=opts.get("out_dir"),
                                                       plot=plot)
                        msg = "(%d/%d) ✔ %s  →  %s   [CSV→PNG, %d 列]" % (
                            i, len(files), os.path.basename(src), os.path.basename(dst), nser)
                        ok += 1
                    else:
                        file_opts = opts
                        if ov:
                            file_opts = dict(opts)
                            plot = dict(opts.get("plot") or {})
                            plot.update(ov)
                            file_opts["plot"] = plot
                        dst, spec, skipped = convert_file(src, **file_opts)
                        if skipped:
                            msg = "(%d/%d) – 已存在，跳过：%s" % (i, len(files),
                                                            os.path.basename(dst))
                            skip += 1
                        else:
                            fmt_txt = "+".join(sorted(opts.get("formats", ("csv",))))
                            msg = "(%d/%d) ✔ %s  →  %s   [输出 %s, %d 通道, %d 点, %.3f~%.3f]" % (
                                i, len(files), os.path.basename(src), os.path.basename(dst),
                                fmt_txt, spec.channel_number, spec.npoints,
                                spec.start, spec.end)
                            ok += 1
                except Exception as exc:
                    msg = "(%d/%d) ✘ %s  →  %s" % (i, len(files), os.path.basename(src), exc)
                    if not isinstance(exc, JwsError):
                        msg += "\n        " + traceback.format_exc().strip().splitlines()[-1]
                    fail += 1
                self.q.put(("log", msg))
                self.q.put(("progress", i))
            self.q.put(("done", (ok, skip, fail)))

        def _poll(self):
            try:
                while True:
                    kind, payload = self.q.get_nowait()
                    if kind == "log":
                        self._log(payload)
                    elif kind == "progress":
                        self.progress.configure(value=payload)
                    elif kind == "batchprog":
                        i, n = payload
                        self.progress.configure(maximum=n, value=i)
                    elif kind == "batchdone":
                        if payload is None:
                            messagebox.showwarning("批处理", "批处理中断，详见日志。")
                            continue
                        res = payload
                        self._log("-" * 48)
                        self._log(T("批处理完成：共 %d 个，成功 %d，跳过 %d，失败 %d")
                                  % (res["total"], res["ok"], res["skip"], res["fail"]))
                        for name, msg in res["errors"][:10]:
                            self._log("   ✘ %s：%s" % (name, msg))
                        if res["summary"]:
                            self._log(T("   汇总统计：%s") % res["summary"])
                        for r in res["rows"][:10]:
                            self._log(T("   %-34s 峰数 %-4d 主峰 %s")
                                      % (r["file"][:34], r["peaks"],
                                         "" if r["x"] is None else "%.1f cm-1" % r["x"]))
                        if res["fail"] == 0:
                            messagebox.showinfo(
                                "批处理完成",
                                "成功 %d 个，跳过 %d 个。\n汇总统计：%s"
                                % (res["ok"], res["skip"], res["summary"] or "（无）"))
                        else:
                            messagebox.showwarning(
                                "批处理完成",
                                "成功 %d 个，跳过 %d 个，失败 %d 个。\n详见日志。"
                                % (res["ok"], res["skip"], res["fail"]))
                    elif kind == "done":
                        ok, skip, fail = payload
                        self._log("-" * 48)
                        self._log(T("全部完成：成功 %d 个，跳过 %d 个，失败 %d 个。") % (ok, skip, fail))
                        self.run_btn.configure(state="normal")
                        if fail == 0:
                            messagebox.showinfo("完成", "转换完成！成功 %d 个，跳过 %d 个。"
                                                % (ok, skip))
                        else:
                            messagebox.showwarning("完成", "成功 %d 个，跳过 %d 个，失败 %d 个。\n详见日志。"
                                                   % (ok, skip, fail))
            except queue.Empty:
                pass
            self.after(100, self._poll)

        def _open_folder(self, path):
            if path and os.path.isdir(path):
                try:
                    os.startfile(path)
                    return
                except OSError:
                    pass
            messagebox.showinfo("提示", "目录不存在：%s" % path)

        def open_data_manager(self):
            win = tk.Toplevel(self)
            win.title("数据文件夹 · 位置 / 占用 / 清理")
            win.transient(self)
            win.resizable(False, False)

            root_var = tk.StringVar()
            info_var = tk.StringVar()
            ttk.Label(win, text="工具自身的数据（下载的参考谱、RRUFF 数据包、分析结果）统一保存在：",
                      font=("Microsoft YaHei UI", 10, "bold")).grid(
                row=0, column=0, columnspan=4, sticky="w", padx=12, pady=(12, 2))
            ttk.Label(win, textvariable=root_var, foreground="#1a4a8a", wraplength=560,
                      justify="left").grid(row=1, column=0, columnspan=4, sticky="w", padx=12)
            ttk.Label(win, textvariable=info_var, foreground="#444").grid(
                row=2, column=0, columnspan=4, sticky="w", padx=12, pady=(2, 2))

            cache_var = tk.StringVar()
            ttk.Label(win, textvariable=cache_var, foreground="#444").grid(
                row=3, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 4))
            lim_row = ttk.Frame(win)
            lim_row.grid(row=4, column=0, columnspan=4, sticky="w", padx=12, pady=(0, 10))
            ttk.Label(lim_row, text="数据包容量上限(MB，0=不限制)：").pack(side="left")
            limit_var = tk.StringVar(value="%.0f" % cache_limit_mb())
            ttk.Entry(lim_row, textvariable=limit_var, width=8).pack(side="left", padx=4)

            def save_limit():
                try:
                    set_cache_limit_mb(float(limit_var.get() or 0))
                except ValueError:
                    messagebox.showwarning("提示", "容量上限必须是数字（MB，0 = 不限制）。")
                    return
                self._log(T("数据包容量上限已设为 %s MB。") % limit_var.get())
                refresh()

            def cleanup_dialog():
                pkgs = package_sizes()
                if not pkgs:
                    messagebox.showinfo("提示", "目前没有已下载的数据包。")
                    return
                listing = "\n".join("   %-22s %8.1f MB" % (key, sz / 1048576.0)
                                    for key, _path, sz in pkgs)
                total = sum(sz for _k, _p, sz in pkgs) / 1048576.0
                if not messagebox.askyesno(
                        "清理数据包",
                        "将删除以下已下载数据包（含索引），共 %.0f MB：\n%s\n\n"
                        "本地参考谱库（已导出的 CSV）不受影响。\n确定继续吗？"
                        % (total, listing)):
                    return
                removed, freed = cleanup_packages()
                self._log(T("已清理 %d 个数据包，释放 %.1f MB。") % (removed, freed / 1048576.0))
                refresh()

            ttk.Button(lim_row, text="保存上限", command=save_limit).pack(side="left", padx=4)
            ttk.Button(lim_row, text="清理已下载数据包…", command=cleanup_dialog).pack(
                side="left", padx=4)

            stat_vars = []
            for i, (label, path) in enumerate(data_folders()):
                row = 6 + i
                ttk.Label(win, text=label + "：").grid(row=row, column=0, sticky="w", padx=12, pady=3)
                var = tk.StringVar()
                stat_vars.append((path, var))
                ttk.Label(win, textvariable=var, width=24, foreground="#333").grid(
                    row=row, column=1, sticky="w")
                ttk.Button(win, text="打开",
                           command=lambda p=path: self._open_folder(p)).grid(
                    row=row, column=2, padx=4)
                ttk.Button(win, text="清空",
                           command=lambda l=label, p=path: clear_folder(l, p)).grid(
                    row=row, column=3, padx=4)

            def refresh():
                root_var.set(data_root())
                total_n = 0
                total_s = 0
                for path, var in stat_vars:
                    n, s = dir_stats(path)
                    total_n += n
                    total_s += s
                    var.set(T("%d 个文件 / %.1f MB") % (n, s / 1048576.0))
                cache_n, cache_s = cache_usage()
                limit = cache_limit_mb()
                over = bool(limit) and cache_s > limit * 1048576.0
                free, total = disk_free()
                free_txt = "" if free is None else T("　磁盘可用 %.1f GB / 共 %.1f GB") % (
                    free / 1073741824.0, total / 1073741824.0)
                info_var.set(T("合计 %d 个文件，占用 %.1f MB%s") % (
                    total_n, total_s / 1048576.0, free_txt))
                cache_var.set(T("数据包 %d 个 / %.1f MB　上限 %s%s") % (
                    cache_n, cache_s / 1048576.0,
                    "不限制" if not limit else "%.0f MB" % limit,
                    "（已超限，建议清理）" if over else ""))

            def clear_folder(label, path):
                if not messagebox.askyesno(
                        "确认清理",
                        "确定要清空「%s」吗？\n%s\n\n此操作不可撤销。" % (label, path)):
                    return
                removed = 0
                try:
                    names = os.listdir(path)
                except OSError as exc:
                    messagebox.showerror("清理失败", str(exc))
                    return
                for name in names:
                    target = os.path.join(path, name)
                    try:
                        if os.path.isdir(target):
                            shutil.rmtree(target)
                        else:
                            os.remove(target)
                        removed += 1
                    except OSError as exc:
                        self._log(T("清理失败 %s：%s") % (name, exc))
                self._log(T("已清空「%s」：删除 %d 项。") % (label, removed))
                refresh()

            def change_root():
                folder = filedialog.askdirectory(title=T("选择新的数据文件夹（建议放在本工具目录内）"))
                if not folder:
                    return
                new_root = os.path.join(folder, _DATA_ROOT_NAME)
                if os.path.exists(new_root) and os.listdir(new_root):
                    new_root = folder
                move = messagebox.askyesno(
                    "是否搬移已有数据",
                    "是否把现有数据一并移动到新位置？\n（数据包较大时可能要一会儿）")
                if move:
                    for _label, old in data_folders():
                        dst = os.path.join(new_root, os.path.basename(old))
                        os.makedirs(dst, exist_ok=True)
                        try:
                            names = os.listdir(old)
                        except OSError:
                            continue
                        for name in names:
                            src = os.path.join(old, name)
                            target = os.path.join(dst, name)
                            if os.path.exists(target):
                                continue
                            try:
                                shutil.move(src, target)
                            except OSError as exc:
                                self._log(T("搬移失败 %s：%s") % (name, exc))
                set_data_root(new_root)
                self._log(T("数据文件夹已切换到：%s") % data_root())
                refresh()

            ttk.Label(win, text="提示：默认数据文件夹就在工具（exe）所在目录；若该目录不可写，会自动改用系统用户目录。"
                                "下载与分析结果都不会写到别的位置；数据包超过上限时会有提醒。",
                      foreground="#888", wraplength=580, justify="left").grid(
                row=9, column=0, columnspan=4, sticky="w", padx=12, pady=(8, 2))

            btns = ttk.Frame(win)
            btns.grid(row=10, column=0, columnspan=4, sticky="we", padx=12, pady=(4, 12))
            ttk.Button(btns, text="打开数据文件夹",
                       command=lambda: self._open_folder(data_root())).pack(side="left")
            ttk.Button(btns, text="更改位置…", command=change_root).pack(side="left", padx=6)
            ttk.Button(btns, text="刷新", command=refresh).pack(side="left", padx=6)
            ttk.Button(btns, text="关闭", command=win.destroy).pack(side="right")
            refresh()

        def _dl_status(self, label, got, total, speed, eta):
            """下载状态行：名称 · 已下/总数 · 速率 · 剩余时间。"""
            parts = [label]
            if total:
                parts.append("%s / %s" % (format_mb(got), format_mb(total)))
            else:
                parts.append(format_mb(got))
            if speed and speed > 0:
                parts.append("%.2f MB/s" % (speed / 1048576.0))
            if eta:
                parts.append(T("剩余 %s") % format_eta(eta))
            return " · ".join(parts)

        def _bg_run(self, win, work, status, bar=None, cancel_btn=None,
                    on_done=None, on_error=None, interval=200, busy_text=""):
            """把耗时工作放后台线程跑，主线程刷新状态 / 进度条 / 取消按钮。

            work(report, cancel)：report(text=, got=, total=, speed=, eta=) 更新进度，
            cancel 是 threading.Event。这样下载十几分钟窗口也不会假死。
            """
            state = {"text": busy_text, "got": 0, "total": 0, "speed": 0.0,
                     "eta": None, "done": False, "result": None, "error": None,
                     "cancelled": False, "cancel": threading.Event()}
            lock = threading.Lock()

            def report(text=None, got=None, total=None, speed=None, eta=None):
                with lock:
                    if text is not None:
                        state["text"] = text
                    if got is not None:
                        state["got"] = got
                    if total is not None:
                        state["total"] = total
                    if speed is not None:
                        state["speed"] = speed
                    if eta is not None:
                        state["eta"] = eta

            def runner():
                try:
                    state["result"] = work(report, state["cancel"])
                except DownloadCancelled:
                    state["cancelled"] = True
                except Exception as exc:
                    state["error"] = exc
                finally:
                    state["done"] = True

            def poll():
                if not win.winfo_exists():
                    return
                with lock:
                    snap = dict(state)
                if cancel_btn is not None:
                    cancel_btn.state(["disabled"] if snap["done"] else ["!disabled"])
                if snap["text"]:
                    status.set(snap["text"])
                if bar is not None:
                    if snap["total"]:
                        if str(bar.cget("mode")) != "determinate":
                            bar.stop()
                            bar.configure(mode="determinate")
                        bar.configure(value=int(1000.0 * snap["got"] / max(1, snap["total"])))
                    elif not snap["done"]:
                        if str(bar.cget("mode")) != "indeterminate":
                            bar.configure(mode="indeterminate")
                            bar.start(12)
                if not snap["done"]:
                    win.after(interval, poll)
                    return
                if bar is not None:
                    bar.stop()
                    bar.configure(mode="determinate", value=0)
                if snap["cancelled"]:
                    status.set(T("已取消。已下载的部分会保留，下次接着下。"))
                    if on_error:
                        on_error(DownloadCancelled())
                    return
                if snap["error"] is not None:
                    if on_error:
                        on_error(snap["error"])
                    return
                if on_done:
                    on_done(snap["result"])

            if cancel_btn is not None:
                cancel_btn.configure(command=state["cancel"].set)
            threading.Thread(target=runner, daemon=True).start()
            win.after(interval, poll)

        def _activity_bar(self, parent):
            """对话框底部的「进度条 + 取消」栏，返回 (frame, bar, cancel_btn)。"""
            row = ttk.Frame(parent)
            bar = ttk.Progressbar(row, mode="determinate", maximum=1000, length=380)
            bar.pack(side="left", fill="x", expand=True, padx=(0, 8))
            cbtn = ttk.Button(row, text=T("取消"))
            cbtn.pack(side="left")
            cbtn.state(["disabled"])
            return row, bar, cbtn

        def open_db_search(self):
            win = tk.Toplevel(self)
            win.title("在线检索数据库 · ROD（RRUFF 拉曼数据的官方开放库）")
            win.transient(self)
            win.geometry("880x480")

            top = ttk.Frame(win)
            top.pack(fill="x", padx=10, pady=8)
            ttk.Label(top, text="按矿物名 / 化学式检索：").pack(side="left")
            query = tk.StringVar(value="")
            entry = ttk.Entry(top, textvariable=query, width=26)
            entry.pack(side="left", padx=6)
            entry.focus_set()

            cols = ("mineral", "formula", "chem", "wave", "device", "id")
            heads = ("矿物", "化学式", "化学名称", "波长(nm)", "仪器", "编号")
            widths = (150, 120, 170, 80, 190, 90)
            tree = ttk.Treeview(win, columns=cols, show="headings", selectmode="extended")
            for col, head, width in zip(cols, heads, widths):
                tree.heading(col, text=T(head))
                tree.column(col, width=width, anchor="w")
            tree.pack(fill="both", expand=True, padx=10, pady=(0, 4))

            status = tk.StringVar(value=T("数据来源：ROD（Raman Open Database，RRUFF 拉曼数据）。")
                                        + T("下载后统一存到工具的数据文件夹（参考谱库），可离线使用。"))
            ttk.Label(win, textvariable=status, foreground="#555", wraplength=840,
                      justify="left").pack(anchor="w", padx=10)

            def do_search():
                key = query.get().strip()
                if not key:
                    status.set(T("请输入检索词，例如 quartz / SiO2 / 石英"))
                    return

                def work(report, cancel):
                    report(text=T("正在检索 …  这一步要连 solsa.crystallography.net"),
                           got=0, total=0)
                    return rod_search(key)

                def done(rows):
                    tree.delete(*tree.get_children())
                    for row in rows:
                        tree.insert("", "end", values=(
                            row.get("mineral") or "", row.get("formula") or "",
                            row.get("chemname") or "", row.get("wavelength") or "",
                            row.get("devicecompany") or "", row.get("file") or ""))
                    status.set(T("共 %d 条结果。按住 Ctrl 多选后点“下载选中谱线”。")
                               % len(rows))

                def failed(exc):
                    status.set(T("检索失败：%s（请检查网络）") % exc)

                before_busy()
                self._bg_run(win, work, status, bar=bar, cancel_btn=cancel_btn,
                             on_done=done, on_error=failed,
                             busy_text=T("正在检索 …"))

            def _download_items(items, tag, specs=None):
                """把选中的数据库谱线逐条下回来；可在下载途中取消。"""
                def work(report, cancel):
                    paths = []
                    failed_n = 0
                    for k, item in enumerate(items):
                        if cancel.is_set():
                            raise DownloadCancelled()
                        values = tree.item(item, "values")
                        rod_id, mineral, wave = values[5], values[0], values[3]
                        report(text=T("正在下载 %d/%d：%s") % (k + 1, len(items), rod_id),
                               got=k + 1, total=len(items))
                        try:
                            dst, _meta = db_download(rod_id, mineral, wave)
                            paths.append(dst)
                        except Exception as exc:
                            failed_n += 1
                            self._log(T("下载失败 %s：%s") % (rod_id, exc))
                    return paths, failed_n

                def done(result):
                    paths, failed_n = result
                    after_busy()
                    if not paths:
                        status.set(T("下载失败，没有取到任何谱线。"))
                        return
                    if tag == "pair":
                        if not specs:
                            status.set(T("没有可配对的实测谱，已下载的谱线存在本地库。"))
                            return
                        name, xs, ys = specs[0]
                        self._do_pair(name, (xs, ys), paths,
                                      "手动配对（数据库 %d 条）" % len(paths))
                        status.set(T("已完成与 %d 条数据库谱线的配对，详见主界面日志与右侧叠加图。")
                                   % len(paths))
                        return
                    msg = T("已下载 %d 条到：%s") % (len(paths), db_dir())
                    if failed_n:
                        msg += T("（%d 条失败，详见日志）") % failed_n
                    status.set(msg)

                def failed(exc):
                    after_busy()
                    if isinstance(exc, DownloadCancelled):
                        self._log(T("下载已取消。"))
                        return
                    status.set(T("下载失败：%s") % exc)

                before_busy()
                self._bg_run(win, work, status, bar=bar, cancel_btn=cancel_btn,
                             on_done=done, on_error=failed,
                             busy_text=T("正在连接数据库 …"))

            def do_download():
                items = tree.selection()
                if not items:
                    status.set(T("请先在列表中选中条目。"))
                    return
                _download_items(list(items), "save")

            def do_add_local():
                files = db_list()
                if not files:
                    status.set(T("本地数据库还是空的，请先下载谱线。"))
                    return
                added = self._add_paths(files)
                if added:
                    self._select_first_new()
                self._log(T("已从本地数据库加入 %d 个文件到列表。") % added)
                status.set(T("已加入 %d 个文件到左侧列表，可绘图 / 比对 / 平均 / 相减。") % added)

            def do_pair_selected():
                items = tree.selection()
                if not items:
                    status.set(T("请先在列表中选中条目。"))
                    return
                specs = self._selected_spectra()
                if not specs:
                    status.set(T("请先关闭本窗口，在主界面左侧选择 1 个实测谱。"))
                    return
                _download_items(list(items), "pair", specs)

            def after_busy():
                for b in odb_btns:
                    b.state(["!disabled"])

            def before_busy():
                for b in odb_btns:
                    b.state(["disabled"])

            sbtn = ttk.Button(top, text="检索", command=do_search)
            sbtn.pack(side="left", padx=4)
            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=(4, 2))
            odb_btns = [sbtn]
            for text, cmd, pad in (("下载选中谱线", do_download, 0),
                                   ("下载并与实测谱配对", do_pair_selected, 6),
                                   ("把本地库加入文件列表", do_add_local, 6)):
                b = ttk.Button(btns, text=text, command=cmd)
                b.pack(side="left", padx=(0, pad)) if pad else b.pack(side="left")
                odb_btns.append(b)
            ttk.Button(btns, text="打开数据库文件夹",
                       command=lambda: self._open_folder(db_dir())).pack(side="left", padx=6)
            ttk.Button(btns, text="关闭", command=win.destroy).pack(side="right")
            abar, bar, cancel_btn = self._activity_bar(win)
            abar.pack(fill="x", padx=10, pady=(0, 10))
            entry.bind("<Return>", lambda _e: do_search())

        def open_rruff(self):
            win = tk.Toplevel(self)
            win.title("RRUFF 全量拉曼库 · 下载 / 按矿物检索")
            win.transient(self)
            win.geometry("940x640")

            status = tk.StringVar(
                value=T("第一步：下载一个数据包（自动建索引）。第二步：按矿物名或 RRUFF 编号检索。")
                      + T("第三步：导出到本地库，即可绘图 / 配对比较。")
                      + T("红外、XRD、化学成分包同样适用。"))
            ttk.Label(win, textvariable=status, foreground="#444", wraplength=900,
                      justify="left").pack(anchor="w", padx=10, pady=(8, 2))

            pf = ttk.LabelFrame(win, text="① 数据包（来源 www.rruff.net/zipped_data_files）")
            pf.pack(fill="x", padx=10, pady=4)
            ptree = ttk.Treeview(pf, columns=("state", "kind", "name", "size", "eta", "idx"),
                                 show="headings", height=min(12, len(all_packages()) + 1),
                                 selectmode="browse")
            for col, head, width in (("state", "状态", 80), ("kind", "类型", 60),
                                     ("name", "数据包", 300), ("size", "大小", 80),
                                     ("eta", "预计耗时", 110), ("idx", "索引条数", 90)):
                ptree.heading(col, text=T(head))
                ptree.column(col, width=width, anchor="w")
            ptree.pack(fill="x", padx=6, pady=6)
            pkg_info = tk.StringVar(value="")
            ttk.Label(pf, textvariable=pkg_info, foreground="#666").pack(anchor="w", padx=8)

            def refresh_pkgs():
                ptree.delete(*ptree.get_children())
                for pkg in all_packages():
                    key = pkg["key"]
                    have = os.path.exists(rruff_zip_path(key))
                    idx = _read_index(key)
                    eta = format_eta(estimate_download_seconds(package_bytes(pkg))) or "-"
                    ptree.insert("", "end", iid=key, values=(
                        T("已下载") if have else T("未下载"), T(pkg.get("kind", "")),
                        "%s  (%s)" % (pkg["label"], key), pkg["size"], eta,
                        len(idx) if idx else "-"))
                n, s = cache_usage()
                free, total = disk_free()
                limit = cache_limit_mb()
                extra = "" if free is None else T("　磁盘可用 %.1f GB") % (free / 1073741824.0)
                if limit and s > limit * 1048576.0:
                    extra += T("　⚠ 已超过容量上限 %.0f MB，建议清理") % limit
                pkg_info.set(T("数据包占用 %.1f MB / %d 个文件%s") % (s / 1048576.0, n, extra)
                             + T("　（预计耗时按上次实测速率 %.2f MB/s 估算）")
                             % (download_speed_bps() / 1048576.0)
                             + T("　并发 %d 路 · %s")
                             % (download_connections(),
                                (T("代理 %s") % proxy_url()) if proxy_url()
                                else T("直连")))

            def after_busy():
                refresh_pkgs()
                for b in pkg_btns:
                    b.state(["!disabled"])

            def before_busy():
                for b in pkg_btns:
                    b.state(["disabled"])

            def dl_failed(exc):
                after_busy()
                if isinstance(exc, DownloadCancelled):
                    self._log(T("下载已取消。"))
                    return
                status.set(T("下载失败：%s") % exc)
                self._log(T("下载失败：%s") % exc)

            def do_download():
                sel = ptree.selection()
                if not sel:
                    status.set(T("请先选中一个数据包。"))
                    return
                save_net()                 # 把刚才改的并发 / 代理落盘，本次下载就生效
                key = sel[0]
                pkg = rruff_package(key)
                size = package_bytes(pkg)
                if size and size >= 50 * 1048576:
                    eta = format_eta(estimate_download_seconds(size))
                    if not messagebox.askyesno(
                            T("这个包比较大"),
                            T("%s 约 %s，按实测速率预计要 %s。\n\n"
                              "下载在后台进行，窗口不会卡住，随时可以点【取消】；"
                              "中断后已下载的部分会保留，下次接着下，不会从头再来。\n\n"
                              "现在开始下载吗？")
                            % (pkg["label"], format_mb(size), eta)):
                        return
                existing = rruff_zip_path(key)
                if os.path.exists(existing) and not messagebox.askyesno(
                        T("确认覆盖"), T("%s 已经下载过了，要重新下载吗？") % pkg["label"]):
                    return

                def work(report, cancel):
                    def prog(got, total, speed, eta):
                        report(text=self._dl_status(
                            T("正在下载 %s") % pkg["label"], got, total, speed, eta),
                            got=got, total=total, speed=speed, eta=eta)
                    rruff_download(key, prog, cancel=cancel)
                    report(text=T("下载完成，正在建立索引（谱线较多时需一会儿）…"),
                           got=0, total=0, speed=0.0)
                    return rruff_index(key, rebuild=True)

                def done(rows):
                    after_busy()
                    status.set(T("完成：%s 共 %d 条，现在可以检索了。")
                               % (pkg["label"], len(rows or [])))
                    self._log(T("数据包 %s 下载完成，索引 %d 条")
                              % (pkg["label"], len(rows or [])))

                before_busy()
                self._bg_run(win, work, status, bar=bar, cancel_btn=cancel_btn,
                             on_done=done, on_error=dl_failed,
                             busy_text=T("正在连接 %s …") % pkg["label"])

            def do_import():
                path = filedialog.askopenfilename(
                    title=T("选择自己下载的 RRUFF 数据包 zip"),
                    filetypes=[(T("ZIP 压缩包"), "*.zip"), (T("所有文件"), "*.*")])
                if not path:
                    return

                def work(report, cancel):
                    report(text=T("正在导入并建立索引 …"), got=0, total=0)
                    return rruff_import_zip(path)

                def done(result):
                    after_busy()
                    key, count = result
                    status.set(T("导入成功：%s（索引 %d 条）。") % (key, count))
                    self._log(T("已导入 RRUFF 数据包：%s（索引 %d 条）") % (key, count))

                def failed(exc):
                    after_busy()
                    status.set(T("导入失败：%s") % exc)

                before_busy()
                self._bg_run(win, work, status, bar=bar, cancel_btn=cancel_btn,
                             on_done=done, on_error=failed,
                             busy_text=T("正在导入 …"))

            def do_index():
                sel = ptree.selection()
                if not sel:
                    return
                key = sel[0]

                def work(report, cancel):
                    def prog(i, n):
                        report(text=T("正在建立索引 %d/%d …") % (i + 1, n),
                               got=i + 1, total=n)
                    return rruff_index(key, progress=prog, rebuild=True)

                def done(rows):
                    after_busy()
                    status.set(T("索引完成：%d 条。") % len(rows or []))

                before_busy()
                self._bg_run(win, work, status, bar=bar, cancel_btn=cancel_btn,
                             on_done=done, on_error=lambda exc: (
                                 after_busy(),
                                 status.set(T("索引失败：%s") % exc)),
                             busy_text=T("正在建立索引 %s …") % key)

            pbtn = ttk.Frame(pf)
            pbtn.pack(fill="x", padx=6, pady=(0, 6))
            pkg_btns = []
            for text, cmd, pad in (("下载选中数据包并建索引", do_download, 0),
                                   ("重建索引", do_index, 6),
                                   ("导入本地 zip…", do_import, 6)):
                b = ttk.Button(pbtn, text=text, command=cmd)
                b.pack(side="left", padx=(0, pad)) if pad else b.pack(side="left")
                pkg_btns.append(b)
            ttk.Button(pbtn, text="清理数据包…",
                       command=lambda: self.open_data_manager()).pack(side="left", padx=6)

            netf = ttk.Frame(pf)
            netf.pack(fill="x", padx=6, pady=(0, 4))
            ttk.Label(netf, text=T("并发连接数：")).pack(side="left")
            conn_var = tk.StringVar(value=str(download_connections()))
            conn_entry = ttk.Spinbox(netf, from_=1, to=DOWNLOAD_CONNECTIONS_MAX, width=5,
                                     textvariable=conn_var)
            conn_entry.pack(side="left")
            ttk.Label(netf, text=T("（1~%d，越大越快）") % DOWNLOAD_CONNECTIONS_MAX,
                      foreground="#888").pack(side="left", padx=4)
            ttk.Label(netf, text=T("代理：")).pack(side="left", padx=(10, 2))
            proxy_var = tk.StringVar(value=str(_load_settings().get("proxy") or ""))
            proxy_entry = ttk.Entry(netf, textvariable=proxy_var, width=20)
            proxy_entry.pack(side="left")
            ttk.Label(netf, text=T("（留空 = 自动用系统代理；none = 强制直连）"),
                      foreground="#888").pack(side="left", padx=4)

            def save_net():
                data = _load_settings()
                try:
                    value = int(float(conn_var.get()))
                    if 1 <= value <= DOWNLOAD_CONNECTIONS_MAX:
                        data["download_conns"] = str(value)
                except (TypeError, ValueError):
                    pass
                px = proxy_var.get().strip()
                if px:
                    data["proxy"] = px
                else:
                    data.pop("proxy", None)
                _save_settings(data)
                refresh_pkgs()

            for _w in (conn_entry, proxy_entry):
                _w.bind("<Return>", lambda _e: save_net())
                _w.bind("<FocusOut>", lambda _e: save_net())

            abar, bar, cancel_btn = self._activity_bar(pf)
            abar.pack(fill="x", padx=6, pady=(0, 6))

            sf = ttk.LabelFrame(win, text="② 按矿物名 / RRUFF 编号检索（需先下载任一数据包）")
            sf.pack(fill="both", expand=True, padx=10, pady=4)
            srow = ttk.Frame(sf)
            srow.pack(fill="x", padx=6, pady=6)
            ttk.Label(srow, text="检索：").pack(side="left")
            query = tk.StringVar(value="")
            entry = ttk.Entry(srow, textvariable=query, width=24)
            entry.pack(side="left", padx=4)
            rtree = ttk.Treeview(sf, columns=("min", "id", "wave", "type", "pkg"),
                                 show="headings", height=9, selectmode="extended")
            for col, head, width in (("min", "矿物", 150), ("id", "编号", 90),
                                     ("wave", "波长(nm)", 80), ("type", "类型", 140),
                                     ("pkg", "数据包", 160)):
                rtree.heading(col, text=T(head))
                rtree.column(col, width=width, anchor="w")
            rtree.pack(fill="both", expand=True, padx=6, pady=(0, 6))
            cache = {"rows": []}

            def do_search():
                key = query.get().strip()
                if not key:
                    status.set(T("请输入矿物名（如 Zircon / Quartz）或 RRUFF 编号（如 R050034）。"))
                    return
                rows = rruff_search(key)
                cache["rows"] = rows
                rtree.delete(*rtree.get_children())
                for i, row in enumerate(rows):
                    rtree.insert("", "end", iid=str(i),
                                 values=(row[1], row[2], row[3], row[4], row[0]))
                status.set(T("检索到 %d 条。") % len(rows))

            def picked():
                out = []
                for iid in rtree.selection():
                    try:
                        out.append(cache["rows"][int(iid)])
                    except (ValueError, IndexError):
                        continue
                return out

            def do_export():
                sel = picked()
                if not sel:
                    status.set(T("请先在检索结果里选中条目（可 Ctrl/Shift 多选）。"))
                    return
                grouped = {}
                for row in sel:
                    grouped.setdefault(row[0], []).append(row[5])
                saved = 0
                for key, entries in grouped.items():
                    status.set(T("正在导出 %d 条 …") % len(entries))
                    win.update_idletasks()
                    try:
                        saved += len(rruff_export(key, entries))
                    except Exception as exc:
                        status.set(T("导出失败：%s") % exc)
                        return
                status.set(T("已导出 %d 条到本地库：%s") % (saved, db_dir()))

            def do_export_pair():
                sel = picked()
                specs = self._selected_spectra()
                if not specs:
                    status.set(T("请先关闭本窗口，在主界面左侧选择 1 个实测谱。"))
                    return
                if not sel:
                    status.set(T("请先在检索结果里选中条目。"))
                    return
                paths = []
                for row in sel:
                    status.set(T("正在导出 %s …") % row[2])
                    win.update_idletasks()
                    try:
                        paths.extend(rruff_export(row[0], [row[5]]))
                    except Exception as exc:
                        self._log(T("导出失败 %s：%s") % (row[2], exc))
                if not paths:
                    status.set(T("没有导出成功，无法配对。"))
                    return
                name, xs, ys = specs[0]
                self._do_pair(name, (xs, ys), paths, "RRUFF 配对（%d 条）" % len(paths))
                status.set(T("已用 %d 条 RRUFF 参考谱完成配对，见主界面日志与报告图。") % len(paths))

            def do_fetch():
                mineral = query.get().strip()
                if not mineral:
                    status.set(T("请输入矿物名，例如 Zircon / Quartz。"))
                    return
                save_net()

                def plan():
                    """主线程里定好要下哪些包（可能要弹窗确认），再交给后台跑。"""
                    done_saved, done_hits = rruff_export_matches(mineral)
                    if done_hits:
                        return None, done_saved, done_hits
                    if not rruff_available_keys():
                        if not messagebox.askyesno(
                                "需要先下载数据包",
                                "本地还没有 RRUFF 数据包。\n"
                                "是否现在下载最小的数据包（未评级·非定向，12 MB）？"):
                            return [], [], 0
                        return [rruff_package("unrated_unoriented")], [], 0
                    missing = [p for p in rruff_packages_sorted()
                               if not os.path.exists(rruff_zip_path(p["key"]))]
                    if not missing:
                        return [], [], 0
                    listing = "\n".join("   %s（%s）" % (p["label"], p["size"])
                                        for p in missing)
                    if not messagebox.askyesno(
                            "未找到，是否下载更大的数据包？",
                            "已下载的数据包里没有 \"%s\"。\n\n"
                            "可依次下载以下数据包并自动重试：\n%s" % (mineral, listing)):
                        return [], [], 0
                    return missing, [], 0

                todo, saved, hits = plan()
                if todo is None:
                    do_search()
                    status.set(T("已按矿物批量抓取 \"%s\"：命中 %d 条，导出 %d 条到本地库。")
                               % (mineral, hits, len(saved)))
                    return
                if not todo:
                    if not hits:
                        status.set(T("所有数据包中都没有找到 \"%s\"。") % mineral)
                    return

                def work(report, cancel):
                    total_mb = sum(package_bytes(p) for p in todo) / 1048576.0
                    report(text=T("批量抓取 \"%s\"：需下载 %d 个包，共约 %.0f MB")
                           % (mineral, len(todo), total_mb), got=0, total=0)
                    out_saved, out_hits = [], 0
                    for k, pkg in enumerate(todo):
                        if cancel.is_set():
                            raise DownloadCancelled()
                        report(text=T("正在下载 %s（第 %d/%d 个包，%s）")
                               % (pkg["label"], k + 1, len(todo), pkg["size"]),
                               got=0, total=0)

                        def prog(got, total, speed, eta, label=pkg["label"]):
                            report(text=self._dl_status(
                                T("正在下载 %s") % label, got, total, speed, eta),
                                got=got, total=total, speed=speed, eta=eta)

                        rruff_download(pkg["key"], prog, cancel=cancel)
                        report(text=T("正在建立索引 %s …") % pkg["label"],
                               got=0, total=0)
                        rruff_index(pkg["key"], rebuild=True)
                        out_saved, out_hits = rruff_export_matches(mineral)
                        if out_hits:
                            break
                    return out_saved, out_hits

                def done(result):
                    after_busy()
                    got_saved, got_hits = result
                    refresh_pkgs()
                    if got_hits:
                        do_search()
                        status.set(T("已按矿物批量抓取 \"%s\"：命中 %d 条，导出 %d 条到本地库。")
                                   % (mineral, got_hits, len(got_saved)))
                        self._log(T("RRUFF 按矿物抓取 \"%s\"：命中 %d 条，导出 %d 条到 %s")
                                  % (mineral, got_hits, len(got_saved), db_dir()))
                        for path in got_saved:
                            self._log("   " + os.path.basename(path))
                    else:
                        status.set(T("所有数据包中都没有找到 \"%s\"。") % mineral)

                def failed(exc):
                    after_busy()
                    refresh_pkgs()
                    if isinstance(exc, DownloadCancelled):
                        status.set(T("已取消。已下载的部分会保留，下次接着下。"))
                        self._log(T("按矿物抓取已取消。"))
                        return
                    status.set(T("抓取失败：%s") % exc)
                    self._log(T("抓取失败：%s") % exc)

                before_busy()
                self._bg_run(win, work, status, bar=bar, cancel_btn=cancel_btn,
                             on_done=done, on_error=failed,
                             busy_text=T("正在准备批量抓取 …"))

            ttk.Button(srow, text="检索", command=do_search).pack(side="left", padx=4)
            ttk.Button(srow, text="按矿物批量抓取（自动下载+导出）", command=do_fetch).pack(
                side="left", padx=4)
            entry.bind("<Return>", lambda _e: do_search())
            rbtn = ttk.Frame(sf)
            rbtn.pack(fill="x", padx=6, pady=(0, 8))
            ttk.Button(rbtn, text="导出选中到本地库", command=do_export).pack(side="left")
            ttk.Button(rbtn, text="导出并与实测谱配对", command=do_export_pair).pack(
                side="left", padx=6)
            ttk.Button(rbtn, text="打开本地库文件夹",
                       command=lambda: self._open_folder(db_dir())).pack(side="left")
            ttk.Button(rbtn, text="关闭", command=win.destroy).pack(side="right")
            refresh_pkgs()

        def db_compare(self):
            specs = self._selected_spectra()
            if not specs:
                messagebox.showinfo("提示", "请先在左侧选择 1 个待比对文件。")
                return
            files = db_list()
            if not files:
                messagebox.showinfo(
                    "提示",
                    "本地数据库为空。\n请先用“数据库 → 在线检索数据库…”下载参考谱线。")
                return
            name, xs, ys = specs[0]
            results = db_match((xs, ys), top=20)
            if not results:
                messagebox.showwarning("提示", "与本地数据库谱线没有重叠波数区间，无法比对。")
                return
            src = self.files[self.listbox.curselection()[0]]
            base = os.path.splitext(os.path.basename(src))[0]
            dst = os.path.join(self._tool_out_dir(src), base + "_数据库比对.csv")
            with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                f.write(T("排名,数据库谱线,相关系数,谱角(度)\n"))
                for i, r in enumerate(results, 1):
                    f.write("%d,%s,%.4f,%.2f\n" % (i, r["name"], r["corr"], r["angle"]))
            self._log(T("与本地数据库（%d 条）比对 → %s") % (len(files), os.path.basename(dst)))
            for r in results[:6]:
                self._log(T("   %-34s 相关系数 %.4f  谱角 %.2f°") % (
                    r["name"][:34], r["corr"], r["angle"]))

        def _show_image(self, path, title="预览"):
            try:
                from PIL import Image as _I
                from PIL import ImageTk
            except Exception:
                self._log(T("未安装 Pillow（或缺少 ImageTk），无法在窗口内预览：%s") % path)
                return
            try:
                img = _I.open(path)
            except Exception as exc:
                self._log(T("打开图片失败：%s") % exc)
                return
            win = tk.Toplevel(self)
            win.title(title)
            win.transient(self)
            scale = min(1.0, 1180.0 / float(img.width), 780.0 / float(img.height))
            if scale < 1.0:
                resample = _I.LANCZOS if hasattr(_I, "LANCZOS") else _I.BILINEAR
                img = img.resize((max(1, int(img.width * scale)),
                                  max(1, int(img.height * scale))), resample)
            photo = ImageTk.PhotoImage(img)
            cv = tk.Canvas(win, width=img.width, height=img.height, background="white",
                           highlightthickness=0)
            cv.pack()
            cv.create_image(0, 0, anchor="nw", image=photo)
            cv.image = photo
            row = ttk.Frame(win)
            row.pack(fill="x", pady=6)
            ttk.Button(row, text="打开所在文件夹",
                       command=lambda: self._open_folder(os.path.dirname(path))).pack(side="left",
                                                                                     padx=8)
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right", padx=8)

        def tool_cluster(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "请先在左侧选中至少 2 条光谱（可 Ctrl / Shift 多选）。")
                return
            cut = simpledialog.askstring(
                "聚类分析", "自动分割阈值（1 − 相关系数，留空 = 自动）：", parent=self)
            if cut is None:
                return
            cut_val = None
            if cut.strip():
                try:
                    cut_val = float(cut)
                except ValueError:
                    messagebox.showwarning("提示", "阈值必须是数字。")
                    return
            self._log(T("正在聚类分析（%d 条光谱，重采样 300 点）…") % len(specs))
            self.update_idletasks()
            try:
                res = cluster_spectra(specs, cut=cut_val)
            except Exception as exc:
                messagebox.showerror("聚类失败", str(exc))
                return
            if res is None:
                messagebox.showwarning("提示", "所选光谱没有共同的波数区间，无法聚类。")
                return
            src = self.files[self.listbox.curselection()[0]]
            out = self._tool_out_dir(src)
            csv_path = os.path.join(out, "聚类分析_分组.csv")
            with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(T("文件,簇号,PC1,PC2,PC3\n"))
                for i, name in enumerate(res["names"]):
                    sc = (res["scores"][i] if i < len(res["scores"]) else []) + [0.0, 0.0, 0.0]
                    f.write("%s,%d,%.4f,%.4f,%.4f\n" % (name, res["labels"][i],
                                                        sc[0], sc[1], sc[2]))
            tree = os.path.join(out, "聚类分析_树状图.png")
            render_dendrogram(tree, res["names"], res["merges"], "层次聚类树状图",
                              cut=res["cut"], groups=res["labels"])
            self._log(T("聚类完成：%d 条光谱 → %d 个簇（自动分割阈值 %.3f）")
                      % (len(res["names"]), len(res["groups"]), res["cut"]))
            for cid in sorted(res["groups"]):
                members = res["groups"][cid]
                shown = "、".join(res["names"][i] for i in members[:6])
                self._log(T("   簇%d（%d 条）：%s%s") % (cid, len(members), shown,
                                                 " …" if len(members) > 6 else ""))
            self._log(T("   分组表：%s") % os.path.basename(csv_path))
            self._log(T("   树状图：%s") % os.path.basename(tree))
            if res["scores"] and len(res["scores"][0]) >= 2:
                xlab, ylab = "PC1", "PC2"
                if res["explained"]:
                    xlab = "PC1（%.1f%%）" % res["explained"][0]
                    if len(res["explained"]) > 1:
                        ylab = "PC2（%.1f%%）" % res["explained"][1]
                png = os.path.join(out, "聚类分析_主成分散点.png")
                render_scatter(png, [(s[0], s[1]) for s in res["scores"]],
                               "主成分散点（PC1-PC2）", xlab, ylab,
                               groups=res["labels"], labels=res["names"])
                self._log(T("   主成分散点：%s") % os.path.basename(png))
            self._show_image(tree, T("聚类树状图"))

        def tool_map(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "二维成像需要选中多个点位的扫描光谱（按文件顺序排布）。")
                return
            n = len(specs)
            best = (n, 1)
            for r in range(2, int(math.sqrt(n)) + 1):
                if n % r == 0:
                    best = (r, n // r)
            win = tk.Toplevel(self)
            win.title("二维成像 · 点阵扫描热图")
            win.transient(self)
            win.resizable(False, False)
            ttk.Label(win, text="把选中的 %d 条光谱按“行优先”顺序铺成点阵。" % n,
                      font=("Microsoft YaHei UI", 10, "bold")).grid(
                row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 4))
            ttk.Label(win, text="行数：").grid(row=1, column=0, sticky="w", padx=12, pady=3)
            rows_var = tk.StringVar(value=str(best[0]))
            ttk.Entry(win, textvariable=rows_var, width=8).grid(row=1, column=1, sticky="w", pady=3)
            ttk.Label(win, text="列数：").grid(row=2, column=0, sticky="w", padx=12, pady=3)
            cols_var = tk.StringVar(value=str(n // best[0]))
            ttk.Entry(win, textvariable=cols_var, width=8).grid(row=2, column=1, sticky="w", pady=3)
            ttk.Label(win, text="热图指标：").grid(row=3, column=0, sticky="w", padx=12, pady=3)
            metric_labels = {"main_peak": "主峰位 (cm-1)", "intensity": "主峰强度",
                             "fwhm": "主峰半高宽 FWHM", "peaks": "识别峰数",
                             "total": "总强度", "at": "指定波数附近的峰强度"}
            metric_shown = {k: T(v) for k, v in metric_labels.items()}
            mvar = tk.StringVar(value=metric_shown["main_peak"])
            ttk.Combobox(win, textvariable=mvar, state="readonly",
                         values=list(metric_shown.values()), width=22).grid(
                row=3, column=1, columnspan=2, sticky="w", pady=3)
            ttk.Label(win, text="指定波数(cm-1)：").grid(row=4, column=0, sticky="w", padx=12, pady=3)
            at_var = tk.StringVar(value="1008")
            ttk.Entry(win, textvariable=at_var, width=10).grid(row=4, column=1, sticky="w", pady=3)
            status = tk.StringVar(value="")
            ttk.Label(win, textvariable=status, foreground="#1a4a8a").grid(
                row=5, column=0, columnspan=3, sticky="w", padx=12, pady=(4, 8))

            def run():
                try:
                    rows_n = max(1, int(rows_var.get()))
                    cols_n = max(1, int(cols_var.get()))
                except ValueError:
                    status.set(T("行数 / 列数必须是整数。"))
                    return
                if rows_n * cols_n > n:
                    status.set(T("需要 %d 个点位，当前只有 %d 条光谱。") % (rows_n * cols_n, n))
                    return
                label = mvar.get()
                key = [k for k, v in metric_shown.items() if v == label][0]
                target = None
                if key == "at":
                    try:
                        target = float(at_var.get())
                    except ValueError:
                        status.set(T("指定波数必须是数字。"))
                        return
                plot = self.plot_options()
                chosen = specs[:rows_n * cols_n]
                values = []
                for name, xs, ys in chosen:
                    peaks = analyze_peaks(xs, ys, plot)
                    if not peaks:
                        values.append(None)
                        continue
                    main = max(peaks, key=lambda pk: pk["prominence"])
                    if target is not None:
                        near = min(peaks, key=lambda pk: abs(pk["x"] - target))
                        values.append(near["y"] if abs(near["x"] - target) <= 20 else None)
                    elif key == "intensity":
                        values.append(main["y"])
                    elif key == "fwhm":
                        values.append(main["fwhm"])
                    elif key == "peaks":
                        values.append(float(len(peaks)))
                    elif key == "total":
                        values.append(sum(ys))
                    else:
                        values.append(main["x"])
                grid = [values[r * cols_n:(r + 1) * cols_n] for r in range(rows_n)]
                cbar = label if key != "at" else T("%.1f cm-1 附近峰强度") % (target or 0.0)
                src = self.files[self.listbox.curselection()[0]]
                out = self._tool_out_dir(src)
                png = os.path.join(out, "二维成像_%s.png" % _safe_name(cbar))
                render_heatmap(png, grid, T("二维成像 · %s") % cbar, cbar, plot,
                               row_labels=[T("第%d行") % (r + 1) for r in range(rows_n)])
                csv_path = os.path.join(out, "二维成像_%s.csv" % _safe_name(cbar))
                with open(csv_path, "w", encoding="utf-8-sig", newline="") as f:
                    f.write(T("行\\列,") + ",".join(str(c + 1) for c in range(cols_n)) + "\n")
                    for r in range(rows_n):
                        cells = ["" if v is None else "%.6g" % v for v in grid[r]]
                        f.write(T("第%d行,") % (r + 1) + ",".join(cells) + "\n")
                valid = [v for v in values if v is not None]
                if valid:
                    self._log(T("二维成像：%d 个点位（%d×%d），%s 范围 %.4g ~ %.4g")
                              % (len(chosen), rows_n, cols_n, cbar, min(valid), max(valid)))
                else:
                    self._log(T("二维成像：%d 个点位（无有效数值）") % len(chosen))
                self._log(T("   热图：%s") % os.path.basename(png))
                self._log(T("   数值表：%s") % os.path.basename(csv_path))
                status.set(T("完成，已保存热图与数值表。"))
                self._show_image(png, T("二维成像 · %s") % cbar)

            row = ttk.Frame(win)
            row.grid(row=6, column=0, columnspan=3, sticky="e", padx=12, pady=(0, 12))
            ttk.Button(row, text="生成热图", command=run).pack(side="right")
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right", padx=6)

        def tool_waterfall(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "瀑布图需要选中至少 2 条光谱。")
                return
            src = self.files[self.listbox.curselection()[0]]
            out = self._tool_out_dir(src)
            dst = os.path.join(out, "瀑布图_%d条.png" % len(specs))
            try:
                p = self.plot_options()
                render_waterfall(dst, specs, T("瀑布图（%d 条，归一化后纵向错开）") % len(specs),
                                 _DEFAULT_X_HEADER, T("归一化强度"), p, offset=p["stack_offset"])
            except Exception as exc:
                messagebox.showerror("生成失败", str(exc))
                return
            self._log(T("瀑布图：%s") % os.path.basename(dst))
            self._show_image(dst, T("瀑布图（%d 条）") % len(specs))

        def tool_overlay(self):
            """多数据图叠加：先开交互式预览，用户可手动增减标注峰位，确认后再导出。"""
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo(T("提示"), T("多数据图叠加需要选中至少 2 条光谱。"))
                return
            try:
                from PIL import Image as _I
                from PIL import ImageTk
            except Exception:
                messagebox.showerror(T("缺少组件"), T("预览需要 Pillow（pip install pillow）。"))
                return
            n = len(specs)
            adv = self.adv_values
            title_cn = "多数据图叠加（%d 条）" % n

            win = tk.Toplevel(self)
            win.title(T("多数据图叠加（堆叠排布 · 预览）"))
            win.transient(self)
            adv["stack_offset"] = float(adv.get("stack_offset", 1.0) or 1.0)

            # ---- 参数行 ----
            p1 = ttk.Frame(win)
            p1.pack(fill="x", padx=12, pady=(10, 2))
            nvar = tk.BooleanVar(value=True)
            ttk.Checkbutton(p1, text=T("各条归一化到最大值 = 1"),
                            variable=nvar).pack(side="left")
            ttk.Label(p1, text=T("谱线偏移：")).pack(side="left", padx=(10, 0))
            off_var = tk.StringVar(value=str(adv.get("stack_offset", 1.0)))
            ttk.Entry(p1, textvariable=off_var, width=6).pack(side="left")
            ttk.Label(p1, text=T("峰位合并容差(cm-1)：")).pack(side="left", padx=(10, 0))
            tol_var = tk.StringVar(value="")
            ttk.Entry(p1, textvariable=tol_var, width=6).pack(side="left")
            ttk.Label(p1, text=T("(留空 = 最小峰间距)"),
                      foreground="#888").pack(side="left", padx=4)
            p2 = ttk.Frame(win)
            p2.pack(fill="x", padx=12, pady=(0, 4))
            avar = tk.BooleanVar(value=bool(self.annotate_peaks.get()))
            ttk.Checkbutton(p2, text=T("标注峰位"), variable=avar).pack(side="left")
            mvar = tk.BooleanVar(value=True)
            ttk.Checkbutton(p2, text=T("峰位跨谱合并（一峰一线一值）"), variable=mvar).pack(
                side="left", padx=(10, 0))
            lvar = tk.BooleanVar(value=bool(self.peak_labels.get()))
            ttk.Checkbutton(p2, text=T("显示峰位数值"),
                            variable=lvar).pack(side="left", padx=(10, 0))
            tvar = tk.BooleanVar(value=bool(self.show_title.get()))
            ttk.Checkbutton(p2, text=T("标题"), variable=tvar).pack(side="left", padx=(10, 0))
            ttk.Button(p2, text=T("刷新预览"),
                       command=lambda: refresh()).pack(side="left", padx=12)

            # ---- 图例位置（可以选，也可以直接在图上把图例拖走）----
            # lg 是当前生效的位置；拖动时写成 inside + 比例坐标，选完/拖完都存成全局默认。
            lg = {"pos": adv.get("legend_pos") or "right", "xy": adv.get("legend_xy")}
            legend_labels = {
                "right": T("绘图区右侧留白（推荐）"),
                "top": T("绘图区上方"),
                "bottom": T("绘图区下方"),
                "inside": T("图内自由位置（可拖动）"),
                "none": T("不显示图例"),
            }
            lpos_inv = {v: k for k, v in legend_labels.items()}
            p3 = ttk.Frame(win)
            p3.pack(fill="x", padx=12, pady=(0, 4))
            ttk.Label(p3, text=T("图例位置：")).pack(side="left")
            lpos_var = tk.StringVar(
                value=legend_labels.get(lg["pos"], legend_labels["right"]))
            lpos_box = ttk.Combobox(p3, textvariable=lpos_var, state="readonly",
                                    values=list(legend_labels.values()), width=22)
            lpos_box.pack(side="left", padx=(2, 10))
            ttk.Label(p3, text=T("（也可以在图上直接按住图例拖到想要的位置）"),
                      foreground="#888").pack(side="left")

            # ---- 画布 ----
            sh = win.winfo_screenheight()
            max_h = max(300, min(600, sh - 270))
            cw = tk.Canvas(win, background="#f2f2f2", highlightthickness=1,
                           highlightbackground="#c9c9c9")
            cw.pack(padx=12, pady=(2, 4))

            lo, hi = overlay_range(specs)
            rng = "（%.1f ~ %.1f cm-1）" % (lo, hi) if lo is not None else ""
            info = tk.StringVar(value=T("横坐标取各条谱波数范围的交集%s") % rng)
            ttk.Label(win, textvariable=info, foreground="#777").pack(anchor="w", padx=12)
            tip = tk.StringVar(value="")
            ttk.Label(win, textvariable=tip, foreground="#1a4a8a").pack(anchor="w", padx=12)

            row = ttk.Frame(win)
            row.pack(fill="x", padx=12, pady=(4, 10))
            ttk.Button(row, text=T("导出 PNG"), command=lambda: do_export()).pack(side="right")
            ttk.Button(row, text=T("重新检测"), command=lambda: reset_marks()).pack(
                side="right", padx=6)
            ttk.Button(row, text=T("关闭"), command=win.destroy).pack(side="right")

            # ---- 状态 ----
            state = {"marks": None, "geom": None, "scale": 1.0, "photo": None,
                     "auto": [], "drag": None}
            tmp_png = os.path.join(tempfile.gettempdir(), "_raman_overlay_preview.png")

            def plot_opts():
                op = self.plot_options()
                op["annotate_peaks"] = bool(avar.get())
                op["peak_labels"] = bool(lvar.get())
                op["show_title"] = bool(tvar.get())
                op["stack_offset"] = _offset()
                op["peak_merge_tol"] = _tol()
                op["legend_pos"] = lg["pos"]
                op["legend_xy"] = lg["xy"]
                return op

            def _offset():
                try:
                    v = float(str(off_var.get()).strip() or 1.0)
                except ValueError:
                    return float(adv.get("stack_offset", 1.0) or 1.0)
                return max(0.2, min(v, 3.0))

            def _tol():
                txt = str(tol_var.get()).strip()
                if not txt:
                    return None
                try:
                    return max(0.001, float(txt))
                except ValueError:
                    return None

            def refresh(note=""):
                """重新渲染预览。回到没手动改过时，同时记下自动检测的峰位。"""
                op = plot_opts()
                if state["marks"] is None or not mvar.get():
                    state["auto"] = [{"x": c["x"], "manual": c["manual"]}
                                     for c in peak_clusters(specs, op)]
                ylab = T("归一化强度") if nvar.get() else T("强度")
                try:
                    geom = render_overlay(tmp_png, specs, T(title_cn),
                                          _DEFAULT_X_HEADER, ylab, op,
                                          normalize=bool(nvar.get()),
                                          merge_peaks=bool(mvar.get()),
                                          marks=state["marks"],
                                          return_geometry=True)
                except Exception as exc:
                    tip.set(T("生成失败：%s") % exc)
                    return
                state["geom"] = geom
                img = _I.open(tmp_png)
                scale = min(1.0, 1120.0 / img.width, float(max_h) / img.height)
                if scale < 1.0:
                    resample = _I.LANCZOS if hasattr(_I, "LANCZOS") else _I.BILINEAR
                    img = img.resize((max(1, int(img.width * scale)),
                                      max(1, int(img.height * scale))), resample)
                state["scale"] = scale
                state["photo"] = ImageTk.PhotoImage(img)
                cw.configure(width=img.width, height=img.height)
                cw.delete("all")
                cw.create_image(0, 0, anchor="nw", image=state["photo"])
                show_mark_count(note)

            def marks_now():
                if not mvar.get():
                    return []
                if state["marks"] is None:
                    return list(state["auto"])
                return state["marks"]

            def show_mark_count(note=""):
                cnt = len(marks_now())
                tip.set(T("当前标注 %d 个峰位。") % cnt
                        + ("　" + T(note) if note else "")
                        + ("　" + T("左键点图＝加一个峰位，右键＝删掉最近的一条")
                           if mvar.get() else "　" + T("未开启跨谱合并，不能手动增减")))

            def _to_data(event):
                """画布坐标 → 波数。落在绘图区外返回 None。"""
                g = state["geom"]
                if not g:
                    return None
                ix = event.x / state["scale"]
                iy = event.y / state["scale"]
                if not (g["ml"] - 2 <= ix <= g["ml"] + g["pw"] + 2):
                    return None
                if not (g["mt"] - 2 <= iy <= g["mt"] + g["ph"] + 2):
                    return None
                if g["xmax"] <= g["xmin"]:
                    return None
                return g["xmin"] + (ix - g["ml"]) / float(g["pw"]) * (g["xmax"] - g["xmin"])

            def on_left(event):
                if not mvar.get():
                    tip.set(T("请先勾选“峰位跨谱合并”，再手动增减峰位。"))
                    return
                x = _to_data(event)
                if x is None:
                    return
                if state["marks"] is None:
                    state["marks"] = list(state["auto"])
                tol = _merge_tol(plot_opts())
                for m in state["marks"]:
                    if abs(m["x"] - x) <= tol:
                        tip.set(T("附近已有标注（%.1f），没有重复添加。") % m["x"])
                        return
                state["marks"].append({"x": x, "manual": True})
                state["marks"].sort(key=lambda m: m["x"])
                refresh(T("已手动添加峰位 %.1f。") % x)

            def on_right(event):
                if not mvar.get():
                    tip.set(T("请先勾选“峰位跨谱合并”，再手动增减峰位。"))
                    return
                x = _to_data(event)
                if x is None:
                    return
                marks = marks_now()
                if not marks:
                    tip.set(T("当前没有可删除的标注。"))
                    return
                g = state["geom"]
                # 命中范围按屏幕 20 px 折算成波数，点得准一点
                lim = (20.0 / max(state["scale"], 1e-6)) / float(g["pw"]) * (
                    g["xmax"] - g["xmin"])
                near = min(marks, key=lambda m: abs(m["x"] - x))
                if abs(near["x"] - x) > lim:
                    tip.set(T("附近没有标注峰位，请点在虚线上再右键。"))
                    return
                if state["marks"] is None:
                    state["marks"] = list(state["auto"])
                state["marks"] = [m for m in state["marks"]
                                  if abs(m["x"] - near["x"]) > 1e-9]
                refresh(T("已删除峰位标注 %.1f。") % near["x"])

            def reset_marks():
                state["marks"] = None
                refresh(T("已恢复为自动检测的峰位。"))

            def do_export():
                src = self.files[self.listbox.curselection()[0]]
                out = self._tool_out_dir(src)
                dst = os.path.join(out, "叠加图_%d条.png" % n)
                op = plot_opts()
                ylab = T("归一化强度") if nvar.get() else T("强度")
                try:
                    render_overlay(dst, specs, T(title_cn), _DEFAULT_X_HEADER, ylab, op,
                                   normalize=bool(nvar.get()),
                                   merge_peaks=bool(mvar.get()),
                                   marks=state["marks"] if mvar.get() else None)
                except Exception as exc:
                    tip.set(T("生成失败：%s") % exc)
                    return
                adv["stack_offset"] = op["stack_offset"]
                self._log(T("叠加图：%s（标注 %d 个峰位）")
                          % (os.path.basename(dst), len(marks_now())))
                self._show_image(dst, T(title_cn))

            # ---- 图例拖动 ----
            # 按在图例框里＝拖图例；按在别处还是“加峰位”，两者互不干扰。
            def on_press(event):
                g = state["geom"]
                if g and legend_hit(g.get("legend_box"), event.x / state["scale"],
                                    event.y / state["scale"]):
                    box = g["legend_box"]
                    state["drag"] = {"off": (event.x / state["scale"] - box[0],
                                             event.y / state["scale"] - box[1]),
                                     "box": None}
                    tip.set(T("按住拖动图例，松手即定位。"))
                    return
                on_left(event)

            def on_motion(event):
                d = state["drag"]
                g = state["geom"]
                if not d or not g or not g.get("legend_box"):
                    return
                box = g["legend_box"]
                bw, bh = box[2] - box[0], box[3] - box[1]
                s = state["scale"]
                ix = event.x / s - d["off"][0]
                iy = event.y / s - d["off"][1]
                ix = min(max(g["ml"] + 4, ix), max(g["ml"] + 4, g["ml"] + g["pw"] - bw - 4))
                iy = min(max(g["mt"] + 4, iy), max(g["mt"] + 4, g["mt"] + g["ph"] - bh - 4))
                d["box"] = (ix, iy, ix + bw, iy + bh)
                cw.delete("legendghost")
                cw.create_rectangle(ix * s, iy * s, (ix + bw) * s, (iy + bh) * s,
                                    outline="#1a4a8a", dash=(4, 3), tags="legendghost")

            def on_release(_event=None):
                d = state["drag"]
                state["drag"] = None
                cw.delete("legendghost")
                g = state["geom"]
                if not d or not d.get("box") or not g or g["pw"] <= 0 or g["ph"] <= 0:
                    return
                box = d["box"]
                lg["xy"] = [min(1.0, max(0.0, (box[0] - g["ml"]) / float(g["pw"]))),
                            min(1.0, max(0.0, (box[1] - g["mt"]) / float(g["ph"])))]
                lg["pos"] = "inside"
                lpos_var.set(legend_labels["inside"])
                adv["legend_pos"], adv["legend_xy"] = lg["pos"], lg["xy"]
                save_legend_defaults(lg["pos"], lg["xy"])
                refresh(T("图例已挪到图内 %.0f%% × %.0f%%（导出和下次预览都用这个位置）。")
                        % (lg["xy"][0] * 100, lg["xy"][1] * 100))

            def on_pos_pick(_event=None):
                key = lpos_inv.get(lpos_var.get(), "right")
                lg["pos"] = key
                if key == "inside" and not lg["xy"]:
                    lg["xy"] = list(LEGEND_XY_DEFAULT)
                adv["legend_pos"], adv["legend_xy"] = lg["pos"], lg["xy"]
                save_legend_defaults(lg["pos"], lg["xy"])
                refresh(T("图例位置：%s") % lpos_var.get())

            lpos_box.bind("<<ComboboxSelected>>", on_pos_pick)
            cw.bind("<Button-1>", on_press)
            cw.bind("<B1-Motion>", on_motion)
            cw.bind("<ButtonRelease-1>", on_release)
            cw.bind("<Button-3>", on_right)
            # 勾选框即时重绘；两个输入框按回车或点【刷新预览】才重绘，
            # 免得每敲一个字符就重画一次
            for v in (nvar, mvar, lvar, tvar):
                v.trace_add("write", lambda *a: refresh())
            for e in p1.winfo_children():
                if e.winfo_class() == "TEntry":
                    e.bind("<Return>", lambda ev: refresh())
            win.bind("<Escape>", lambda e: win.destroy())
            refresh()

        def tool_replace(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "谱段替换需要选中至少 2 条光谱（第 1 条为目标谱）。")
                return
            names = [s[0] for s in specs]
            win = tk.Toplevel(self)
            win.title("谱段替换 / 拼接")
            win.transient(self)
            win.resizable(False, False)
            ttk.Label(win, text="把目标谱的一段波数区间替换成另一条谱的对应区间。",
                      font=("Microsoft YaHei UI", 10, "bold")).grid(
                row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 4))
            ttk.Label(win, text="目标谱：").grid(row=1, column=0, sticky="w", padx=12, pady=3)
            tgt_var = tk.StringVar(value=names[0])
            ttk.Combobox(win, textvariable=tgt_var, state="readonly",
                         values=names, width=40).grid(row=1, column=1, columnspan=2,
                                                      sticky="w", pady=3)
            ttk.Label(win, text="来源谱：").grid(row=2, column=0, sticky="w", padx=12, pady=3)
            src_var = tk.StringVar(value=names[1])
            ttk.Combobox(win, textvariable=src_var, state="readonly",
                         values=names, width=40).grid(row=2, column=1, columnspan=2,
                                                      sticky="w", pady=3)
            ttk.Label(win, text="替换区间(cm-1)：").grid(row=3, column=0, sticky="w",
                                                        padx=12, pady=3)
            lo_var = tk.StringVar(value="")
            hi_var = tk.StringVar(value="")
            ttk.Entry(win, textvariable=lo_var, width=10).grid(row=3, column=1, sticky="w", pady=3)
            ttk.Label(win, text="至").grid(row=3, column=2, sticky="w")
            ttk.Entry(win, textvariable=hi_var, width=10).grid(row=3, column=3, sticky="w", pady=3)
            ttk.Label(win, text="过渡宽度(cm-1)：").grid(row=4, column=0, sticky="w",
                                                        padx=12, pady=3)
            blend_var = tk.StringVar(value="10")
            ttk.Entry(win, textvariable=blend_var, width=10).grid(row=4, column=1, sticky="w",
                                                                  pady=3)
            ttk.Label(win, text="（边缘线性过渡，避免出现台阶）", foreground="#888").grid(
                row=4, column=2, columnspan=2, sticky="w")
            status = tk.StringVar(value="")
            ttk.Label(win, textvariable=status, foreground="#1a4a8a", wraplength=520,
                      justify="left").grid(row=5, column=0, columnspan=4, sticky="w",
                                           padx=12, pady=(4, 8))

            def compute():
                tgt_i = names.index(tgt_var.get())
                src_i = names.index(src_var.get())
                if tgt_i == src_i:
                    status.set(T("目标谱与来源谱不能相同。"))
                    return None
                try:
                    lo = float(lo_var.get())
                    hi = float(hi_var.get())
                except ValueError:
                    status.set(T("请填写有效的替换区间（起、止波数）。"))
                    return None
                try:
                    blend = float(blend_var.get() or 0)
                except ValueError:
                    blend = 0.0
                _n, txs, tys = specs[tgt_i]
                _n2, sxs, sys_ = specs[src_i]
                out, count = replace_segment(txs, tys, sxs, sys_, lo, hi, blend)
                if count == 0:
                    status.set(T("该区间内没有数据点，请检查波数范围。"))
                    return None
                return tgt_i, txs, out, count

            def preview():
                got = compute()
                if not got:
                    return
                tgt_i, txs, out, count = got
                self._draw([(names[tgt_i], txs, out, _PALETTE[0])],
                           "谱段替换预览（替换 %d 点）" % count,
                           _DEFAULT_X_HEADER, "Intensity")
                status.set(T("已在右侧预览（尚未保存）。"))

            def export():
                got = compute()
                if not got:
                    return
                tgt_i, txs, out, count = got
                src_path = self.files[self.listbox.curselection()[tgt_i]]
                dst = os.path.join(self._tool_out_dir(src_path),
                                   os.path.splitext(specs[tgt_i][0].split("|")[0])[0] +
                                   "_替换拼接.csv")
                with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                    f.write("%s,Intensity\n" % _DEFAULT_X_HEADER)
                    for x, y in zip(txs, out):
                        f.write("%.6f,%.6g\n" % (x, y))
                self._log(T("谱段替换完成（替换 %d 点）→ %s") % (count, os.path.basename(dst)))
                status.set(T("已保存：%s") % os.path.basename(dst))

            row = ttk.Frame(win)
            row.grid(row=6, column=0, columnspan=4, sticky="e", padx=12, pady=(0, 12))
            ttk.Button(row, text="预览", command=preview).pack(side="right")
            ttk.Button(row, text="导出 CSV", command=export).pack(side="right", padx=6)
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right")

        def tool_subtract_interactive(self):
            specs = self._selected_spectra()
            if len(specs) < 2:
                messagebox.showinfo("提示", "A − k·B 需要选中 2 条光谱（按选中顺序，A 在前）。")
                return
            name_a, axs, ays = specs[0]
            name_b, bxs, bys = specs[1]
            win = tk.Toplevel(self)
            win.title("交互式相减 · A − k·B")
            win.transient(self)
            win.resizable(False, False)
            ttk.Label(win, text="A：%s" % name_a, font=("Microsoft YaHei UI", 10, "bold")).grid(
                row=0, column=0, columnspan=3, sticky="w", padx=12, pady=(12, 2))
            ttk.Label(win, text="B：%s" % name_b, font=("Microsoft YaHei UI", 10, "bold")).grid(
                row=1, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 6))
            norm_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(win, text="先各自按最大值归一化再相减（推荐，便于找平）",
                            variable=norm_var).grid(row=2, column=0, columnspan=3,
                                                    sticky="w", padx=12, pady=3)
            k_var = tk.DoubleVar(value=1.0)
            info = tk.StringVar(value="")
            ttk.Label(win, text="k 值：").grid(row=3, column=0, sticky="w", padx=12)
            scale = ttk.Scale(win, from_=0.0, to=2.0, orient="horizontal",
                              variable=k_var, length=360)
            scale.grid(row=3, column=1, columnspan=2, sticky="w", pady=6)

            def current():
                a = (axs, _norm_max(ays) if norm_var.get() else ays)
                b = (bxs, _norm_max(bys) if norm_var.get() else bys)
                return subtract_scaled(a, b, k_var.get())

            def refresh(*_a):
                got = current()
                if got is None:
                    info.set(T("两条光谱没有重叠波数区间。"))
                    return
                xs, ys = got
                info.set(T("k = %.2f   区间 %.1f ~ %.1f cm-1   剩余最大强度 %.4g")
                         % (k_var.get(), xs[0], xs[-1], max(ys) if ys else 0.0))
                self._draw([("A − %.2f·B" % k_var.get(), xs, ys, _PALETTE[3])],
                           "交互相减：A − k·B", _DEFAULT_X_HEADER, "Intensity")

            def export():
                got = current()
                if got is None:
                    return
                xs, ys = got
                src_path = self.files[self.listbox.curselection()[0]]
                dst = os.path.join(self._tool_out_dir(src_path), "相减_A减kB.csv")
                with open(dst, "w", encoding="utf-8-sig", newline="") as f:
                    f.write("%s,A-kB(k=%.2f)\n" % (_DEFAULT_X_HEADER, k_var.get()))
                    for x, y in zip(xs, ys):
                        f.write("%.6f,%.6g\n" % (x, y))
                self._log(T("交互相减已保存（k=%.2f）→ %s") % (k_var.get(), os.path.basename(dst)))

            scale.configure(command=refresh)
            norm_var.trace_add("write", lambda *a: refresh())
            ttk.Label(win, textvariable=info, foreground="#1a4a8a").grid(
                row=4, column=0, columnspan=3, sticky="w", padx=12)
            row = ttk.Frame(win)
            row.grid(row=5, column=0, columnspan=3, sticky="e", padx=12, pady=12)
            ttk.Button(row, text="导出 CSV", command=export).pack(side="right")
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right", padx=6)
            refresh()

        def open_mineral_info(self):
            win = tk.Toplevel(self)
            win.title("矿物信息与拉曼峰归属")
            win.transient(self)
            win.geometry("760x620")
            ttk.Label(win, text="内置常见矿物特征峰表（用于峰位归属参考，可离网使用）。",
                      font=("Microsoft YaHei UI", 10, "bold")).pack(anchor="w",
                                                                   padx=12, pady=(12, 4))
            top = ttk.Frame(win)
            top.pack(fill="x", padx=12)
            ttk.Label(top, text="检索：").pack(side="left")
            q = tk.StringVar(value="")
            ttk.Entry(top, textvariable=q, width=22).pack(side="left", padx=4)
            mode_labels = {"name": "名称 / 化学式", "formula": "化学式包含",
                           "element": "元素（空格分隔，全部含）"}
            mode_shown = {k: T(v) for k, v in mode_labels.items()}
            mvar = tk.StringVar(value=mode_shown["name"])
            ttk.Combobox(top, textvariable=mvar, state="readonly", width=20,
                         values=list(mode_shown.values())).pack(side="left", padx=4)
            tv = ttk.Treeview(win, columns=("en", "cn", "formula", "system"),
                              show="headings", height=8, selectmode="browse")
            for col, head, w in (("en", "英文名", 170), ("cn", "中文名", 150),
                                 ("formula", "化学式", 160), ("system", "晶系", 120)):
                tv.heading(col, text=head)
                tv.column(col, width=w, anchor="w")
            tv.pack(fill="both", expand=False, padx=12, pady=8)
            detail = tk.Text(win, height=14, wrap="none", font=("Consolas", 10))
            detail.pack(fill="both", expand=True, padx=12)
            hit_cache = {"rows": []}

            def do_search():
                mode = [k for k, v in mode_shown.items() if v == mvar.get()][0]
                rows = mineral_search(q.get(), mode)
                if not q.get().strip():
                    rows = list(_MINERAL_DB)
                hit_cache["rows"] = rows
                tv.delete(*tv.get_children())
                for i, rec in enumerate(rows):
                    tv.insert("", "end", iid=str(i),
                              values=(rec["en"], rec["cn"], rec["formula"], rec["system"]))
                detail.delete("1.0", "end")
                detail.insert("end", T("命中 %d 种矿物。\n%s") % (
                    len(rows), T("用法：选中一行看详情，再点“作为峰位归属参考”。")))

            def show_detail(_e=None):
                sel = tv.selection()
                if not sel:
                    return
                try:
                    rec = hit_cache["rows"][int(sel[0])]
                except (ValueError, IndexError):
                    return
                detail.delete("1.0", "end")
                detail.insert("end", mineral_info_text(rec["en"]))
                detail.insert("end", T("\n\n（正在查找 RRUFF 本地样品记录…）"))
                win.update_idletasks()
                extra = ""
                try:
                    sample = rruff_sample_info(rec["en"])
                except Exception:
                    sample = None
                if sample:
                    extra = "\n\n" + rruff_info_text(sample)
                else:
                    extra = T("\n\n（本地数据包里暂无该矿物的样品记录；"
                              "可用“数据库 → RRUFF 数据源”下载后自动补全）")
                detail.delete("1.0", "end")
                detail.insert("end", mineral_info_text(rec["en"]) + extra)

            def use_as_ref():
                sel = tv.selection()
                if not sel:
                    return
                try:
                    rec = hit_cache["rows"][int(sel[0])]
                except (ValueError, IndexError):
                    return
                self.adv_values["mineral_name"] = rec["en"]
                self._log(T("已把「%s」设为峰位归属参考；峰列表/报告会多一列“归属”。") % rec["en"])
                win.destroy()

            tv.bind("<<TreeviewSelect>>", show_detail)
            row = ttk.Frame(win)
            row.pack(fill="x", padx=12, pady=(0, 12))
            ttk.Button(row, text="检索", command=do_search).pack(side="left")
            ttk.Button(row, text="作为峰位归属参考", command=use_as_ref).pack(side="left", padx=6)
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right")
            do_search()

        def tool_report(self):
            sel = self.listbox.curselection()
            if not sel:
                messagebox.showinfo("提示", "请先在左侧选择要写入报告的光谱文件。")
                return
            paths = [self.files[i] for i in sel]
            first = paths[0]
            out = self._tool_out_dir(first)
            dst = os.path.join(out, "光谱分析报告_%s.html" % time.strftime("%Y%m%d_%H%M%S"))
            try:
                path, count = build_report(paths, dst, plot=self.plot_options())
            except Exception as exc:
                messagebox.showerror("生成报告失败", str(exc))
                return
            self._log(T("分析报告已生成（%d 条光谱）：%s") % (count, path))
            if messagebox.askyesno("报告已生成",
                                   "已生成自包含 HTML 报告（图片已内嵌）：\n%s\n\n"
                                   "现在打开查看吗？\n（浏览器里可“打印 → 另存为 PDF”）" % path):
                try:
                    os.startfile(path)
                except OSError:
                    self._open_folder(os.path.dirname(path))

        def open_batch(self):
            folder = filedialog.askdirectory(title=T("选择要批处理的文件夹（会递归查找光谱文件）"))
            if not folder:
                return
            win = tk.Toplevel(self)
            win.title("文件夹批处理")
            win.transient(self)
            win.resizable(False, False)
            ttk.Label(win, text="批处理目录：%s" % folder, wraplength=560,
                      justify="left").grid(row=0, column=0, columnspan=3, sticky="w",
                                           padx=12, pady=(12, 4))
            ttk.Label(win, text="输出内容：").grid(row=1, column=0, sticky="w", padx=12, pady=3)
            fmt_row = ttk.Frame(win)
            fmt_row.grid(row=1, column=1, columnspan=2, sticky="w", pady=3)
            f_excel = tk.BooleanVar(value=True)
            f_png = tk.BooleanVar(value=True)
            f_csv = tk.BooleanVar(value=False)
            f_peaks = tk.BooleanVar(value=True)
            f_jcamp = tk.BooleanVar(value=False)
            ttk.Checkbutton(fmt_row, text="Excel", variable=f_excel).pack(side="left")
            ttk.Checkbutton(fmt_row, text="PNG", variable=f_png).pack(side="left", padx=6)
            ttk.Checkbutton(fmt_row, text="CSV", variable=f_csv).pack(side="left")
            ttk.Checkbutton(fmt_row, text="峰列表", variable=f_peaks).pack(side="left", padx=6)
            ttk.Checkbutton(fmt_row, text="JCAMP-DX", variable=f_jcamp).pack(side="left")
            ttk.Label(win, text="输出位置：").grid(row=2, column=0, sticky="w", padx=12, pady=3)
            dest_list = ["源文件夹下的 _转换结果", "分析结果文件夹",
                         "源文件夹下（与源文件混放）"]
            dest_var = tk.StringVar(value=T(dest_list[0]))
            ttk.Combobox(win, textvariable=dest_var, state="readonly", width=32,
                         values=[T(v) for v in dest_list]).grid(row=2, column=1,
                                                                columnspan=2, sticky="w", pady=3)
            status = tk.StringVar(value="")
            ttk.Label(win, textvariable=status, foreground="#1a4a8a", wraplength=560,
                      justify="left").grid(row=3, column=0, columnspan=3, sticky="w",
                                           padx=12, pady=(6, 8))

            def start():
                fmts = []
                if f_excel.get():
                    fmts.append("xlsx")
                if f_png.get():
                    fmts.append("png")
                if f_csv.get():
                    fmts.append("csv")
                if f_peaks.get():
                    fmts.append("peaks")
                if f_jcamp.get():
                    fmts.append("jcamp")
                if not fmts:
                    status.set(T("请至少勾选一种输出内容。"))
                    return
                dest = dest_var.get()
                for _v in dest_list:
                    if T(_v) == dest:
                        dest = _v
                        break
                if dest == "分析结果文件夹":
                    target = results_dir()
                elif dest == "源文件夹下（与源文件混放）":
                    target = folder
                else:
                    target = os.path.join(folder, "_转换结果")
                status.set(T("正在批处理，请稍候…（详细进度见主界面日志）"))
                self._log("=" * 48)
                self._log(T("批处理开始：%s") % folder)
                self._log(T("   输出到：%s（%s）") % (target, "+".join(fmts)))
                plot_opts = self.plot_options()

                def work():
                    def prog(i, n, src):
                        self.q.put(("log", "   (%d/%d) %s" % (i, n, os.path.basename(src))))
                        self.q.put(("batchprog", (i, n)))
                    try:
                        res = batch_convert(folder, out_dir=target, formats=tuple(fmts),
                                            plot=plot_opts, progress=prog)
                    except Exception as exc:
                        self.q.put(("log", "批处理失败：%s" % exc))
                        self.q.put(("batchdone", None))
                        return
                    self.q.put(("batchdone", res))

                threading.Thread(target=work, daemon=True).start()

            row = ttk.Frame(win)
            row.grid(row=4, column=0, columnspan=3, sticky="e", padx=12, pady=(0, 12))
            ttk.Button(row, text="开始批处理", command=start).pack(side="right")
            ttk.Button(row, text="关闭", command=win.destroy).pack(side="right", padx=6)

        def open_out_dir(self):
            target = None
            if self.same_dir.get():
                if self.files:
                    target = os.path.dirname(self.files[0])
            else:
                target = self.dir_var.get().strip() or None
            if target and os.path.isdir(target):
                try:
                    os.startfile(target)
                    return
                except OSError:
                    pass
            messagebox.showinfo("提示", "还没有可打开的输出目录。")

    app = App()
    try:
        localize_menu(app.nametowidget(app.cget("menu")))
    except Exception:
        pass
    try:
        localize_tree(app)
    except Exception:
        pass
    app.mainloop()


_MANUAL_VERSION = "2.5"
_MANUAL_TITLE = "拉曼光谱工具 · 使用说明书"
_MANUAL_MARKER = "（说明书版本：%s）" % _MANUAL_VERSION
_MANUAL_MARKER_EN = "(guide version: %s)" % _MANUAL_VERSION

_MANUAL_HEAD = """拉曼光谱工具 · Raman Spectrum Toolkit  使用说明书
（JASCO .jws 光谱转换 · 拉曼峰分析 · 矿物鉴定）
""" + _MANUAL_MARKER + """
============================================================

一句话介绍：
  把 JASCO 光谱仪 .jws 等光谱文件转成能看的图和数据，
  并内置拉曼光谱的常用处理、分析、比对与数据库功能。

本说明书怎么看：
  · 新手：先看第 1 章，5 分钟就能出第一张图。
  · 会用了：直接跳到第 17 章“常见任务速查”，按任务找步骤。
  · 界面里也能看：菜单【帮助】→【使用说明（完整手册）】，
    左侧点章节即可跳转，可搜索关键词。
  · 本文件是纯文本，用记事本 / VS Code 打开都行。

语言（中英双语）：
  界面、命令行输出、分析报告与图注都支持中文 / 英文；
  在菜单【设置】→【语言】里切换，或命令行加 --lang en|zh。
  首次启动跟随 Windows 显示语言，之后记住你的选择；
  说明书中英各一份（使用说明.txt / User_Guide.txt）。
  注意：磁盘上的数据文件夹名（工具数据 / 参考谱库 / RRUFF数据包 / 分析结果）
  保持中文，这样两种语言下数据可以互通。

目录
------------------------------------------------------------
"""

_I18N_RAW = """
# ===== 界面：控件文案 =====
(空=自动)||(blank = auto)
CSV→图片||CSV to image
CSV（纯数据）||CSV (data only)
Excel（含图表）||Excel (with chart)
PNG 图片||PNG image
PNG 文件名与源文件一致||PNG file name matches the source file
RRUFF 数据源：拉曼 / 红外 / XRD / 成分…||RRUFF data sources: Raman / IR / XRD / chemistry
k 值：||k value:
① 待鉴定的未知光谱||1. Unknown spectrum to identify
① 数据包（来源 www.rruff.net/zipped_data_files）||1. Data packages (source www.rruff.net/zipped_data_files)
① 选择文件：.jws 转数据/图表，.csv 转图片（可多选或整个文件夹）||1. Choose files: .jws to data/chart, .csv to image (multi-select or a whole folder)
② 参考库来源（勾选参与检索的数据包）||2. Reference library (tick the packages to search)
② 按矿物名 / RRUFF 编号检索（需先下载任一数据包）||2. Search by mineral name / RRUFF ID (download a package first)
② 输出位置||2. Output location
③ 候选矿物（按可信度排序）||3. Candidate minerals (ranked by confidence)
③ 表头与列名||3. Header and column names
④ 运行日志||4. Run log
下一个||Next
下载并与实测谱配对||Download and pair with the measured spectrum
下载选中数据包并建索引||Download selected packages and index them
下载选中谱线||Download selected spectra
不知道样品是什么？选中一条光谱，工具拿它的峰去参考库里比对，||Not sure what the sample is? Select a spectrum and the tool matches its peaks against the reference library,
二维成像热图（所选点位）…||2D imaging heat map (selected points)…
交互式相减 A−k·B（拖动找平）…||Interactive subtraction A−k·B (drag to flatten)…
从数据库配对（在线检索选择）…||Pair from database (search online and choose)…
作为峰位归属参考||Use as peak-assignment reference
使用说明||Guide
使用说明书||User guide
使用说明（完整手册）…||Guide (full manual)…
保存上限||Save limit
保存为 PNG||Save as PNG
保存到源文件所在目录（推荐）||Save next to the source file (recommended)
保存预览为 PNG||Save preview as PNG
候选个数：||Candidates:
先各自按最大值归一化再相减（推荐，便于找平）||Normalize each to its maximum before subtracting (recommended)
光谱处理（默认仅用于出图与峰识别）||Spectrum processing (applies to plots and peak finding by default)
光谱比对（参考谱文件）…||Compare with a reference file…
光谱相减（A−B）||Subtract spectra (A−B)
光谱预览||Spectrum preview
关于||About
关闭||Close
内置常见矿物特征峰表（用于峰位归属参考，可离网使用）。||Built-in characteristic-peak table of common minerals (offline peak assignment).
写入表头行||Write a header row
分析工具||Analysis
列数：||Columns:
刷新||Refresh
取消||Cancel
同时应用到导出的数据（CSV/Excel）||Also apply to exported data (CSV/Excel)
图幅（PNG 输出像素）||Figure size (PNG pixels)
图例位置：||Legend position:
图例位置：%s||Legend position: %s
绘图区右侧留白（推荐）||Right of the plot (recommended)
绘图区右侧留白（不压谱线，推荐）||Right of the plot (never covers the spectra; recommended)
绘图区上方||Above the plot
绘图区下方||Below the plot
图内自由位置（可拖动）||Free position inside the plot (draggable)
不显示图例||Hide the legend
（也可以在图上直接按住图例拖到想要的位置）||(You can also drag the legend anywhere on the chart)
按住拖动图例，松手即定位。||Hold and drag the legend; release to drop it.
图例已挪到图内 %.0f%% × %.0f%%（导出和下次预览都用这个位置）。||Legend moved inside the plot to %.0f%% x %.0f%% (used for the export and the next preview).
几条谱线就列几条图例；选“图内自由位置”后可以在叠加图预览窗口里拖着放，位置会记住||Every spectrum gets a legend entry; pick "free position inside the plot" and drag it in the overlay preview - the position is remembered
…还有 %d 条||...and %d more
…还有 %d 条没列（把图例位置改到右侧留白，或把图放大）||...and %d more not shown (move the legend to the right-hand gutter, or enlarge the figure)
图表设置||Chart settings
在线检索数据库（ROD / RRUFF 数据）…||Search online databases (ROD / RRUFF)…
坐标轴范围（留空 = 自动）||Axis range (blank = auto)
字号：||Font size:
对前列候选读原始谱精算（更准，稍慢）||Re-score top candidates from raw spectra (more accurate, slower)
导入本地 zip…||Import a local zip…
导出 CSV||Export CSV
导出为 TXT…||Export as TXT…
导出候选 CSV||Export candidates to CSV
导出分析报告（选中文件）…||Export analysis report (selected files)…
导出并与实测谱配对||Export and pair with the measured spectrum
导出说明书||Export guide
导出说明书为 TXT…||Export the guide as TXT…
导出选中到参考谱库||Export selected to the reference library
导出选中到本地库||Export selected to the local library
尖峰去除（宇宙射线 / 坏点）||Spike removal (cosmic rays / dead pixels)
峰位检索（参考峰表）…||Peak-position search (reference peak table)…
峰列表||Peak list
峰拟合||Peak fitting
峰拟合（选中文件）||Peak fitting (selected files)
峰标签同时显示相对强度(%)||Show relative intensity (%) with peak labels
峰识别与峰拟合||Peak detection and fitting
工具自身的数据（下载的参考谱、RRUFF 数据包、分析结果）统一保存在：||All tool data (downloaded references, RRUFF packages, analysis results) is kept in:
左键点图=补标峰，右键=删掉最近的峰（自动 / 手动都可）||Left-click the plot = add a peak, right-click = delete the nearest peak (auto or manual)
帮助||Help
平均所选光谱||Average selected spectra
应用||Apply
建立 / 更新特征索引||Build / update feature index
开始批处理||Start batch
开始检索||Search
开始转换||Start conversion
强制重建特征索引||Force rebuild the feature index
必须不含：||Must not contain:
必须含元素：||Must contain elements:
手动配对（选参考谱文件 / 文件夹）…||Pair manually (choose reference files / folder)…
打开||Open
打开 CSV 看图||Open a CSV as a chart
打开参考谱库||Open the reference library
打开所在文件夹||Open containing folder
打开数据库文件夹||Open the database folder
打开数据文件夹||Open the data folder
打开本地库文件夹||Open the local library folder
打开说明书所在文件夹||Open the folder containing the guide
打开输出目录||Open the output folder
批处理文件夹…（预处理 + 峰表 + 出图 + 汇总）||Batch a folder… (processing + peak tables + plots + summary)
批处理目录：%s||Batch folder: %s
把本地库加入文件列表||Add the local library to the file list
把目标谱的一段波数区间替换成另一条谱的对应区间。||Replace a wavenumber range of the target spectrum with the matching range of another spectrum.
把选中的 %d 条光谱按“行优先”顺序铺成点阵。||Lay the selected %d spectra out as a grid in row-major order.
拉曼位移校准（实测峰位 → 标准峰位，留空 = 不校准）||Raman shift calibration (measured peak -> standard peak; blank = none)
指定波数(cm-1)：||Wavenumber (cm-1):
按矿物名 / 化学式检索：||Search by mineral name / formula:
按矿物批量抓取（自动下载+导出）||Bulk fetch by mineral (auto download + export)
提示：Ctrl+A 全选后 Ctrl+C 可复制任意段落||Tip: Ctrl+A then Ctrl+C to copy any part
提示：默认数据文件夹就在工具（exe）所在目录；若该目录不可写，会自动改用系统用户目录。||Note: the default data folder is next to the tool (exe); if it is not writable the system user folder is used instead.
撤销手动峰||Undo manual peak
数据包容量上限(MB，0=不限制)：||Package size limit (MB, 0 = unlimited):
数据库||Database
数据文件夹…（位置 / 占用 / 容量上限 / 清理）||Data folder… (location / usage / limit / cleanup)
文件||File
显示纵坐标数值||Show Y-axis values
暂无预览\\n选中左侧文件，或点“打开 CSV 看图”||No preview yet\\nSelect a file on the left or click "Open a CSV as a chart"
更改位置…||Change location…
替换区间(cm-1)：||Replacement range (cm-1):
最佳候选对比图||Best-candidate overlay
最小峰间距(cm-1)：||Minimum peak separation (cm-1):
未知光谱检索（全库鉴定）…||Unknown-spectrum search (whole-library identification)…
本地还没有数据包。请先到【数据库 → RRUFF 数据源…】||No data packages yet. Go to [Database -> RRUFF data sources...]
下载（建议至少下载“fair / excellent”等级，谱质更好）。||to download one (prefer the "fair / excellent" grades for better spectra).
来源谱：||Source spectrum:
查找：||Find:
标注峰位||Annotate peaks
标题||Title
检索||Search
检索类型：||Search type:
检索：||Search:
横坐标列名：||X-axis column name:
横坐标刻度间隔：||X-axis tick interval:
浏览…||Browse…
添加文件||Add files
添加文件…||Add files…
添加文件夹||Add folder
添加文件夹…||Add folder…
清理已下载数据包…||Clean up downloaded packages…
清理数据包…||Clean up packages…
清空||Clear
清空列表||Clear list
清空手动峰||Clear manual peaks
清空手动峰标注||Clear manual peak annotations
清空视图||Clear view
瀑布图（所选光谱）||Waterfall (selected spectra)
点左侧章节跳转||Click a chapter on the left to jump
热图指标：||Heat-map metric:
生成热图||Build heat map
用候选做配对报告||Pair report with the candidate
用左侧选中的光谱||Use the spectrum selected on the left
用本地数据库比对（选中文件）||Compare with the local database (selected files)
目标谱：||Target spectrum:
相似度矩阵（所选文件）||Similarity matrix (selected files)
矿物信息与拉曼峰归属库…||Mineral info and Raman peak-assignment library…
确定||OK
移除选中||Remove selected
章节目录||Contents
纵坐标列名：||Y-axis column name:
网格线||Grid lines
聚类树状图||Cluster dendrogram
聚类分析 + 主成分（所选文件）…||Cluster analysis + PCA (selected files)…
自动识别光谱列名（波长 / 吸光度 / CD 等）||Auto-detect spectral column names (wavelength / absorbance / CD ...)
自动配对（在线检索并下载）…||Auto-pair (search online and download)…
自动配对（本地数据库）||Auto-pair (local database)
至||to
行数：||Rows:
设置||Settings
谱段替换 / 拼接…||Range replacement / stitching…
起始刻度：||Start tick:
跳过已存在的 CSV（不覆盖）||Skip existing CSV files (no overwrite)
输出位置：||Output location:
源文件夹下的 _转换结果||_converted subfolder in the source folder
分析结果文件夹||Analysis results folder
源文件夹下（与源文件混放）||Source folder (mixed with source files)
数据来源：ROD（Raman Open Database，RRUFF 拉曼数据）。||Source: ROD (Raman Open Database, RRUFF Raman data).
下载后统一存到工具的数据文件夹（参考谱库），可离线使用。||Downloads are stored in the tool's data folder (Reference library) for offline use.
输出内容：||What to output:
输出文件夹：||Output folder:
过渡宽度(cm-1)：||Transition width (cm-1):
退出||Exit
选择光谱文件（.jws / .csv / .spc / .jdx / .txt / .xlsx）||Choose spectrum files (.jws / .csv / .spc / .jdx / .txt / .xlsx)
选择参考峰位表（CSV：名称,峰位1,峰位2,…）||Choose a reference peak table (CSV: name,peak1,peak2,...)
选择参考谱所在文件夹||Choose the folder containing reference spectra
选择参考谱文件（可多选，支持 RRUFF 导出的 .txt / .csv）||Choose reference files (multi-select; RRUFF .txt / .csv supported)
选择参考谱文件（可多选；取消则改为选文件夹）||Choose reference files (multi-select; cancel to pick a folder instead)
选择文件…||Choose a file…
选择文件夹（会递归查找 .jws / .csv / .spc / .jdx）||Choose a folder (searched recursively for .jws / .csv / .spc / .jdx)
选择新的数据文件夹（建议放在本工具目录内）||Choose the new data folder (inside the tool folder is recommended)
选择自己下载的 RRUFF 数据包 zip||Choose an RRUFF package zip you downloaded
选择要批处理的文件夹（会递归查找光谱文件）||Choose the folder to batch (searched recursively for spectra)
选择要查看的 CSV 文件||Choose the CSV file to view
选择要转成图片的 CSV 文件（可多选）||Choose the CSV files to turn into images (multi-select)
选择要鉴定的光谱文件||Choose the spectrum file to identify
选择输出文件夹||Choose the output folder
配对比较（手动 / 自动）||Pairing (manual / auto)
重建索引||Rebuild index
预览||Preview
预览选中文件||Preview selected files
高级设置…||Advanced settings…
高级设置…（处理 / 校准 / 坐标轴 / 峰 / 图幅）||Advanced settings… (processing / calibration / axes / peaks / size)
（如 Zr Si，空格分隔；不知道就留空）||(e.g. Zr Si, space separated; leave blank if unknown)
（边缘线性过渡，避免出现台阶）||(linear edge transition to avoid steps)

# ===== 界面：下拉与列表项 =====
全部||All
分析结果文件夹||Results folder in the data folder
基线校正：||Baseline:
导数：||Derivative:
平滑方式：||Smoothing:
归一化：||Normalization:
成分||Chemistry
拉曼||Raman
拟合峰形：||Peak shape:
源文件夹下的 _转换结果||_转换结果 under the source folder
红外||IR

# ===== 界面：文件类型 =====
CSV 文件||CSV files
CSV/文本||CSV / text
ZIP 压缩包||ZIP archive
光谱文件||Spectrum files
所有文件||All files
文本文件||Text files
谱文件||Spectrum files

# ===== 界面：提示框 =====
A − k·B 需要选中 2 条光谱（按选中顺序，A 在前）。||A - k·B needs 2 spectra selected (in selection order, A first).
与参考谱没有重叠的波数区间，无法配对。||No overlapping wavenumber range with the reference; cannot pair.
与本地数据库谱线没有重叠波数区间，无法比对。||No overlapping wavenumber range with the local database spectra; cannot compare.
两条光谱没有重叠区间，无法相减。||The two spectra do not overlap; cannot subtract.
二维成像需要选中多个点位的扫描光谱（按文件顺序排布）。||2D imaging needs several point spectra (laid out in file order).
参考谱与待测谱没有重叠的波数区间，无法比对。||The reference and the measured spectrum do not overlap; cannot compare.
完成||Done
容量上限必须是数字（MB，0 = 不限制）。||The size limit must be a number (MB, 0 = unlimited).
导出 PNG 需要 Pillow（pip install pillow）。||Exporting PNG requires Pillow (pip install pillow).
导出失败||Export failed
导出完成||Export finished
当前没有可导出的预览图，请先选择文件或打开 CSV。||There is no preview to export; select a file or open a CSV first.
成功 %d 个，跳过 %d 个，失败 %d 个。\\n详见日志。||Succeeded %d, skipped %d, failed %d.\\nSee the log for details.
完成：成功 %d 个，跳过 %d 个，失败 %d 个||Done: succeeded %d, skipped %d, failed %d
下载参考谱：--db-get 编号[,编号...]||Download references: --db-get ID[,ID...]
完成：成功 %d / %d，保存在 %s||Done: succeeded %d / %d, saved to %s
最佳配对：%s||Best match: %s
报告图已保存：%s||Report figure saved: %s
现在可以鉴定未知光谱了：--identify 你的谱.csv||You can now identify an unknown spectrum: --identify your.csv
清理全部数据包：--cleanup ；只释放指定空间：--cleanup 500（MB）||Clear all packages: --cleanup ; free a specific amount: --cleanup 500 (MB)
所选光谱没有共同的波数区间，无法聚类。||The selected spectra share no common wavenumber range; cannot cluster.
所选光谱没有共同的重叠区间，无法平均。||The selected spectra share no common range; cannot average.
批处理||Batch
批处理中断，详见日志。||Batch interrupted; see the log.
报告已生成||Report generated
提示||Notice
数据库中没有找到“%s”。||"%s" was not found in the database.
文件里没有可用的光谱||No usable spectrum in the file
无法访问数据库：%s||Cannot access the database: %s
无法读取 CSV||Cannot read the CSV
未设置输出目录||No output folder set
未识别到峰，无法检索。可降低“峰灵敏阈值”。||No peaks found, cannot search. Try lowering the peak sensitivity threshold.
未选择输出内容||No output type selected
本地数据库为空||The local database is empty
检索失败||Search failed
没有可用的参考谱。||No usable reference spectra.
没有文件||No files
没能下载到参考谱。||Could not download reference spectra.
清理失败||Cleanup failed
瀑布图需要选中至少 2 条光谱。||The waterfall needs at least 2 spectra.
生成失败||Generation failed
生成报告失败||Report generation failed
目前没有已下载的数据包。||No data packages downloaded yet.
目录不存在：%s||Folder does not exist: %s
缺少组件||Missing component
聚类失败||Clustering failed
说明书已保存到：\\n%s||Guide saved to:\\n%s
请先在主界面左侧选中 1 条光谱，||Select 1 spectrum on the left first,
请先在左侧选中至少 2 条光谱（可 Ctrl / Shift 多选）。||Select at least 2 spectra on the left (Ctrl / Shift for multi-select).
请先在左侧选择 1 个待比对文件。||Select 1 file to compare on the left first.
请先在左侧选择 1 个待识别文件。||Select 1 file to search on the left first.
请先在左侧选择 1 个待配对的实测谱。||Select 1 measured spectrum to pair on the left first.
请先在左侧选择文件。||Select a file on the left first.
请先在左侧选择要写入报告的光谱文件。||Select the spectra to include in the report on the left first.
请先添加要转换的 .jws 文件或文件夹。||Add the .jws files or a folder to convert first.
请先选择 2 个文件（第一个作为 A，第二个作为 B）。||Select 2 files first (the first is A, the second is B).
请先选择至少 2 个文件。||Select at least 2 files first.
读取失败||Read failed
谱段替换需要选中至少 2 条光谱（第 1 条为目标谱）。||Range replacement needs at least 2 spectra (the first is the target).
转换完成！成功 %d 个，跳过 %d 个。||Conversion finished. Succeeded %d, skipped %d.
还没有可打开的输出目录。||No output folder to open yet.
重叠区间采样点太少。||Too few points in the overlapping range.
重叠区间采样点太少，无法平均。||Too few points in the overlapping range; cannot average.
阈值必须是数字。||The threshold must be a number.

# ===== 界面：运行日志 =====
   %-22s 最相似：%s（r=%.4f）||   %-22s most similar: %s (r=%.4f)
   %-26s 相关系数 %.4f  谱角 %.2f°||   %-26s corr %.4f  angle %.2f°
   %-30s F1 %5.1f%% (命中%d/%d 参考%d)  r=%.4f  谱角 %.2f°||   %-30s F1 %5.1f%% (hit %d/%d ref %d)  r=%.4f  angle %.2f°
   %-34s 峰数 %-4d 主峰 %s||   %-34s peaks %-4d main %s
   %-34s 相关系数 %.4f  谱角 %.2f°||   %-34s corr %.4f  angle %.2f°
   %8.1f  中心 %7.2f  FWHM %6.2f  面积 %9.1f  R2 %.3f||   %8.1f  center %7.2f  FWHM %6.2f  area %9.1f  R2 %.3f
   %d. %-24s 匹配 %.0f%% (%d/%d)||   %d. %-24s match %.0f%% (%d/%d)
   下载 %s → %s||   download %s -> %s
   下载失败 %s：%s||   download failed %s: %s
   下载的参考谱/数据包与分析结果都只写到这里；数据包超上限或磁盘不足会提醒。||   Downloaded references, packages and analysis results are written here only; you will be warned if the limit or disk space is exceeded.
   主成分散点：%s||   PCA scatter: %s
   分组表：%s||   group table: %s
   峰位匹配：%d/%d 已在容差 %.1f cm-1 内||   peak match: %d/%d within %.1f cm-1
   已把旧的“光谱数据库”内容迁移进数据文件夹：%d 项。||   Migrated the old "光谱数据库" content into the data folder: %d item(s).
   报告 → %s||   report -> %s
   报告图 → %s||   report figure -> %s
   报告图生成失败：%s||   report figure failed: %s
   数值表：%s||   value table: %s
   新手建议先读一遍：菜单【帮助】→【使用说明（完整手册）】。||   New users: please read [Help] -> [Guide (full manual)] first.
   树状图：%s||   dendrogram: %s
   汇总统计：%s||   summary stats: %s
   热图：%s||   heat map: %s
   相似度矩阵 / 平均 / 相减 / 谱段替换 / A−k·B / 瀑布图 / 聚类 / 二维成像。||   similarity matrix / average / subtract / range replacement / A−k·B / waterfall / clustering / 2D imaging.
   簇%d（%d 条）：%s%s||   cluster %d (%d spectra): %s%s
   输出到：%s（%s）||   output to: %s (%s)
   迭代多项式与滚动球基线）、矿物信息与峰归属库、数据文件夹与容量上限。||   iterative & rolling-ball baselines), mineral info and peak assignment, data folder and size limit.
CSV → PNG 失败：%s（%s）||CSV -> PNG failed: %s (%s)
CSV → PNG：%s（%d 列）||CSV -> PNG: %s (%d columns)
CSV 转图片完成，共 %d 个。||CSV to image finished, %d file(s).
RRUFF 按矿物抓取 "%s"：命中 %d 条，导出 %d 条到 %s||RRUFF bulk fetch "%s": %d hit, exported %d to %s
① 可读：.jws / .csv / .spc / .jdx / .txt / .xlsx；可出 Excel(图表) / PNG / CSV / 峰列表 / 峰拟合 / JCAMP-DX。||1. Reads .jws / .csv / .spc / .jdx / .txt / .xlsx; outputs Excel (chart) / PNG / CSV / peak list / peak fit / JCAMP-DX.
② 漏标的峰：在右侧图上左键点击即可手动补标（蓝色方块）；标错的峰（含自动峰）对准它右键即可删掉。||2. Missed peaks: left-click the plot to add one (blue square); to drop a wrong peak, including automatic ones, right-click on it.
③ 菜单“分析工具”：未知光谱检索（全库鉴定）/ 峰拟合 / 比对 / 峰位检索 /||3. "Analysis" menu: unknown-spectrum search (whole library) / peak fitting / compare / peak-position search /
④ 菜单“文件”：文件夹批处理、导出分析报告；菜单“设置”：高级设置（含尖峰去除、||4. "File" menu: batch folder, export analysis report; "Settings" menu: advanced settings (spike removal,
⑤ 数据文件夹：%s||5. Data folder: %s
⑥ 说明书：%s||6. Guide: %s
下载失败 %s：%s||Download failed %s: %s
与本地数据库（%d 条）比对 → %s||Compared with the local database (%d spectra) -> %s
二维成像：%d 个点位（%d×%d），%s 范围 %.4g ~ %.4g||2D imaging: %d points (%dx%d), %s range %.4g ~ %.4g
二维成像：%d 个点位（无有效数值）||2D imaging: %d points (no valid values)
交互相减已保存（k=%.2f）→ %s||Interactive subtraction saved (k=%.2f) -> %s
从文件夹添加了 %d 个光谱文件。||Added %d spectrum file(s) from the folder.
候选参考谱已导出：%s||Candidate reference exported: %s
光谱比对：%s 与 %d 个参考谱 → %s||Compare: %s vs %d reference spectra -> %s
全部完成：成功 %d 个，跳过 %d 个，失败 %d 个。||All done: succeeded %d, skipped %d, failed %d.
分析报告已生成（%d 条光谱）：%s||Analysis report generated (%d spectra): %s
导出失败 %s：%s||Export failed %s: %s
峰位检索（用 %d 个峰，容差 5 cm-1）→ %s||Peak-position search (%d peaks, tolerance 5 cm-1) -> %s
峰拟合失败：%s（%s）||Peak fitting failed: %s (%s)
峰拟合（%s，%d 个峰）→ %s||Peak fitting (%s, %d peaks) -> %s
已下载数据库谱线：%s||Downloaded database spectra: %s
已下载：%s||Downloaded: %s
已从本地数据库加入 %d 个文件到列表。||Added %d file(s) from the local database to the list.
已保存预览图片：%s||Preview image saved: %s
已在预览中打开 CSV：%s||Opened the CSV in the preview: %s
已导入 RRUFF 数据包：%s（索引 %d 条）||Imported the RRUFF package: %s (%d indexed)
已把「%s」设为峰位归属参考；峰列表/报告会多一列“归属”。||Set "%s" as the peak-assignment reference; the peak list / report gains an "assignment" column.
已添加 %d 个文件。||Added %d file(s).
已清理 %d 个数据包，释放 %.1f MB。||Removed %d package(s), freed %.1f MB.
已清空「%s」：删除 %d 项。||Cleared "%s": %d item(s) deleted.
已清空该文件的手动标注。||Cleared the manual annotations of this file.
已移除手动峰，剩余 %d 个。||Manual peak removed, %d left.
已删除手动峰：%.2f（剩余 %d 个）||Manual peak deleted: %.2f (%d left)
已删除自动峰：%.2f（该文件共隐藏 %d 个）||Automatic peak deleted: %.2f (%d hidden for this file)
已恢复 %d 个被删除的自动峰。||Restored %d deleted automatic peaks.
当前文件没有被删除的自动峰||No deleted automatic peaks for the current file.
图上没有可删除的标注峰||There is no annotated peak to delete on the chart.
附近没有峰，请点在峰上再右键||No peak nearby - click on a peak first, then right-click.
删除峰标注请只选择 1 个文件||Select exactly 1 file to delete peak annotations.
显示峰位数值||Show peak values
显示峰位数值（取消勾选只留虚线与标记）||Show peak values (untick to keep only the dashed lines and markers)
横坐标取各条谱波数范围的交集%s||X axis uses the intersection of every spectrum's wavenumber range %s
恢复自动峰||Restore auto peaks
恢复被删的自动峰||Restore deleted automatic peaks
平均完成（%d 条，%d 点）→ %s||Average done (%d spectra, %d points) -> %s
手动标注峰：%.2f（共 %d 个）||Manual peak: %.2f (%d total)
打开图片失败：%s||Cannot open the image: %s
批处理完成：共 %d 个，成功 %d，跳过 %d，失败 %d||Batch finished: %d total, %d succeeded, %d skipped, %d failed
批处理开始：%s||Batch started: %s
搬移失败 %s：%s||Move failed %s: %s
数据包容量上限已设为 %s MB。||Package size limit set to %s MB.
数据文件夹已切换到：%s||Data folder switched to: %s
未安装 Pillow（或缺少 ImageTk），无法在窗口内预览：%s||Pillow (or ImageTk) is missing; cannot preview in the window: %s
未知光谱检索结果：%s||Unknown-spectrum search result: %s
未知光谱检索（%s）：比对 %d 条，最佳候选 %s||Unknown-spectrum search (%s): compared %d, best candidate %s
检索到 %d 条，下载前 %d 条用于配对 …||Found %d records, downloading the top %d for pairing ...
正在检索数据库：%s …||Searching the database: %s ...
正在聚类分析（%d 条光谱，重采样 300 点）…||Clustering (%d spectra, resampled to 300 points)...
清理失败 %s：%s||Cleanup failed %s: %s
瀑布图：%s||Waterfall: %s
叠加图：%s||Overlay: %s
叠加图（%d 条）||Overlay (%d spectra)
多数据图叠加（%d 条）||Multi-dataset overlay (%d spectra)
ERR 叠加图 -> %s||ERR overlay -> %s
叠加图至少需要 2 条光谱，已跳过。||The overlay needs at least 2 spectra; skipped.
多数据图叠加（所选光谱）…||Multi-dataset overlay (selected spectra)…
多数据图叠加||Multi-dataset overlay
多数据图叠加需要选中至少 2 条光谱。||The overlay needs at least 2 spectra.
把选中的 %d 条光谱上下错开排列，每条一种颜色。||Stack the %d selected spectra one above another, one colour per dataset.
各条归一化到最大值 = 1（便于比较谱型）||Normalize each curve to max = 1 (easier to compare shapes)
谱线偏移：||Stack offset:
1.0 = 相邻谱线刚好不压线，越大越分开||1.0 = neighbouring curves just touch; larger means more separation
峰位合并容差(cm-1)：||Peak merging tolerance (cm-1):
留空 = 用“最小峰间距”；此距离内的峰算同一个||Blank = use "minimum peak separation"; peaks within this distance count as one
峰位跨谱合并：邻近的峰只画一条虚线、只标一个平均值||Merge peaks across spectra: nearby peaks share one dashed line and one averaged value
谱线偏移必须是数字。||The stack offset must be a number.
峰位合并容差必须是数字。||The peak merging tolerance must be a number.
多数据图叠加（堆叠排布）||Multi-dataset overlay (stacked)
堆叠偏移（瀑布图 / 叠加图）：||Stack offset (waterfall / overlay):
0.2 ~ 2.0，1.0 = 谱线刚好不压线||0.2 ~ 2.0; 1.0 = curves just touch
生成失败：%s||Render failed: %s
峰位虚线引到横坐标轴（自动峰 + 手动峰）||Draw a dashed line from each peak down to the x axis (auto + manual peaks)
多数据图叠加（堆叠排布 · 预览）||Multi-dataset overlay (stacked · preview)
各条归一化到最大值 = 1||Normalize each curve to max = 1
峰位跨谱合并（一峰一线一值）||Merge peaks across spectra (one peak, one line, one value)
刷新预览||Refresh preview
重新检测||Re-detect
导出 PNG||Export PNG
自动检测||Auto-detected
已手动添加峰位 %.1f。||Manually added peak at %.1f.
附近已有标注（%.1f），没有重复添加。||There is already a mark at %.1f; nothing else added.
请先勾选“峰位跨谱合并”，再手动增减峰位。||Tick "Merge peaks across spectra" before adding or removing peaks.
当前没有可删除的标注。||There is no mark to remove.
附近没有标注峰位，请点在虚线上再右键。||No mark nearby; right-click right on a dashed line.
已删除峰位标注 %.1f。||Removed the peak mark at %.1f.
已恢复为自动检测的峰位。||Restored the auto-detected peaks.
当前标注 %d 个峰位。||Currently %d peak(s) marked.
左键点图＝加一个峰位，右键＝删掉最近的一条||Left-click the plot = add a peak; right-click = remove the nearest one
未开启跨谱合并，不能手动增减||Merging is off, so peaks cannot be added or removed by hand
叠加图：%s（标注 %d 个峰位）||Overlay: %s (%d peaks marked)
预览需要 Pillow（pip install pillow）。||The preview needs Pillow (pip install pillow).
(留空 = 最小峰间距)||(blank = minimum peak separation)
不到 1 分钟||under a minute
约 %.0f 分钟||about %.0f min
约 %d 小时||about %d h
约 %d 小时 %d 分钟||about %d h %d min
剩余 %s||%s left
取消下载||Cancel download
状态||Status
类型||Type
数据包||Package
大小||Size
预计耗时||Estimated time
索引条数||Indexed rows
已下载||Downloaded
未下载||Not downloaded
　（预计耗时按上次实测速率 %.2f MB/s 估算）||  (estimated at the last measured %.2f MB/s)
正在下载 %s||Downloading %s
正在连接 %s …||Connecting to %s ...
正在连接数据库 …||Connecting to the database ...
正在导入 …||Importing ...
正在建立索引 %d/%d …||Building the index %d/%d ...
正在检索 …  这一步要连 solsa.crystallography.net||Searching ... this step contacts solsa.crystallography.net
正在检索数据库：%s …（在后台进行，界面不会卡住）||Searching the database: %s ... (runs in the background; the window stays responsive)
正在下载 %d/%d：%s||Downloading %d/%d: %s
正在下载候选谱线 %d/%d：%s||Downloading candidate spectra %d/%d: %s
检索到 %d 条，下载了 %d 条用于配对。||Found %d entries; downloaded %d of them for pairing.
下载已取消。||Download cancelled.
已取消。已下载的部分会保留，下次接着下。||Cancelled. What was already downloaded is kept and the next run continues from there.
在线检索 / 下载已取消。||Online search / download cancelled.
下载失败，没有取到任何谱线。||Download failed; no spectrum was retrieved.
正在准备批量抓取 …||Preparing the bulk fetch ...
批量抓取 "%s"：需下载 %d 个包，共约 %.0f MB||Bulk fetch of "%s": %d package(s) to download, about %.0f MB in total
正在下载 %s（第 %d/%d 个包，%s）||Downloading %s (package %d/%d, %s)
按矿物抓取已取消。||Bulk fetch by mineral cancelled.
抓取失败：%s||Bulk fetch failed: %s
并发连接数：||Parallel connections:
（1~%d，越大越快）||(1~%d; higher is faster)
代理：||Proxy:
（留空 = 自动用系统代理；none = 强制直连）||(blank = use the Windows system proxy automatically; "none" = force a direct connection)
直连||direct
直连（不用代理）||direct (no proxy)
　并发 %d 路 · %s||  %d connections · %s
代理 %s||proxy %s
服务器多次限流（HTTP %s），请把并发连接数调小些再试。||  The server throttled us repeatedly (HTTP %s); lower the number of parallel connections and retry.
并发连接数要在 1 ~ %d 之间。||Parallel connections must be between 1 and %d.
下载并发连接数：%d（可用 --dl-conns 1~%d 调整）||Download concurrency: %d (change it with --dl-conns 1~%d)
下载代理：%s||Download proxy: %s
（检测到 Windows 系统代理：%s，已自动使用）||(Windows system proxy detected: %s - used automatically)
提示：跨境高丢包链路下并发数几乎决定速度；有代理 / VPN 时走代理通常更快。||Tip: on a long, lossy international link the number of parallel connections almost determines the speed; if you have a proxy or VPN, going through it is usually much faster.
（%d 条失败，详见日志）||(%d failed; see the log)
没有可配对的实测谱，已下载的谱线存在本地库。||No measured spectrum to pair with; the downloaded spectra are in the local library.
这个包比较大||This package is large
%s 已经下载过了，要重新下载吗？||%s has already been downloaded. Download it again?
确认覆盖||Confirm overwrite
数据包 %s 下载完成，索引 %d 条||Package %s downloaded; %d rows indexed
服务器不支持断点续传（HTTP %s）||The server does not support resuming (HTTP %s)
连接提前中断（本段还差 %d 字节）||The connection dropped early (%d bytes still missing for this part)
下载失败（重试 %d 次仍未完成）：%s||Download failed (still incomplete after %d retries): %s
合并下载分片失败：%s||Merging the download parts failed: %s
下载不完整：应为 %d 字节，实际收到 %d 字节（已保留分片，可重试续传）||Incomplete download: expected %d bytes, got %d (parts kept; the next run resumes)
下载的压缩包打不开（传输中损坏），已删除，请重新下载。||The downloaded archive cannot be opened (damaged in transit); it has been deleted, please download again.
响应过大（超过 %d MB），已放弃||Response too large (over %d MB); aborted
网络请求失败（重试 %d 次）：%s||Network request failed (after %d retries): %s
光谱数据不完整：应为 %d 点，实际收到 %d 点（可稍后重试）||Incomplete spectrum data: expected %d points, got %d (try again later)
%s 约 %s，按实测速率预计要 %s。\\n\\n下载在后台进行，窗口不会卡住，随时可以点【取消】；中断后已下载的部分会保留，下次接着下，不会从头再来。\\n\\n现在开始下载吗？||%s is about %s; at the measured rate that takes roughly %s.\\n\\nThe download runs in the background so the window stays responsive, and you can press "Cancel" at any time. Whatever has already been downloaded is kept, so the next run continues from there instead of starting over.\\n\\nStart the download now?
== 批量配对汇总 ==\\n||== Batch pairing summary ==\\n
== 批量未知谱鉴定汇总 ==\\n||== Batch identification summary ==\\n
批量配对汇总||Batch pairing summary
批量未知谱鉴定汇总||Batch identification summary
参考谱集：%s||Reference set: %s
参考谱集：%s\\n||Reference set: %s\\n
实测谱||Measured spectra
实测谱：%d 条\\n||Measured spectra: %d\\n
实测谱：%d 条；峰位匹配容差 %.0f cm-1\\n||Measured spectra: %d; peak matching tolerance %.0f cm-1\\n
参与检索条目数：%d||Reference records searched: %d
综合分 = 0.5×F1 + 0.5×强峰命中率（与「未知谱鉴定」同一口径）；这些只是辅助指标，最终是哪个物相请自行核对峰位与谱型。||Score = 0.5 x F1 + 0.5 x strong-peak hit rate (the same scoring as "Unknown spectrum search"); these are advisory metrics only - check the peaks and the profile yourself before deciding the phase.
峰位匹配容差||Peak matching tolerance
排序||Ordering
按综合分升序（最可疑的在前）||Ascending composite score (most doubtful first)
按最佳候选综合分升序（最可疑的在前）||Ascending best-candidate composite score (most doubtful first)
逐条配对结果||Per-spectrum pairing result
逐条鉴定结果||Per-spectrum identification result
（按综合分升序排列：最可疑的排在最前面）\\n||(sorted by ascending composite score: the most doubtful come first)\\n
（按最佳候选的综合分升序排列：最可疑的排在最前面）\\n||(sorted by ascending best-candidate composite score: the most doubtful come first)\\n
匹配良好，可作同一物相（参考判读）||Good match; can be treated as the same phase (advisory only)
部分匹配，建议核对峰位对照（参考判读）||Partial match; check the peak table (advisory only)
匹配很差，很可能不是同一物相（参考判读）||Poor match; probably not the same phase (advisory only)
没有可比对的参考谱||No reference spectrum to compare against
与参考谱没有可比对的峰（波数区间不重叠或参考谱没有峰）||No comparable peaks with the reference (wavenumber ranges do not overlap, or the reference has no peaks)
这条光谱没有识别到峰||No peaks detected in this spectrum
没有找到候选||No candidate found
参考判读||Advisory reading
最佳参考谱||Best reference
最佳候选||Best candidate
强峰总数||Strong peaks total
命中峰数||Matched peaks
实测峰数||Measured peaks
参考峰数||Reference peaks
平均偏差(cm-1)||Mean deviation (cm-1)
谱角(度)||Spectral angle (deg)
领先第二名(分)||Lead over runner-up (pts)
文件名||File name
数据点数||Data points
识别峰数||Peaks detected
F1 为峰位匹配 F1（容差 %.0f cm-1，命中<2 记 0）。||F1 is the peak-matching F1 (tolerance %.0f cm-1; fewer than 2 matches scores 0).
这些只是辅助指标，最终是哪个物相请自行核对峰位与谱型。||These are advisory metrics only; please check the peaks and the profile yourself before deciding the phase.
这些只是辅助指标，最终是哪个物相请自行核对峰位与谱型。\\n||These are advisory metrics only; please check the peaks and the profile yourself before deciding the phase.\\n
批量配对（文件夹 × 参考谱）||Batch pairing (folder x reference spectra)
批量鉴定（文件夹逐条鉴定）||Batch identification (identify a whole folder)
批量配对（文件夹 × 参考谱）…||Batch pairing (folder x reference spectra)...
批量鉴定（文件夹逐条鉴定）…||Batch identification (identify a whole folder)...
把一整个文件夹的实测谱逐条与参考谱打分，汇总成一张表；||Scores every measured spectrum in a folder against the reference set and summarises them in one table;
表里只有辅助指标，最后是哪个物相由你自己判断。||The table holds advisory metrics only; you decide which phase it really is.
把一整个文件夹的谱逐条在全库里检索，汇总成一张表；||Searches every spectrum in a folder against the whole library and summarises them in one table;
每条给出最佳候选与参考判读，最终判断由你自己做。||Each row gives the best candidate and an advisory reading; the final call is yours.
实测谱文件夹：||Measured-spectra folder:
选择放实测谱的文件夹||Choose the folder holding the measured spectra
浏览…||Browse...
（留空 = 用主界面左侧选中的光谱）||(blank = use the spectra selected in the main window)
参考谱：||Reference spectra:
本地参考谱库||Local reference library
本地参考谱库（%d 条）||Local reference library (%d)
指定文件夹 / 文件：||Specified folder / file:
选择参考谱文件夹||Choose the reference folder
候选个数：||Candidates:
开始批量处理||Run batch
打开汇总表 CSV||Open summary CSV
打开汇总报告 HTML||Open summary report (HTML)
选好文件夹后点【开始批量处理】。||Pick a folder, then press "Run batch".
结果（按综合分升序，最可疑的在前）||Results (ascending composite score, most doubtful first)
主界面选中||main-window selection
批量结果||batch result
正在处理 %d/%d：%s||Processing %d/%d: %s
正在批量配对：%d 条实测谱 × %d 条参考谱…||Batch pairing: %d measured spectra x %d reference spectra...
正在批量鉴定：%d 条谱 × %d 个数据包…||Batch identification: %d spectra x %d data packages...
没有读到实测谱：请选一个文件夹，或先在主界面左侧选中光谱。||No spectra loaded: choose a folder, or select spectra in the main window first.
请先指定参考谱文件夹或文件。||Specify the reference folder or file first.
文件不存在：%s||File not found: %s
批量处理失败：%s||Batch run failed: %s
批量配对完成：%d 条实测谱 × %d 条参考谱，其中 %d 条匹配很差（<35 分）||Batch pairing done: %d measured spectra x %d references, %d of them matched poorly (< 35)
完成：%d 条；其中 %d 条匹配很差（<35 分），已排在最前。汇总表：%s||Done: %d spectra; %d matched poorly (< 35) and are listed first. Summary: %s
批量鉴定完成：%d 条谱，比对 %d 条参考记录，其中 %d 条没有可靠匹配（<35 分）||Batch identification done: %d spectra, %d reference records compared, %d without a reliable match (< 35)
完成：%d 条；其中 %d 条没有可靠匹配（<35 分），已排在最前。汇总表：%s||Done: %d spectra; %d without a reliable match (< 35) and listed first. Summary: %s
还没有可用于检索的特征索引：请先在【未知光谱检索】窗口里建立索引。||No feature index is available yet: build one in the "Unknown spectrum search" window first.
批量配对：%d 条实测谱 × %d 条参考谱（%s）||Batch pairing: %d measured spectra x %d reference spectra (%s)
批量配对结果（按综合分升序，最可疑的排在最前）：||Batch pairing result (ascending composite score, most doubtful first):
批量鉴定：%d 条未知谱；检索范围 %s；已建索引的数据包 %d 个：%s||Batch identification: %d unknown spectra; search scope %s; %d packages indexed: %s
批量鉴定结果（按最佳候选综合分升序，最可疑的排在最前）：||Batch identification result (ascending best-candidate score, most doubtful first):
汇总表：%s||Summary table: %s
汇总报告：%s||Summary report: %s
没有找到参考谱：用 --pair-ref 指定参考谱文件或文件夹，||No reference spectra found: use --pair-ref to point at a file or folder,
或先用 --db-get 把参考谱下载到本地库。||or download references into the local library with --db-get first.
参考谱一条也读不出来（需要两列文本 / CSV）：%s||Not one reference spectrum could be read (two-column text / CSV required): %s
没有读到待配对的实测谱：%s||No measured spectra to pair were loaded: %s
HTML 汇总生成失败：%s||Generating the HTML summary failed: %s
特征索引：%s -> %d 条||Feature index: %s -> %d entries
相似度矩阵（%d×%d）→ %s||Similarity matrix (%dx%d) -> %s
相减完成：A=%s  B=%s（%d 点）→ %s||Subtraction done: A=%s  B=%s (%d points) -> %s
聚类完成：%d 条光谱 → %d 个簇（自动分割阈值 %.3f）||Clustering done: %d spectra -> %d clusters (auto cut %.3f)
该位置已有自动峰（%.1f），无需重复标注。||An automatic peak already exists there (%.1f); no need to annotate.
说明书中没有找到“%s”。||"%s" was not found in the guide.
说明书已导出：%s||Guide exported: %s
读取失败：%s（%s）||Read failed: %s (%s)
谱段替换完成（替换 %d 点）→ %s||Range replacement done (%d points) -> %s
配对完成（%s）：最佳 %s，峰位匹配 F1 %.0f%%（命中 %d/%d 实测峰，参考 %d 峰），相关系数 %.4f||Pairing done (%s): best %s, peak-match F1 %.0f%% (hit %d/%d measured peaks, %d reference peaks), corr %.4f

# ===== 界面：状态提示 =====
%d 个文件 / %.1f MB||%d file(s) / %.1f MB
%s   |   %d 列数据  ·  %d 点||%s   |   %d data columns  ·  %d points
%s（%d 点，%d 个峰）||%s (%d points, %d peaks)
k = %.2f   区间 %.1f ~ %.1f cm-1   剩余最大强度 %.4g||k = %.2f   range %.1f ~ %.1f cm-1   residual max %.4g
下载 %s %.1f MB||Downloading %s %.1f MB
下载中 %.1f%s MB||Downloading %.1f%s MB
下载失败 %s：%s||Download failed %s: %s
下载失败，无法配对。||Download failed; cannot pair.
下载失败：%s||Download failed: %s
下载完成，正在建立索引（谱线较多时需一会儿）…||Download finished, building the index (may take a while for many spectra)...
两条光谱没有重叠波数区间。||The two spectra share no overlapping wavenumber range.
候选导出失败，无法配对。||Cannot export the candidate; pairing aborted.
候选清单已保存：%s||Candidate list saved: %s
共 %d 条结果。按住 Ctrl 多选后点“下载选中谱线”。||%d results. Ctrl-click to multi-select, then "Download selected spectra".
合计 %d 个文件，占用 %.1f MB%s||%d file(s) in total, %.1f MB%s
完成，已保存热图与数值表。||Done; heat map and value table saved.
完成：%s 共 %d 条，现在可以检索了。||Done: %s with %d entries; you can search now.
对比图已保存：%s||Overlay saved: %s
对比图生成失败：%s||Overlay failed: %s
导入失败：%s||Import failed: %s
导入成功：%s（索引 %d 条）。||Imported: %s (%d indexed).
导出候选失败：%s||Exporting candidates failed: %s
导出失败：%s||Export failed: %s
尚未选择文件||No file selected
已下载 %d 条到：%s||Downloaded %d to: %s
已保存：%s||Saved: %s
已加入 %d 个文件到左侧列表，可绘图 / 比对 / 平均 / 相减。||Added %d file(s) to the list; you can plot / compare / average / subtract.
已取消。可在上方选择数据包手动下载。||Cancelled. You can tick packages above and download manually.
已在右侧预览（尚未保存）。||Shown in the preview (not saved yet).
已完成与 %d 条数据库谱线的配对，详见主界面日志与右侧叠加图。||Paired with %d database spectra; see the main log and the overlay on the right.
已导出 %d 条到本地库：%s||Exported %d to the local library: %s
已导出到参考谱库：%s（可在【分析工具 → 配对比较】里继续比对）||Exported to the reference library: %s (continue in [Analysis -> Pairing])
已按矿物批量抓取 "%s"：命中 %d 条，导出 %d 条到本地库。||Bulk fetch "%s": %d hit, exported %d to the local library.
已用 %d 条 RRUFF 参考谱完成配对，见主界面日志与报告图。||Paired using %d RRUFF references; see the main log and the report figure.
建立索引失败（%s）：%s||Index building failed (%s): %s
建立索引失败：%s||Index building failed: %s
所有数据包中都没有找到 "%s"。||"%s" was not found in any package.
所选的包还没有特征索引，正在先建立…||The selected package has no feature index yet; building it first...
手动标注请只选择 1 个文件||Select exactly 1 file for manual annotation
指定波数必须是数字。||The wavenumber must be a number.
数据包 %d 个 / %.1f MB　上限 %s%s||%d package(s) / %.1f MB　limit %s%s
数据包占用 %.1f MB / %d 个文件%s||Packages use %.1f MB / %d file(s)%s
无法读取：%s||Cannot read: %s
最佳候选没有读到原始谱（可能未精算），先勾选“精算”重跑一次。||No raw spectrum loaded for the best candidate (probably not re-scored); tick "re-score" and run again.
本地数据库还是空的，请先下载谱线。||The local database is still empty; download some spectra first.
检索到 %d 条。||Found %d records.
检索失败：%s||Search failed: %s
检索失败：%s（请检查网络）||Search failed: %s (check your network)
检索完成：共比对 %d 条参考记录，列出 %d 个候选。||Search finished: compared %d references, listing %d candidates.
正在下载 %s …||Downloading %s ...
正在下载 %s … %.1f%s MB||Downloading %s ... %.1f%s MB
正在下载 %s 并配对 …||Downloading %s and pairing ...
正在下载 12 MB 数据包 …||Downloading the 12 MB package ...
正在为 %s 建立特征索引（每条谱都要解析一次）…||Building the feature index for %s (every spectrum is parsed once)...
正在导入并建立索引 …||Importing and building the index ...
正在导出 %d 条 …||Exporting %d ...
正在导出 %s …||Exporting %s ...
正在建立索引 %s …||Building the index for %s ...
正在批处理，请稍候…（详细进度见主界面日志）||Batching, please wait... (see the main log for details)
正在检索 …||Searching ...
正在检索…||Searching...
没有导出成功，无法配对。||Nothing was exported; cannot pair.
特征索引建立失败或不完整，无法检索。||The feature index is missing or incomplete; cannot search.
目标谱与来源谱不能相同。||The target and source spectra must differ.
索引失败：%s||Indexing failed: %s
索引完成：%d 条。||Indexed: %d entries.
索引已就绪，可以开始检索。||Index ready; you can search now.
结论：||Conclusion:
行数 / 列数必须是整数。||Rows / columns must be integers.
该区间内没有数据点，请检查波数范围。||No data points in that range; check the wavenumber range.
请先关闭本窗口，在主界面左侧选择 1 个实测谱。||Please close this window and select 1 measured spectrum on the left.
请先勾选至少一个数据包。||Tick at least one data package first.
请先在主界面左侧选中 1 条未知光谱，或点【选择文件…】。||Select 1 unknown spectrum on the left, or click [Choose a file...].
请先在候选表里选中一条。||Select a row in the candidate table first.
请先在列表中选中条目。||Select an item in the list first.
请先在检索结果里选中条目。||Select an item in the results first.
请先在检索结果里选中条目（可 Ctrl/Shift 多选）。||Select items in the results first (Ctrl/Shift for multi-select).
请先选中一个数据包。||Select a data package first.
请填写有效的替换区间（起、止波数）。||Enter a valid replacement range (start and end wavenumber).
请至少勾选一种输出内容。||Tick at least one output type.
请输入检索词，例如 quartz / SiO2 / 石英||Enter a search term, e.g. quartz / SiO2 / 石英
请输入矿物名（如 Zircon / Quartz）或 RRUFF 编号（如 R050034）。||Enter a mineral name (e.g. Zircon / Quartz) or an RRUFF ID (e.g. R050034).
请输入矿物名，例如 Zircon / Quartz。||Enter a mineral name, e.g. Zircon / Quartz.
还没有可用的候选。||No candidates available yet.
还没有检索结果。||No search results yet.
需要 %d 个点位，当前只有 %d 条光谱。||%d points are needed but only %d spectra are available.

# ===== 命令行输出 =====
      特征峰：%s||      characteristic peaks: %s
   %-10s %5d 个文件  %9.2f MB   %s||   %-10s %5d file(s)  %9.2f MB   %s
   合计：%d 个文件，%.2f MB||   total: %d file(s), %.2f MB
   完成：%d 条特征 -> %s||   done: %d features -> %s
  ! 出图失败 %s：%s||  ! plot failed %s: %s
  ! 跳过 %s：%s||  ! skipped %s: %s
  %2d. %-32s F1 %5.1f%% (命中%d/%d 参考%d)  r=%.4f  谱角 %.2f°||  %2d. %-32s F1 %5.1f%% (hit %d/%d ref %d)  r=%.4f  angle %.2f°
  %2d. %-40s 相关系数 %.4f  谱角 %.2f°||  %2d. %-40s corr %.4f  angle %.2f°
  主峰位示例：||  example main peak positions:
  实测峰位    参考峰位    偏差     匹配||  measured    reference   dev       match
  容量上限：%.0f MB%s||  size limit: %.0f MB%s
  容量上限：不限制||  size limit: unlimited
  所在磁盘可用：%.1f GB / 共 %.1f GB||  disk free: %.1f GB / %.1f GB total
  数据包：%d 个文件，占用 %.1f MB||  packages: %d file(s), %.1f MB
  簇%d（%d 条）：%s%s||  cluster %d (%d spectra): %s%s
ERR 瀑布图 -> %s||ERR waterfall -> %s
OK  %s  ->  %s  (%d 条曲线, %d 点)||OK  %s  ->  %s  (%d curves, %d points)
OK  %s  ->  %s  (%d 通道, %d 点, %.3f ~ %.3f)||OK  %s  ->  %s  (%d channels, %d points, %.3f ~ %.3f)
RRUFF 数据包（来源 www.rruff.net/zipped_data_files）：||RRUFF data packages (source www.rruff.net/zipped_data_files):
下载失败：%s||Download failed: %s
与本地数据库比对（库内 %d 条）：||Compared with the local database (%d spectra):
与本地数据库没有可比较的重叠区间。||No comparable overlap with the local database.
主成分散点：%s||PCA scatter: %s
二维成像完成：%d 个点位，%s 范围 %.4g ~ %.4g||2D imaging done: %d points, %s range %.4g ~ %.4g
二维成像网格无效：--map 行,列（例如 --map 5,5）。||Invalid imaging grid: --map rows,cols (e.g. --map 5,5).
二维成像需要 %d×%d = %d 个点位，当前只读到 %d 条。||2D imaging needs %dx%d = %d points but only %d spectra were read.
候选清单：%s||Candidate list: %s
元素筛选：必须含 [%s]；必须不含 [%s]||Element filter: must contain [%s]; must not contain [%s]
共 %d 条结果（数据来源：ROD / RRUFF 拉曼库）：||%d results (source: ROD / RRUFF Raman library):
共比对 %d 条参考记录，以下按可信度排序：||Compared %d reference records; ranked by confidence below:
内置矿物表里没有「%s」。可用 --mineral-search <名称> 查找。||"%s" is not in the built-in mineral table. Use --mineral-search <name>.
内置矿物表里没有匹配（检索方式：%s）：%s||No match in the built-in mineral table (mode: %s): %s
内置表共 %d 种常见矿物，可用 --mineral-search 名称 模糊查找。||The built-in table has %d common minerals; use --mineral-search <name>.
分组表：%s||Group table: %s
可下载更大的数据包后重试（每个约 1~2 分钟）：||Download a larger package and retry (each takes about 1-2 minutes):
命中 %d 条，已导出 %d 条（同类优先 Processed）到本地库：%s||%d hits, exported %d (Processed preferred) to the local library: %s
大小||Size
完成：%d 条，保存在 %s||Done: %d entries saved to %s
完成：共 %d 个文件，成功 %d，跳过 %d，失败 %d||Done: %d file(s) total, %d succeeded, %d skipped, %d failed
容量上限必须是数字（单位 MB，0 表示不限制）。||The size limit must be a number in MB (0 = unlimited).
对比图生成失败：%s||Overlay failed: %s
对比图：  %s||Overlay:  %s
导入失败：%s → %s||Import failed: %s -> %s
导入成功：%s → %s（索引 %d 条）||Imported: %s -> %s (%d indexed)
导入自己的包：--import-pkg 路径.zip||Import your own package: --import-pkg path.zip
导出到本地库：--rruff-export Zircon||Export to the local library: --rruff-export Zircon
导出说明书失败：%s||Exporting the guide failed: %s
已下载的数据包里没有找到 "%s"。||"%s" was not found in the downloaded packages.
已删除 %d 个数据包，释放 %.1f MB。||Removed %d package(s), freed %.1f MB.
所选光谱没有共同的波数区间，无法聚类。||The selected spectra share no common wavenumber range; cannot cluster.
批处理失败：%s||Batch failed: %s
批处理目录：%s||Batch folder: %s
报告已保存：%s||Report saved: %s
报告已生成（%d 条光谱）：%s||Report generated (%d spectra): %s
按%s检索 "%s"：命中 %d 种矿物||Search %s "%s": %d minerals found
按矿物批量抓取 "%s"：命中 %d 条，导出 %d 条到 %s||Bulk fetch "%s": %d hit, exported %d to %s
按矿物检索：  --rruff-search Zircon（也可直接输 RRUFF 编号，如 R050034）||Search by mineral:  --rruff-search Zircon (or enter an RRUFF ID such as R050034)
数值表：%s||Value table: %s
数据包 key||Package key
数据包目录：%s||Package folder: %s
数据库中没有找到：%s||Not found in the database: %s
数据文件夹已设置为：%s||Data folder set to: %s
数据文件夹：%s||Data folder: %s
文件夹不存在：%s||Folder does not exist: %s
未找到 "%s"。请先用 --rruff-get 下载数据包（或 --rruff-list 查看）。||"%s" was not found. Use --rruff-get to download a package (or --rruff-list to see the list).
未知光谱：%s（%d 点，%d 个峰）||Unknown spectrum: %s (%d points, %d peaks)
未知数据包：%s||Unknown package: %s
未知的数据包 key：%s（用 --rruff-list 查看可选值）||Unknown package key: %s (use --rruff-list for valid values)
本地数据包里没有 "%s" 的记录，或尚未下载数据包。||No "%s" records in the local packages, or no package downloaded yet.
本地数据库为空，请先用 --db-get 下载参考谱。||The local database is empty; use --db-get to download references first.
本地数据库：%s（%d 条）||Local database: %s (%d spectra)
本地还没有 RRUFF 数据包，先下载最小的（未评级·非定向，12 MB）…||No RRUFF package yet; downloading the smallest one (unrated/unoriented, 12 MB)...
本地还没有数据包。请先 --rruff-get（或 --import-pkg）下载。||No packages yet. Use --rruff-get (or --import-pkg) first.
查看占用/清理：--cache-limit ；--cleanup 500||Usage / cleanup: --cache-limit ; --cleanup 500
树状图：%s||Dendrogram: %s
检索 "%s" -> %d 条||Search "%s" -> %d records
检索失败：%s||Search failed: %s
检索范围：%s；已建索引的数据包 %d 个：%s||Search scope: %s; %d indexed package(s): %s
正在下载 %s（%s）…||Downloading %s (%s)...
正在为 %s（%s）建立特征索引…||Building the feature index for %s (%s)...
正在建立索引 …||Building the index ...
正在检索…||Searching...
汇总统计：%s||Summary stats: %s
没有可清理的数据包。||No packages to clean up.
没有找到任何光谱文件（支持 .jws / .csv / .spc / .jdx / .txt / .xlsx）||No spectrum files found (supported: .jws / .csv / .spc / .jdx / .txt / .xlsx)
没有读到光谱：%s||No spectrum read: %s
瀑布图至少需要 2 条光谱，已跳过。||The waterfall needs at least 2 spectra; skipped.
热图：%s||Heat map: %s
状态||Status
生成报告失败：%s||Report generation failed: %s
类型||Type
索引完成：%d 条||Indexed: %d entries
结　论：%s||Conclusion: %s
聚类完成：%d 条光谱 → %d 个簇（自动分割阈值 %.3f）||Clustering done: %d spectra -> %d clusters (auto cut %.3f)
聚类至少需要 2 条光谱（请给出文件或文件夹）。||Clustering needs at least 2 spectra (give files or a folder).
说明||Description
说明书已导出：%s||Guide exported: %s
请先：--rruff-get <数据包key> 然后 --identify-index（为已下载包建索引）||First: --rruff-get <package key>, then --identify-index (to index downloaded packages)
读取待比对文件失败：%s||Failed to read the file to compare: %s
读取待配对文件失败：%s||Failed to read the file to pair: %s
输出内容：  %s||Output:  %s
输出目录：  %s||Output folder:  %s
还没有可用于检索的特征索引。||No feature index available for searching yet.
配对排名（本地库 %d 条，按峰位匹配 F1 优先）：||Pairing ranking (local library, %d spectra, by peak-match F1):
（下载的参考谱、RRUFF 数据包、分析结果都保存在此目录）||(downloaded references, RRUFF packages and analysis results are all kept here)
（已超限，建议清理）||(over limit; cleanup recommended)

# ===== 错误消息 =====
JCAMP-DX 文件里没有可用的数据点||No usable data points in the JCAMP-DX file
SPC 声明了 X 数组但数据段不足||SPC declares an X array but the data block is too short
SPC 大端格式（fversn=0x4C）暂不支持||Big-endian SPC (fversn=0x4C) is not supported yet
SPC 子文件 X 数组不完整||Incomplete X array in the SPC subfile
SPC 数据段长度不足（16 位）||SPC data block too short (16-bit)
SPC 数据段长度不足（32 位）||SPC data block too short (32-bit)
SPC 数据段长度不足（浮点）||SPC data block too short (float)
SPC 文件过小或已损坏||SPC file is too small or corrupted
SPC 未声明数据点数，无法解析||SPC does not declare the number of points; cannot parse
SPC 缺少子文件头，文件可能被截断||Missing SPC subfile header; the file may be truncated
不是有效的 OLE2/复合文档，可能不是 .jws 文件||Not a valid OLE2 compound document; this may not be a .jws file
参考谱数据不足：%s||Not enough reference data: %s
导入成功但建立索引失败：%s||Imported but indexing failed: %s
导出 Excel 需要 openpyxl 组件；请改用 CSV，或先安装 openpyxl||Exporting Excel needs openpyxl; use CSV instead or install openpyxl
导出 PNG 需要 Pillow 组件（pip install pillow）||Exporting PNG needs Pillow (pip install pillow)
导出报告图需要 Pillow 组件（pip install pillow）||Exporting the report figure needs Pillow (pip install pillow)
尚未下载数据包：%s||Package not downloaded yet: %s
数据库返回内容无法解析（可能网络被拦截）||Cannot parse the database response (the network may be blocked)
数据点太少，无法导出 JCAMP-DX||Too few data points to export JCAMP-DX
文件不存在：%s||File does not exist: %s
文件里没有找到可用的数值数据||No usable numeric data in the file
文件里没有找到可用的数值数据（需第一列为数字）||No usable numeric data in the file (the first column must be numeric)
暂不支持该 SPC 版本（fversn=0x%02X，仅支持新版 0x4B）||Unsupported SPC version (fversn=0x%02X; only 0x4B is supported)
未取得有效光谱数据：%s||No valid spectral data: %s
未找到 Root Entry，文件结构异常||Root Entry not found; the file structure is abnormal
未知的 RRUFF 数据包：%s||Unknown RRUFF package: %s
没有可绘制的数据||No data to plot
没有读到任何光谱数据||No spectral data was read
没有读到任何光谱，无法生成报告||No spectra were read; cannot generate the report
没有读到数据||No data was read
没有配对结果可绘制||No pairing result to plot
缺少 DataInfo 流，无法确定横坐标轴||Missing the DataInfo stream; cannot determine the X axis
缺少 Y-Data 流，文件里没有光谱数据||Missing the Y-Data stream; no spectral data in the file
至少需要 2 条光谱才能聚类||At least 2 spectra are needed for clustering
读取 Excel 需要 openpyxl 组件（pip install openpyxl）||Reading Excel needs openpyxl (pip install openpyxl)
这条光谱没有识别到峰，无法检索（可调低峰灵敏阈值后重试）||No peaks were found in this spectrum; cannot search (try lowering the peak sensitivity threshold)

# ===== 导出文件：表头与图注 =====
== 未知光谱检索：%s ==\\n||== Unknown-spectrum search: %s ==\\n
== 配对排名（%s，共 %d 条参考谱；按峰位匹配 F1 优先排序）==\\n||== Pairing ranking (%s, %d reference spectra; ranked by peak-match F1) ==\\n
== 配对排名（本地库 %d 条；按峰位匹配 F1 优先排序）==\\n||== Pairing ranking (local library, %d spectra; ranked by peak-match F1) ==\\n
二维成像 · %s||2D imaging · %s
参与检索条目数：%d\\n||Records searched: %d\\n
实测峰位,参考峰位,偏差(cm-1),是否匹配\\n||measured,reference,deviation(cm-1),matched\\n
层次聚类树状图||Hierarchical clustering dendrogram
峰位,拟合中心,拟合峰高,拟合半高宽FWHM,峰面积,峰形,混合系数,拟合R2,来源\\n||position,center,height,FWHM,area,shape,eta,R2,source\\n
排名,候选,匹配率(%),匹配峰数,参考峰数\\n||rank,candidate,match rate(%),matched peaks,reference peaks\\n
排名,参考谱,峰位匹配F1(%),命中峰数,实测峰数,参考峰数,相关系数,谱角(度)\\n||rank,reference,peak-match F1(%),hits,measured peaks,reference peaks,corr,angle(deg)\\n
排名,参考谱,相关系数,谱角(度),重叠下限,重叠上限\\n||rank,reference,corr,angle(deg),overlap min,overlap max\\n
排名,数据库谱线,相关系数,谱角(度)\\n||rank,database spectrum,corr,angle(deg)\\n
文件,数据点数,识别峰数,主峰位(cm-1),主峰强度,主峰相对强度(%),主峰半高宽FWHM\\n||file,points,peaks,main position(cm-1),main intensity,main rel. intensity(%),main FWHM\\n
文件,簇号,PC1,PC2,PC3\\n||file,cluster,PC1,PC2,PC3\\n
瀑布图（%d 条）||Waterfall (%d spectra)
瀑布图（%d 条，归一化后纵向错开）||Waterfall (%d spectra, normalized and offset)
相似度(相关系数),||similarity (corr),
第%d行,||row %d,
结论：%s\\n||Conclusion: %s\\n
行\\列,||row\\col,
评分说明：综合分 = 0.5×F1 + 0.5×强峰命中率；||Scoring: score = 0.5×F1 + 0.5×strong-peak hit rate;
通道,峰位,强度,相对强度(%),半高宽FWHM,峰突出度,来源\\n||channel,position,intensity,rel. intensity(%),FWHM,prominence,source\\n
通道,峰位,拟合中心,拟合峰高,拟合半高宽FWHM,峰面积,峰形,混合系数,拟合R2,来源\\n||channel,position,center,height,FWHM,area,shape,eta,R2,source\\n

# ===== 其它界面文案 =====
 | 通道%d|| | channel %d
%.1f cm-1 附近峰强度||intensity of the peak near %.1f cm-1
%d 个文件||%d file(s)
%d 通道  ·  %d 点  ·  %.3f ~ %.3f  ·  强度 %.3f ~ %.3f||%d channels  ·  %d points  ·  %.3f ~ %.3f  ·  intensity %.3f ~ %.3f
%d 阶导数||derivative order %d
%s · 峰表||%s · peak table
%s（参考 %.1f cm-1）||%s (reference %.1f cm-1)
%s（未命名样品）||%s (unnamed sample)
(%d/%d) – 已存在，跳过：%s||(%d/%d) - already exists, skipped: %s
(%d/%d) ✔ %s  →  %s   [CSV→PNG, %d 列]||(%d/%d) OK %s  ->  %s   [CSV->PNG, %d columns]
(%d/%d) ✔ %s  →  %s   [输出 %s, %d 通道, %d 点, %.3f~%.3f]||(%d/%d) OK %s  ->  %s   [output %s, %d channels, %d points, %.3f~%.3f]
,归属||,assignment
16 位定点||16-bit fixed point
1=不平滑||1 = no smoothing
32 位定点||32-bit fixed point
3~7 即可||3-7 is fine
<div class="note">图片缺失：%s</div>||<div class="note">Missing image: %s</div>
<div class="note">（无数据）</div>||<div class="note">(no data)</div>
<h2>说明</h2><ul>||<h2>Notes</h2><ul>
F1 为峰位匹配 F1（容差 %.0f cm-1，命中<2 记 0）；||F1 is the peak-position match F1 (tolerance %.0f cm-1; 0 if fewer than 2 hits);
JASCO 光谱(*.jws)||JASCO spectra (*.jws)
拉曼光谱工具 · 转换 / 分析 / 鉴定||Raman Spectrum Toolkit · Convert / Analyze / Identify
RRUFF 全量拉曼库 · 下载 / 按矿物检索||RRUFF full Raman library · download / search by mineral
RRUFF 数据包已占用 %.0f MB，超过设定的上限 %.0f MB。\\n||RRUFF packages use %.0f MB, over the %.0f MB limit.\\n
RRUFF 配对（%d 条）||RRUFF pairing (%d)
RRUFF编号||RRUFF ID
SPC 光谱(*.spc)||SPC spectra (*.spc)
XRD 峰位列表（DIF）||XRD peak list (DIF)
XRD 粉末衍射（XY 处理后）||XRD powder diffraction (processed XY)
excellent · 定向||excellent · oriented
excellent · 非定向||excellent · unoriented
fair · 定向||fair · oriented
fair · 非定向||fair · unoriented
poor · 非定向||poor · unoriented
── RRUFF 样品记录（来源数据包：%s）──||-- RRUFF sample record (from package: %s) --
　⚠ 已超过容量上限 %.0f MB，建议清理||  ! Over the %.0f MB limit; cleanup recommended
　磁盘可用 %.1f GB||  disk free %.1f GB
　磁盘可用 %.1f GB / 共 %.1f GB||  disk free %.1f GB / %.1f GB total
一阶导数||first derivative
下载与分析结果都不会写到别的位置；数据包超过上限时会有提醒。||Downloads and analysis results are written nowhere else; you will be warned when the limit is exceeded.
下载候选谱线后与实测谱自动配对（可修改）：||Download candidate spectra and pair them with the measured spectrum automatically (editable):
下载后统一存到工具的数据文件夹（参考谱库），可离线使用。||Everything is stored in the tool data folder (reference library) for offline use.
下载（建议至少下载“fair / excellent”等级，谱质更好）。||Download (at least the "fair / excellent" grades for better quality).
不限制||unlimited
两点校准才需填||only needed for two-point calibration
中文名||Chinese name
英文名||English name
主峰FWHM||main peak FWHM
主峰位 (cm-1)||main peak (cm-1)
主峰位(cm-1)||main peak (cm-1)
主峰半高宽 FWHM||main peak FWHM
主峰半高宽FWHM||main peak FWHM
主峰强度||main peak intensity
主成分散点（PC1-PC2）||PCA scatter (PC1-PC2)
也可直接用 Word / WPS 打开。||You can also open it with Word / WPS.
二维成像 · 点阵扫描热图||2D imaging · grid-scan heat map
二维成像完成：%d 个点位（无有效数值）||2D imaging done: %d points (no valid values)
二阶导数||second derivative
交互式相减 · A − k·B||Interactive subtraction · A - k·B
交互相减：A − k·B||Interactive subtraction: A - k·B
产地||Locality
仪器||Instrument
伪Voigt||pseudo-Voigt
位移校准:%s||shift calibration: %s
低分辨率 LR||low resolution LR
例 520.6 → 520.7||e.g. 520.6 -> 520.7
偏差||Deviation
偶然概率||Chance probability
元素||Elements
元素组成：%s||Elements: %s
元素（空格分隔，全部含）||Elements (space separated, all required)
光谱分析报告||Spectral analysis report
光谱图 · %s||Spectrum · %s
光谱条数||Spectra
光谱转换||Spectrum conversion
拉曼光谱工具 · 使用说明书||Raman Spectrum Toolkit · User Guide
拉曼光谱工具 · Raman Spectrum Toolkit||Raman Spectrum Toolkit
（JASCO .jws 光谱转换 · 拉曼峰分析 · 矿物鉴定）||(JASCO .jws conversion · Raman peak analysis · mineral identification)
读入：.jws / .csv / .spc / .jdx / .txt / .xlsx||Reads: .jws / .csv / .spc / .jdx / .txt / .xlsx
输出：Excel（含图表）/ PNG / CSV / 峰列表 / 峰拟合 / JCAMP-DX||Writes: Excel (with chart) / PNG / CSV / peak table / peak fit / JCAMP-DX
预处理：尖峰去除 · 基线校正 · 平滑 · 导数 · 归一化 · 位移校准||Preprocessing: despike · baseline · smoothing · derivative · normalize · shift calibration
分析：未知谱检索 · 峰拟合 · 峰位检索 · 相似度矩阵 · 谱运算 · 瀑布图 · 聚类 · 二维成像 · 报告||Analysis: unknown-spectrum search · peak fitting · peak-position search · similarity matrix · spectrum arithmetic · waterfall · clustering · 2D imaging · reports
数据库：ROD 在线检索 · RRUFF 拉曼 / 红外 / XRD / 化学成分数据包||Databases: ROD online search · RRUFF Raman / IR / XRD / chemistry packages
矿物：内置特征峰归属库 + RRUFF 真实样品记录||Minerals: built-in characteristic-peak library + real RRUFF sample records
版本：%s||Version: %s
全部||All
关于||About
其他||Other
分析：  未知光谱检索（全库鉴定）、峰拟合(高斯/洛伦兹/伪Voigt)、\\n||Analysis:  unknown-spectrum search (whole library), peak fitting (Gaussian/Lorentzian/pseudo-Voigt),\\n
化学名称||Chemical name
化学式||Formula
化学式包含||Formula contains
化学式：%s||Formula: %s
化学成分（电子探针）||Chemistry (electron microprobe)
匹配||Matched
半高宽FWHM||FWHM
单侧||one-sided
占强度范围，越小越灵敏||of the intensity range; smaller is more sensitive
参考 ||Reference 
参考峰位||Reference peak
参考谱||Reference spectrum
叠加显示前 %d 个文件||Overlay the first %d files
另有「未命名样品 %s」（分数 %.0f）更接近，其峰位 %s；||Another unnamed sample %s (score %.0f) is even closer, with peaks %s;
可下载更大 / 质量更高的数据包后再检索以提高把握。||Download a larger / better package and search again to gain confidence.
可作价态/物相线索，但该记录没有矿物名，无法据此定名。||It can serve as a phase/valence clue, but the record has no mineral name so it cannot name your mineral.
可信候选：%s（%s，库内 %d 条同名谱；领先第二名 %.0f 分）。||Reliable candidate: %s (%s, %d spectra of this mineral in the library; leads the runner-up by %.0f points).
可先用 --rruff-get 或 --import-pkg 准备数据。||Use --rruff-get or --import-pkg to prepare the data first.
可在“设置 → 数据文件夹…”中清理不再需要的数据包，或调高上限。||Clean up unneeded packages in "Settings -> Data folder..." or raise the limit.
可用“数据库 → RRUFF 数据源”下载后自动补全）||download it from "Database -> RRUFF data sources" to complete the index)
可能的候选：%s（%s），置信度一般。||Possible candidate: %s (%s), moderate confidence.
同一矿物只列最佳一条（库内同名谱数为该矿物在库中的条数）。\\n||Only the best entry per mineral is listed (the "library spectra" column counts that mineral in the library).\\n
名称||Name
名称 / 化学式||Name / formula
命中||Hits
命中 %d 种矿物。\\n%s||%d minerals found.\\n%s
命中/参考||Hits / reference
命中/参考峰数||Hits / reference peaks
在线检索数据库 · ROD（RRUFF 拉曼数据的官方开放库）||Search online databases · ROD (the open Raman database of RRUFF)
基线校正:||Baseline:
基线窗口(cm-1)：||Baseline window (cm-1):
处理参数：||Processing parameters:
多项式2阶||polynomial, order 2
多项式（2阶）||polynomial (order 2)
大小||Size
如 Zircon / 锆石，留空 = 不归属||e.g. Zircon, blank = no assignment
实测 ||Measured 
实测/参考||Measured / reference
实测化学式||Measured formula
实测峰位,参考峰位,偏差(cm-1),是否匹配||measured position,reference position,deviation (cm-1),matched
实测峰位||Measured peak
实测：%s        最佳匹配：%s        F1 %.0f%%    r=%.4f    谱角 %.1f°||Measured: %s        best match: %s        F1 %.0f%%    r=%.4f    angle %.1f°
宽：||Width:
对应标准值：||Standard value:
将删除以下已下载数据包（含索引），共 %.0f MB：\\n%s\\n\\n||The following downloaded packages (including indexes) will be deleted, %.0f MB in total:\\n%s\\n\\n
将在 ROD 数据库（RRUFF 拉曼数据）检索以下关键词，\\n||The following keywords will be searched in ROD (RRUFF Raman data),\\n
尖峰去除(窗口%d, 阈值%.1f)||spike removal (window %d, threshold %.1f)
尖峰窗口(点)：||Spike window (points):
尖峰阈值(噪声倍数)：||Spike threshold (x noise):
峰位(cm-1)||Position (cm-1)
峰位、FWHM 由本工具自动识别，请在发表前人工核对。||Positions and FWHMs are detected automatically; please verify manually before publication.
峰位归属参考：%s||Peak-assignment reference: %s
峰位归属矿物：||Assign peaks to mineral:
峰灵敏阈值(%)：||Peak sensitivity threshold (%):
峰识别阈值 %.1f%% 强度范围，最小峰间距 %.1f cm-1||peak threshold %.1f%% of the intensity range, minimum separation %.1f cm-1
（无数据）||(no data)
图片缺失：%s||missing image: %s
峰位(cm-1),强度,相对强度(%),半高宽FWHM,峰突出度,来源||position (cm-1),intensity,relative intensity (%),FWHM,prominence,source
已下载||Downloaded
已下载的数据包里没有 "%s"。\\n\\n可依次下载以下数据包并自动重试：\\n%s||"%s" is not in the downloaded packages.\\n\\nThe following packages can be downloaded in order and retried automatically:\\n%s
已下载，未建索引||downloaded, not indexed
已下载，索引 %d 条||downloaded, %d entries indexed
已生成自包含 HTML 报告（图片已内嵌）：\\n%s\\n\\n||Self-contained HTML report generated (images embedded):\\n%s\\n\\n
平均偏差||Mean deviation
平均强度(%d条)||mean intensity (%d)
平滑:%s 窗口%d点||smoothing: %s, window %d points
平滑窗口(点)：||Smoothing window (points):
库内同名谱数||Library spectra
库内谱数||Library spectra
库里没有命名的矿物峰位接近（%s）||No named mineral has similar peak positions in the library (%s)
库里没有找到任何峰位接近的参考谱：可能该矿物不在已下载的数据包里，||No reference spectrum with similar peaks was found: the mineral may not be in the downloaded packages,
建议先清理旧数据包（设置 → 数据文件夹… → 清空），||Cleaning up old packages first is recommended (Settings -> Data folder... -> Clear),
建议再用配对报告核对峰位对照表。||Verify the peak comparison table with a pairing report.
强峰命中||Strong-peak hits
强峰命中 = 未知谱最强 %d 个峰里命中几个。||Strong-peak hits = how many of the %d strongest peaks of the unknown spectrum were found.
强度||Intensity
强度差||Intensity difference
归一化:||Normalization:
归一化强度||Normalized intensity
归属||Assignment
归属文本参考内置矿物特征峰表（容差 8 cm-1），仅供参考。||Assignments refer to the built-in characteristic-peak table (tolerance 8 cm-1) for reference only.
总强度||Total intensity
总览||Overview
成分||Chemistry
成功 %d 个，跳过 %d 个。\\n汇总统计：%s||Succeeded %d, skipped %d.\\nSummary: %s
成功 %d 个，跳过 %d 个，失败 %d 个。\\n详见日志。||Succeeded %d, skipped %d, failed %d.\\nSee the log for details.
或把数据文件夹改到空间更大的分区。||Or move the data folder to a partition with more space.
或激光波长/谱区与库中不同，或这不是拉曼谱。||or the laser wavelength / spectral range differs from the library, or this is not a Raman spectrum.
或点“选择文件…”直接指定。||or click "Choose a file..." to pick one directly.
或这不是拉曼谱（峰位单位/谱区不同）。||or this is not a Raman spectrum (different units or range).
手动||manual
手动配对||Manual pairing
手动配对（数据库 %d 条）||Manual pairing (database %d)
批量||bulk
报告为单个 HTML 文件，图片已内嵌；可用浏览器“打印 → 另存为 PDF”，||The report is a single HTML file with embedded images; in a browser use "Print -> Save as PDF",
拉曼||Raman
拉曼光谱配对对比报告||Raman pairing comparison report
指定波数附近的峰强度||intensity of the peak near a given wavenumber
按可信度给出候选矿物。||and lists candidate minerals by confidence.
排名||Rank
提示||Notice
数据包||Packages
数据库：ROD 在线检索、RRUFF 拉曼/红外/XRD/化学成分数据包、\\n||Database: ROD online search, RRUFF Raman/IR/XRD/chemistry packages,\\n
数据文件夹||Data folder
数据文件夹 · 位置 / 占用 / 清理||Data folder · location / usage / cleanup
数据来源：ROD（Raman Open Database，RRUFF 拉曼数据）。||Source: ROD (Raman Open Database, RRUFF Raman data).
数据点数||Points
数据链接||Data link
文件||File
文件夹批处理||Batch a folder
文本(*.txt)||Text (*.txt)
方式||Mode
无||None
是否匹配||Matched
是否把现有数据一并移动到新位置？\\n（数据包较大时可能要一会儿）||Move the existing data to the new location?\\n(May take a while for large packages)
是否搬移已有数据||Move existing data
是否现在下载最小的数据包（未评级·非定向，12 MB）？||Download the smallest package now (unrated/unoriented, 12 MB)?
晶系||Crystal system
晶系：%s||Crystal system: %s
晶胞参数 / 晶系||Cell parameters / crystal system
最佳配对峰位对照（容差 %.1f cm-1）：%s||Best peak comparison (tolerance %.1f cm-1): %s
最佳配对峰位对照（容差 5 cm-1）：%s||Best peak comparison (tolerance 5 cm-1): %s
最大值=1||max = 1
最小-最大||min-max
未下载||not downloaded
未命名||unnamed
未建索引||not indexed
未归属||unassigned
未找到，是否下载更大的数据包？||Not found. Download a larger package?
未知||unknown
未知光谱 vs 最佳候选||Unknown spectrum vs best candidate
未知光谱 vs 最佳候选 %s||Unknown spectrum vs best candidate %s
未知光谱 × %s||Unknown spectrum x %s
未知光谱检索 · 全库鉴定||Unknown-spectrum search · whole-library identification
未评级 · 定向||unrated · oriented
未评级 · 非定向||unrated · unoriented
本地参考谱库（已导出的 CSV）不受影响。\\n确定继续吗？||The local reference library (exported CSVs) is not affected.\\nContinue?
本地数据库为空。\\n请先用“数据库 → 在线检索数据库…”下载参考谱线。||The local database is empty.\\nUse "Database -> Search online databases..." to download reference spectra.
本地数据库还没有参考谱线。\\n是否现在在线检索并下载？||There are no reference spectra yet.\\nSearch and download them now?
本地还没有 RRUFF 数据包。\\n||No RRUFF packages yet.\\n
来源||Source
样品描述||Sample description
样品：%s   |   ||Sample: %s   |   
横坐标最大值：||X max:
横坐标最小值：||X min:
纵坐标最大值：||Y max:
纵坐标最小值：||Y min:
没有可靠匹配（有名字的最高分 %.0f）。可能该矿物不在已下载数据包里，||No reliable match (best named score %.0f). The mineral may not be in the downloaded packages,
波长||Wavelength
波长(nm)||Wavelength (nm)
注意前几名分数很接近（仅领先 %.0f 分），||Note that the top scores are very close (only %.0f points apart);
洛伦兹||Lorentzian
清理数据包||Clean up packages
滚动最小值||rolling minimum
滚动最小值 / 滚动球用||for rolling minimum / rolling ball
滚动球||rolling ball
瀑布图 / 叠加图偏移：||Waterfall / overlay offset:
特征拉曼峰（cm-1）：||Characteristic Raman peaks (cm-1):
状态||Status
现在打开查看吗？\\n（浏览器里可“打印 → 另存为 PDF”）||Open it now?\\n(In a browser you can "Print -> Save as PDF")
理想化学式||Ideal formula
生成时间||Generated
用法：选中一行看详情，再点“作为峰位归属参考”。||How to use: select a row for details, then click "Use as peak-assignment reference".
命中 %d 种矿物。\\n%s||%d minerals matched.\\n%s
\\n\\n（正在查找 RRUFF 本地样品记录…）||\\n\\n(looking up the local RRUFF sample record ...)
\\n\\n（本地数据包里暂无该矿物的样品记录；可用“数据库 → RRUFF 数据源”下载后自动补全）||\\n\\n(no sample record for this mineral in the local packages; download one from "Database -> RRUFF data sources" to complete it automatically)
已生成自包含 HTML 报告（图片已内嵌）：\\n%s\\n\\n现在打开查看吗？\\n（浏览器里可“打印 → 另存为 PDF”）||Self-contained HTML report generated (images embedded):\\n%s\\n\\nOpen it now?\\n(in the browser you can use "Print -> Save as PDF")
由 拉曼光谱工具 生成 · 共 %d 条光谱||Generated by Raman Spectrum Toolkit · %d spectra
拉曼光谱工具||Raman Spectrum Toolkit
相关系数||Correlation
相对强度(%)||Rel. intensity (%)
矿物||Mineral
矿物信息与拉曼峰归属||Mineral info and Raman peak assignment
矿物名||Mineral
确定要清空「%s」吗？\\n%s\\n\\n此操作不可撤销。||Clear "%s"?\\n%s\\n\\nThis cannot be undone.
确认清理||Confirm cleanup
磁盘剩余空间不足：本次约需 %.0f MB，可用仅 %.0f MB。\\n||Not enough disk space: about %.0f MB is needed but only %.0f MB is free.\\n
移动平均||moving average
突出度||Prominence
第%d行||row %d
第1对 实测/标准：||Pair 1  measured / standard:
第2对 实测：||Pair 2  measured:
第一次使用请先【建立 / 更新特征索引】。||On first use, click [Build / update feature index] first.
第一步：下载一个数据包（自动建索引）。第二步：按矿物名或 RRUFF 编号检索。||Step 1: download a package (indexed automatically). Step 2: search by mineral name or RRUFF ID.
第三步：导出到本地库，即可绘图 / 配对比较。||Step 3: export to the local library to plot / pair-compare.
簇%d（%d 条）||cluster %d (%d)
簇：||Clusters:
类型||Type
精算||re-scored
索引||indexed
索引 %d 条||%d indexed
索引条数||Index entries
红外、XRD、化学成分包同样适用。||The IR, XRD and chemistry packages work the same way.
红外光谱 IR（RAW）||IR spectra (RAW)
线性||linear
线性（端点连线）||linear (endpoints joined)
综合分||Score
编号||ID
聚类分析||Cluster analysis
聚类距离（1 − 相关系数）||cluster distance (1 - correlation)
能力参考 RRUFF 项目所列工具（RamanCrystalHunter、RamanLab 等）||Feature set inspired by RRUFF-listed tools (RamanCrystalHunter, RamanLab, ...)
自动||auto
自动分割阈值 %.3f||auto cut %.3f
自动分割阈值（1 − 相关系数，留空 = 自动）：||Auto cut (1 - correlation; blank = auto):
自动配对 · 在线检索||Auto-pair · online search
自动配对（在线，新下载 %d 条）||Auto-pair (online, %d newly downloaded)
自动配对（本地库 %d 条）||Auto-pair (local library, %d)
评分方式||Scored by
识别峰数||Peaks found
说明书结束。祝实验顺利！||End of the guide. Good luck with your experiments!
请至少勾选一种输出：Excel / PNG / CSV / 峰列表。||Tick at least one output: Excel / PNG / CSV / peak list.
请选择输出文件夹，或勾选“保存到源文件所在目录”。||Choose an output folder or tick "Save next to the source file".
读入：.jws / .csv / .spc / .jdx / .txt / .xlsx\\n||Reads: .jws / .csv / .spc / .jdx / .txt / .xlsx\\n
谱段替换 / 拼接||Range replacement / stitching
谱段替换预览（替换 %d 点）||Range replacement preview (%d points replaced)
谱角||Spectral angle
越小越激进，建议 8~15||smaller is more aggressive; 8-15 recommended
输出：Excel(含图表) / PNG / CSV / 峰列表 / 峰拟合 / JCAMP-DX\\n\\n||Outputs: Excel (chart) / PNG / CSV / peak list / peak fit / JCAMP-DX\\n\\n
这种情况也可能是库中没有对应矿物，请结合峰位人工判断。||This can also happen when the mineral is not in the library; please judge from the peak positions.
迭代多项式||iterative polynomial
迭代多项式用||for iterative polynomial
迭代多项式（推荐）||iterative polynomial (recommended)
迭代次数：||Iterations:
迭代阶数：||Order:
选择好未知光谱后，点【开始检索】。||After choosing the unknown spectrum, click [Search].
通道%d||channel %d
配对||Pairing
配对排名（共 %d 条参考谱，按峰位匹配 F1 排序）||Pairing ranking (%d reference spectra, by peak-match F1)
配对：%s  ↔  %s   (r=%.3f)||Pairing: %s  <->  %s   (r=%.3f)
鉴定状态||Identification status
需要先下载数据包||Download a package first
预处理：尖峰去除、基线(线性/多项式/迭代多项式/滚动最小值/滚动球)、\\n||Preprocessing: spike removal, baseline (linear/polynomial/iterative/rolling minimum/rolling ball),\\n
高斯||Gaussian
高级设置 · 处理 / 校准 / 坐标轴 / 峰 / 图幅||Advanced settings · processing / calibration / axes / peaks / size
高：||Height:
（含 %d 个子谱，取第 1 个）||(%d sub-spectra; using the first)
（尚未选择）||(nothing selected)
（已超限，建议清理）||(over limit; cleanup recommended)
（无）||(none)
（未命名样品）||(unnamed sample)
（说明书版本：%s）||(guide version: %s)
语言 / Language||Language
中文||Chinese
界面语言已切换为 %s||UI language switched to %s
名称||name
化学式||formula
元素||element
参考谱库||Reference library
RRUFF数据包||RRUFF packages
分析结果||Analysis results
三斜晶系||triclinic
三方晶系||trigonal
六方晶系||hexagonal
单斜/正交||monoclinic/orthorhombic
单斜晶系||monoclinic
四方晶系||tetragonal
正交晶系||orthorhombic
立方晶系||cubic
下载并建索引：--rruff-get 数据包key||Download and index: --rruff-get package-key
使用说明书 · 拉曼光谱工具||User Guide · Raman Spectrum Toolkit

# ===== 矿物特征峰归属文本 =====
晶格振动||lattice vibration
外部振动||external mode
外部振动/晶格模式||external mode / lattice mode
Zr-O 振动||Zr-O vibration
Zr-O 伸缩||Zr-O stretching
Zr-O 伸缩（特征强峰）||Zr-O stretching (strong characteristic peak)
Si-O 伸缩||Si-O stretching
Si-O 伸缩（特征峰）||Si-O stretching (characteristic peak)
Si-O 弯曲||Si-O bending
Si-O 反对称伸缩||Si-O antisymmetric stretching
ν1(SiO4) 对称伸缩||nu1(SiO4) symmetric stretching
ν1(SiO4) 对称伸缩（双峰）||nu1(SiO4) symmetric stretching (doublet)
ν2(SiO4) 弯曲振动||nu2(SiO4) bending
ν3(SiO4) 反对称伸缩（最强峰）||nu3(SiO4) antisymmetric stretching (strongest)
ν1(CO3) 对称伸缩（最强峰）||nu1(CO3) symmetric stretching (strongest)
ν1(CO3)（最强峰）||nu1(CO3) (strongest)
ν3(CO3) 反对称伸缩||nu3(CO3) antisymmetric stretching
ν4(CO3) 弯曲||nu4(CO3) bending
ν1(SO4) 对称伸缩（最强峰）||nu1(SO4) symmetric stretching (strongest)
ν1(SO4)（最强峰）||nu1(SO4) (strongest)
ν2(SO4) 弯曲||nu2(SO4) bending
ν3(SO4) 反对称伸缩||nu3(SO4) antisymmetric stretching
ν1(PO4) 对称伸缩（最强峰）||nu1(PO4) symmetric stretching (strongest)
ν3(PO4) 反对称伸缩||nu3(PO4) antisymmetric stretching
ν1 A1 Si-O 对称伸缩（最强峰）||nu1 A1 Si-O symmetric stretching (strongest)
OH 伸缩||OH stretching
内层 OH 伸缩||inner OH stretching
表面 OH 伸缩||surface OH stretching
Al-OH 弯曲||Al-OH bending
Al-OH（宽）||Al-OH (broad)
Fe-OH 弯曲||Fe-OH bending
Al-O 振动||Al-O vibration
S-S 伸缩||S-S stretching
Ca-F 伸缩||Ca-F stretching
双磁子散射（宽峰）||two-magnon scattering (broad)
D 带（缺陷诱导）||D band (defect induced)
G 带（sp2 骨架）||G band (sp2 network)
2D 带（二阶）||2D band (second order)
sp3 C-C 伸缩（特征单峰）||sp3 C-C stretching (single characteristic peak)
E2g Mo-S 面内伸缩（最强峰）||E2g Mo-S in-plane stretching (strongest)
A1g Mo-S 面外伸缩||A1g Mo-S out-of-plane stretching
TO/LO Zn-S 伸缩||TO/LO Zn-S stretching
Ag S-S 伸缩（最强峰）||Ag S-S stretching (strongest)
Eg S-S 伸缩||Eg S-S stretching
Eg Ti-O 伸缩||Eg Ti-O stretching
Eg Ti-O 伸缩（很强）||Eg Ti-O stretching (very strong)
A1g Ti-O 伸缩（最强峰）||A1g Ti-O stretching (strongest)
B2g Ti-O 伸缩（弱）||B2g Ti-O stretching (weak)
A1g Al-O（最强）||A1g Al-O (strongest)
A1g Sn-O（最强）||A1g Sn-O (strongest)
B1g 晶格振动（弱）||B1g lattice vibration (weak)
Eu 晶格||Eu lattice
T2g（特征）||T2g (characteristic)
A1g（最强）||A1g (strongest)
用法：||Usage:
    双击“启动拉曼光谱工具.bat”或直接运行（不带参数）= 打开图形界面||    Double-click "启动拉曼光谱工具.bat" or run with no arguments = open the GUI
    jws2csv.py 文件/文件夹 [--xlsx --png --csv --peaks]    批量转换||    jws2csv.py FILE/FOLDER [--xlsx --png --csv --peaks]    batch convert
    jws2csv.py --identify 文件           未知光谱全库鉴定||    jws2csv.py --identify FILE        identify an unknown spectrum
    jws2csv.py --pair 文件               与本地库自动配对||    jws2csv.py --pair FILE            auto-pair against the local library
    jws2csv.py --pair-batch 文件夹       批量配对（找对不上的）||    jws2csv.py --pair-batch FOLDER    batch pairing (spot the mismatches)
    jws2csv.py --identify-batch 文件夹   批量鉴定（逐条给最佳候选）||    jws2csv.py --identify-batch FOLDER  batch identification (best candidate each)
    jws2csv.py --rruff-list              查看 RRUFF 数据包||    jws2csv.py --rruff-list           list RRUFF data packages
    jws2csv.py --mineral-search 名称     查内置矿物特征峰表||    jws2csv.py --mineral-search NAME  query the built-in mineral table
    jws2csv.py --manual [路径]           打印 / 导出说明书||    jws2csv.py --manual [PATH]        print / export the guide
    jws2csv.py --lang en|zh              切换界面语言||    jws2csv.py --lang en|zh           switch the UI language
    更多参数见说明书第 16 章（命令行速查）||    See chapter 16 of the guide for all options
"""


_MANUAL_SECTIONS = [
    ("1. 五分钟上手", """【第 1 步】双击 RamanSpectrumToolkit.exe 打开界面（也可双击
          “启动拉曼光谱工具.bat”，它会自动找 Python）。

【第 2 步】点左侧【添加文件】选 .jws 文件；文件多就点【添加文件夹】
          一次性递归添加。选中文件后，右侧立即显示光谱图。

【第 3 步】在【② 输出位置】里勾选需要的“输出内容”：
          Excel（含图表）  最省事，双击用 WPS/Excel 打开就有图
          PNG 图片         出图片，文件名与源文件一致
          CSV（纯数据）    只要数字
          峰列表           峰位、强度、相对强度、FWHM、突出度
          峰拟合           高斯 / 洛伦兹 / 伪Voigt 拟合结果

【第 4 步】点【开始转换】。左下“运行日志”会逐条显示结果，
          完成后弹提示。

【第 5 步】要和标准谱对比：
          菜单【数据库】→【RRUFF 数据源：拉曼/红外/XRD/成分…】
          → 输入矿物名（如 Zircon）→【按矿物批量抓取】
          → 回主界面菜单【分析工具】→【自动配对（本地数据库）】。

【常见疑问】为什么 CSV 打开没有图？
          CSV 只能存数字。要图请勾选“Excel（含图表）”或“PNG 图片”。"""),

    ("2. 界面总览", """窗口分左右两半：左边操作、右边看图。

左侧（自上而下）
  ① 选择文件       文件列表框 + 添加文件 / 添加文件夹 / 移除选中 / 清空列表
  ② 输出位置       是否与源文件同目录、输出文件夹、输出内容勾选项
  ③ 表头与列名     是否写表头、是否自动识别列名、横纵坐标列名
  ④ 运行日志       每一步的结果都在这里
  开始转换 / 进度条 / 打开输出目录

右侧（光谱预览）
  预览选中文件  打开 CSV 看图  CSV→图片  保存为 PNG  清空视图
  图表设置：横坐标刻度间隔、起始刻度、显示纵坐标数值、网格线、
            标题、标注峰位、显示峰位数值、最小峰间距(cm-1)、高级设置…
  撤销手动峰 / 清空手动峰 / 恢复自动峰
  下方大图 = 画布，左键点图补标峰、右键删掉最近的峰（自动 / 手动都可）

顶部菜单
  文件     添加文件 / 添加文件夹 / 批处理文件夹… / 导出分析报告… / 退出
  分析工具 未知光谱检索（全库鉴定）、配对比较、峰拟合、峰位检索、
           相似度矩阵、聚类分析、二维成像、平均、相减、交互式相减、
           谱段替换、瀑布图、多数据图叠加、光谱比对
  设置     高级设置… / 矿物信息与拉曼峰归属库… / 数据文件夹… /
           清空手动峰标注 / 恢复被删的自动峰
  数据库   在线检索、RRUFF 数据源、本地库比对、打开数据文件夹
  帮助     使用说明（完整手册）… / 关于"""),

    ("3. 支持的格式", """能读进来的文件
  .jws              JASCO 光谱仪原始文件（本工具的主打格式）
  .csv              两列或多列数值表（第一列为横坐标）
  .spc              Galactic / Thermo GRAMS 格式（新版 0x4B）
  .jdx / .dx        JCAMP-DX 标准交换格式
  .txt .dat .asc .xy 两列数值文本（制表符 / 逗号 / 空格分隔均可）
  .xlsx / .xlsm     Excel（读第一个工作表）

读入时的小规则
  · 只要求第一列是数字；末尾的说明列（如“来源”“备注”）会自动丢弃。
  · 表头行有没有都行，工具会自己判断。
  · .txt 若以 ## 开头，按 JCAMP-DX 解析。

文件夹递归查找时，会自动跳过工具自身的输出目录（_转换结果、
工具数据、分析结果等）与衍生表（*_peaks.csv、*_fit.csv、
批处理汇总、配对报告），不会把它们误当成光谱。

能导出的文件
  .xlsx 数据 + 内嵌图表      .csv 纯数据
  .png  光谱图               *_peaks.csv 峰列表
  *_fit.csv 峰拟合表         .jdx JCAMP-DX（等间距用 X++ 写法）"""),

    ("4. 输出内容与文件位置", """输出位置有两种选择，在【② 输出位置】里切换：
  ☑ 保存到源文件所在目录（推荐）—— 输出与源文件放在一起
  ☐ 取消勾选后，可自己指定“输出文件夹”（点【浏览…】）

输出内容可多选，各格式说明：
  Excel（含图表）  一个文件里既有数据表又有折线图，默认勾选
  PNG 图片         文件名与源文件一致，图上自动标注峰位
  CSV（纯数据）    第一列横坐标，之后每个通道一列
  峰列表           <名字>_peaks.csv，含归属列（若设了归属矿物）
  峰拟合           <名字>_fit.csv，含峰面积、峰形、R²

☑ 跳过已存在的 CSV（不覆盖）：重复转换时不会覆盖已有结果。

注意：这里说的是“你自己样品”的转换结果，默认与样品放在一起；
      工具自身的分析输出（聚类、成像、报告等）统一进“工具数据”
      文件夹，不会污染样品目录（见下一章）。"""),

    ("5. 数据文件夹与容量管理", """所有工具自己下载/生成的数据，只写在一个地方：

  拉曼光谱工具（Raman Spectrum Toolkit）/
  ├─ RamanSpectrumToolkit.exe
  ├─ 使用说明.txt / User_Guide.txt   ← 中英两份说明书
  ├─ assets/icon.png                 ← 应用图标
  ├─ 工具数据/
  │   ├─ 参考谱库/          ROD 下载 + RRUFF 导出的拉曼参考谱
  │   │   └─ IR / XRD/      其他模态参考谱单独存放，不混用
  │   ├─ RRUFF数据包/       数据包 zip、检索索引与特征索引（可清理）
  │   └─ 分析结果/          峰拟合、配对、聚类、成像、报告等输出
  └─ 源码/                  可删除

管理入口：菜单【设置】→【数据文件夹…】
  · 查看根目录、每个子目录的文件数与占用、磁盘可用空间
  · 打开 / 清空任一子目录（清空前有确认）
  · 更改位置…：把数据文件夹换到别的盘，可选一并搬移已有数据
  · 数据包容量上限（默认 2048 MB）：超限会提醒
  · 清理已下载数据包…：批量删除数据包，参考谱库 CSV 不受影响

三道防占空间的闸门
  ① 容量上限：超限提醒
  ② 下载前预检：磁盘不够直接拒绝，并告诉你差多少
  ③ 一键清理：按数据包大小批量删除

若工具目录不可写（例如放在只读位置），会自动改用系统用户目录，
实际路径在“数据文件夹…”窗口里显示。"""),

    ("6. 图表设置与手动补标峰", """右侧“图表设置”改完立即刷新预览：

  横坐标刻度间隔   默认 200，刻度取整（拉曼位移推荐 200）
  起始刻度         留空 = 自动；想让刻度落在 0/500/1000 就填 0
  显示纵坐标数值   默认关闭（拉曼强度是相对值，隐藏更清爽）
  网格线 / 标题    按需开关
  标注峰位         自动识别并标注峰位
  显示峰位数值     默认开启；取消勾选后峰位数字不再画出，
                   但峰位标记与虚线还在，适合谱线密集时看走势
  最小峰间距(cm-1) 默认 20，太小会把噪声当峰
  高级设置…        预处理、校准、峰、图幅都在里面（见第 7、15 章）

峰位波长虚线
  自动标注的峰和手动补标的峰，都会从峰顶往下画一条虚线引到横坐标轴，
  一眼就能读出这个峰对应的波数。不想显示就去
  【设置】→【高级设置…】→“峰识别与峰拟合”，取消勾选
  “峰位虚线引到横坐标轴”（命令行 --no-peak-dash）。

手动补标峰（自动漏掉的峰）
  · 在图上左键点一下 = 补标一个峰（蓝色方块，会自动吸附到峰顶）
  · 每个文件单独记忆，导出峰列表与出图时一并标注
  · 手动峰在峰列表里“来源”列显示为“手动”
  · 点【撤销手动峰】撤销最后补的那个，【清空手动峰】一次清光
  · 也可以命令行一次给多个峰位：--manual 1007,974

删除不想要的峰（自动标错的也能删）
  · 在图上对准某个峰右键 = 删掉离鼠标最近的那个峰，
    自动峰（红色圆点）和手动峰（蓝色方块）都适用
  · 自动峰不会真的“消失”，而是记进该文件的“已删除”名单，
    出图、峰列表、导出都按删除后的结果算；
    峰位删除要求点在峰附近（横向约 ±1/40 图宽），
    否则只是提示“附近没有峰”，避免在空白处误删
  · 点【恢复自动峰】把该文件删掉的自动峰一次全恢复
    （菜单【设置】→【恢复被删的自动峰】同样效果）
  · 手动峰被删掉就是真的移除了，用【撤销手动峰】或重新左键点回来
  · 每个文件各记各的，删错了随时能恢复

保存图片：点【保存为 PNG】；或直接勾选输出内容“PNG 图片”。
把 CSV 当图看：点【打开 CSV 看图】，再用【保存为 PNG】存图。"""),

    ("7. 预处理（基线·平滑·尖峰·校准）", """全部在菜单【设置】→【高级设置…】里，默认只影响出图与峰识别；
若要写进导出的 CSV/Excel，勾选“同时应用到导出的数据”。

尖峰去除（宇宙射线 / 坏点）
  ☑ 尖峰去除（宇宙射线 / 坏点）
  尖峰窗口(点)：5 左右
  尖峰阈值(噪声倍数)：8~15，越小越激进
  判定同时看两条，所以不会削平真实窄峰：
    ①偏离中值滤波超过阈值倍数  ②偏差要达邻域动态范围一定比例
  另外要求“连续命中不超过 3 点”，真实峰即便很窄也不止 3 点。

基线校正（六选一）
  无                    不做
  线性（端点连线）      两端连线，最粗暴但快
  多项式（2阶）         适合轻微弯曲的背底
  迭代多项式（推荐）    反复拟合并剔掉基线以上的点，最贴谷底，
                        适合荧光背景强的拉曼谱，阶数 5、迭代 20 即可
  滚动最小值            按窗口取谷底，适合基线起伏不规则
  滚动球                形态学开运算近似球体下切，参数同上

平滑
  移动平均 / Savitzky-Golay，窗口(点)：1=不平滑，常用 5~11
  要保峰形（不压峰高）用 Savitzky-Golay

导数：无 / 一阶 / 二阶（二阶常用于找肩峰）
归一化：无 / 最大值=1 / 最小-最大

拉曼位移校准（用标准样品校正峰位）
  第1对 实测/标准：例如 520.6 → 520.7
  第2对 实测/标准：填了就是两点校准（平移 + 缩放），不填只平移

峰位归属矿物：填 Zircon（或菜单里的矿物信息库点“作为峰位归属参考”），
  峰列表与报告就会多一列“归属”。"""),

    ("8. 峰识别与峰拟合", """峰是怎么找出来的
  在预处理后的曲线上找局部极大，并用“噪声 + 强度范围”双重门槛筛选，
  再按最小峰间距去重。调【峰灵敏阈值(%)】（高级设置里）：
    默认 7%，越小越灵敏（能出小峰，也可能带噪声）
    噪声大的谱建议 8~15%

峰列表（*_peaks.csv）里有什么
  峰位(cm-1)  强度  相对强度(%)  半高宽FWHM  峰突出度  来源  归属
  · 相对强度：以最强峰为 100%
  · FWHM：半高宽，结晶度/应力分析常用
  · 来源：自动 / 手动
  · 归属：设了归属矿物才有

峰拟合（菜单【分析工具】→【峰拟合（选中文件）】）
  峰形可选：高斯 / 洛伦兹 / 伪Voigt（含混合系数 η）
  输出 <名字>_fit.csv：
    峰位、拟合中心、拟合峰高、FWHM、峰面积、峰形、混合系数、R²
  判断拟合好坏看 R²，越接近 1 越好。

峰位检索（菜单【分析工具】→【峰位检索（参考峰表）…】）
  自己准备一个参考峰位表（每行：名称,峰位1,峰位2,…），
  工具用实测峰位去匹配（默认容差 5 cm-1），给出命中率排名。"""),

    ("9. 矿物信息与峰位归属", """菜单【设置】→【矿物信息与拉曼峰归属库…】

能做什么
  · 内置 48 种常见矿物的特征拉曼峰与振动归属，完全离线可用
    覆盖：锆石及各相氧化锆、SiO₂ 各相、氧化物与氢氧化物、
          碳酸盐、硫酸盐、磷酸盐、硅酸盐、硫化物、石墨、金刚石等
  · 三种检索方式：名称 / 化学式包含 / 元素（空白分隔，要求全部含）
    例：元素填 “Zr Si”；化学式填 “SiO2”
  · 选中一条，右侧显示：化学式、晶系、元素组成、特征峰与振动归属
  · 若本地已有 RRUFF 数据包，还会显示该矿物的真实样品记录：
    理想化学式 / 实测化学式 / 晶胞参数与晶系 / 产地 / 鉴定状态 / 链接

怎么用上归属
  选中矿物后点【作为峰位归属参考】，之后：
    峰列表、批处理汇总、分析报告都会多一列“归属”，
    例如 “ν3(SiO4) 反对称伸缩（参考 1008.0 cm-1）”
  归属按容差 8 cm-1 匹配，超出显示“未归属”——这是正常提示，不是错误。

命令行
  --mineral-search Zircon     按名称检索
  --by-element "Zr Si"        按元素检索
  --by-formula SiO2           按化学式检索
  --mineral-info Zircon       打印信息卡（含 RRUFF 样品记录）
  --rruff-info Zircon         只打印本地数据包里的 RRUFF 样品记录"""),

    ("10. 数据库：ROD 与 RRUFF", """两个数据来源
  ROD（Raman Open Database）   RRUFF 拉曼数据的官方开放库，按名称在线检索
  RRUFF 官网                   rruff.net/zipped_data_files，
                              共四类数据包：拉曼 / 红外 / XRD / 化学成分

◆ 在线检索：菜单【数据库】→【在线检索数据库（ROD / RRUFF 数据）…】
  输入矿物名或化学式 → 列表 → 下载选中谱线 / 下载并与实测谱配对

◆ RRUFF 数据包：菜单【数据库】→【RRUFF 数据源：拉曼/红外/XRD/成分…】
  ① 数据包表：类型 / 名称 / 大小 / 是否已下载 / 索引条数
     下方显示占用、磁盘可用空间与超限提醒
  ② 按钮：
     下载选中数据包并建索引   首次使用必须做（建索引需一会儿）
     重建索引                 索引丢失或数据包更新后用
     导入本地 zip…            自己从官网下载的 zip，直接导入
     清理数据包…              跳转到数据文件夹管理界面
  ③ 按矿物名或 RRUFF 编号检索（编号可直接反查，如 R050034）
  ④ 导出选中到本地库 / 导出并与实测谱配对

  另外：第一次用【分析工具 → 未知光谱检索】时，工具会为数据包建立
  “特征索引”（每条谱的峰位/化学式/晶系），这是未知谱鉴定的基础，
  建一次即可（600 条约 2 秒），删除数据包时会一并清掉。

导出规则（重要）
  · 同一种矿物有多个数据源时按“类型”分别去重，同类优先 Processed
  · 红外、XRD 参考谱会存到 参考谱库/IR、参考谱库/XRD 子目录，
    不会和拉曼谱混在一起配对（跨模态比对没有意义）
  · 化学成分包的电子探针表会原样提取到“分析结果”目录

随包附带：RRUFF 未评级·非定向数据包（12 MB，639 条，含真正的
锆石 Zircon R050034）+ 20 余条已导出的参考谱，开箱即可离线检索。

数据包比较大（红外 14 MB、XRD 67 MB、成分 19 MB，拉曼 12~229 MB），
下载前工具会先检查磁盘空间，超容量上限也会提醒。

下载慢 / 掉线怎么办（重要）
  rruff.net 在国外，瓶颈是这条跨境链路（实测 ping 丢包 25%、RTT 236 ms），
  不是工具的问题。实测吞吐几乎正比于同时在跑的连接数（服务器全程没有拒绝）：
  1 路约 0.03、8 路约 0.32、32 路约 0.89、64 路约 1.28、96 路约 1.45 MB/s
  （同一次测量的结果；这条链路随时段波动很大，绝对值能差 2~3 倍，
  但“连接数越多越快”这个趋势是稳定的）。
  所以以前下 227 MB 要 40 多分钟，中途一断还从头再来。现在下载改成了这样：
    · 后台下载：进度条 + 已下/总数 + 实时速率 + 剩余时间，窗口照常能点，
      不会再假死（以前用的是 update_idletasks，整段时间鼠标和关闭都不响应）；
    · 【取消】随时可点。中断后**已下载的部分会保留**，下次接着下，不从 0 重来；
    · 并发连接数默认 32 路，可在下载对话框里调（1~128），命令行 --dl-conns 96 也行。
      为了不把时间都花在握手上，每个连接至少分到 256 KB 的活，小包不会硬开几十路；
    · 动态领活：不是“把文件平均分给 N 条连接”，而是切成小段排队、谁下完谁再领。
      同一个 33 MB 的包，32 等分里最快那段 4.3 s、最慢那段 79 s——静态均分等于
      花 79 s 等一条卡住的连接；真实链路同一时段交错 A/B（12 MB 的包、都是 32 路）
      显示：静态均分平均 39.1 s → 动态领活平均 23.2 s，快 1.69 倍；
    · 代理：这条链路丢包严重，走代理 / VPN 往往比堆并发更快。
      在下载对话框里填代理地址（留空则沿用 Windows 系统代理），或命令行 --proxy；
    · 自动重试：断线按 1.5×n 秒退避重试，并且从断点继续，不是整段重下；
      服务器返回 429/503 会识别出来，并提示把并发调小些；
    · 完整性校验：长度对不上就不算下载成功，**残缺文件不会被当成已下载**
      （以前会——建索引时才报错，看着就像又崩了一次）；
      下完再验一次 zip 能不能打开，坏了自动删掉并提示重试；
    · 取消或强杀留下的 .part0 / .part1 分片就是断点，不会被当成正式数据包。
  数据包表里的「预计耗时」按你上次实测到的速率估算，会越用越准；
  下大包（≥ 50 MB）前会弹窗告知预计时长，可以先挑小包用。
  如果确实太慢：先下 fair_oriented（271 KB）或 powder_DIF（7.6 MB）练手，
  或者把并发拉到 64~96 试试，再或者在能直连的网络上用浏览器从官网下载 zip，
  再用【导入本地 zip…】。"""),

    ("11. 配对比较与未知谱鉴定", """◆ 未知光谱检索（全库鉴定，不依赖矿物名）

  手里一条谱完全不知道是什么？用这个功能：
  菜单【分析工具】→【未知光谱检索（全库鉴定）…】

  第一步：建立特征索引
    首次使用要先做（把库里每条谱的峰位、化学式、晶系提取出来，
    600 条约 2 秒，之后一直缓存复用）。在窗口里勾选数据包 →
    点【建立 / 更新特征索引】。
  第二步：指定待鉴定的光谱
    打开窗口时会自动带入你在左侧选中的那条谱；
    也可以点【选择文件…】直接指定（.jws/.csv/.spc/.jdx/.txt 都行）。
  第三步：点【开始检索】
    工具拿未知谱的峰去库里逐条比对，先粗筛再对前列候选读原始谱精算，
    最后按可信度给出候选矿物表。

  结果各列含义
    综合分   0.5×F1 + 0.5×强峰命中率，用来排序
    F1       峰位匹配 F1（容差 5 cm-1；命中少于 2 个峰记 0）
    强峰命中 未知谱最强的 3 个峰里，有几个在候选里找到（如 3/3）
    命中/参考 匹配上的峰数 / 该参考谱的峰数
    平均偏差 匹配峰位的平均差（越小越好，正常 1~3 cm-1）
    相关系数 两条曲线形状的相似度（精算才有）
    库内谱数 库里这个矿物共有几条谱
    偶然概率 纯属巧合也能匹配到这个程度的概率（越小越好）

  结论怎么读（三档）
    可信候选   综合分 ≥ 60、强峰命中 ≥ 2，且明显领先第二名（≥15 分）
                例：可信候选：Quartz（R110104，库内 6 条同名谱；领先第二名 31 分）
    可能的候选 综合分 ≥ 35，但证据不够强
                若前几名分数很接近，会明确提醒“也可能是库中没有对应矿物”
    没找到     最高分 < 35

  小技巧
    · 有 EDS / 电子探针结果时，在“必须含元素 / 必须不含”里填上
      （如必须含 Zr Si，或必须不含 Ca），能大幅提高准确率。
    · 检索类型默认“拉曼”。红外、XRD、成分包按需选择，别混着搜。
    · 候选表里可以【导出选中到参考谱库】或【用候选做配对报告】
      看峰位对照表，作为最终判断依据。
    · 库质很重要：unrated/poor 包的谱噪声大，建议下载 fair / excellent 包。

  命令行
    --identify 文件 --identify-top 12           检索（默认同时精算前 20 名）
    --identify-must "Zr Si"  --identify-not Ca  元素筛选
    --identify-kind 拉曼|红外|XRD|全部           限定检索类型
    --identify-index [数据包key]                建立 / 重建特征索引
    --identify-all                              不按矿物合并，列出每条谱
    --identify-no-exact                         只做索引粗筛（更快）

◆ 配对比较（手动 / 自动）

手动配对
  选参考谱文件 / 文件夹…   自己指定参考谱（RRUFF 的 txt、csv 都行）
  从数据库配对…            在线检索 → 选谱线 → 下载并与实测谱配对

自动配对
  自动配对（本地数据库）      与本地全部参考谱比对，给出最像的一条
  自动配对（在线检索并下载）  自动推测关键词 → 在线检索 → 下载前 8 条
                              → 自动配对

结果怎么看
  主排序：峰位匹配 F1（容差 5 cm-1）
    F1 同时考虑“命中率”和“误配”，避免峰多的参考谱靠巧合胜出；
    命中少于 2 个峰不计为识别。
  参考指标：相关系数（整条曲线形状）、谱角
  输出（放在“工具数据/分析结果”）：
    报告 CSV（排名表 + 最佳配对的峰位对照）
    报告图 PNG（叠加曲线 + 峰位配对连线 + 排名表）

批量配对（一整个文件夹 × 参考谱）
  菜单【分析工具】→【配对比较（手动 / 自动）】→【批量配对（文件夹 × 参考谱）…】，
  命令行 --pair-batch 文件夹 [--pair-ref 参考谱文件夹]。
  上面那些配对一次只处理一条谱，这个是外层批量：
    · 实测谱：指定一个文件夹（留空则用主界面左侧选中的光谱）；
    · 参考谱：本地参考谱库，或自己指定的文件夹 / 文件；
    · 逐条打分后汇总成一张表，一条实测谱一行。
  典型用法：手里一批锆石，拿锆石标准谱当参考集，汇总表里综合分低的
  就是“可能混了别的晶”的，它们会被排在最前面。

批量鉴定（一整个文件夹逐条鉴定）
  菜单【分析工具】→【批量鉴定（文件夹逐条鉴定）…】，
  命令行 --identify-batch 文件夹 --identify-top 5。
  手上是一堆不认识的谱时用：逐条在全库里找最像的矿物，每条给出最佳候选、
  综合分、F1 与结论文本，汇总成一张表。
  需要先有特征索引（在【未知光谱检索】窗口里建，或命令行 --identify-index）。

汇总表里列了什么
  · 综合分 = 0.5×F1 + 0.5×强峰命中率（与「未知谱鉴定」同一口径）；
  · F1、强峰命中、命中 / 实测 / 参考峰数、平均偏差、相关系数、谱角、
    偶然概率、领先第二名（参考谱多于一条时才有意义）。
  · “参考判读”一列只是按分数给的提示（匹配良好 / 部分匹配 / 匹配很差），
    不下最终结论 —— 到底是不是同一种物相，请自己核对峰位与谱型。
  输出（放在“工具数据/分析结果”）：
    汇总表 CSV（一条谱一行，按综合分升序：最可疑的排最前）
    汇总报告 HTML（同一张表，浏览器里可直接打印为 PDF）
  小技巧：表里综合分低的行，可以再单独对它跑一次【自动配对】，
          会生成峰位对照表与报告图，细看差在哪几个峰。

  两个口径说明：
  · 只识别到 1 个峰的谱，F1 一律记 0（命中<2 不算识别），综合分因此封顶
    50，会一直停在“部分匹配”档 —— 这是故意的，单峰不足以定案；
  · 命令行 --pair-batch / --identify-batch 是批处理，没有界面窗口。

提醒：配对是筛选辅助，最终请结合峰位对照表与谱图判断。
      实测与参考常有 3~5 cm-1 校准差，超过容差会判“未匹配”，
      可用第 7 章的“拉曼位移校准”先修正再配对。"""),

    ("12. 谱运算", """都在菜单【分析工具】里。

平均所选光谱
  多条谱插值到公共波数区间后逐点平均，输出 CSV 并画图。
  适合同一点位多次扫描去噪。

光谱相减（A−B）
  两条谱在重叠区间直接相减，快，但要求两条谱强度可比。

交互式相减 A−k·B（拖动找平）
  拖动滑块实时调 k（默认先各自归一化），把 B 的贡献拉平，
  一边看曲线一边找 k，满意后【导出 CSV】。
  适合从两相混合谱里剥离已知成分。

谱段替换 / 拼接
  把目标谱的一段波数区间换成另一条谱的对应区间，
  边缘可设“过渡宽度”，避免出现台阶。
  常用于替换荧光过强或坏点的区段。

瀑布图
  多条光谱归一化后纵向错开堆叠，便于横向比较谱型（菜单里一键出图，
  命令行加 --waterfall）。

多数据图叠加（堆叠排布）
  菜单【分析工具】→【多数据图叠加（所选光谱）…】，命令行加 --overlay。
  参照 stacked spectra 的画法：各条谱先归一化到最大值 = 1，
  再按“谱线偏移”纵向错开 k×偏移，所以**谱线彼此分开、不压在一起**，
  每条一种颜色，谱型仍可横向比较。偏移 1.0 = 相邻谱线刚好不压线，越大越分开。
    · **峰位跨谱合并**：同一个峰在每条谱上都标一遍会糊成一片。
      勾选后把各条谱上邻近的峰（相差不超过容差）归成同一个峰，
      只画一条虚线、只标一个**平均波数**。容差留空时沿用“最小峰间距”，
      也可以在对话框里单独填（命令行 --peak-merge 30）。
      各条谱自己的峰位标记（红点 / 蓝方块）仍然保留，
      方便看出这个峰是哪几条谱贡献的。
    · 横坐标取所有数据图波数范围的**交集** —— 只比大家都有数据的波段，
      某条谱短一截时右边就不会空出一段白。对话框里会写明实际取到的范围。
      若在【高级设置】里手动填过横坐标范围，则以手动的为准。
    · 默认各条归一化到最大值 = 1；取消勾选则按原始强度错开，
      用来看真实的相对强弱。
    · 可取消勾选“显示峰位数值”，只留虚线和标记 —— 十几条谱叠在一起时，
      数字会糊成一片，关掉更清爽（同样受主界面【显示峰位数值】影响）。
    · 图例：**几条谱线就列几条**（不再只列前 12 条），放不下会自动分列；
      默认排在**绘图区右侧的留白里**，横向跟谱线错开，谱线再密也压不到你。
      位置可选：右侧留白 / 上方 / 下方 / 图内自由位置 / 不显示，
      也可以在预览窗口里**直接按住图例拖到想要的位置**（松手即定位）。
      位置会存进设置文件，导出图和下次打开都用同一个位置。
    · 打开后先出**预览窗口**：改参数、增减峰位都只重画预览，
      不点【导出 PNG】就不会往结果目录写文件。
      - 左键点图 = 在点击处加一个峰位（蓝色虚线 + 蓝色数值）；
      - 右键点虚线 = 删掉离点击处最近的峰位（自动的和手动的都能删）；
      - 【重新检测】= 丢掉全部手动改动，回到自动检测的峰位；
      - 按住图例拖动 = 把图例挪到任意位置（按在图上别处仍然是加峰位，互不干扰）；
      - 改完参数按回车或点【刷新预览】才重画，免得每敲一个字符就重绘。
      手动增减要求勾选“峰位跨谱合并（一峰一线一值）”，
      因为只有合并模式才是“一个峰一条线一个数值”。
      命令行 --overlay 是批处理，没有预览窗口，直接按自动检测的峰位出图；
      命令行可以用 --legend-pos right|top|bottom|inside|none 指定图例位置。
  导出为 叠加图_N条.png，存放在分析结果目录。

  与瀑布图的区别：瀑布图只把各条错开看“有哪些峰”；
  这里额外做了峰位跨谱合并标注，并且每条曲线带颜色、带图例。

光谱比对（参考谱文件）
  选一个参考谱，算相关系数与谱角，快速看像不像。"""),

    ("13. 统计分析（聚类·二维成像）", """◆ 聚类分析 + 主成分
  菜单【分析工具】→【聚类分析 + 主成分（所选文件）…】
  先在左侧多选光谱（或先添加整个文件夹再全选），然后：
    · 自动重采样到公共波数区间（点数不粗于最密的那条谱）
    · 每条谱按最小值-最大值归一
    · 以“1 − 相关系数”为距离做平均连接层次聚类
    · 自动按最大间隔确定分割阈值；若所有谱几乎一致，直接判 1 簇
  输出（在“分析结果”里）：
    聚类分析_分组.csv        每条谱的簇号与 PC1~PC3 坐标
    聚类分析_树状图.png      叶子在左、距离在右，标注自动分割阈值
    聚类分析_主成分散点.png  按簇着色，含 PC1/PC2 方差贡献
  命令行：--cluster 文件夹 [--cluster-cut 0.2 手动指定阈值]

◆ 二维成像热图
  菜单【分析工具】→【二维成像热图（所选点位）…】
  把点阵扫描的多个点位光谱按“行优先”铺成 行×列 热图，
  指标可选：
    主峰位 / 主峰强度 / 主峰 FWHM / 识别峰数 / 总强度 /
    指定波数附近的峰强度（例如 at:1008）
  输出热图 PNG + 数值表 CSV。
  命令行：--map 5,5 --map-metric main_peak 文件夹
          --map 2,4 --map-metric at:1008 文件夹

为什么有用：把“峰位漂移”“峰宽变化”“某相强度”变成一张图，
物相分布、应力/蜕晶化趋势一眼就能看出来。"""),

    ("14. 批量处理与分析报告", """◆ 文件夹批处理
  菜单【文件】→【批处理文件夹…（预处理 + 峰表 + 出图 + 汇总）】
  选文件夹 → 勾选输出内容（Excel / PNG / CSV / 峰列表 / JCAMP-DX）
  → 选输出位置 → 点【开始批处理】。
  输出位置三选一：
    源文件夹下的 _转换结果（默认，推荐）
    分析结果文件夹
    源文件夹下（与源文件混放）
  完成后生成 <文件夹名>_批处理汇总.csv，每行一个文件：
    文件、数据点数、识别峰数、主峰位、主峰强度、主峰相对强度、主峰 FWHM
  这是最省事的“整批样品”流程：跑完直接得到一张统计表。
  命令行：--batch 文件夹 [--out 目录] [--peaks --baseline iterpoly …]

◆ 分析报告
  菜单【文件】→【导出分析报告（选中文件）…】
  在左侧选中若干文件后生成自包含 HTML 报告：
    · 总览表（每条谱的数据点数、峰数、主峰位…）
    · 每条谱的峰表（含归属列）
    · 每条谱的光谱图
    · 处理参数说明（用了哪些预处理、阈值多少，方便复现）
  图片已内嵌进 HTML，单文件即可发人；
  浏览器里“打印 → 另存为 PDF”，或用 Word / WPS 直接打开。
  命令行：--report 文件夹 [--out 目录]"""),

    ("15. 高级设置一览", """菜单【设置】→【高级设置…】，分五组：

坐标轴范围
  横坐标最小值 / 最大值、纵坐标最小值 / 最大值（留空 = 自动）

光谱处理
  尖峰去除 + 窗口 + 阈值（见第 7 章）
  基线校正（六种）+ 基线窗口 + 迭代阶数 + 迭代次数
  平滑方式（移动平均 / Savitzky-Golay）+ 平滑窗口(点)
  导数（无 / 一阶 / 二阶）
  归一化（无 / 最大值=1 / 最小-最大）
  峰位归属矿物（填 Zircon 之类，留空 = 不归属）
  堆叠偏移（瀑布图 / 叠加图）
  ☐ 同时应用到导出的数据（CSV/Excel）

拉曼位移校准
  第1对 / 第2对 实测值与标准值

峰识别与峰拟合
  峰灵敏阈值(%)、拟合峰形（高斯/洛伦兹/伪Voigt）
  ☑ 峰标签同时显示相对强度(%)
  ☑ 峰位虚线引到横坐标轴（自动峰 + 手动峰）

图幅与图例
  PNG 输出宽 / 高（默认 1600×900）
  图例位置：右侧留白（默认，不压谱线）/ 上方 / 下方 /
            图内自由位置（可在叠加图预览窗口里拖着放）/ 不显示
  几条谱线就列几条图例，放不下自动分列；图例位置会存成全局默认

改完点【应用】只看效果（不关窗），点【确定】保存并关窗。"""),

    ("16. 命令行速查", """图形界面够用，但批量/自动化时命令行更快。
（把 RamanSpectrumToolkit.exe 换成 python jws2csv.py 效果一样）

基本转换
  --all 文件.jws                导出 CSV+Excel+PNG+峰列表+峰拟合
  --png / --xlsx / --csv / --peaks / --fit / --jcamp
  某文件夹                      递归转换
  --waterfall 文件夹            瀑布图
  --overlay 文件夹              多数据图叠加（每条一色，批处理不带预览）
  --out 输出目录                指定输出位置

出图参数
  --x-step 200 --x-start 0 --x-min --x-max --y-min --y-max
  --y-ticks --no-grid --no-title --no-peaks --no-peak-dash
  --no-peak-labels --peak-dist 20 --peak-thresh 7 --peak-label-rel
  --peak-merge 30 --no-peak-merge --stack-offset 1.0
  --fig-width 1600 --fig-height 900 --manual 1007,974
  --legend-pos right|top|bottom|inside|none          图例位置（几条谱线列几条）

预处理与归属
  --despike --despike-thresh 10 --despike-window 5
  --baseline iterpoly|rolling_ball|rolling|poly2|linear
  --baseline-degree 5 --baseline-iters 20 --baseline-window 60
  --smooth 9 --normalize max --calib 520.6,520.7 --mineral Zircon

分析
  --identify 文件 --identify-top 12              未知谱全库鉴定
  --identify-must "Zr Si" --identify-not Ca      元素筛选
  --identify-kind 拉曼|红外|XRD|全部              限定检索类型
  --identify-index [数据包key]                   建立 / 更新特征索引
  --identify-all / --identify-no-exact           不过滤同名谱 / 只粗筛
  --cluster [文件或文件夹] [--cluster-cut 0.2]
  --map 行,列 --map-metric main_peak|intensity|fwhm|peaks|total|at:1008
  --batch 文件夹 [--out 目录]
  --report [文件或文件夹] [--out 目录]
  --pair 文件 / --db-match 文件
  --pair-batch 文件夹 [--pair-ref 参考谱文件夹]   批量配对（一条谱一行汇总）
  --identify-batch 文件夹 [--identify-top 5]      批量鉴定（逐条给最佳候选）

数据库与矿物
  --db-search 矿物名  --db-get 编号  --db-ls
  --rruff-list  --rruff-get 数据包key  --rruff-search 矿物名或编号
  --rruff-export 矿物名  --rruff-fetch 矿物名  --import-pkg 路径.zip
  --mineral-search 名称  --by-element "Zr Si"  --by-formula SiO2
  --mineral-info 名称  --rruff-info 名称

数据目录与容量
  --data-dir [路径]  --cache-limit [MB]  --cleanup [MB]

说明书
  --manual                      在终端打印本说明书
  --manual 路径.txt             导出说明书到文件

界面语言
  --lang zh|en                  中文 / 英文界面（命令行输出、报告、图注一并切换）

提示：命令行配对报告默认写在“待测文件所在目录”（便于脚本取用），
      界面运行时统一写在“工具数据/分析结果”。"""),

    ("17. 常见任务速查", """我想做的事 → 怎么做

要把 .jws 变成一张能直接看的图
  → 勾选“Excel（含图表）”或“PNG 图片”，点【开始转换】

要峰位、强度、FWHM 的数字表
  → 勾选“峰列表”，得到 *_peaks.csv

自动漏掉的峰要补上
  → 在右侧图上左键点峰顶补标（蓝色方块），右键删错标的

峰位要带振动归属（如 ν3(SiO4)）
  → 菜单【设置】→【矿物信息与峰位归属库…】→ 选矿物 →
    【作为峰位归属参考】→ 重新导出峰列表

荧光背景太强，峰被抬起来
  → 高级设置 → 基线校正选“迭代多项式（推荐）”，阶数 5、迭代 20

谱上有几个尖刺（宇宙射线）
  → 高级设置 → 勾选“尖峰去除”，窗口 5、阈值 8~15

确认我的样品是不是锆石
  → 菜单【数据库】→【RRUFF 数据源…】→ 检索 Zircon → 批量抓取
    → 菜单【分析工具】→【自动配对（本地数据库）】→ 看 F1 排名

完全不知道这条谱是什么矿物
  → 菜单【分析工具】→【未知光谱检索（全库鉴定）…】
    → 先【建立 / 更新特征索引】→ 选好未知光谱 → 【开始检索】
    → 看“结论”与候选表；有 EDS 结果就填“必须含元素”提高准确率

一次处理几十个样品
  → 菜单【文件】→【批处理文件夹…】→ 开始批处理 →
    得到“_批处理汇总.csv”

把一批样品按谱型分组
  → 全选文件 → 菜单【分析工具】→【聚类分析 + 主成分…】→ 看树状图

看某个峰强度在样品上的分布
  → 全选点位谱 → 菜单【分析工具】→【二维成像热图…】→
    指标选“指定波数附近的峰强度”，填 1008

要一份能发人的结果
  → 选中文件 → 菜单【文件】→【导出分析报告…】→ 得到单个 HTML

数据包太占空间
  → 菜单【设置】→【数据文件夹…】→ 看占用 → 清理已下载数据包

仪器峰位有小偏差
  → 高级设置 → 拉曼位移校准填“实测 → 标准”（1 对平移，2 对带缩放）"""),

    ("18. 常见问题", """Q1 数据会不会被写到别的盘？
   不会。下载的参考谱、RRUFF 数据包、分析结果全部在工具目录的
   “工具数据”里；转换你自己的样品时才按“输出位置”保存。
   若工具目录不可写，会自动改用系统用户目录并显示实际路径。

Q2 数据包占空间太大怎么办？
   菜单【设置】→【数据文件夹…】看占用与上限，可一键清理数据包；
   命令行 --cache-limit 设上限、--cleanup 释放指定空间。
   下载前会先检查磁盘剩余空间，不够会直接拒绝。

Q3 双击打开 CSV 只有数据没有图？
   CSV 存不了图。请在“输出内容”里勾选 Excel（含图表）或 PNG 图片。

Q4 为什么横坐标刻度不是整数？
   “横坐标刻度间隔”设为 200、“起始刻度”留空即可；
   想从 0 开始就填 0。

Q5 开了尖峰去除会不会削平窄峰？
   不会。判定要求“偏差达邻域动态范围一定比例”且“连续宽度不超过 3 点”，
   真实窄峰峰顶的偏差相对整段信号起伏很小。不放心就把阈值调大（如 15）。

Q6 归属列显示“未归属”？
   内置库是常见矿物的参考峰位（容差 8 cm-1），个别矿物种或固溶体
   峰会落在容差外。可先确认矿物名，再看 RRUFF 样品记录里的实际峰位。

Q7 自动配对结果不是期望的矿物？
   先看“峰位匹配 F1”而不是相关系数；确认本地库里确有该矿物参考谱；
   校准差超过 5 cm-1 会判未匹配，可用位移校准修正后再配对。
   注意红外/XRD 参考谱存在子目录，不参与拉曼配对。

Q8 抓取时提示“未找到”？
   该矿物可能在更大的数据包里，按提示继续下载即可；
   也可以自己从官网下载 zip 后用【导入本地 zip…】导入。

Q9 杀毒软件报警？
   PyInstaller 打包程序有时被误报，选“信任 / 允许”即可。

Q10 界面上的图能导出成矢量图吗？
    目前导出 PNG（默认 1600×900，可改）；需要矢量图建议用 Excel
    里的图表或 JCAMP-DX 数据在其它软件里重绘。

Q11 未知谱检索说“没找到”，是不是工具不准？
    先确认三件事：①是否已建特征索引；②检索类型是否选对（拉曼谱别去比红外/XRD）；
    ③库里有这个矿物吗——库里没有就一定找不到。unrated 小包只有 600 多条，
    建议下载 fair / excellent 包再试（谱质也更好）。
    另外若你的谱峰位整体偏移较大，先在高级设置里做“拉曼位移校准”再检索。

Q12 “可能的候选”和“可信候选”差别在哪？
    可信候选要求三条同时成立：综合分 ≥ 60、未知谱最强的 3 个峰至少命中 2 个、
    且明显领先第二名（≥15 分）。只要有一条不满足就只报“可能”，
    并在前几名分数接近时明确提醒可能是库里没有对应矿物。
    无论哪一档，最终都建议用【用候选做配对报告】看峰位对照表人工确认。

Q13 能把界面换成英文吗？
    能。菜单【设置】→【语言】里选中 English，或命令行加 --lang en。
    界面、日志、命令行输出、分析报告、图注和说明书都会一起切换；
    说明书同时提供英文版（User_Guide.txt），工具内的【使用说明】窗口
    右上角也有“English / 中文”按钮可即时切换。
    只有磁盘上的数据文件夹名保持中文，保证两种语言下数据互通。"""),

    ("19. 与 RRUFF 工具的关系", """RRUFF 官网“Tools”页面列出的工具，本工具的实现情况：

  RamanCrystalHunter  基线校正（含迭代多项式）、平滑(SG)、归一化、
                      位移校准、导数、加减谱、谱段替换、峰拟合、
                      交互式 A−k·B、库比对识别              已实现
  RamanLab            峰拟合、谱库比对、相似度矩阵、
                      聚类分析（PCA + 层次聚类 + 树状图）、
                      二维成像热图、库管理（导出/导入/清理）  已实现
  Crystal Sleuth      峰位检索 + 配对比较（峰位 F1）        已实现思路
                      （原版基于 XRD 粉末数据，非拉曼）
  RRUFF / ROD 数据库  在线检索 + 按矿物批量抓取 +
                      拉曼/红外/XRD/成分四类数据源 + 编号反查  已接入
  Razor 谱库          曲线比对、谱角、库检索排序             已实现
  XtalDraw            3D 晶体结构绘图                       未实现
                      （属晶体学工具，与拉曼数据处理不同线）
  USPEX               第一性原理 / 进化算法结构预测          无法内置
  Mindat.org          在线矿产地服务（免登录接口不稳定），
                      已改用 RRUFF 样品记录提供化学式/晶系/产地

自定义扩展提示
  · 矿物归属库在源码 jws2csv.py 的 _MINERAL_DB 里，
    照格式加一条即可扩展（含中英文名、化学式、晶系、特征峰与归属）。
  · 参考谱库就是普通 CSV（第一列横坐标、第二列强度），
    任何来源的标准谱放进去都能参与比对与配对。"""),
]


_MANUAL_HEAD_EN = """Raman Spectrum Toolkit  User Guide
(JASCO .jws conversion - Raman peak analysis - mineral identification)
""" + _MANUAL_MARKER_EN + """
============================================================

In one sentence:
  Turns JASCO .jws and other spectrum files into readable charts and data,
  finds and labels peaks, fits them, assigns them to minerals, searches the
  ROD / RRUFF reference libraries (including identifying spectra whose
  mineral is unknown), and writes batch conversions and analysis reports.

How to start (the 30-second version)
  1. Launch the tool: RamanSpectrumToolkit.exe  (or run jws2csv.py with no arguments)
  2. Click [Add files] and choose your .jws file(s), or run the batch entry
  3. Look at the chart on the right; peaks are labelled automatically
  4. Click [Start conversion] and choose Excel / PNG / CSV as needed
  5. Do not know what the sample is? Use
     [Analysis -> Unknown-spectrum search (whole-library identification)...]

Marking conventions in this guide
  [Button]      a button or menu item in the interface
  ->            menu path, e.g. [Settings -> Advanced settings...]
  Values such   recommended values are given directly, e.g. "order 5, 20 iterations"

Language
  The interface, command line and this guide are bilingual (Chinese / English).
  Switch at [Settings -> Language]; the CLI accepts --lang en|zh.
  The first launch follows your Windows display language and remembers your choice.
  Folder names on disk (工具数据 / 参考谱库 / RRUFF数据包 / 分析结果) stay in
  Chinese so that data folders remain portable between the two modes.
"""


_MANUAL_SECTIONS_EN = [
    ("1. Five-minute start", """[Step 1] Launch RamanSpectrumToolkit.exe (or run jws2csv.py) to open the window.
[Step 2] Click [Add files] and choose the .jws you exported from the
  spectrometer; add a whole folder with [Add folder] if you like
  (searched recursively for .jws / .csv / .spc / .jdx / .txt / .xlsx).
[Step 3] Click a file in the left list: the right side draws it at once and
  labels the detected peaks with their wavenumbers.
[Step 4] Choose "What to output" (Excel with chart / PNG / CSV / peak list /
  peak fit / JCAMP-DX), pick an output location, then [Start conversion].
[Step 5] The log below shows the result, and [Open the output folder] takes
  you to the files.

Why does the CSV I double-clicked in WPS show only numbers, no chart?
  A CSV file is plain text: it can hold numbers but not a chart. Please choose
  "Excel (with chart)" (or PNG) as the output, then the chart is embedded.
  Excel output needs the openpyxl component, PNG needs Pillow; if a component
  is missing the tool tells you how to install it (pip install openpyxl).

Peaks were missed
  Left-click on the chart adds a peak manually (blue square). Manual peaks are
  remembered per file and appear in the peak list with source "manual".
  You can also lower the peak sensitivity threshold in [Advanced settings...].

Peaks were detected that should not be there
  Right-click on the peak to delete it - this works for automatic peaks (red
  dots) as well as manual ones. A deleted automatic peak goes onto a per-file
  "deleted" list and is left out of the chart, the peak table and every export.
  Click near the peak (about +-1/40 of the chart width) or the tool will only
  report that no peak is nearby, so right-clicking on empty space never deletes
  anything. [Restore auto peaks] (also under [Settings]) brings them all back,
  and the list is per file, so nothing is unrecoverable."""),

    ("2. Interface overview", """The window is split in two: operations on the left, charts on the right.

Left side
  [1. Choose files]   add files / add a folder / remove / clear, with the hint
                     that .jws becomes data+chart and .csv becomes a chart
  File list          multi-select works (Ctrl / Shift); the preview follows
  [2. Output location] save next to the source file (recommended) or a folder
  [3. Header and column names]  write a header row, custom or auto-detected
                     column names, "skip existing CSV" to avoid overwriting
  [4. Run log]       progress and per-file results

Right side
  [Chart settings]   X tick interval (integer, 200 by default), start tick,
                     hide Y values, grid lines, title, figure size, show peak
                     values, undo / clear manual peaks, restore auto peaks
  Spectrum preview   up to 8 curves overlaid, [Clear view]
  [Save preview as PNG]  saves the current view

Menu bar
  File      batch a folder..., export analysis report..., exit
  Analysis  unknown-spectrum search (whole-library identification), peak
            fitting, compare, peak-position search, similarity matrix,
            clustering, 2D imaging, average, subtract, range replacement,
            waterfall, multi-dataset overlay, spectrum comparison
  Settings  advanced settings..., mineral info and Raman peak assignment...,
            data folder..., clear manual peak annotations, restore deleted
            automatic peaks, language
  Database  search online databases (ROD / RRUFF)..., RRUFF data sources...,
            open data folders
  Help      guide (full manual)..., export the guide as TXT..., about"""),

    ("3. Supported formats", """Read
  .jws    JASCO spectra (OLE2 compound document; DataInfo + Y-Data streams,
          multi-channel supported)
  .csv    two columns of numbers, optional header; a trailing text column
          (notes, source) is dropped automatically
  .spc    Galactic / Thermo SPC (16-bit, 32-bit and float, little endian)
  .jdx    JCAMP-DX, including ##XFACTOR / ##DELTAX / ##XYPOINTS forms
  .txt    .dat .asc .xy two-column text
  .xlsx   .xlsm Excel (needs openpyxl)

Write
  .xlsx   data plus an embedded scatter chart (ScatterChart with a numeric
          X axis, so the tick interval really applies)
  .csv    plain data
  .png    chart image (1600x900 by default, adjustable)
  _peaks.csv    peak list (position, intensity, relative intensity, FWHM,
                prominence, source)
  _fit.csv      peak fitting results (centre, height, FWHM, area, shape, eta, R2)
  .jdx    JCAMP-DX export

Notes
  When a folder is scanned, the tool's own outputs (peak tables, summaries)
  are recognised and skipped, so they are never mistaken for spectra.
  Files in the reference library can themselves be selected as references."""),

    ("4. What is output, and where", """Two choices in [2. Output location]:
  Save next to the source file (recommended)   outputs go beside the .jws
  A folder I choose                            outputs go to that folder

Analysis results (peak fits, pairing reports, similarity matrices, clustering,
imaging, reports, unknown-spectrum searches) always go to
  工具数据/分析结果            inside the tool data folder,
so your sample folders are never cluttered by analysis output.

File naming
  Conversion     <name>.xlsx / .csv / .png
  Peak list      <name>_peaks.csv
  Peak fit       <name>_fit.csv
  Batch summary  <folder>_批处理汇总.csv
  Identification <name>_未知光谱检索.csv  (+ _对比.png overlay)
  Report         光谱分析报告.html  (self-contained, images embedded)

Tip: peak positions and FWHMs are detected automatically. Please check them
manually before publication."""),

    ("5. Data folder and size management", """Every file the tool downloads or produces lives in ONE place:

  <tool folder>/RamanSpectrumToolkit.exe
              User_Guide.txt / 使用说明.txt   (this guide, in EN / ZH)
              assets/icon.png                 (app icon)
              <tool folder>/工具数据/
      ├─ 参考谱库/       reference spectra downloaded from ROD or exported
      │                  from RRUFF (plain CSV, usable offline)
      ├─ RRUFF数据包/    downloaded package zips, search index and feature index
      └─ 分析结果/       peak fits, pairing reports, matrices, imaging, reports

If the tool folder is not writable (for example C:\\Program Files, or a
read-only USB stick) the tool automatically uses
  %LOCALAPPDATA%\\拉曼光谱工具\\工具数据
so it never fails and never writes anywhere else.

Managing it
  [Settings -> Data folder...] shows the location, the file count and the size
  of each subfolder, lets you open or clear a folder individually, and lets you
  move the whole data folder elsewhere (optionally moving the existing data).
  Size limit: the default is 2 GB; before every download the free disk space is
  checked and you are warned if the limit would be exceeded. Cleaning up
  packages also deletes their index and feature index.

Command line
  --data-dir                 show location, file counts and usage
  --data-dir D:\\my\\folder    move the data folder
  --cache-limit 500          set the limit to 500 MB (0 = unlimited)
  --cleanup 500              remove packages until 500 MB have been freed"""),

    ("6. Chart settings and manual peak annotation", """The [Chart settings] panel on the right refreshes the preview immediately:
  X tick interval   integer ticks, 200 cm-1 by default (coordinate snapping)
  Start tick        first labelled tick
  Show Y values     off by default (intensity is relative, so numbers rarely help)
  Grid lines        on/off
  Title             chart title
  Figure size       PNG pixels, 1600x900 by default
  Peak labels       label every detected peak; optionally show relative
                    intensity (%) next to the wavenumber
  Show peak values  on by default; untick it and the wavenumber numbers are no
                    longer drawn while the markers and dashed lines stay - handy
                    when many curves are on the chart

Peak wavelength dashed lines
  Every annotated peak - detected automatically or added by hand - also gets a
  dashed line from the peak top straight down to the x axis, so the wavenumber
  of the peak can be read off at a glance. To hide them, untick "draw a dashed
  line from each peak down to the x axis (auto + manual peaks)" under
  [Settings -> Advanced settings...] -> "Peak detection and fitting"
  (command line: --no-peak-dash).

Manual annotation
  Left-click the chart   add a peak at the nearest peak top (blue square)
  [Undo manual peak]     remove the last one you added
  [Clear manual peaks]   remove them all for the current file
  Manual peaks are kept per file and are marked "manual" in the peak list.

Deleting peaks you do not want (including wrong automatic ones)
  Right-click on a peak deletes the annotated peak nearest to the pointer -
  both automatic peaks (red dots) and manual ones (blue squares).
  An automatic peak is not really "gone": its position goes onto a per-file
  "deleted" list, and the chart, the peak table and every export are then
  computed without it. You must click near the peak (about +-1/40 of the chart
  width); otherwise the tool just says no peak is nearby, so that a stray
  right-click on empty space cannot delete anything.
  [Restore auto peaks] puts every deleted automatic peak of that file back
  (same as [Settings -> Restore deleted automatic peaks]).
  A deleted manual peak is removed for real - add it back with a left-click or
  use [Undo manual peak].
  The lists are per file, so a mistake is always undoable.

When a peak sits on a shoulder, the tool estimates its prominence against a
local baseline so that the annotation and the table stay sensible."""),

    ("7. Preprocessing (baseline, smoothing, spikes, calibration)", """Everything is in [Settings -> Advanced settings...]. By default preprocessing
only affects the chart and peak detection; tick "Also apply to exported data"
if you want it written into the CSV/Excel as well.

Spike removal (cosmic rays / dead pixels)
  Removes narrow spikes using two conditions (deviation from a moving median
  beyond a threshold, and deviation relative to the local dynamic range) plus a
  width limit of three points, so genuine narrow peaks are not flattened.
  Recommended: window 5, threshold 8-15.

Baseline
  linear                  straight line between the two end regions
  polynomial (order 2)    least squares, good for gentle backgrounds
  iterative polynomial    repeatedly fit and drop points above the baseline to
                          follow the valley floor; recommended for strong
                          fluorescence: order 5, 20 iterations
  rolling minimum         centred moving minimum, window in cm-1
  rolling ball            morphological opening, similar but smoother
  none

Smoothing
  moving average, or Savitzky-Golay (preserves peak height and width better).
  Window in points; 3-7 is usually enough. 1 in the SG order box means no
  smoothing.

Derivative
  first or second derivative (Savitzky-Golay based; linear extrapolation at the
  borders so linear signals stay exact).

Normalization
  none / max = 1 / min-max.

Raman shift calibration
  Measured peak -> standard peak. One pair performs a shift; two pairs also
  scale the axis. Example 520.6 -> 520.7 (silicon)."""),

    ("8. Peak detection and peak fitting", """How peaks are found
  A moving average suppresses noise, the noise level is estimated with the MAD
  (median absolute deviation), and a peak must exceed
  max(3 x noise, sensitivity% x intensity range) with a minimum separation in
  cm-1. Default sensitivity is 7% of the intensity range; lower it to find
  smaller peaks. Every peak gets position, intensity, relative intensity (%),
  FWHM, prominence and source.

Peak fitting
  [Analysis -> Peak fitting (selected files)...] chooses Gaussian, Lorentzian
  or pseudo-Voigt per peak. The result contains the fitted centre, height,
  FWHM, area, shape, mixing coefficient eta and R2, and is written to
  <name>_fit.csv.

Peak-position search
  [Analysis -> Peak-position search (reference peak table)...] matches the peak
  table of the spectrum against your own reference table (CSV: name, peak1,
  peak2, ...) with a 5 cm-1 tolerance and ranks the candidates.

Peak list
  [File -> ...] or the conversion options write <name>_peaks.csv; if an
  assignment reference is set, the table gains an "assignment" column."""),

    ("9. Mineral info and peak assignment", """[Settings -> Mineral info and Raman peak assignment...]

  Built-in table: 48 common minerals with 224 characteristic peaks and
  vibrational assignments (e.g. v3(SiO4)). It works offline.
  Search by name, by chemical formula, or by elements (for example "Zr Si",
  meaning all of these elements must be present).

  Selecting a row shows the formula, crystal system, element list and
  characteristic peaks. [Use as peak-assignment reference] makes it the
  reference for the current session: the peak list, the batch summary and the
  HTML report then contain an "assignment" column.

  Assignments use a tolerance of 8 cm-1 and are a guide only.

  RRUFF sample records: if a matching RRUFF package is downloaded, the info
  panel also shows the real sample record - ideal formula, measured formula,
  cell parameters and crystal system, locality, identification status and a
  data link."""),

    ("10. Databases: ROD and RRUFF", """Two data sources
  ROD (Raman Open Database)  the open Raman database of the RRUFF project;
      searched online, one spectrum at a time
  RRUFF data packages        the official zipped collections at
      www.rruff.net/zipped_data_files (Raman by quality grade, plus IR, XRD and
      chemistry packages)

Three steps in the GUI ([Database -> RRUFF data sources...])
  1. Tick the packages you want and download them (the index is built
     automatically). Packages range from 12 MB (unrated, about 600 spectra) to
     229 MB. Prefer the "fair / excellent" grades for better spectra.
  2. Search by mineral name or RRUFF ID (for example Zircon or R050034), then
     download the selected spectra or bulk-fetch by mineral.
  3. Export to the local reference library, then plot / pair-compare.

Where the files go
  Everything is stored under 工具数据 (see chapter 5): the package zips and the
  indexes in RRUFF数据包, the exported reference spectra (plain CSV) in 参考谱库.
  You can also import a zip you downloaded yourself, and IR / XRD / chemistry
  packages are kept in separate subfolders so that spectra of different kinds
  are never mixed up during matching.

An index is built when a package is downloaded; you can rebuild it with
[Database -> ...] or with --identify-index. Deleting a package also deletes its
indexes.

When downloads are slow or keep dropping (important)
  rruff.net is outside China, so the bottleneck is this cross-border link
  (measured: 25% packet loss, 236 ms RTT) rather than a bug in the tool.
  Measured throughput is roughly proportional to the number of connections
  running at the same time (the server never refused): about 0.03 MB/s at 1,
  0.32 at 8, 0.89 at 32, 1.28 at 64 and 1.45 MB/s at 96 connections (all from
  one measurement session; this link swings 2-3x with the time of day, but the
  "more connections is faster" trend is stable).
  That is why the 227 MB package used to take over 40 minutes and had to start
  over after any interruption. Now:
    . the download runs in the BACKGROUND: progress bar + downloaded/total +
      live speed + time left, and the window stays clickable (it used to call
      update_idletasks, so the mouse and the close button were dead the whole time);
    . "Cancel" works at any moment. Whatever has been downloaded is KEPT and
      the next run continues from there instead of starting over;
    . 32 parallel connections by default, adjustable in the download dialog
      (1-128) or with --dl-conns 96. Each connection is guaranteed at least
      256 KB of work so a small package does not spend everything on handshakes;
    . work-stealing: instead of splitting the file into N equal shards (one per
      connection) it is cut into work units in a queue and whichever connection
      finishes grabs the next one. On the same 33 MB package the fastest of 32
      equal shards needed 4.3 s and the slowest needed 79 s - equal shards mean
      waiting 79 s for one stalled connection. Interleaved A/B on the real link
      (12 MB package, 32 connections both ways) gave 39.1 s average for equal
      shards vs 23.2 s with work stealing, i.e. 1.69x faster;
    . proxy: on this lossy link a proxy or VPN is often faster than piling on
      connections. Enter one in the download dialog (blank reuses the Windows
      system proxy) or use --proxy;
    . automatic retries: a dropped connection is retried after 1.5 x n seconds
      and resumes from where it stopped, not from the beginning of the part;
      a 429/503 from the server is detected and reported with a hint to lower
      the concurrency;
    . integrity check: if the byte count does not match it is NOT counted as a
      successful download (it used to be - the failure only surfaced later when
      the index was built, which looked like another crash); the finished file is
      also checked as a zip, and a damaged one is deleted with a message to retry;
    . the .part0 / .part1 files left by a cancel or a force-kill are the resume
      points and are never mistaken for a real package.
  The "estimated time" column uses the speed measured on YOUR last download, so
  it gets more accurate over time. Packages of 50 MB or more ask for confirmation
  first, and you can start with a small one.
  If it is genuinely too slow: try fair_oriented (271 KB) or powder_DIF (7.6 MB)
  first, or raise the concurrency to 64-96, or download the zip in a browser on
  a well-connected network and use "Import local zip..."."""),

    ("11. Pairing and unknown-spectrum identification", """Unknown-spectrum search (whole-library identification, no mineral name needed)

  You have a spectrum and no idea what it is? Use this:
  [Analysis -> Unknown-spectrum search (whole-library identification)...]

  Step 1  Build the feature index (once). It extracts the peak positions,
    relative intensities, ideal formula and crystal system of every spectrum in
    the packages (about 2 seconds for 600 spectra) and caches the result.
    Tick the packages and click [Build / update feature index].
  Step 2  Choose the unknown spectrum. It is taken automatically from the
    selection on the left; [Choose a file...] picks any file directly
    (.jws / .csv / .spc / .jdx / .txt).
  Step 3  Click [Search]. The tool matches the peaks of the unknown spectrum
    against the whole library (fast index pass), then re-scores the top
    candidates from their raw spectra, and lists candidate minerals by
    confidence.

  What the columns mean
    Score       0.5 x F1 + 0.5 x strong-peak hit rate, used for ranking
    F1          peak-position match F1 (5 cm-1 tolerance; 0 if fewer than 2 hits)
    Strong hits how many of the 3 strongest peaks of the unknown were found
    Hits/ref    matched peaks / peaks of that reference
    Mean dev    mean position difference (1-3 cm-1 is normal)
    Correlation curve-shape similarity (only for re-scored candidates)
    Library     how many spectra of that mineral exist in the library
    Chance      probability that a coincidence would match this well

  How to read the conclusion (three levels)
    Reliable candidate  score >= 60, at least 2 strong-peak hits AND clearly
                        ahead of the runner-up (>= 15 points)
    Possible candidate  score >= 35 but the evidence is weaker; if the top
                        scores are very close the tool says so explicitly and
                        reminds you that the mineral may simply not be in the
                        library
    Not found           best score < 35
  An unnamed library record ("unknown") is never used to name your mineral; if
  one is spectrally closer it is reported separately with its peak positions as
  a phase clue.

  Tips
    - If you have EDS / microprobe data, fill in "must contain elements" /
      "must not contain" (e.g. Zr Si) to narrow the search a lot.
    - Search type defaults to Raman; IR, XRD and chemistry packages are
      separate and should not be mixed.
    - Library quality matters: unrated / poor packages are noisy, so download a
      fair / excellent package when accuracy matters.
    - [Export selected to the reference library] and [Pair report with the
      candidate] let you inspect the peak comparison table yourself.

  Command line
    --identify FILE --identify-top 12           search (top 20 re-scored)
    --identify-must "Zr Si" --identify-not Ca   element filter
    --identify-kind Raman|IR|XRD|all            restrict the search type
    --identify-index [PACKAGE_KEY]              build / rebuild the feature index
    --identify-all                              list every spectrum, not per mineral
    --identify-no-exact                         index pass only (faster)

Pairing (manual / automatic)

  [Analysis -> Pairing (manual / auto)] compares measured spectra against
  reference spectra; auto-pairing can search online and download candidates
  first. Scoring is peak-match F1 (the primary criterion, with a minimum of two
  matched peaks so that peak-rich references cannot win by coincidence),
  followed by the correlation coefficient and the spectral angle. The result is
  a ranking CSV plus a pairing report PNG that shows the overlaid curves, the
  matched peaks, the ranking table and the peak comparison table.

  Batch pairing (a whole folder x reference set)
    [Analysis -> Pairing (manual / auto) -> Batch pairing (folder x reference
    spectra)...], or --pair-batch FOLDER [--pair-ref REF_FOLDER].
    The pairing entries above handle ONE spectrum at a time; this is the batch
    wrapper around them:
      . measured spectra: point at a folder (leave blank to use the spectra
        selected in the main window);
      . references: the local reference library, or a folder / file you pick;
      . every spectrum is scored against the reference set and summarised as
        one row per measured spectrum.
    Typical use: a batch of zircons scored against zircon references - the rows
    with a low score are the ones that may contain another crystal, and they are
    listed first.

  Batch identification (identify a whole folder)
    [Analysis -> Batch identification (identify a whole folder)...], or
    --identify-batch FOLDER --identify-top 5.
    For a pile of spectra you do not recognise: each one is searched against the
    whole library and the table lists the best candidate, the score, F1 and the
    conclusion text.
    Needs a feature index first (build it in the "Unknown spectrum search"
    window, or with --identify-index).

  What the summary table contains
    . score = 0.5 x F1 + 0.5 x strong-peak hit rate (the same scoring as
      "Unknown spectrum search");
    . F1, strong-peak hits, matched / measured / reference peak counts, mean
      deviation, correlation, spectral angle, chance probability and the lead
      over the runner-up (meaningful only when the reference set has more than
      one entry);
    . the "Advisory reading" column is only a hint derived from the score
      (good match / partial match / poor match). It does NOT decide for you:
      check the peaks and the profile yourself before calling the phase.
    Output (into 工具数据/分析结果):
      summary CSV (one row per spectrum, ascending score, most doubtful first)
      summary report HTML (same table, printable to PDF from a browser)
    Tip: for a low-scoring row, run the single-spectrum [Auto pairing] on it to
         get the peak comparison table and the report image, and see which peaks
         disagree.

    Two notes on the scoring:
    . a spectrum with only ONE detected peak always scores F1 = 0 (fewer than
      two matches does not count), so its score caps at 50 and it stays in the
      "partial match" band - on purpose, since one peak cannot settle anything;
    . --pair-batch / --identify-batch are batch-only; they open no window."""),

    ("12. Spectrum arithmetic", """All in the [Analysis] menu.

  Average selected spectra     mean of the selected spectra on a common grid
  Subtract spectra (A-B)       B from A (both resampled on the overlap)
  Interactive subtraction A-kB drag the slider for k until the difference
                               flattens; normalizing each to its maximum first
                               is recommended
  Range replacement/stitching  replace a wavenumber range of the target
                               spectrum with the matching range of another
                               spectrum, with a linear transition at the edges
  Waterfall                    many spectra stacked with a vertical offset
  Multi-dataset overlay        [Analysis -> Multi-dataset overlay (selected
  (stacked)                    spectra)...] or --overlay on the command line.
                               Drawn the way stacked-spectra figures are: each
                               curve is normalized to max = 1 and then offset
                               vertically by k x "stack offset", so the curves
                               stay clearly separated instead of piling up on
                               one baseline. One colour per dataset, and the
                               shapes can still be compared horizontally.
                               Offset 1.0 = neighbouring curves just touch;
                               larger separates them further.
                                 . peaks are MERGED ACROSS SPECTRA: the same
                                   peak labelled once per spectrum turns into a
                                   mess, so nearby peaks (within the tolerance)
                                   are treated as one peak - a single dashed
                                   line and a single AVERAGED wavenumber.
                                   The tolerance falls back to "minimum peak
                                   separation" when left blank, or set it in
                                   the dialog (command line: --peak-merge 30).
                                   Each spectrum keeps its own peak markers
                                   (red dots / blue squares) so you can see
                                   which spectra contributed.
                                 . the x axis uses the INTERSECTION of every
                                   dataset's wavenumber range, so a shorter
                                   spectrum does not leave a blank strip on the
                                   right; the dialog states the range it used.
                                   An explicit range in [Advanced settings...]
                                   still wins.
                                 . each curve is normalized to max = 1 by
                                   default; untick it to keep raw intensities
                                   and compare real relative strengths
                                 . untick "show peak values" to drop the
                                   numbers and keep only the dashed lines and
                                   markers - much clearer with a dozen curves
                                   (the main panel's "Show peak values" does
                                   the same)
                                 . the legend lists EVERY curve (not just the
                                   first 12); when it does not fit it splits
                                   into several columns. By default it sits in
                                   a reserved gutter to the RIGHT of the plot,
                                   so however dense the curves are it never
                                   covers them. Choose right / above / below /
                                   free position inside / none, or just DRAG
                                   the legend in the preview window and drop
                                   it where you want. The position is saved in
                                   the settings file and reused for the export
                                   and the next session
                                 . the dialog opens as a PREVIEW first:
                                   changing settings or editing peaks only
                                   redraws the preview - nothing is written to
                                   the results folder until you press
                                   "Export PNG"
                                     - left-click the plot = add a peak there
                                       (blue dashed line + blue value)
                                     - right-click a dashed line = remove the
                                       nearest peak (auto-detected or manual)
                                     - "Re-detect" = drop every manual edit and
                                       go back to the auto-detected peaks
                                     - hold and drag the legend = move it
                                       anywhere (clicking elsewhere on the plot
                                       still adds a peak)
                                     - press Enter or "Refresh preview" to
                                       redraw after typing a parameter
                                   Adding or removing peaks requires "Merge
                                   peaks across spectra (one peak, one line,
                                   one value)", because only merging gives one
                                   line and one value per peak
                                   The command line (--overlay) is batch mode
                                   with no preview window: it draws the
                                   auto-detected peaks straight away. Use
                                   --legend-pos right|top|bottom|inside|none to
                                   pick the legend position on the CLI
                               Exported as 叠加图_N条.png into the results folder.

                               Difference from the waterfall: the waterfall just
                               offsets the curves to show WHICH peaks are there;
                               this one also merges and labels the peaks across
                               spectra, and every curve is coloured and named in
                               a legend.
  Compare with a reference     compare one spectrum with a single reference
  Pairing with the local library  compare against every spectrum in 参考谱库

Outputs are CSV plus a PNG preview for each operation."""),

    ("13. Statistics (clustering and 2D imaging)", """Clustering + PCA
  [Analysis -> Cluster analysis + PCA (selected files)...]
  Hierarchical clustering with average linkage on 1 - correlation, with PCA for
  the principal components. Auto cut: the split threshold is chosen
  automatically (when the spectra are nearly identical you get a single
  cluster); you can also enter a value yourself.
  Outputs: a grouping table, a dendrogram PNG and a PCA scatter PNG (with the
  variance explained).

2D imaging
  [Analysis -> 2D imaging heat map (selected points)...]
  The selected point spectra are laid out row-major into a grid; choose the
  metric: main peak position / intensity / FWHM / number of peaks / total
  intensity / intensity at a given wavenumber. Outputs: a heat-map PNG and a
  value table CSV."""),

    ("14. Batch processing and analysis reports", """Batch a folder
  [File -> Batch a folder...] runs preprocessing + peak tables + plots +
  conversion over a whole folder, and writes <folder>_批处理汇总.csv with one
  row per file (number of peaks, main peak position, intensity, FWHM).
  Useful even for thousands of files.

Analysis report
  [File -> Export analysis report (selected files)...] writes a single
  self-contained 光谱分析报告.html: an overview table, the peak table of every
  spectrum (with assignments when a reference is set), each spectrum chart and
  the processing parameters. Images are embedded, so the file can be e-mailed
  as it is, printed to PDF from the browser, or opened in Word / WPS."""),

    ("15. Advanced settings at a glance", """[Settings -> Advanced settings...], five groups:

  Axis range        X min / max, Y min / max (blank = automatic)
  Spectrum processing  spike removal (window, threshold), baseline (linear /
                    polynomial order 2 / iterative polynomial / rolling
                    minimum / rolling ball + window), smoothing (moving average /
                    Savitzky-Golay + window), derivative, normalization
  Calibration       Raman shift calibration (one or two measured -> standard
                    pairs; the second pair is only needed for two-point scaling)
  Peaks             sensitivity threshold (% of the intensity range, smaller is
                    more sensitive), minimum separation (cm-1), peak shape for
                    fitting, assignment reference mineral, "show relative
                    intensity with peak labels", "draw a dashed line from each
                    peak down to the x axis (auto + manual peaks)"
  Figure size       PNG pixels
  Legend            position: right-hand gutter (default, never covers the
                    curves) / above / below / free position inside the plot
                    (draggable in the overlay preview) / hidden.
                    Every curve gets an entry; extra entries wrap into more
                    columns. The position is saved as a global default

  "Also apply to exported data" decides whether preprocessing is written into
  the CSV/Excel, or only used for the chart and peak detection."""),

    ("16. Command-line quick reference", """The GUI is usually enough, but the command line is faster for batches and
automation. Run  jws2csv.py --help  for a short summary.

Conversion
  jws2csv.py sample.jws                 convert using the defaults
  jws2csv.py folder --png --peaks       whole folder, PNG + peak list
  jws2csv.py --xlsx --out D:\\out a.jws  Excel with chart into a folder
  jws2csv.py --skip-existing            never overwrite existing CSV
  jws2csv.py --waterfall folder         stacked waterfall chart
  jws2csv.py --overlay folder           overlaid chart, one colour per dataset
                                        (batch mode, no preview)

Processing (same names as the advanced settings)
  --despike --baseline iterative --order 5 --iters 20
  --smooth sg --window 7 --savgol-order 2
  --deriv 1 --norm max --calib 520.6:520.7
  --x-min 100 --x-max 2000 --tick 200 --fig-width 1600 --fig-height 900
  --legend-pos right|top|bottom|inside|none   legend position
  --no-peaks --no-peak-dash --no-peak-labels
  --peak-merge 30 --no-peak-merge --stack-offset 1.0
  --header-custom "Wavenumber,Intensity"

Analysis
  --identify FILE --identify-top 12              unknown-spectrum search
  --identify-must "Zr Si" --identify-not Ca      element filter
  --identify-kind Raman|IR|XRD|all               restrict the search type
  --identify-index [PACKAGE_KEY]                 build / update the feature index
  --identify-all / --identify-no-exact           all entries / index pass only
  --cluster [FILE|FOLDER] [--cluster-cut 0.2]    clustering + PCA
  --map 5,5 [--map-metric main_peak|intensity|fwhm|peaks|total|at:1000]
  --pair FILE                                    auto-pair against the local library
  --pair-batch FOLDER [--pair-ref REF_FOLDER]    batch pairing (one row per spectrum)
  --identify-batch FOLDER [--identify-top 5]     batch identification (best candidate each)
  --compare FILE --ref FILE                      compare with one reference
  --average FILES / --subtract A B / --waterfall FILES
  --report FILES                                 HTML report

Database
  --rruff-list                            list packages and their state
  --rruff-get KEY                         download a package (index built automatically)
  --rruff-search Zircon                   search in the local packages
  --rruff-fetch Zircon                    bulk fetch by mineral and export
  --rruff-export Zircon                   export to the local reference library
  --import-pkg path.zip                   import a package you downloaded
  --db-search quartz                      search ROD online
  --db-get IDS                            download spectra from ROD
  --mineral-search Zircon                 built-in mineral table
  --mineral-search --by-element "Zr Si"   search by elements

Data and documentation
  --data-dir [PATH]        show or change the data folder
  --cache-limit 500        size limit in MB (0 = unlimited)
  --cleanup 500            free 500 MB by removing packages
  --manual [PATH]          print the guide, or export it to a file
  --lang en|zh             interface language
  --gui                    force the graphical interface"""),

    ("17. Common tasks", """I want to do this -> how

Convert a whole batch of .jws to Excel charts
  -> [File -> Batch a folder...], or jws2csv.py folder --xlsx --png --peaks

Check whether my sample is zircon
  -> [Database -> RRUFF data sources...] -> search Zircon -> bulk fetch
  -> [Analysis -> Auto-pair (local database)] -> read the F1 ranking

I have no idea what this spectrum is
  -> [Analysis -> Unknown-spectrum search (whole-library identification)...]
  -> [Build / update feature index] -> choose the spectrum -> [Search]
  -> read the conclusion and the candidate table; fill in "must contain
     elements" if you have EDS results

The fluorescence background is strong and peaks sit on a slope
  -> Advanced settings -> Baseline: "iterative polynomial (recommended)",
     order 5, 20 iterations

There are a few spikes (cosmic rays) on the spectrum
  -> Advanced settings -> tick "spike removal", window 5, threshold 8-15

Peaks are missing
  -> lower the peak sensitivity threshold in the advanced settings (for
     example from 7% to 4%), or click them onto the chart manually

The X axis is slightly off compared with the standard
  -> use a standard sample (silicon 520.7) and fill in the calibration pair,
     e.g. 520.6 -> 520.7

I want a figure for publication
  -> set the figure size (for example 2400x1350), hide the Y values, use a
     200 cm-1 tick interval, then [Save preview as PNG]
  -> or export the peak data as JCAMP-DX and redraw it elsewhere

I want to compare many spectra at once
  -> select them and use the similarity matrix, or the waterfall plot, or
     clustering + PCA

My peaks are wider after irradiation
  -> fit them (Gaussian / Lorentzian / pseudo-Voigt), look at the FWHM and the
     area in the fit table, and compare samples with the similarity matrix"""),

    ("18. Frequently asked questions", """Q1 Will data be written to another disk?
  No. Downloads (reference spectra and RRUFF packages), analysis results and
  reports all go to 工具数据 inside the tool folder (or to
  %LOCALAPPDATA%\\拉曼光谱工具\\工具数据 when the tool folder is read-only). The
  only exception is the conversion output, which by default goes next to your
  source file because that is usually what you want; change [2. Output location]
  if you prefer somewhere else.

Q2 Why does my CSV show only numbers in WPS?
  A CSV cannot contain a chart. Choose "Excel (with chart)" or PNG instead.

Q3 The chart is blank / all zeros
  Check whether the file really contains spectral data (open it as text: two
  columns of numbers), whether the X column is numeric, and whether the axis
  range in the advanced settings excludes your data.

Q4 Excel export failed with a message about openpyxl
  Install it:  pip install openpyxl  (PNG export needs  pip install pillow).

Q5 RRUFF download fails
  Check the network and the proxy; the tool prints the reason in the log. If a
  package is too large you can download a smaller one first (12 MB). You can
  also download a zip yourself and use [Database -> ... -> Import a local zip...].

Q6 Why does the pairing report show a low F1 even for the right mineral?
  F1 is strict: peaks that the reference has but your spectrum does not lower
  the score. Check the peak comparison table and the correlation coefficient as
  well, and try a fair / excellent package (better spectra). Slight systematic
  axis differences also matter - use the calibration.

Q7 Why is the peak position slightly different from the literature?
  Instrument calibration and the laser wavelength shift peaks by a few cm-1.
  Use the calibration pair to align with a standard, and always compare within
  the same laser wavelength.

Q8 Can several .jws channels be handled?
  Yes; each channel is exported and plotted separately and the peak table has a
  "channel" column.

Q9 Can I use my own reference spectra?
  Yes. The reference library is a plain CSV folder (first column X, second
  column intensity). Drop files into 工具数据/参考谱库, or export spectra from
  the database; they can then be used for pairing, comparison and combination.

Q10 Can charts be exported as vector graphics?
  Currently PNG (1600x900 by default, adjustable). For vector output use the
  Excel chart or redraw the JCAMP-DX data in another program.

Q11 The unknown-spectrum search says "not found" - is the tool wrong?
  Check three things: (1) is the feature index built; (2) is the search type
  correct (do not compare a Raman spectrum with IR/XRD packages); (3) does the
  library contain the mineral at all - if it does not, nothing can find it. The
  unrated package has only about 600 spectra; download a fair / excellent
  package and try again. If your peaks are systematically shifted, calibrate the
  spectrum first (see chapter 7).

Q12 What is the difference between "possible" and "reliable" candidate?
  "Reliable" requires all three: score >= 60, at least 2 of the 3 strongest
  peaks matched, and a clear lead (>= 15 points) over the runner-up. If any of
  them fails you only get "possible", and when the top scores are close the tool
  says explicitly that the mineral may not be in the library. In either case the
  final check is the peak comparison table from the pairing report.

Q13 Can the interface be in English?
  Yes: [Settings -> Language], or --lang en on the command line. The
  full manual and this guide exist in both languages; folder names on disk stay
  in Chinese so that data remains portable."""),

    ("19. Relation to the RRUFF tools", """The tools listed on the RRUFF "Tools" page, and how this tool covers them:

  Crystal Sleuth      reading and plotting spectra, peak tables, JCAMP-DX
                      import/export            -> covered (chapters 3, 4)
  RamanLab            baseline correction, smoothing, derivatives, peak
                      fitting, peak assignment to minerals, batch processing
                      -> covered (chapters 7, 8, 9, 14)
  RamanCrystalHunter  comparing spectra with a reference library, peak-match
                      scoring and candidate ranking -> covered, with F1 scoring
                      and whole-library unknown-spectrum identification
                      (chapter 11)
  XtalDraw / USPEX    three-dimensional crystal structure drawing and structure
                      prediction -> NOT covered: they are crystallography and
                      first-principles modelling tools, a different line of work
                      from Raman data processing

Features added beyond those tools
  - RRUFF packages (Raman / IR / XRD / chemistry) with local search, bulk fetch
    and an offline reference library
  - unknown-spectrum identification with confidence levels and an element filter
  - 2D imaging heat maps and clustering + PCA
  - self-contained HTML reports and folder batch processing

Extending the tool
  - The mineral assignment table is _MINERAL_DB in jws2csv.py; add one record
    (English and Chinese names, formula, crystal system, characteristic peaks
    and assignments) to extend it.
  - The reference library is plain CSV (first column X, second column
    intensity); standards from any source can take part in comparison and
    pairing."""),
]


def manual_sections(lang=None):
    """取指定语言的章节表（默认当前界面语言）。"""
    code = str(lang or ui_lang()).lower()
    return _MANUAL_SECTIONS_EN if code.startswith("en") else _MANUAL_SECTIONS


def manual_text(lang=None):
    """生成完整说明书文本（工具内阅读与导出 TXT 共用同一份内容，中英各一份）。"""
    code = "en" if str(lang or ui_lang()).lower().startswith("en") else "zh"
    head = _MANUAL_HEAD_EN if code == "en" else _MANUAL_HEAD
    sections = manual_sections(code)
    parts = [head]
    for i, (title, _body) in enumerate(sections, 1):
        parts.append("%2d. %s" % (i, title.split(". ", 1)[-1]))
    parts.append("")
    for title, body in sections:
        parts.append("")
        parts.append("=" * 60)
        parts.append(title)
        parts.append("=" * 60)
        parts.append("")
        parts.append(body.strip("\n"))
    parts.append("")
    parts.append("=" * 60)
    parts.append("End of the guide. Good luck with your experiments!"
                 if code == "en" else "说明书结束。祝实验顺利！")
    parts.append("=" * 60)
    return "\n".join(parts) + "\n"


def manual_path(lang=None):
    """说明书文件路径：中文 使用说明.txt / 英文 User_Guide.txt。"""
    code = "en" if str(lang or ui_lang()).lower().startswith("en") else "zh"
    return os.path.join(_app_dir(), "User_Guide.txt" if code == "en" else "使用说明.txt")


def ensure_manual_file(force=False):
    """说明书文件与程序同步：中英各生成一份；用户自建同名文件不动。

    返回当前语言的说明书路径（供界面日志显示）。
    """
    for code in ("zh", "en"):
        path = manual_path(code)
        text = manual_text(code)
        if not force and os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8-sig", errors="replace") as f:
                    head = f.read(400)
            except OSError:
                head = ""
            if _MANUAL_MARKER in head or _MANUAL_MARKER_EN in head:
                continue
            if head and ("光谱转换" not in head and "使用说明" not in head
                         and "拉曼光谱工具" not in head
                         and "Raman Spectrum Toolkit" not in head):
                continue
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(text)
        except OSError:
            pass
    return manual_path()


def _attach_parent_console():
    """打包成窗口程序后，带命令行参数运行时把输出接回父控制台。"""
    if os.name != "nt":
        return False
    try:
        import ctypes
        if not ctypes.windll.kernel32.AttachConsole(-1):
            return False
        stream = open("CONOUT$", "w", encoding="utf-8", errors="replace", buffering=1)
        sys.stdout = stream
        sys.stderr = stream
        return True
    except Exception:
        return False


def _prepare_console():
    """命令行输出准备：控制台切到 UTF-8，且即使切不动也绝不因编码崩溃。"""
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleOutputCP(65001)
            ctypes.windll.kernel32.SetConsoleCP(65001)
        except Exception:
            pass
    for name in ("stdout", "stderr"):
        stream = getattr(sys, name, None)
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass


def main():
    args = sys.argv[1:]
    if args and getattr(sys, "frozen", False) and sys.stdout is None:
        _attach_parent_console()
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = sys.stdout
    if args:
        _prepare_console()
        sys.exit(_cli(args))
    _run_gui()


if __name__ == "__main__":
    main()
