# -*- coding: utf-8 -*-
"""图例摆放自测：不许压谱线、条目数要跟谱线数对得上、位置能选也能拖。

分两段：
* 纯几何（不开窗）——用 render_overlay 返回的绘图区几何，
  逐项检查图例框与绘图区**不重叠**（right/top/bottom）、inside 落在图内、
  条目数与谱线数一致（40 条会自动分列而不是砍掉）；
* 真实 Tk 窗口——在叠加图预览窗口里拖动图例，检查位置被写进设置文件，
  并且导出图沿用同一个位置。
"""
import math
import os
import sys
import tempfile
import tkinter as tk
import tkinter.messagebox as mb

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
import jws2csv as T

mb.showinfo = lambda *a, **k: None
mb.showwarning = lambda *a, **k: None
mb.showerror = lambda *a, **k: None
mb.askyesno = lambda *a, **k: False

ok = [0, 0]


def check(label, cond, extra=""):
    ok[0 if cond else 1] += 1
    print("  [%s] %s %s" % ("OK" if cond else "FAIL", label, extra))


def walk(w, out=None):
    if out is None:
        out = []
    for c in w.winfo_children():
        out.append(c)
        walk(c, out)
    return out


def make_series(n, pts=300):
    out = []
    for k in range(n):
        xs = [400.0 + i * (1400.0 / pts) for i in range(pts)]
        ys = []
        for x in xs:
            v = 0.0
            for c in (520, 700, 900, 1100, 1300):
                v += math.exp(-((x - (c + k * 6)) ** 2) / (2 * 18.0 ** 2))
            ys.append(v + 0.02 * ((k % 3) - 1))
        out.append(("谱线%02d_Sample_%d" % (k, k), xs, ys))
    return out


def overlap(box, rect):
    """盒子与绘图区是否有交叠（都按 (x0,y0,x1,y1) 算）。"""
    if not box:
        return False
    return not (box[2] <= rect[0] or box[0] >= rect[2]
                or box[3] <= rect[1] or box[1] >= rect[3])


tmp = os.path.join(tempfile.gettempdir(), "_raman_legend_test")
os.makedirs(tmp, exist_ok=True)
png = os.path.join(tmp, "legend.png")

print("=== 图例几何（不压谱线）===")
for n in (2, 5, 12, 20):
    series = make_series(n)
    for pos in ("right", "top", "bottom"):
        g = T.render_overlay(png, series, "叠加图 %d 条" % n, "波数", "强度",
                             {"legend_pos": pos, "show_y_ticks": True},
                             return_geometry=True)
        rect = (g["ml"], g["mt"], g["ml"] + g["pw"], g["mt"] + g["ph"])
        check("%2d 条 · %-6s 图例不压绘图区" % (n, pos),
              g["legend_box"] and not overlap(g["legend_box"], rect),
              "图例 %s 绘图区 %s" % (g["legend_box"], rect))
        check("%2d 条 · %-6s 图例条目数与谱线数一致" % (n, pos),
              g["legend_rows"] == n, "图例 %d 条 / 谱线 %d 条" % (g["legend_rows"], n))

g = T.render_overlay(png, make_series(40), "叠加图 40 条", "波数", "强度",
                     {"legend_pos": "right", "show_y_ticks": True},
                     return_geometry=True)
rect = (g["ml"], g["mt"], g["ml"] + g["pw"], g["mt"] + g["ph"])
check("40 条 · right 自动分列后条目齐全（不再只列前 12 条）",
      g["legend_rows"] == 40, "图例 %d 条" % g["legend_rows"])
check("40 条 · right 分列后依然不压绘图区", not overlap(g["legend_box"], rect),
      str(g["legend_box"]))
check("40 条 · right 分列后绘图区还留得下（没被图例挤没）", g["pw"] >= 500,
      "绘图区宽 %d" % g["pw"])

g = T.render_overlay(png, make_series(6), "叠加图 6 条", "波数", "强度",
                     {"legend_pos": "none"}, return_geometry=True)
check("图例位置选“不显示”时不画图例", g["legend_box"] is None and g["legend_rows"] == 0,
      str(g["legend_box"]))

print()
print("=== inside：图内自由位置 ===")
g = T.render_overlay(png, make_series(4), "叠加图 4 条", "波数", "强度",
                     {"legend_pos": "inside", "legend_xy": [0.05, 0.62],
                      "show_y_ticks": True}, return_geometry=True)
rect = (g["ml"], g["mt"], g["ml"] + g["pw"], g["mt"] + g["ph"])
box = g["legend_box"]
check("inside 图例框落在绘图区内",
      box and box[0] >= rect[0] and box[1] >= rect[1]
      and box[2] <= rect[2] and box[3] <= rect[3], str(box))
check("inside 图例框听 legend_xy（左上角按比例落位）",
      box and abs(box[0] - (g["ml"] + 0.05 * g["pw"])) <= 2
      and abs(box[1] - (g["mt"] + 0.62 * g["ph"])) <= 2,
      "期望 x=%.0f y=%.0f，实际 x=%.0f y=%.0f"
      % (g["ml"] + 0.05 * g["pw"], g["mt"] + 0.62 * g["ph"], box[0], box[1]))

# 拖动往返：把画出来的框反算成比例再渲染一次，位置必须稳住（不然拖完会跳）
fx = (box[0] - g["ml"]) / float(g["pw"])
fy = (box[1] - g["mt"]) / float(g["ph"])
g2 = T.render_overlay(png, make_series(4), "叠加图 4 条", "波数", "强度",
                      {"legend_pos": "inside", "legend_xy": [fx, fy],
                       "show_y_ticks": True}, return_geometry=True)
check("按画出来的框反算比例再渲染，位置不跳（拖动才不会漂）",
      abs(g2["legend_box"][0] - box[0]) <= 2 and abs(g2["legend_box"][1] - box[1]) <= 2,
      "%s → %s" % (box, g2["legend_box"]))

check("legend_hit 能命中图例框、也能区分框外",
      T.legend_hit(box, (box[0] + box[2]) / 2, (box[1] + box[3]) / 2)
      and not T.legend_hit(box, box[2] + 40, box[3] + 40)
      and not T.legend_hit(None, 10, 10), "")

print()
print("=== 其它图：配对报告 5 种位置都能出图，且不压图表 ===")
from PIL import Image  # noqa: E402
series = make_series(2)
results = [{"name": "Zircon_R050034", "ref_xy": series[1][1:], "peak_score": 92.0,
            "peak_matched": 5, "peak_left": 1, "peak_right": 5, "corr": 0.98,
            "angle": 4.2,
            "pair_rows": [[521.3, 520.9, 0.4, True], [698.7, 699.2, -0.5, True],
                          [899.3, 900.1, -0.8, True]]}]
for pos in T.LEGEND_POSITIONS:
    p = os.path.join(tmp, "pair_%s.png" % pos)
    g = T.render_pair_report(p, "Sample_A", (series[0][1], series[0][2]), results,
                             {"legend_pos": pos}, return_geometry=True)
    check("配对报告图 · %-6s 能出图且非空" % pos,
          os.path.isfile(p) and Image.open(p).size == (1700, 1180),
          "%s" % (os.path.getsize(p) if os.path.isfile(p) else "-"))
    if pos == "none":
        check("配对报告图 · none 不画图例", g["legend_box"] is None, str(g["legend_box"]))
        continue
    if pos != "inside":
        check("配对报告图 · %-6s 图例不压图表" % pos,
              not overlap(g["legend_box"], g["chart"]),
              "图例 %s 图表 %s" % (g["legend_box"], g["chart"]))
    if pos == "bottom":
        check("配对报告图 · bottom 图例排在排名表上方（不跟表叠字）",
              g["legend_box"][3] <= g["table_top"], "图例下沿 %.0f / 表顶 %d"
              % (g["legend_box"][3], g["table_top"]))
    check("配对报告图 · %-6s 排名表与峰位对照表不重叠" % pos,
          g["pair_top"] - (g["table_top"] + 34 + 26 * (len(results) + 1)) >= 20,
          "第二张表 %d / 第一张表底 %d"
          % (g["pair_top"], g["table_top"] + 34 + 26 * (len(results) + 1)))

print()
print("=== 设置文件：图例位置是可持久化的全局默认 ===")
before = T._load_settings()
orig_pos, orig_xy = before.get("legend_pos"), before.get("legend_xy")
check("未设置过时默认给 right", T.legend_defaults()[0] in T.LEGEND_POSITIONS, "")
T.save_legend_defaults("inside", [0.25, 0.5])
check("存进去再读回来一致（含自由位置的比例坐标）",
      T.legend_defaults() == ("inside", [0.25, 0.5]), str(T.legend_defaults()))
T.save_legend_defaults("top", None)
check("改成预设位置时不再残留自由坐标",
      T.legend_defaults() == ("top", None), str(T.legend_defaults()))
check("写进设置文件的键名是 legend_pos / legend_xy",
      T._load_settings().get("legend_pos") == "top"
      and "legend_xy" not in T._load_settings(), "")
# 还原用户原来的设置
if orig_pos is None and orig_xy is None:
    data = T._load_settings()
    data.pop("legend_pos", None)
    data.pop("legend_xy", None)
    T._save_settings(data)
else:
    T.save_legend_defaults(orig_pos or "right", None)


def find_pos_combo(dlg):
    """找到图例位置那个只读下拉框（它带 5 个固定选项）。"""
    for w in walk(dlg):
        try:
            if w.winfo_class() == "TCombobox":
                vals = list(w.cget("values"))
                if len(vals) == len(T.LEGEND_POSITIONS):
                    return w, vals
        except Exception:
            continue
    return None, []


def tip_text(root):
    for w in walk(root):
        try:
            if w.winfo_class() == "TLabel":
                t = str(w.cget("text"))
                if t.startswith("图例") or "图例" in t:
                    return t
        except Exception:
            continue
    return ""


def mainloop(self):
    app = self
    app.update()

    lib = os.path.join(ROOT, "工具数据", "参考谱库")
    names = sorted(n for n in os.listdir(lib) if n.lower().endswith(".csv"))[:3] \
        if os.path.isdir(lib) else []
    if len(names) < 3:
        check("参考谱库至少 3 条", False, "只有 %d 条" % len(names))
        app.destroy()
        return
    paths = [os.path.join(lib, n) for n in names]
    outdir = os.path.join(tempfile.gettempdir(), "_raman_legend_out")
    os.makedirs(outdir, exist_ok=True)
    for f in os.listdir(outdir):
        try:
            os.remove(os.path.join(outdir, f))
        except OSError:
            pass
    app.same_dir.set(False)
    app.dir_var.set(outdir)
    app.files = list(paths)
    app.listbox.delete(0, "end")
    for n in names:
        app.listbox.insert("end", n)
    app.listbox.selection_set(0, "end")
    app.update()

    # ---------- 主界面预览：图例同样画在绘图区外面 ----------
    def right_side_lines():
        cv_main = app.canvas
        cv_main.update_idletasks()
        view = app._view
        if not view:
            return None
        hits = 0
        for item in cv_main.find_all():
            if cv_main.type(item) != "line":
                continue
            xs = cv_main.coords(item)[0::2]
            if xs and min(xs) > view["ml"] + view["pw"]:
                hits += 1
        return hits

    app.adv_values["legend_pos"] = "right"
    app.preview_selected()
    app.update()
    hits_right = right_side_lines()
    check("主界面预览：图例画在绘图区右侧留白里（不压谱线）",
          hits_right is not None and hits_right > 0, "绘图区右侧的色块线 %s 条" % hits_right)
    app.adv_values["legend_pos"] = "none"
    app.preview_selected()
    app.update()
    check("主界面预览：选“不显示图例”后绘图区右侧不再有图例色块",
          right_side_lines() == 0, "%s 条" % right_side_lines())
    app.adv_values["legend_pos"] = "right"
    app.preview_selected()
    app.update()

    before_wins = {id(w) for w in app.winfo_children() if isinstance(w, tk.Toplevel)}
    app.tool_overlay()
    app.update()
    new = [w for w in app.winfo_children()
           if isinstance(w, tk.Toplevel) and id(w) not in before_wins]
    if not new:
        check("叠加图预览窗口已打开", False, "")
        app.destroy()
        return
    dlg = new[-1]
    for _ in range(10):
        dlg.update_idletasks()
        app.update()

    cvs = [w for w in walk(dlg) if w.winfo_class() == "Canvas"]
    cv = cvs[0] if cvs else None
    check("预览窗口里有画布", cv is not None, "")
    combo, vals = find_pos_combo(dlg)
    check("预览窗口里有【图例位置】下拉框", combo is not None, str(vals))
    check("下拉框提供 5 个位置选项", len(vals) == 5, str(vals))
    if cv is None or combo is None:
        dlg.destroy()
        app.destroy()
        return

    # 用同一套参数离线重算一遍图例框，好算出“该按在哪儿”才能抓住图例
    pos_now = {v: k for k, v in {
        "right": "绘图区右侧留白（推荐）", "top": "绘图区上方",
        "bottom": "绘图区下方", "inside": "图内自由位置（可拖动）",
        "none": "不显示图例"}.items()}.get(str(combo.get()), "right")
    opts = app.plot_options()
    opts.update({"annotate_peaks": True, "peak_labels": True, "show_title": True,
                 "stack_offset": 1.0, "peak_merge_tol": None,
                 "legend_pos": pos_now, "legend_xy": None})
    probe = os.path.join(tmp, "probe.png")
    geom = T.render_overlay(probe, app._selected_spectra(),
                            T.T("多数据图叠加（%d 条）" % len(paths)),
                            T._DEFAULT_X_HEADER, T.T("归一化强度"), opts,
                            normalize=True, merge_peaks=True, return_geometry=True)
    box = geom["legend_box"]
    scale = cv.winfo_width() / float(geom["width"]) if geom["width"] else 1.0
    cx = int((box[0] + box[2]) / 2 * scale)
    cy = int((box[1] + box[3]) / 2 * scale)
    check("图例框（right）在画布上有对应的可点区域",
          box and 0 < cx < cv.winfo_width() and 0 < cy < cv.winfo_height(),
          "框 %s → 画布 (%d, %d)" % (box, cx, cy))

    # ---------- 拖动图例 ----------
    cv.event_generate("<Button-1>", x=cx, y=cy)
    app.update()
    cv.event_generate("<B1-Motion>", x=cx - int(cv.winfo_width() * 0.28), y=cy + 40)
    app.update()
    cv.event_generate("<ButtonRelease-1>", x=cx - int(cv.winfo_width() * 0.28), y=cy + 40)
    app.update()

    saved = T._load_settings()
    check("拖完图例位置变成 inside（图内自由位置）",
          saved.get("legend_pos") == "inside", str(saved.get("legend_pos")))
    xy = T.legend_defaults()[1]
    check("拖完把比例坐标写进了设置文件",
          xy and 0.0 <= xy[0] <= 1.0 and 0.0 <= xy[1] <= 1.0, str(xy))
    check("拖到偏左下方，x 比例确实比原来小",
          xy and xy[0] < (box[0] + box[2]) / 2.0 / geom["width"], str(xy))
    check("下拉框同步显示为“图内自由位置”",
          "图内" in str(combo.get()), str(combo.get()))

    # ---------- 导出沿用拖好的位置 ----------
    btns = {}
    for w in walk(dlg):
        try:
            if w.winfo_class() == "TButton":
                btns[str(w.cget("text"))] = w
        except Exception:
            continue
    if "导出 PNG" in btns:
        btns["导出 PNG"].invoke()
        app.update()
        made = [f for f in os.listdir(outdir)
                if f.startswith("叠加图_") and f.endswith(".png")]
        check("拖完图例后仍能正常导出叠加图", len(made) == 1, str(made))
        if made:
            exp = T.render_overlay(
                os.path.join(tmp, "expect.png"), app._selected_spectra(),
                T.T("多数据图叠加（%d 条）" % len(paths)), T._DEFAULT_X_HEADER,
                T.T("归一化强度"), dict(opts, legend_pos="inside", legend_xy=xy),
                normalize=True, merge_peaks=True, return_geometry=True)
            ef = (exp["legend_box"][0] - exp["ml"]) / float(exp["pw"])
            check("导出图用的就是拖好的位置（按同一比例重算，框位置一致）",
                  abs(ef - xy[0]) <= 0.05, "%.3f vs %.3f" % (ef, xy[0]))

    for f in os.listdir(outdir):
        try:
            os.remove(os.path.join(outdir, f))
        except OSError:
            pass
    try:
        os.rmdir(outdir)
    except OSError:
        pass
    try:
        dlg.destroy()
    except Exception:
        pass
    app.update()
    app.destroy()


tk.Tk.mainloop = mainloop
# 开窗之前先把图例设置清干净，免得受上一次跑测试留下的位置影响（跑完再还原）
_orig_settings = T._load_settings()
_orig_legend = (_orig_settings.get("legend_pos"), _orig_settings.get("legend_xy"))
_clean = dict(_orig_settings)
_clean.pop("legend_pos", None)
_clean.pop("legend_xy", None)
T._save_settings(_clean)

T._run_gui()

_restore = T._load_settings()
_restore.pop("legend_pos", None)
_restore.pop("legend_xy", None)
if _orig_legend[0]:
    _restore["legend_pos"] = _orig_legend[0]
if _orig_legend[1]:
    _restore["legend_xy"] = _orig_legend[1]
T._save_settings(_restore)

import shutil  # noqa: E402
shutil.rmtree(tmp, ignore_errors=True)

print()
print("图例摆放自测：通过 %d 项，失败 %d 项" % (ok[0], ok[1]))
sys.exit(1 if ok[1] else 0)

