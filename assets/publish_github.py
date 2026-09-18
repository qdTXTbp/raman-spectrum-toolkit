# -*- coding: utf-8 -*-
"""把当前源码 + 绿色版压缩包发布到 GitHub 仓库。

用法：
    set GH_TOKEN_FILE=C:\\path\\to\\token.txt      &  python assets/publish_github.py
    python assets/publish_github.py --token-file C:\\path\\to\\token.txt

令牌来源优先级：
    --token-file 参数  >  环境变量 GH_TOKEN_FILE  >  环境变量 GH_TOKEN
令牌只用于 HTTP 头，全程不打印、不写入任何受版本管理的文件。

流程：
    1. 取令牌 → GET /user 确认身份
    2. 逐个文件上传为 blob（含二进制）→ tree → commit → 更新 main 分支
    3. 设置仓库 topics
    4. 建 Release（已存在则沿用）并上传 dist/RamanSpectrumToolkit-v<版本>.zip
"""
import argparse
import base64
import json
import os
import re
import sys
import time
import urllib.error
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OWNER = "qdTXTbp"
REPO = "raman-spectrum-toolkit"
API = "https://api.github.com"

DESC = ("Raman Spectrum Toolkit: convert JASCO .jws / CSV / SPC / JCAMP-DX spectra to CSV, "
        "Excel, PNG; detect, delete and fit Raman peaks, assign mineral bands, search ROD / "
        "RRUFF reference libraries, identify unknown spectra one by one or a whole folder in "
        "batch, pair measured spectra against a reference set, overlay multiple datasets with "
        "draggable legends. Bilingual EN/ZH Windows desktop tool. 拉曼光谱工具：转换 · 分析 · 矿物鉴定")

TOPICS = ["raman-spectroscopy", "raman", "spectroscopy", "spectrum-converter", "jasco",
          "jws", "mineral-identification", "rruff", "rod-database", "peak-fitting",
          "xrd", "infrared", "desktop-app", "tkinter", "python", "bilingual",
          "geoscience", "materials-science"]

FILES = [
    ".gitignore", "README.md", "LICENSE", "CHANGELOG.md", "jws2csv.py",
    "使用说明.txt", "User_Guide.txt", "启动拉曼光谱工具.bat",
    "assets/icon.png", "assets/icon.ico", "assets/icon_64.b64",
    "assets/make_icon.py", "assets/embed_icon.py",
    "assets/make_readme_images.py", "assets/make_release.py",
    "assets/publish_github.py",
    "docs/example_spectrum.png", "docs/example_pairing.png", "docs/example_overlay.png",
    "tests/test_i18n_kernel.py", "tests/test_ui_english.py",
    "tests/test_output_english.py", "tests/test_cli_english.py",
    "tests/test_dialog_layout.py", "tests/test_overlay_dash.py",
    "tests/test_overlay_preview.py", "tests/test_batch_pair_identify.py",
    "tests/test_download_engine.py", "tests/test_legend_layout.py",
]

TOK = None


def version():
    src = open(os.path.join(ROOT, "jws2csv.py"), encoding="utf-8").read()
    m = re.search(r'_MANUAL_VERSION\s*=\s*"([^"]+)"', src)
    return m.group(1) if m else "0.0"


def read_token(path):
    if path:
        if not os.path.isfile(path):
            raise SystemExit("找不到令牌文件：%s" % path)
        with open(path, "r", encoding="utf-8-sig") as f:
            t = f.read().strip()
    else:
        t = (os.environ.get("GH_TOKEN") or "").strip()
    t = t.strip().strip('"').strip("'")
    if not t:
        raise SystemExit("没有拿到令牌。请用 --token-file 指定文件，"
                         "或设置环境变量 GH_TOKEN_FILE / GH_TOKEN。")
    print("· 令牌已读入（长度 %d，不回显任何字符）" % len(t))
    return t


def call(method, url, data=None, raw=None, ctype="application/json",
         ok=(200, 201, 204), tries=4):
    """发一个 GitHub API 请求。网络抖动（连接被重置 / 超时）自动重试。"""
    global TOK
    headers = {"Authorization": "Bearer " + TOK,
               "Accept": "application/vnd.github+json",
               "X-GitHub-Api-Version": "2022-11-28",
               "User-Agent": "raman-spectrum-toolkit-publish"}
    body = None
    if raw is not None:
        body = raw
        headers["Content-Type"] = ctype
    elif data is not None:
        body = json.dumps(data).encode("utf-8")
        headers["Content-Type"] = ctype

    last = None
    for attempt in range(1, tries + 1):
        req = urllib.request.Request(url, data=body, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=900) as r:
                txt = r.read().decode("utf-8", "replace")
                return r.status, (json.loads(txt) if txt.strip() else {})
        except urllib.error.HTTPError as e:
            txt = e.read().decode("utf-8", "replace")
            if e.code in ok:
                return e.code, (json.loads(txt) if txt.strip() else {})
            # 5xx 多为服务端抖动，值得重试；4xx 是真错，直接报
            if e.code < 500:
                raise SystemExit("HTTP %s %s %s\n%s" % (e.code, method, url, txt[:400]))
            last = "HTTP %s %s" % (e.code, txt[:160])
        except Exception as e:                      # 网络类异常
            last = "%s: %s" % (type(e).__name__, e)
        if attempt < tries:
            wait = 3 * attempt
            print("   ~ 第 %d 次失败（%s），%d 秒后重试 %s"
                  % (attempt, last, wait, url.rsplit("/", 1)[-1][:40]))
            time.sleep(wait)
    raise SystemExit("多次重试仍失败 %s %s\n%s" % (method, url, last))


def main():
    global TOK
    ap = argparse.ArgumentParser()
    ap.add_argument("--token-file", default=os.environ.get("GH_TOKEN_FILE") or "")
    args = ap.parse_args()

    tag = "v" + version()
    zip_path = os.path.join(ROOT, "dist", "RamanSpectrumToolkit-%s.zip" % tag)
    TOK = read_token(args.token_file)

    _st, me = call("GET", API + "/user")
    login = me.get("login")
    email = "%s@users.noreply.github.com" % (me.get("id") or login)
    print("· 已认证：%s" % login)
    if login != OWNER:
        raise SystemExit("令牌身份是 %s，与预期 %s 不一致" % (login, OWNER))

    # 1) 仓库已存在则沿用（本仓库是早先建好的，这里只做兜底）
    #    细粒度令牌通常没有“创建仓库”权限，403 也当已存在处理，后面几步会真正报错
    st, r = call("POST", API + "/user/repos",
                 {"name": REPO, "description": DESC, "private": False,
                  "auto_init": False, "has_issues": True, "has_wiki": False},
                 ok=(201, 422, 403))
    if st == 201:
        print("· 仓库已创建：%s" % r.get("html_url"))
    elif st == 403:
        print("· 令牌无建仓权限，按“仓库已存在”继续")

    # 2) 空仓库必须先用 Contents API 打一个初始提交，否则 Git Data API 会 409
    st, commits = call("GET", "%s/repos/%s/%s/commits?per_page=1" % (API, OWNER, REPO),
                       ok=(200, 409))
    if st == 409:
        with open(os.path.join(ROOT, ".gitignore"), "rb") as f:
            seed = f.read()
        st, init = call("PUT", "%s/repos/%s/%s/contents/.gitignore" % (API, OWNER, REPO),
                        {"message": "chore: 初始化仓库",
                         "content": base64.b64encode(seed).decode("ascii"),
                         "branch": "main"})
        parent = init["commit"]["sha"]
        print("· 空仓库已初始化：%s" % parent[:8])
    else:
        parent = commits[0]["sha"] if commits else None
        print("· 已有提交，接续：%s" % (parent[:8] if parent else "无"))

    # 3) 逐个文件上传为 blob，然后一次性提交
    blobs = []
    for rel in FILES:
        full = os.path.join(ROOT, rel.replace("/", os.sep))
        if not os.path.isfile(full):
            print("   ! 跳过缺失文件 %s" % rel)
            continue
        with open(full, "rb") as f:
            content = f.read()
        _st, b = call("POST", "%s/repos/%s/%s/git/blobs" % (API, OWNER, REPO),
                      {"content": base64.b64encode(content).decode("ascii"),
                       "encoding": "base64"})
        blobs.append({"path": rel, "mode": "100644", "type": "blob", "sha": b["sha"]})
        print("   + %-42s %8d B" % (rel, len(content)))

    _st, tree = call("POST", "%s/repos/%s/%s/git/trees" % (API, OWNER, REPO),
                     {"tree": blobs})
    _st, commit = call("POST", "%s/repos/%s/%s/git/commits" % (API, OWNER, REPO),
                       {"message": "%s 图例摆放：几条谱线列几条、改到绘图区外不压谱线，"
                                   "可在预览窗口拖动并记住位置" % tag,
                        "tree": tree["sha"],
                        "parents": ([parent] if parent else []),
                        "author": {"name": login, "email": email},
                        "committer": {"name": login, "email": email}})
    st, _ref = call("POST", "%s/repos/%s/%s/git/refs" % (API, OWNER, REPO),
                    {"ref": "refs/heads/main", "sha": commit["sha"]}, ok=(201, 422))
    if st == 422:
        call("PATCH", "%s/repos/%s/%s/git/refs/heads/main" % (API, OWNER, REPO),
             {"sha": commit["sha"], "force": True})
    print("· 已提交 %d 个文件，commit %s" % (len(blobs), commit["sha"][:8]))

    # 4) topics（细粒度令牌可能没有该权限；不是必需步骤，403 只提示不中断）
    st, _t = call("PUT", "%s/repos/%s/%s/topics" % (API, OWNER, REPO),
                  {"names": TOPICS}, ok=(200, 201, 403))
    if st == 403:
        print("· 令牌无 topics 权限，跳过（仓库原有话题保持不变）")
    else:
        print("· topics 已设置：%d 个" % len(TOPICS))

    # 5) Release（已存在则沿用）
    st, rel = call("GET", "%s/repos/%s/%s/releases/tags/%s" % (API, OWNER, REPO, tag),
                   ok=(200, 404))
    if st == 200:
        print("· Release 已存在，沿用：%s" % rel.get("html_url"))
    else:
        _st, rel = call("POST", "%s/repos/%s/%s/releases" % (API, OWNER, REPO),
                        {"tag_name": tag, "target_commitish": "main",
                         "name": "%s · 拉曼光谱工具 Raman Spectrum Toolkit" % tag,
                         "body": RELEASE_BODY.replace("__TAG__", tag),
                         "draft": False, "prerelease": False})
        print("· Release 已创建：%s" % rel.get("html_url"))

    # 6) 上传压缩包
    if os.path.isfile(zip_path):
        size = os.path.getsize(zip_path)
        up = ("https://uploads.github.com/repos/%s/%s/releases/%d/assets?name=%s"
              % (OWNER, REPO, rel["id"], os.path.basename(zip_path)))
        print("· 正在上传 %.1f MB …" % (size / 1048576.0))
        with open(zip_path, "rb") as f:
            _st, asset = call("POST", up, raw=f.read(),
                              ctype="application/zip", ok=(201, 422))
        print("· 附件：%s（%s）" % (asset.get("name"), asset.get("state")))
    else:
        print("! 没找到压缩包 %s（先跑 assets/make_release.py）" % zip_path)

    print("\n完成 → https://github.com/%s/%s" % (OWNER, REPO))


RELEASE_BODY = """## 拉曼光谱工具 · Raman Spectrum Toolkit __TAG__

JASCO `.jws` 光谱转换 · 拉曼峰分析 · 矿物鉴定（中英双语，Windows 绿色版）

### 下载
下载下面的压缩包，解压到任意目录，双击 `RamanSpectrumToolkit.exe` 即可。
免安装、不写注册表，所有数据都写在解压目录的 `工具数据/` 里。

### 本次新增
- **图例不再压谱线，位置也由你定**。
  以前图例固定画在**绘图区里的右上角**，靠一层白底卡片“盖住但读得清”——
  堆叠图上曲线从下到上铺满整幅，图例必然压掉一块谱线。
  - **位置可选**：右侧留白（默认）/ 上方 / 下方 / 图内自由位置 / 不显示；
    选“上方”“下方”时绘图区、标题、横坐标标题会一起让开，互不叠字。
  - **可以直接拖**：在叠加图预览窗口里按住图例拖到任意位置，松手即定位；
    按在图上别处仍然是“加峰位”，两种操作互不干扰（拖动时有虚线预览框）。
  - **位置会记住**：写进设置文件，导出 PNG 和下次预览都用同一个位置；
    【设置 → 高级设置 → 图幅与图例】里也能选。
  - **条目数不再被砍**：以前“图例只列前 12 条”，20 条谱线就有 8 条对不上号。
    现在**几条谱线就列几条**，竖着放不下自动分列（最多 4 列），
    实在放不下才省略并写明“还有 N 条没列”。
  - 对所有出图生效：叠加图、瀑布图、鉴定对比图、配对对比报告、主界面预览，
    命令行 `--legend-pos right|top|bottom|inside|none`。
  - 修掉连带 bug：配对对比报告选“下方”时排名表让开了、
    但“最佳配对峰位对照”表没跟着让，两张表会叠字——现在一起下移。

### 顺带回顾上一版
- **下载数据库又快又稳**：后台下载 + 进度条 + 实时速率 + 剩余时间 + 取消；
  断点续传（切法记在 `.part.meta`，中途改并发数也能续）；并发默认 32 路（1~128）；
  **动态领活**消除“慢尾”；可填代理；自动重试；长度校验。
  真实链路同一时段 A/B：静态均分 39.1 s → 动态领活 23.2 s，快 1.69 倍。
- **批量配对**：一个文件夹 × 一组参考谱，**对不上的排在最前面**；
  命令行 `--pair-batch 文件夹 [--pair-ref 参考谱文件夹]`。
- **批量鉴定**：命令行 `--identify-batch 文件夹 [--identify-top 5]`。
  汇总表只给辅助数据（综合分 / F1 / 强峰命中 / 相关系数 / 谱角 / 领先分），
  判读只是提示，**最终是哪个物相由你自己核对峰位与谱型来判**。
- 叠加图「先预览、左键加峰 / 右键删峰、确认后再导出」。

### 修复
- 配对对比报告选“下方”图例时，排名表与峰位对照表叠字。
- 鉴定结论里“仅领先第二名”偶尔显示负数——现在夹到 0。
- RRUFF 数据包对话框的表头在英文模式下没有翻译。

### 画质与稳定性
- 十套自测合计 **308 项全部通过**，其中新增的 `tests/test_legend_layout.py`
  （62 项）逐项检查 right/top/bottom 的图例框与绘图区**零重叠**、
  inside 落在图内且听比例坐标、40 条谱线自动分列后条目齐全、
  拖动往返位置不漂、图例位置能存能读能还原、配对报告两张表不重叠，
  并在真实 Tk 窗口里**模拟一次鼠标拖动**，验证位置被写进设置且导出沿用同一位置。
- 中英两份说明书（`使用说明.txt` / `User_Guide.txt`）同步更新至 2.5。

### 已知口径（不是 bug）
- 只识别到 **1 个峰**的谱，F1 一律记 0（命中 < 2 不算识别），综合分因此封顶 50、
  会一直停在「部分匹配」档。这是刻意的：单个峰不足以定案，宁可让人工去看。

### 数据说明
参考谱与数据包来自 RRUFF 项目与 Raman Open Database，请遵守其使用条款；
本工具与上述项目无隶属关系。
"""


if __name__ == "__main__":
    main()
