# -*- coding: utf-8 -*-
"""Mermaid → PNG 离线渲染：用本机 Edge 无头模式渲染（mermaid.min.js 已本地化到 vendor）。

render_mermaid_to_png(code, out_png, out_svg=None) -> bool
  两轮 Edge：
  1) --dump-dom 渲染 mermaid，提取 SVG（并校验无语法错误；out_svg 给定则存矢量文件）
  2) SVG 按原生尺寸嵌入极简页面，--force-device-scale-factor=3 输出高清 PNG
任何一步失败返回 False（调用方兜底为留空白占位）。
"""
import json
import os
import re
import subprocess
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
MERMAID_JS = os.path.join(ROOT, "static", "vendor", "mermaid.min.js")

EDGE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

MAX_PIXELS = 16000  # 无头截图单边像素上限（Chromium 纹理上限 16384，留余量）


def find_edge():
    for p in EDGE_CANDIDATES:
        if os.path.isfile(p):
            return p
    return None


def _file_url(path):
    return "file:///" + path.replace("\\", "/")


def _embed_code(code):
    """LLM 生成的 mermaid 代码安全嵌入 JS 字符串。"""
    return json.dumps(code).replace("</", "<\\/")


def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))


def _unescape_html(s):
    return (s.replace("&amp;", "&").replace("&lt;", "<")
             .replace("&gt;", ">").replace("&quot;", '"').strip())


def _label_px(text, fs=16.0):
    """估算文本像素宽：CJK/全角按 1em，其余按 0.56em。"""
    return sum(fs if ord(c) > 0x2E7F else fs * 0.56 for c in text)


def _wrap_label(text, max_w, fs=16.0):
    """按像素宽贪心换行。模拟浏览器在 max-width 内的换行效果（标记里没有 <br/> 可恢复）。"""
    def wch(ch):
        return fs if ord(ch) > 0x2E7F else fs * 0.56
    lines, cur, cw = [], "", 0.0
    for ch in text:
        w = wch(ch)
        if cw + w > max_w and cur:
            lines.append(cur)
            cur, cw = ch, w
        else:
            cur, cw = cur + ch, cw + w
    if cur:
        lines.append(cur)
    return lines


def _svg_textify(svg):
    """把 mermaid 节点标签的 <foreignObject>（HTML 渲染）转成原生 SVG <text>。

    Word/WPS 的 SVG 解析器不支持 foreignObject——矢量嵌入后文字会整体丢失
    （PDF 里只剩框和线）。转成原生 text 后 Word 正常渲染，且 PNG 截图结果不变。
    长文本在 foreignObject 里靠 CSS max-width:200px 视觉换行（无 <br/> 标记），
    必须按框宽重新换行成多个 <text>，否则单行文字横向溢出框线。
    注意：短标签的 foreignObject 宽度恰好等于文本宽度（如「根」w=16），
    估算宽 ≤ w+2 时保持单行，只有触到 200px 上限的长文本才按 w-4 换行。
    """
    def rep(m):
        w, h, inner = float(m.group(1)), float(m.group(2)), m.group(3)
        cm = re.search(r"color:\s*(rgb\([^)]+\)|#[0-9a-fA-F]{3,8})", inner)
        color = cm.group(1) if cm else "#333333"
        text = _unescape_html(re.sub(r"<[^>]+>", "", inner))
        if not text:
            return ""
        fs = 16.0
        lines = [text] if _label_px(text, fs) <= w + 2 else _wrap_label(text, w - 4, fs)
        lh = fs * 1.5  # 与 foreignObject 的 line-height:1.5 一致，保证行数不超框高
        y0 = h / 2 - (len(lines) - 1) / 2.0 * lh + fs * 0.35
        out = []
        for i, ln in enumerate(lines):
            out.append('<text x="%.1f" y="%.1f" text-anchor="middle" '
                       'font-family="Microsoft YaHei, Segoe UI, sans-serif" '
                       'font-size="%d" style="fill:%s">%s</text>'
                       % (w / 2, y0 + i * lh, fs, color, _xml_escape(ln)))
        return "".join(out)
    return re.sub(r'<foreignObject width="([\d.]+)" height="([\d.]+)">([\s\S]*?)</foreignObject>',
                  rep, svg)


def _render_html(code):
    return """<!doctype html><html><head><meta charset="utf-8">
<script src="%s"></script>
</head><body>
<div id="graph"></div>
<script>
var code = %s;
mermaid.initialize({startOnLoad:false, securityLevel:"loose",
  flowchart:{curve:"basis", htmlLabels:false, padding:8, nodeSpacing:18, rankSpacing:32}});
mermaid.render("mm", code).then(function (r) {
  document.getElementById("graph").innerHTML = r.svg;
  document.body.setAttribute("data-done", "1");
}).catch(function (e) {
  document.body.setAttribute("data-done", "err");
});
</script></body></html>""" % (_file_url(MERMAID_JS), _embed_code(code))


def _shot_html(svg, w, h):
    # 去掉 mermaid 的 max-width 限制，按像素固定尺寸
    svg = re.sub(r'style="[^"]*"', "", svg, count=1)
    svg = re.sub(r'width="[^"]*"', 'width="%d"' % w, svg, count=1)
    if 'height="' not in svg.split(">", 1)[0]:
        svg = svg.replace("<svg ", '<svg height="%d" ' % h, 1)
    else:
        svg = re.sub(r'height="[^"]*"', 'height="%d"' % h, svg, count=1)
    return """<!doctype html><html><head><meta charset="utf-8">
<style>html,body{margin:0;padding:0;background:#ffffff;overflow:hidden}
#wrap{padding:12px;display:inline-block}</style>
</head><body><div id="wrap">%s</div></body></html>""" % svg


def _run_edge(args, timeout=60):
    edge = find_edge()
    if not edge:
        raise RuntimeError("未找到 Edge 浏览器")
    return subprocess.run([edge, "--headless", "--disable-gpu", "--hide-scrollbars"] + args,
                          capture_output=True, timeout=timeout)


def _render_code_to_svg(code):
    """第一轮：Edge 渲染 mermaid 代码，返回 (textify 后的干净 svg, w0, h0)；失败返回 None。"""
    tmpdir = tempfile.mkdtemp(prefix="mmd_")
    try:
        h1 = os.path.join(tmpdir, "render.html")
        with open(h1, "w", encoding="utf-8") as f:
            f.write(_render_html(code))
        r = _run_edge(["--virtual-time-budget=20000", "--dump-dom", _file_url(h1)])
        dom = r.stdout.decode("utf-8", "ignore")
        if 'data-done="1"' not in dom:
            return None  # 语法错误或超时
        m = re.search(r"<svg[\s\S]*?</svg>", dom)
        if not m:
            return None
        svg = _svg_textify(m.group(0))  # foreignObject → 原生 <text>（Word 不支持 foreignObject）
        vb = re.search(r'viewBox="[\d.\-]+\s+[\d.\-]+\s+([\d.]+)\s+([\d.]+)"', svg)
        if not vb:
            return None
        w0, h0 = float(vb.group(1)), float(vb.group(2))
        if w0 < 10 or h0 < 10:
            return None
        # 去掉 mermaid 的 max-width 限制
        svg = re.sub(r'style="[^"]*"', "", svg, count=1)
        return svg, w0, h0
    except Exception:
        return None
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)


def _render_svg_to_png(svg, w, h, out_png, scale=3.0):
    """第二轮：SVG 原生尺寸 + 高 device-scale-factor 截图。"""
    sc = min(scale, MAX_PIXELS / (w + 24), MAX_PIXELS / (h + 24))
    tmpdir = tempfile.mkdtemp(prefix="mmdshot_")
    try:
        h2 = os.path.join(tmpdir, "shot.html")
        with open(h2, "w", encoding="utf-8") as f:
            f.write(_shot_html(svg, int(w), int(h)))
        _run_edge(["--virtual-time-budget=5000",
                   "--force-device-scale-factor=%.3f" % sc,
                   "--window-size=%d,%d" % (int(w) + 24, int(h) + 24),
                   "--screenshot=" + os.path.abspath(out_png),
                   _file_url(h2)])
        return os.path.isfile(out_png) and os.path.getsize(out_png) >= 2000
    except Exception:
        return False
    finally:
        import shutil
        shutil.rmtree(tmpdir, ignore_errors=True)


def render_mermaid_to_png(code, out_png, scale=3.0, out_svg=None):
    """把 mermaid 代码渲染成单张 PNG。成功返回 True（保留给演示/调试用）。"""
    if not code or not code.strip() or not os.path.isfile(MERMAID_JS):
        return False
    r = _render_code_to_svg(code)
    if not r:
        return False
    svg, w0, h0 = r
    if out_svg:
        try:
            with open(out_svg, "w", encoding="utf-8") as f:
                f.write(svg)
        except Exception:
            pass
    return _render_svg_to_png(svg, w0, h0, out_png, scale)


# ---------------- 分页切片 ----------------
NODE_HALF_H = 22.0  # 兜底半高（取不到 rect 高度时）：单行标签节点高 40 + 余量

def _node_rows(svg):
    """从 SVG 提取所有节点的 y 区间（节点组 id 形如 flowchart-<id>-<n>）。
    高度取组内 rect 的实际高度——长文本节点会自动换行变高（48/72px），
    用固定半高会把多行节点从中间切断。"""
    rows = []
    for m in re.finditer(
            r'<g class="node [^"]*?"[^>]*?transform="translate\((-?[\d.]+),\s*(-?[\d.]+)\)"[^>]*?>([\s\S]*?)</g>', svg):
        y = float(m.group(2))
        rh = re.search(r'<rect[^>]*height="([\d.]+)"', m.group(3))
        half = float(rh.group(1)) / 2 + 5 if rh else NODE_HALF_H
        rows.append((y - half, y + half))
    return rows


def _slice_cuts(svg, h0, max_h):
    """按节点间隙贪心选择切点：每段高度 ≤ max_h，切点不穿过任何节点。

    切点会切断跨段的长连线（root→远端分支），视觉上连线在切片边缘截断，
    这是分页打印的正常表现。
    """
    rows = _node_rows(svg)
    if not rows:
        return [0.0, h0]
    rows.sort()
    merged = []
    for t, b in rows:
        if merged and t <= merged[-1][1] + 1:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([t, b])
    gaps = []
    for i in range(1, len(merged)):
        g0, g1 = merged[i - 1][1], merged[i][0]
        if g1 - g0 > 6:
            gaps.append((g0 + g1) / 2)
    cuts = [max(0.0, merged[0][0] - 8)]
    while h0 - cuts[-1] > max_h:
        limit = cuts[-1] + max_h
        cand = [g for g in gaps if cuts[-1] + 120 < g <= limit]
        if cand:
            cuts.append(cand[-1])
        else:
            # 窗口内没有间隙（合并块比一页还高）：从 limit 往前找穿过节点数最少的位置
            best, best_n = limit, None
            y = limit
            while y > cuts[-1] + 120:
                n = sum(1 for t, b in merged if t < y < b)
                if best_n is None or n < best_n:
                    best, best_n = y, n
                    if n == 0:
                        break
                y -= 2
            cuts.append(best)
    cuts.append(h0)
    return cuts


def _slice_svg(svg, y0, y1, w0):
    """同一份 SVG 换 viewBox 裁出 [y0, y1) 区段（内容不动，矢量无损）。"""
    def rep(m):
        return re.sub(r'viewBox="[^"]*"',
                      'viewBox="0 %.2f %.2f %.2f"' % (y0, w0, y1 - y0),
                      m.group(0), count=1)
    return re.sub(r"<svg[^>]*>", rep, svg, count=1)


def render_mermaid_slices(code, out_dir, scale=3.0, max_aspect=1.39):
    """渲染并按节点间隙切片，让每段打印时顶满版心宽度（字更大）。

    max_aspect: 每段 高/宽 上限（21.5/15.5 ≈ 1.39，对应版心一页）。
    产出 mindmap_1.png/.svg、mindmap_2.png/.svg……
    返回 [{"png","svg","w","h"}...]；只有一段时也切片返回（由调用方统一按列表处理）。
    失败返回 None。
    """
    if not code or not code.strip() or not os.path.isfile(MERMAID_JS):
        return None
    r = _render_code_to_svg(code)
    if not r:
        return None
    svg, w0, h0 = r
    cuts = _slice_cuts(svg, h0, w0 * max_aspect)
    out = []
    for i in range(len(cuts) - 1):
        y0, y1 = cuts[i], cuts[i + 1]
        ssvg = _slice_svg(svg, y0, y1, w0)
        png_path = os.path.join(out_dir, "mindmap_%d.png" % (i + 1))
        svg_path = os.path.join(out_dir, "mindmap_%d.svg" % (i + 1))
        if not _render_svg_to_png(ssvg, w0, y1 - y0, png_path, scale):
            return None
        try:
            with open(svg_path, "w", encoding="utf-8") as f:
                f.write(ssvg)
        except Exception:
            pass
        out.append({"png": png_path, "svg": svg_path, "w": w0, "h": y1 - y0})
    return out


if __name__ == "__main__":
    demo = """flowchart TB
  root(["声音是什么"]):::rootLevel
  root --> a("产生"):::level1
  root --> b("传播"):::level1
  a --> a1("物体振动发声"):::level2
  b --> b1("需要介质"):::level2
  b1 --> c1("真空不能传声"):::level3
  classDef rootLevel fill:#1e3a8a, stroke:#1e3a8a, stroke-width:2px, color:#ffffff;
  classDef level1 fill:#e0f2fe, stroke:#0284c7, stroke-width:2px, color:#075985;
  classDef level2 fill:#fef3c7, stroke:#d97706, stroke-width:1.5px, color:#92400e;
  classDef level3 fill:#f3f4f6, stroke:#9ca3af, stroke-width:1px, color:#374151;"""
    out = os.path.join(ROOT, "runtime", "_mmd_demo.png")
    ok = render_mermaid_to_png(demo, out)
    print("render:", ok, out if ok else "")
