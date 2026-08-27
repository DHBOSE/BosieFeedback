# -*- coding: utf-8 -*-
"""课堂反馈生成器 · 本地后端（标准库 http.server，零第三方依赖）

多模型供应商架构：DeepSeek / Kimi(Moonshot) / 阿里云百炼 / 自定义 OpenAI 兼容端点，
配置存于 config.json（仅存本地，前端只能看到掩码 key）。

接口：
  GET  /api/providers              → 供应商列表（掩码 key）+ 当前选择
  POST /api/config                 → 保存供应商配置/当前选择
  POST /api/generate               → {overview, notes, mindmap?, meta?, provider?, model?}
  GET  /api/file?job=xx&fmt=docx|pdf[&download=1]
  GET  /                           → static/index.html
启动：python server.py --host 0.0.0.0 --port 7100
"""
import argparse
import base64
import json
import os
import re
import threading
import time
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import llm_client
import mermaid_render
import pdf_convert
import report_builder

ROOT = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(ROOT, "static")
RUNTIME_DIR = os.path.join(ROOT, "runtime")
STUDENTS_DIR = os.path.join(RUNTIME_DIR, "_students")  # 学情总结缓存 / 学情 PDF 产物
BRAND_DIR = os.path.join(ROOT, "brand")
CONFIG_PATH = os.path.join(ROOT, "config.json")
os.makedirs(RUNTIME_DIR, exist_ok=True)
os.makedirs(BRAND_DIR, exist_ok=True)

BRAND_SLOTS = ("cover", "header", "footer")
BRAND_DEFAULTS = {
    "org_name": "", "contact": "",
    "cover": {"enabled": True, "height_cm": 1.2},
    "header": {"enabled": True, "height_cm": 0.7},
    "footer": {"enabled": True, "height_cm": 0.5},
}
# 正文大水印默认配置（type: text/image；size_pct 占版心宽百分比；
# opacity_pct 为不透明度，越小越浅；angle 逆时针为正；font 见 report_builder.WM_FONT_TABLE）
WATERMARK_DEFAULTS = {"enabled": False, "type": "text", "text": "",
                      "size_pct": 60, "opacity_pct": 15, "angle": -45,
                      "font": "msyh", "color": "216873"}

MAX_BODY = 40 * 1024 * 1024  # 40MB（脑图 base64）
_config_lock = threading.Lock()


# ---------------- 配置 ----------------
def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    # 兼容旧版扁平结构（DEEPSEEK_API_KEY / DEEPSEEK_MODEL）
    if "providers" not in cfg:
        cfg = {
            "current": "deepseek",
            "providers": {
                "deepseek": {
                    "label": "DeepSeek",
                    "base_url": "https://api.deepseek.com",
                    "api_key": cfg.get("DEEPSEEK_API_KEY", ""),
                    "model": cfg.get("DEEPSEEK_MODEL", "deepseek-chat"),
                    "models": ["deepseek-chat", "deepseek-reasoner"],
                }
            },
        }
    cfg["brand"] = normalize_brand(cfg.get("brand") or {})
    return cfg


def normalize_watermark(raw):
    """正文水印配置标准化 + 数值夹取。"""
    wm = dict(WATERMARK_DEFAULTS)
    wm.update(raw or {})
    if wm.get("type") not in ("text", "image"):
        wm["type"] = "text"
    wm["enabled"] = bool(wm.get("enabled"))
    wm["text"] = str(wm.get("text") or "").strip()[:60]
    def _num(key, lo, hi, default):
        try:
            v = float(wm.get(key))
        except (TypeError, ValueError):
            v = default
        return max(lo, min(hi, v))
    wm["size_pct"] = _num("size_pct", 10, 200, 60)
    wm["opacity_pct"] = _num("opacity_pct", 3, 100, 15)
    wm["angle"] = _num("angle", -180, 180, -45)
    wm["font"] = str(wm.get("font") or "msyh").strip()
    if wm["font"] not in report_builder.WM_FONT_KEYS:
        wm["font"] = "msyh"
    c = str(wm.get("color") or "").lstrip("#").lower()
    if len(c) != 6 or any(ch not in "0123456789abcdef" for ch in c):
        c = "216873"
    wm["color"] = c
    return wm


def normalize_brand(raw):
    """品牌配置标准化：填补默认值，并迁移旧版字段（show_on_cover/show_in_header）。"""
    out = {"org_name": str(raw.get("org_name") or ""),
           "contact": str(raw.get("contact") or "")}
    old_cover = raw.get("show_on_cover")
    old_header = raw.get("show_in_header")
    for slot in BRAND_SLOTS:
        s = dict(BRAND_DEFAULTS[slot])
        s.update(raw.get(slot) or {})
        if slot == "cover" and old_cover is not None and slot not in raw:
            s["enabled"] = bool(old_cover)
        if slot in ("header", "footer") and old_header is not None and slot not in raw:
            s["enabled"] = bool(old_header)
        out[slot] = s
    out["watermark"] = normalize_watermark(raw.get("watermark"))
    return out


def migrate_brand_files():
    """旧版单一 brand/logo.* → 复制为 cover.* 与 header.*（只执行一次）。"""
    old = None
    for ext in ("png", "jpg"):
        p = os.path.join(BRAND_DIR, "logo." + ext)
        if os.path.isfile(p):
            old = p
            break
    if not old:
        return
    import shutil as _sh
    ext = old.rsplit(".", 1)[-1]
    for slot in ("cover", "header"):
        dst = os.path.join(BRAND_DIR, slot + "." + ext)
        if not os.path.isfile(dst):
            _sh.copy2(old, dst)
    os.remove(old)


def save_config():
    with _config_lock:
        tmp = CONFIG_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
        os.replace(tmp, CONFIG_PATH)


CONFIG = load_config()
migrate_brand_files()


def mask_key(key):
    if not key:
        return ""
    if len(key) <= 12:
        return "****"
    return key[:6] + "…" + key[-4:]


def public_providers():
    out = []
    for name, p in CONFIG["providers"].items():
        out.append({
            "name": name,
            "label": p.get("label", name),
            "base_url": p.get("base_url", ""),
            "model": p.get("model", ""),
            "models": p.get("models", []),
            "has_key": bool(p.get("api_key")),
            "key_masked": mask_key(p.get("api_key", "")),
        })
    return {"current": CONFIG.get("current", ""), "providers": out}


def get_provider(name=None):
    name = name or CONFIG.get("current") or ""
    p = CONFIG["providers"].get(name)
    if not p:
        raise ValueError("未找到模型供应商：%s" % (name or "（未选择）"))
    if not p.get("api_key"):
        raise ValueError("「%s」尚未填写 API Key，请先在右上角 ⚙ 设置中配置" % p.get("label", name))
    if not p.get("model"):
        raise ValueError("「%s」尚未选择模型，请先在右上角 ⚙ 设置中配置" % p.get("label", name))
    return name, p


# ---------------- 机构品牌 / 水印 ----------------
def brand_slot_path(slot):
    for ext in ("png", "jpg"):
        p = os.path.join(BRAND_DIR, slot + "." + ext)
        if os.path.isfile(p):
            return p
    return None


def public_brand():
    b = normalize_brand(CONFIG.get("brand") or {})
    slots = {}
    for slot in BRAND_SLOTS:
        slots[slot] = {"enabled": b[slot]["enabled"],
                       "height_cm": b[slot]["height_cm"],
                       "has_image": brand_slot_path(slot) is not None}
    wm = b["watermark"]
    wm["has_image"] = brand_slot_path("body") is not None
    return {"org_name": b["org_name"], "contact": b["contact"],
            "slots": slots, "watermark": wm,
            "wm_fonts": report_builder.available_wm_fonts()}


def get_brand():
    """传给 report_builder 的品牌信息（全局配置，旧报告重建时用最新品牌）。"""
    b = normalize_brand(CONFIG.get("brand") or {})
    for slot in BRAND_SLOTS:
        b[slot]["path"] = brand_slot_path(slot)
    b["watermark"]["path"] = brand_slot_path("body")
    return b


def _decode_brand_image(b64):
    if "," in b64:
        b64 = b64.split(",", 1)[1]
    raw = base64.b64decode(b64)
    if raw.startswith(b"\x89PNG\r\n\x1a\n"):
        return raw, "png"
    if raw.startswith(b"\xff\xd8\xff"):
        return raw, "jpg"
    raise ValueError("图片仅支持 PNG / JPG 格式")


def _remove_slot_image(slot):
    for ext in ("png", "jpg"):
        p = os.path.join(BRAND_DIR, slot + "." + ext)
        if os.path.isfile(p):
            os.remove(p)


def brand_from_payload(payload):
    """从请求体解析品牌设置（不落盘），供预览/保存共用。
    新上传的图片解码到 brand/_tmp_{slot}.* 临时文件。"""
    b = normalize_brand(CONFIG.get("brand") or {})
    if payload.get("org_name") is not None:
        b["org_name"] = str(payload["org_name"]).strip()
    if payload.get("contact") is not None:
        b["contact"] = str(payload["contact"]).strip()
    slots_in = payload.get("slots") or {}
    for slot in BRAND_SLOTS:
        s = slots_in.get(slot) or {}
        if s.get("enabled") is not None:
            b[slot]["enabled"] = bool(s["enabled"])
        if s.get("height_cm") is not None:
            try:
                h = float(s["height_cm"])
                b[slot]["height_cm"] = max(0.3, min(8.0, h))
            except (TypeError, ValueError):
                pass
        b[slot]["path"] = brand_slot_path(slot)
    if payload.get("watermark") is not None:
        b["watermark"] = normalize_watermark(payload.get("watermark"))
    b["watermark"]["path"] = brand_slot_path("body")
    remove = payload.get("remove_images") or []
    for slot, data_url in (payload.get("images") or {}).items():
        if slot not in BRAND_SLOTS + ("body",) or not data_url:
            continue
        raw, ext = _decode_brand_image(data_url)
        tmp = os.path.join(BRAND_DIR, "_tmp_%s.%s" % (slot, ext))
        with open(tmp, "wb") as f:
            f.write(raw)
        if slot == "body":
            b["watermark"]["path"] = tmp
        else:
            b[slot]["path"] = tmp
    for slot in remove:
        if slot in BRAND_SLOTS:
            b[slot]["path"] = None
        elif slot == "body":
            b["watermark"]["path"] = None
    return b



# ---------------- 工具 ----------------
def safe_name(s):
    return re.sub(r'[\\/:*?"<>|\s（）()·—-]+', "-", s).strip("-")


def report_filename(meta):
    """报告文件名：学科-标题[-学生]-课堂报告。"""
    return "-".join(x for x in [
        safe_name((meta.get("subject") or "").split("（")[0]),
        safe_name(meta.get("title") or ""),
        safe_name(meta.get("student") or ""),
        "课堂报告"] if x)


def job_dir(job):
    d = os.path.join(RUNTIME_DIR, job)
    os.makedirs(d, exist_ok=True)
    return d


MINDMAP_CLASSDEFS = (
    "classDef rootLevel fill:#1e3a8a, stroke:#1e3a8a, stroke-width:2px, color:#ffffff;\n"
    "classDef level1 fill:#e0f2fe, stroke:#0284c7, stroke-width:2px, color:#075985;\n"
    "classDef level2 fill:#fef3c7, stroke:#d97706, stroke-width:1.5px, color:#92400e;\n"
    "classDef level3 fill:#f3f4f6, stroke:#9ca3af, stroke-width:1px, color:#374151;\n"
    "classDef level4 fill:#ffffff, stroke:#d1d5db, stroke-width:1px, color:#4b5563;\n"
    "classDef level5 fill:#fafafa, stroke:#e5e7eb, stroke-width:1px, color:#6b7280;")


def _existing_mindmaps(d):
    """job 目录里已存在的脑图产物：切片清单（mindmap_slices.json）> 单图 mindmap.png。"""
    manifest = os.path.join(d, "mindmap_slices.json")
    if os.path.isfile(manifest):
        try:
            with open(manifest, encoding="utf-8") as f:
                files = json.load(f)
            paths = [os.path.join(d, n) for n in files]
            if paths and all(os.path.isfile(p) for p in paths):
                return paths
        except Exception:
            pass
    single = os.path.join(d, "mindmap.png")
    return [single] if os.path.isfile(single) else None


def ensure_mindmap(d, data, mindmap_paths):
    """未提供脑图时，用模型返回的 mermaid 代码自动渲染脑图（按节点间隙切片，失败则留空）。"""
    if mindmap_paths:
        return mindmap_paths
    code = data.get("mindmap_mermaid")
    if not isinstance(code, str) or not code.strip():
        return None
    try:  # 缺连线的代码渲染出来是无结构单列，直接弃用
        llm_client._check_mindmap_code(code)
    except ValueError:
        return None
    # 模型可能漏写 classDef 样式定义：渲染前按名字逐个兜底补齐（否则 class 绑定无样式）
    add = [ln for ln in MINDMAP_CLASSDEFS.split("\n")
           if ("classDef " + ln.split()[1]) not in code]
    if add:
        code = code.rstrip() + "\n" + "\n".join(add)
        data["mindmap_mermaid"] = code
        try:  # 补齐后的代码写回，供下次重新排版复用
            with open(os.path.join(d, "data.json"), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass
    try:
        slices = mermaid_render.render_mermaid_slices(code, d)
        if slices:
            names = [os.path.basename(s["png"]) for s in slices]
            with open(os.path.join(d, "mindmap_slices.json"), "w", encoding="utf-8") as f:
                json.dump(names, f, ensure_ascii=False)
            return [s["png"] for s in slices]
    except Exception:
        pass
    return None


class Handler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, fmt, *args):
        pass

    # ---------- helpers ----------
    def _send_json(self, obj, status=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_error_json(self, msg, status=500):
        self._send_json({"error": msg}, status=status)

    def _read_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_BODY:
            raise ValueError("请求体过大（>40MB）")
        return self.rfile.read(length)

    def _send_file(self, path, content_type, download_name=None):
        with open(path, "rb") as f:
            body = f.read()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        if download_name:
            from urllib.parse import quote
            self.send_header(
                "Content-Disposition",
                "attachment; filename*=UTF-8''%s" % quote(download_name))
        self.end_headers()
        self.wfile.write(body)

    # ---------- routing ----------
    def do_GET(self):
        route = self.path.split("?", 1)[0].split("#", 1)[0]
        if route == "/" or route == "/index.html":
            return self._send_file(os.path.join(STATIC_DIR, "index.html"),
                                   "text/html; charset=utf-8")
        if route == "/api/providers":
            return self._send_json(public_providers())
        if route == "/api/brand":
            return self._send_json(public_brand())
        if route == "/api/brand/image":
            return self.handle_brand_image()
        if route == "/api/history":
            return self.handle_history()
        if route == "/api/students":
            return self.handle_students()
        if route == "/api/student":
            return self.handle_student()
        if route == "/api/student_pdf":
            return self.handle_student_pdf()
        if route == "/api/data":
            return self.handle_data()
        if route == "/api/file":
            return self.handle_file()
        if route.startswith("/vendor/"):
            return self.handle_vendor()
        self._send_error_json("Not Found", 404)

    def handle_vendor(self):
        """本地前端依赖（jszip / docx-preview），避免 CDN 抖动。"""
        name = self.path[len("/vendor/"):].split("?")[0]
        if not re.fullmatch(r"[0-9a-zA-Z._-]+", name):
            return self._send_error_json("非法路径", 400)
        path = os.path.join(STATIC_DIR, "vendor", name)
        if not os.path.isfile(path):
            return self._send_error_json("Not Found", 404)
        mime = "application/javascript" if name.endswith(".js") else "application/octet-stream"
        return self._send_file(path, mime)

    # ---------- 历史报告 ----------
    def _job_id_from_query(self):
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        job = (qs.get("job") or [""])[0]
        if not re.fullmatch(r"[0-9a-zA-Z-]+", job):
            return None
        return job

    def handle_history(self):
        """扫描 runtime/ 下所有已生成的报告，按时间倒序返回。"""
        items = []
        for name in os.listdir(RUNTIME_DIR):
            d = os.path.join(RUNTIME_DIR, name)
            if not os.path.isdir(d):
                continue
            if not os.path.isfile(os.path.join(d, "report.docx")):
                continue
            filename = "课堂报告"
            fn_file = os.path.join(d, "filename.txt")
            if os.path.isfile(fn_file):
                filename = open(fn_file, encoding="utf-8").read().strip() or filename
            meta = {}
            data_file = os.path.join(d, "data.json")
            if os.path.isfile(data_file):
                try:
                    meta = json.load(open(data_file, encoding="utf-8")).get("meta", {})
                except Exception:
                    pass
            # job 名形如 20260822-021239-xxxxxx
            created = ""
            m = re.match(r"(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})", name)
            if m:
                created = "%s-%s-%s %s:%s" % (m.group(1), m.group(2), m.group(3),
                                              m.group(4), m.group(5))
            items.append({
                "job": name,
                "filename": filename,
                "title": meta.get("title", ""),
                "subject": meta.get("subject", ""),
                "student": meta.get("student", ""),
                "date": meta.get("date", ""),
                "created": created,
                "has_mindmap": _existing_mindmaps(d) is not None,
                "has_pdf": os.path.isfile(os.path.join(d, "report.pdf")),
                "size": os.path.getsize(os.path.join(d, "report.docx")),
            })
        items.sort(key=lambda x: x["job"], reverse=True)
        self._send_json({"items": items})

    def handle_data(self):
        """读取某次报告的结构化 JSON（用于载入编辑器）。"""
        job = self._job_id_from_query()
        if not job:
            return self._send_error_json("非法 job", 400)
        path = os.path.join(RUNTIME_DIR, job, "data.json")
        if not os.path.isfile(path):
            return self._send_error_json("数据不存在", 404)
        with open(path, encoding="utf-8") as f:
            data = json.load(f)
        fn_file = os.path.join(RUNTIME_DIR, job, "filename.txt")
        filename = open(fn_file, encoding="utf-8").read().strip() if os.path.isfile(fn_file) else "课堂报告"
        self._send_json({"job": job, "filename": filename, "data": data})

    def handle_delete(self):
        """删除一次报告（整个 job 目录）。body: {job}"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        job = payload.get("job") or ""
        if not re.fullmatch(r"[0-9a-zA-Z-]+", job):
            return self._send_error_json("非法 job", 400)
        d = os.path.join(RUNTIME_DIR, job)
        if not os.path.isdir(d):
            return self._send_error_json("任务不存在", 404)
        import shutil as _shutil
        try:
            _shutil.rmtree(d)
        except Exception as e:
            return self._send_error_json("删除失败: %s" % e, 500)
        self._send_json({"ok": True})

    # ---------- 学生档案 / 学情趋势 ----------
    def _iter_report_data(self):
        """扫描 runtime/ 下所有报告，yield (job, data)。跳过 _ 开头的内部目录。"""
        for name in sorted(os.listdir(RUNTIME_DIR)):
            if name.startswith("_"):
                continue
            d = os.path.join(RUNTIME_DIR, name)
            data_file = os.path.join(d, "data.json")
            if not os.path.isdir(d) or not os.path.isfile(data_file):
                continue
            try:
                with open(data_file, encoding="utf-8") as f:
                    yield name, json.load(f)
            except Exception:
                continue

    def _student_reports(self, student):
        """某学生的全部报告（按时间正序）：[{job, data}]。"""
        out = []
        for job, data in self._iter_report_data():
            if ((data.get("meta") or {}).get("student") or "").strip() == student:
                out.append({"job": job, "data": data})
        return out

    def _student_name_from_query(self):
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        name = (qs.get("name") or [""])[0].strip()
        if not name or len(name) > 50:
            return None
        return name

    def handle_students(self):
        """GET /api/students → 按学生聚合：姓名 / 报告数 / 最近上课日期。"""
        agg = {}
        for job, data in self._iter_report_data():
            st = ((data.get("meta") or {}).get("student") or "").strip()
            if not st:
                continue
            a = agg.setdefault(st, {"name": st, "count": 0, "last_job": ""})
            a["count"] += 1
            a["last_job"] = max(a["last_job"], job)
        items = sorted(agg.values(), key=lambda x: x["last_job"], reverse=True)
        for it in items:
            m = re.match(r"(\d{4})(\d{2})(\d{2})", it.pop("last_job"))
            it["last_date"] = "%s-%s-%s" % m.groups() if m else ""
        self._send_json({"items": items})

    def handle_student(self):
        """GET /api/student?name=xx → 该学生全部报告（时间正序，含完整 data）。"""
        name = self._student_name_from_query()
        if not name:
            return self._send_error_json("非法学生姓名", 400)
        reports = []
        for item in self._student_reports(name):
            fn_file = os.path.join(RUNTIME_DIR, item["job"], "filename.txt")
            fn = open(fn_file, encoding="utf-8").read().strip() \
                if os.path.isfile(fn_file) else ""
            reports.append({"job": item["job"], "filename": fn,
                            "data": item["data"]})
        self._send_json({"name": name, "reports": reports})

    def handle_student_summary(self):
        """POST /api/student_summary {name, provider?, model?} → AI 跨报告学情总结。

        结果缓存到 runtime/_students/<姓名>.json，供学情 PDF 附带。"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        name = (payload.get("name") or "").strip()
        if not name or len(name) > 50:
            return self._send_error_json("非法学生姓名", 400)
        reports = self._student_reports(name)
        if not reports:
            return self._send_error_json("该学生还没有关联的报告", 404)
        # 可选课程范围：jobs 为选中的 job 列表（前端范围选择器给出），空 = 全部
        jobs = payload.get("jobs") or []
        if jobs:
            wanted = set(str(j) for j in jobs)
            reports = [r for r in reports if r["job"] in wanted]
            if not reports:
                return self._send_error_json("所选课程范围内没有报告", 400)
        try:
            pname, provider = get_provider(payload.get("provider"))
        except ValueError as e:
            return self._send_error_json(str(e), 400)
        model = (payload.get("model") or "").strip() or provider["model"]

        # 学情分析以「课堂表现」为依据：digest 不含 mistakes（误区不作为评价依据）
        digest = []
        for item in reports:
            data = item["data"]
            meta = data.get("meta") or {}
            digest.append({
                "date": meta.get("date", ""),
                "subject": meta.get("subject", ""),
                "title": meta.get("title", ""),
                "mainline": data.get("mainline", ""),
                "performance": data.get("performance") or {},
                "homework": data.get("homework") or [],
                "gains": data.get("gains") or [],
            })
        try:
            summary = llm_client.generate_student_summary(
                provider["base_url"], provider["api_key"], model, name, digest,
                extra_body=provider.get("extra_body"))
        except Exception as e:
            return self._send_error_json("学情总结生成失败（%s · %s）: %s" % (
                provider.get("label", pname), model, e), 502)

        os.makedirs(STUDENTS_DIR, exist_ok=True)
        cache = os.path.join(STUDENTS_DIR, safe_name(name) + ".json")
        with open(cache, "w", encoding="utf-8") as f:
            json.dump({"name": name, "summary": summary,
                       "jobs": [r["job"] for r in reports],
                       "updated": time.strftime("%Y-%m-%d %H:%M")},
                      f, ensure_ascii=False, indent=2)
        self._send_json({"name": name, "summary": summary})

    def handle_student_pdf(self):
        """GET /api/student_pdf?name=xx → 学情趋势 PDF（含已缓存的 AI 总结，如有）。

        每次都重新生成 docx 并转换，保证与最新报告数据一致。"""
        name = self._student_name_from_query()
        if not name:
            return self._send_error_json("非法学生姓名", 400)
        reports = self._student_reports(name)
        if not reports:
            return self._send_error_json("该学生还没有关联的报告", 404)
        # 可选课程范围：?jobs=a,b,c（与前端范围选择一致），空 = 全部
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        jobs = [j for j in (qs.get("jobs") or [""])[0].split(",") if j]
        if jobs:
            wanted = set(jobs)
            reports = [r for r in reports if r["job"] in wanted]
            if not reports:
                return self._send_error_json("所选课程范围内没有报告", 400)

        # 仅当缓存的 AI 总结与本次课程范围一致时才附进 PDF
        summary = ""
        cache = os.path.join(STUDENTS_DIR, safe_name(name) + ".json")
        if os.path.isfile(cache):
            try:
                with open(cache, encoding="utf-8") as f:
                    cached = json.load(f)
                if cached.get("jobs", []) == [r["job"] for r in reports]:
                    summary = cached.get("summary", "")
            except Exception:
                pass

        os.makedirs(STUDENTS_DIR, exist_ok=True)
        docx_path = os.path.join(STUDENTS_DIR,
                                 safe_name(name) + "-学情趋势.docx")
        pdf_path = docx_path[:-5] + ".pdf"
        try:
            report_builder.build_student_report(
                name, [item["data"] for item in reports], summary, docx_path,
                brand=get_brand())
        except Exception as e:
            return self._send_error_json("学情报告生成失败: %s" % e, 500)
        try:
            pdf_convert.convert(docx_path, pdf_path)
        except Exception as e:
            return self._send_error_json("PDF 转换失败: %s" % e, 500)
        return self._send_file(pdf_path, "application/pdf",
                               safe_name(name) + "-学情趋势.pdf")

    def do_POST(self):
        route = self.path.split("?", 1)[0]
        if route == "/api/generate":
            return self.handle_generate()
        if route == "/api/delete":
            return self.handle_delete()
        if route == "/api/update":
            return self.handle_update()
        if route == "/api/config":
            return self.handle_config()
        if route == "/api/brand":
            return self.handle_brand_save()
        if route == "/api/brand/preview":
            return self.handle_brand_preview()
        if route == "/api/batch_zip":
            return self.handle_batch_zip()
        if route == "/api/student_summary":
            return self.handle_student_summary()
        if route == "/api/parent-feedback":
            return self.handle_parent_feedback()
        self._send_error_json("Not Found", 404)

    # ---------- 批量打包 ----------
    def handle_batch_zip(self):
        """把多个 job 的报告打包成 zip。body: {jobs: [...], include_pdf?: bool}"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        jobs = payload.get("jobs") or []
        if not jobs or len(jobs) > 50:
            return self._send_error_json("jobs 数量非法（1–50）", 400)
        for job in jobs:
            if not re.fullmatch(r"[0-9a-zA-Z-]+", str(job)):
                return self._send_error_json("非法 job: %s" % job, 400)
        include_pdf = bool(payload.get("include_pdf"))
        # 前端批量弹窗里用户可能改过家长反馈：优先用随请求传来的文本
        feedbacks = payload.get("feedbacks")
        if not isinstance(feedbacks, dict):
            feedbacks = {}

        import zipfile
        zdir = os.path.join(RUNTIME_DIR, "_zips")
        os.makedirs(zdir, exist_ok=True)
        zpath = os.path.join(zdir, "batch_%s.zip" % time.strftime("%Y%m%d-%H%M%S"))
        used = {}
        packed = 0
        try:
            with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
                for job in jobs:
                    d = os.path.join(RUNTIME_DIR, job)
                    docx_path = os.path.join(d, "report.docx")
                    if not os.path.isfile(docx_path):
                        continue
                    fn_file = os.path.join(d, "filename.txt")
                    base = "课堂报告"
                    if os.path.isfile(fn_file):
                        base = open(fn_file, encoding="utf-8").read().strip() or base
                    # 重名处理：xxx.docx → xxx-2.docx
                    n = used.get(base, 0) + 1
                    used[base] = n
                    if n > 1:
                        base = "%s-%d" % (base, n)
                    z.write(docx_path, base + ".docx")
                    packed += 1
                    # 家长反馈（如有）随包附 txt，方便逐个发微信；优先用前端改过的版本
                    try:
                        pf = feedbacks.get(job)
                        if not isinstance(pf, str):
                            pf = None
                        if pf is None:
                            data_file = os.path.join(d, "data.json")
                            if os.path.isfile(data_file):
                                with open(data_file, encoding="utf-8") as f:
                                    pf = json.load(f).get("parent_feedback")
                        pf = (pf or "").strip()
                        if pf:
                            z.writestr(base + "-家长反馈.txt", pf)
                    except Exception:
                        pass  # 单个反馈附件失败不阻塞打包
                    if include_pdf:
                        pdf_path = os.path.join(d, "report.pdf")
                        if not os.path.isfile(pdf_path):
                            try:
                                pdf_convert.convert(docx_path, pdf_path)
                            except Exception:
                                continue  # 单个 PDF 失败不阻塞打包
                        z.write(pdf_path, base + ".pdf")
        except Exception as e:
            return self._send_error_json("打包失败: %s" % e, 500)
        if not packed:
            return self._send_error_json("没有可打包的报告", 404)
        dl = "批量课堂报告-%s.zip" % time.strftime("%Y%m%d-%H%M")
        return self._send_file(zpath, "application/zip", dl)

    # ---------- 机构品牌 / 水印 ----------
    def handle_brand_image(self):
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        slot = (qs.get("slot") or [""])[0]
        if slot not in BRAND_SLOTS + ("body",):
            return self._send_error_json("非法槽位", 400)
        path = brand_slot_path(slot)
        if not path:
            return self._send_error_json("该位置尚未上传图片", 404)
        mime = "image/png" if path.endswith(".png") else "image/jpeg"
        return self._send_file(path, mime)

    def handle_brand_save(self):
        """保存品牌设置。body: {org_name?, contact?, slots: {cover:{enabled,height_cm},...},
        watermark: {enabled,type,text,size_pct,opacity_pct,angle},
        images: {cover: base64,..., body: base64}, remove_images: ["footer","body",...]}"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        try:
            b = brand_from_payload(payload)
        except Exception as e:
            return self._send_error_json(str(e), 400)

        # 落盘：新图片从 _tmp 挪到正式槽位文件
        for slot in BRAND_SLOTS:
            if slot in (payload.get("remove_images") or []):
                _remove_slot_image(slot)
            path = b[slot].get("path") or ""
            if os.path.basename(path).startswith("_tmp_"):
                ext = path.rsplit(".", 1)[-1]
                _remove_slot_image(slot)
                os.replace(path, os.path.join(BRAND_DIR, "%s.%s" % (slot, ext)))
            b[slot].pop("path", None)
        # 正文水印图片（body 槽位）
        if "body" in (payload.get("remove_images") or []):
            _remove_slot_image("body")
        wm_path = b["watermark"].get("path") or ""
        if os.path.basename(wm_path).startswith("_tmp_"):
            ext = wm_path.rsplit(".", 1)[-1]
            _remove_slot_image("body")
            os.replace(wm_path, os.path.join(BRAND_DIR, "body.%s" % ext))
        b["watermark"].pop("path", None)
        CONFIG["brand"] = b
        try:
            save_config()
        except Exception as e:
            return self._send_error_json("配置保存失败: %s" % e, 500)
        self._send_json(public_brand())

    PREVIEW_DATA = {
        "meta": {"title": "示例课程", "subtitle": "—— 水印预览", "subject": "示例学科",
                 "form": "一对一辅导课", "content": "封面 / 页眉 / 页脚效果预览", "date": "2026 年 8 月 24 日"},
        "mainline": "本页为水印预览样例，展示当前品牌设置在真实 Word 排版中的效果。",
        "content": [{"lead": "预览说明：", "text": "封面顶部显示封面图片与机构名；本页（第二页起）显示页眉与页脚，页码从本页起算第 1 页。"}],
        "performance": {"pros": [{"lead": "页眉：", "text": "机构名与图片靠右排列，下方细分隔线。"}],
                        "cons": [{"lead": "页脚：", "text": "机构名、联系方式、页码与图片靠右排列。"}]},
        "minutes": [{"title": "调整方式", "points": ["在左侧修改图片、高度与开关后，再次点击「刷新预览」即可。"]}],
        "gains": [{"lead": "提示：", "text": "预览不会保存配置，确认效果后请点击「保存设置」。"}],
        "homework": [{"lead": "无", "text": "样例数据"}],
        "mistakes": [{"lead": "无", "text": "样例数据"}],
        "suggestions": [{"lead": "无", "text": "样例数据"}],
    }

    def handle_brand_preview(self):
        """按请求体中的（未保存）设置生成预览 docx，直接返回文件。"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        try:
            b = brand_from_payload(payload)
        except Exception as e:
            return self._send_error_json(str(e), 400)
        tmp_paths = [b[s]["path"] for s in BRAND_SLOTS
                     if b[s].get("path") and os.path.basename(b[s]["path"]).startswith("_tmp_")]
        wm_path = b["watermark"].get("path") or ""
        if os.path.basename(wm_path).startswith("_tmp_"):
            tmp_paths.append(wm_path)
        d = os.path.join(RUNTIME_DIR, "_brand_preview")
        os.makedirs(d, exist_ok=True)
        docx_path = os.path.join(d, "preview.docx")
        try:
            report_builder.build_report(self.PREVIEW_DATA, None, docx_path, brand=b)
        except Exception as e:
            import traceback; traceback.print_exc()
            return self._send_error_json("预览生成失败: %s" % e, 500)
        finally:
            for p in tmp_paths:
                if os.path.isfile(p):
                    os.remove(p)
        return self._send_file(
            docx_path,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document")

    # ---------- API ----------
    def handle_update(self):
        """编辑后的报告 JSON → 重建 docx（并作废旧 pdf）。body: {job, data}"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        job = payload.get("job") or ""
        if not re.fullmatch(r"[0-9a-zA-Z-]+", job):
            return self._send_error_json("非法 job", 400)
        d = os.path.join(RUNTIME_DIR, job)
        if not os.path.isdir(d):
            return self._send_error_json("任务不存在", 404)
        data = payload.get("data")
        if not isinstance(data, dict):
            return self._send_error_json("缺少 data", 400)
        try:
            data = llm_client.validate(data)
        except Exception as e:
            return self._send_error_json("报告数据校验失败: %s" % e, 400)

        with open(os.path.join(d, "data.json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        mindmap_paths = ensure_mindmap(d, data, _existing_mindmaps(d))
        docx_path = os.path.join(d, "report.docx")
        try:
            report_builder.build_report(data, mindmap_paths, docx_path,
                                        brand=get_brand())
        except Exception as e:
            return self._send_error_json("Word 生成失败: %s" % e, 500)
        pdf_path = os.path.join(d, "report.pdf")
        if os.path.isfile(pdf_path):
            os.remove(pdf_path)  # 内容已变，旧 PDF 作废

        m = data["meta"]
        base = report_filename(m)
        with open(os.path.join(d, "filename.txt"), "w", encoding="utf-8") as f:
            f.write(base)
        self._send_json({"ok": True, "filename": base})
    def handle_config(self):
        """保存供应商配置。body: {current?, updates: {name: {base_url?, api_key?, model?, models?}}}"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)

        updates = payload.get("updates") or {}
        for name, u in updates.items():
            p = CONFIG["providers"].setdefault(name, {"label": name, "models": []})
            if u.get("label") is not None:
                p["label"] = u["label"]
            if u.get("base_url") is not None:
                p["base_url"] = u["base_url"].strip()
            if u.get("model") is not None:
                p["model"] = u["model"].strip()
            if u.get("models") is not None:
                models = u["models"]
                if isinstance(models, str):
                    models = [m.strip() for m in models.split(",") if m.strip()]
                p["models"] = models
            key = (u.get("api_key") or "").strip()
            # 空值或掩码值（含 … 或 ****）表示不修改
            if key and "…" not in key and key != "****":
                p["api_key"] = key

        current = (payload.get("current") or "").strip()
        if current and current in CONFIG["providers"]:
            CONFIG["current"] = current
        try:
            save_config()
        except Exception as e:
            return self._send_error_json("配置保存失败: %s" % e, 500)
        self._send_json(public_providers())

    def handle_generate(self):
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)

        overview = (payload.get("overview") or "").strip()
        notes = (payload.get("notes") or "").strip()
        if not overview or not notes:
            return self._send_error_json("请填写「课堂速览」和「课堂纪要」", 400)
        student = (payload.get("student") or "").strip()
        if not student:
            return self._send_error_json("请填写「学生姓名」（必填，用于学生档案归档）", 400)

        try:
            name, provider = get_provider(payload.get("provider"))
        except ValueError as e:
            return self._send_error_json(str(e), 400)
        model = (payload.get("model") or "").strip() or provider["model"]

        job = time.strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:6]
        d = job_dir(job)

        # 脑图（base64，可为空）
        mindmap_path = None
        mindmap_b64 = payload.get("mindmap") or ""
        if mindmap_b64:
            try:
                if "," in mindmap_b64:
                    mindmap_b64 = mindmap_b64.split(",", 1)[1]
                raw = base64.b64decode(mindmap_b64)
                mindmap_path = os.path.join(d, "mindmap.png")
                with open(mindmap_path, "wb") as f:
                    f.write(raw)
            except Exception as e:
                return self._send_error_json("脑图图片解析失败: %s" % e, 400)

        meta_hints = payload.get("meta") or {}
        try:
            data = llm_client.generate_report_json(
                provider["base_url"], provider["api_key"], model,
                overview, notes, meta_hints=meta_hints,
                extra_body=provider.get("extra_body"))
            data = llm_client.validate(data)
            student = (payload.get("student") or "").strip()
            if student:
                data["meta"]["student"] = student
            # 第二次独立调用：未提供脑图时，基于归纳好的报告 JSON 生成完整知识点脑图
            if mindmap_path is None:
                try:
                    data["mindmap_mermaid"] = llm_client.generate_mindmap_mermaid(
                        provider["base_url"], provider["api_key"], model, data,
                        extra_body=provider.get("extra_body"))
                except Exception:
                    pass  # 脑图失败不影响主报告
            # 第三次独立调用：生成给家长的简易文字反馈（失败不影响主报告）
            try:
                data["parent_feedback"] = llm_client.generate_parent_feedback(
                    provider["base_url"], provider["api_key"], model, data,
                    extra_body=provider.get("extra_body"))
            except Exception:
                pass
        except Exception as e:
            return self._send_error_json("模型生成失败（%s · %s）: %s" % (
                provider.get("label", name), model, e), 502)

        with open(os.path.join(d, "data.json"), "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # 未提供脑图时，用 AI 返回的 mermaid 代码自动渲染脑图（切片列表）
        mindmap_paths = ensure_mindmap(d, data, [mindmap_path] if mindmap_path else None)

        docx_path = os.path.join(d, "report.docx")
        try:
            report_builder.build_report(data, mindmap_paths, docx_path,
                                        brand=get_brand())
        except Exception as e:
            return self._send_error_json("Word 生成失败: %s" % e, 500)

        m = data["meta"]
        base = report_filename(m)
        with open(os.path.join(d, "filename.txt"), "w", encoding="utf-8") as f:
            f.write(base)

        self._send_json({"job": job, "filename": base, "data": data,
                         "provider": provider.get("label", name), "model": model})

    def handle_parent_feedback(self):
        """为已有报告（重新）生成家长简易文字反馈。body: {job, provider?, model?}"""
        try:
            payload = json.loads(self._read_body().decode("utf-8"))
        except Exception as e:
            return self._send_error_json("请求解析失败: %s" % e, 400)
        job = payload.get("job") or ""
        if not re.fullmatch(r"[0-9a-zA-Z-]+", job):
            return self._send_error_json("非法 job", 400)
        d = os.path.join(RUNTIME_DIR, job)
        data_path = os.path.join(d, "data.json")
        if not os.path.isfile(data_path):
            return self._send_error_json("任务不存在", 404)
        with open(data_path, encoding="utf-8") as f:
            data = json.load(f)
        try:
            name, provider = get_provider(payload.get("provider"))
        except ValueError as e:
            return self._send_error_json(str(e), 400)
        model = (payload.get("model") or "").strip() or provider["model"]
        try:
            text = llm_client.generate_parent_feedback(
                provider["base_url"], provider["api_key"], model, data,
                extra_body=provider.get("extra_body"))
        except Exception as e:
            return self._send_error_json("家长反馈生成失败（%s · %s）: %s" % (
                provider.get("label", name), model, e), 502)
        data["parent_feedback"] = text
        with open(data_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        self._send_json({"job": job, "parent_feedback": text})

    def handle_file(self):
        from urllib.parse import urlparse, parse_qs
        qs = parse_qs(urlparse(self.path).query)
        job = (qs.get("job") or [""])[0]
        fmt = (qs.get("fmt") or ["docx"])[0]
        if not re.fullmatch(r"[0-9a-zA-Z-]+", job):
            return self._send_error_json("非法 job", 400)
        d = os.path.join(RUNTIME_DIR, job)
        if not os.path.isdir(d):
            return self._send_error_json("任务不存在", 404)

        name_file = os.path.join(d, "filename.txt")
        base = "课堂报告"
        if os.path.isfile(name_file):
            base = open(name_file, encoding="utf-8").read().strip() or base

        if fmt == "docx":
            path = os.path.join(d, "report.docx")
            if not os.path.isfile(path):
                return self._send_error_json("docx 尚未生成", 404)
            dl = base + ".docx" if "download" in qs else None
            return self._send_file(
                path,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                dl)
        if fmt == "pdf":
            docx_path = os.path.join(d, "report.docx")
            pdf_path = os.path.join(d, "report.pdf")
            if not os.path.isfile(pdf_path):
                if not os.path.isfile(docx_path):
                    return self._send_error_json("docx 尚未生成", 404)
                try:
                    pdf_convert.convert(docx_path, pdf_path)
                except Exception as e:
                    return self._send_error_json("PDF 转换失败: %s" % e, 500)
            dl = base + ".pdf" if "download" in qs else None
            return self._send_file(pdf_path, "application/pdf", dl)
        self._send_error_json("未知 fmt", 400)


class Server(ThreadingHTTPServer):
    # Windows 下 SO_REUSEADDR 会允许重复绑定同一端口（产生双实例混乱），这里禁止
    allow_reuse_address = False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--host", default="0.0.0.0")
    ap.add_argument("--port", type=int, default=7100)
    args, _ = ap.parse_known_args()
    import sys
    if sys.stdout is None:  # pythonw 无窗口模式：丢弃 print 输出
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
        sys.stderr = sys.stdout
    try:
        srv = Server((args.host, args.port), Handler)
    except OSError:
        # 端口已被占用 = 服务已在运行，静默退出
        sys.exit(0)
    print("课堂反馈生成器: http://localhost:%d/" % args.port)
    srv.serve_forever()


if __name__ == "__main__":
    main()
