# -*- coding: utf-8 -*-
"""水印功能验证：三槽位独立图片/高度 → 预览接口 → 保存 → 重建 docx 断言 → 清理"""
import base64
import io
import json
import os
import shutil
import subprocess
import sys
import time
import urllib.request
import zipfile

ROOT = os.path.dirname(os.path.abspath(__file__))
BASE = "http://127.0.0.1:7100"
JOB = "brandtest-000000-abcdef"

DATA = {
    "meta": {"title": "品牌测试", "subtitle": "—— 副标题", "subject": "初中物理",
             "form": "一对一", "content": "测试内容", "date": "2026 年 8 月 24 日"},
    "mainline": "主线", "content": [{"lead": "点", "text": "内容"}],
    "performance": {"pros": [{"lead": "优", "text": "好"}], "cons": [{"lead": "改", "text": "进"}]},
    "minutes": [{"title": "纪要", "points": ["要点"]}],
    "gains": [{"lead": "获", "text": "得"}], "homework": [{"lead": "作", "text": "业"}],
    "mistakes": [{"lead": "错", "text": "析"}], "suggestions": [{"lead": "建", "text": "议"}],
}


def post(route, body, timeout=60, raw=False):
    req = urllib.request.Request(BASE + route, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read()
        return data if raw else json.loads(data.decode("utf-8"))


def make_png(color):
    from PIL import Image as PILImage
    buf = io.BytesIO()
    PILImage.new("RGBA", (240, 120), color).save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def main():
    imgs = {"cover": make_png((33, 104, 115, 255)),
            "header": make_png((230, 90, 90, 255)),
            "footer": make_png((60, 160, 90, 255))}

    job_dir = os.path.join(ROOT, "runtime", JOB)
    os.makedirs(job_dir, exist_ok=True)
    with open(os.path.join(job_dir, "data.json"), "w", encoding="utf-8") as f:
        json.dump(DATA, f, ensure_ascii=False)

    saved_cfg = open(os.path.join(ROOT, "config.json"), encoding="utf-8").read()
    brand_dir = os.path.join(ROOT, "brand")
    brand_bak = os.path.join(ROOT, "runtime", "_brand_backup")
    shutil.rmtree(brand_bak, ignore_errors=True)
    if os.path.isdir(brand_dir):
        shutil.copytree(brand_dir, brand_bak)

    proc = subprocess.Popen([sys.executable, os.path.join(ROOT, "server.py"),
                             "--host", "127.0.0.1", "--port", "7100"],
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    try:
        for _ in range(100):
            try:
                urllib.request.urlopen(BASE + "/", timeout=2)
                break
            except Exception:
                time.sleep(0.1)
        else:
            raise RuntimeError("服务器未启动")

        payload = {"org_name": "启航教育", "contact": "400-000-0000",
                   "slots": {"cover": {"enabled": True, "height_cm": 1.5},
                             "header": {"enabled": True, "height_cm": 0.9},
                             "footer": {"enabled": True, "height_cm": 0.4}},
                   "images": imgs}

        # 1. 预览接口（不保存配置，不留 _tmp 文件）
        docx_bytes = post("/api/brand/preview", payload, raw=True)
        assert docx_bytes[:2] == b"PK", "预览 docx 无效"
        leftovers = [f for f in os.listdir(os.path.join(ROOT, "brand")) if f.startswith("_tmp_")]
        assert not leftovers, "预览残留临时文件: %s" % leftovers
        print("[1/6] 预览接口 OK（%d 字节，无临时文件残留）" % len(docx_bytes))

        # 2. 保存品牌
        d = post("/api/brand", payload)
        assert d["org_name"] == "启航教育" and d["slots"]["footer"]["has_image"], d
        assert abs(d["slots"]["cover"]["height_cm"] - 1.5) < 1e-6, d
        print("[2/6] 品牌保存 OK:", json.dumps(d["slots"], ensure_ascii=False))

        # 3. 按槽取图
        with urllib.request.urlopen(BASE + "/api/brand/image?slot=footer", timeout=5) as resp:
            raw = resp.read()
            assert resp.headers["Content-Type"] == "image/png" and raw[:4] == b"\x89PNG"
        print("[3/6] /api/brand/image?slot=footer OK（%d 字节）" % len(raw))

        # 4. 重建 docx 并断言
        d = post("/api/update", {"job": JOB, "data": DATA})
        assert d.get("ok"), d
        z = zipfile.ZipFile(os.path.join(job_dir, "report.docx"))
        names = z.namelist()
        media = [n for n in names if n.startswith("word/media/")]
        assert len(media) == 3, "应嵌入 3 张不同图片（封面/页眉/页脚），实际 %s" % media
        doc_xml = z.read("word/document.xml").decode("utf-8")
        hdr_xml = "".join(z.read(n).decode("utf-8") for n in names if n.startswith("word/header"))
        ftr_xml = "".join(z.read(n).decode("utf-8") for n in names if n.startswith("word/footer"))
        assert "启航教育" in doc_xml and "blip" in doc_xml, "封面缺机构名或图片"
        assert "blip" in hdr_xml and "启航教育" in hdr_xml, "页眉缺图片或机构名"
        assert "blip" in ftr_xml and "PAGE" in ftr_xml and "400-000-0000" in ftr_xml, "页脚缺图片/页码/联系方式"
        print("[4/6] docx 断言 OK：封面/页眉/页脚各有独立图片，页脚含页码+联系方式")

        # 5a. 页码从正文起算：正文节 sectPr 含 pgNumType start=1，且无 titlePg
        import re as _re
        sects = _re.findall(r"<w:sectPr[^>]*>.*?</w:sectPr>", doc_xml, _re.S)
        assert len(sects) == 2, "应为封面+正文两节，实际 %d" % len(sects)
        assert 'w:pgNumType' in sects[1] and 'w:start="1"' in sects[1], "正文节页码未从 1 开始"
        assert "titlePg" not in doc_xml, "不应再有 titlePg（改用两节方案）"
        # 5b. 页眉页脚内容右对齐
        assert 'w:val="right"' in hdr_xml, "页眉未右对齐"
        assert 'w:val="right"' in ftr_xml, "页脚未右对齐"
        # 5c. 高度设置生效（页眉图片高度 0.9cm ≈ 324000 EMU）
        m = _re.findall(r'extent cx="\d+" cy="(\d+)"', hdr_xml)
        assert m and abs(int(m[0]) - 0.9 * 360000) < 20000, "页眉图片高度不符: %s" % m
        print("[5/6] 页码从正文起算第 1 页 ✓ 页眉/页脚右对齐 ✓ 图片高度生效（cy=%s EMU）" % m[0])

        # 6. 清理：删除全部槽位图片 + 清空文字
        d = post("/api/brand", {"org_name": "", "contact": "",
                                "remove_images": ["cover", "header", "footer"]})
        assert not any(d["slots"][s]["has_image"] for s in d["slots"]), d
        print("[6/6] 清理 OK（品牌已重置为空）")
        print("BRAND TEST PASS")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        shutil.rmtree(job_dir, ignore_errors=True)
        shutil.rmtree(os.path.join(ROOT, "runtime", "_brand_preview"), ignore_errors=True)
        # 恢复测试前的 config.json 与 brand/ 图片（测试会清空它们）
        with open(os.path.join(ROOT, "config.json"), "w", encoding="utf-8") as f:
            f.write(saved_cfg)
        if os.path.isdir(brand_bak):
            shutil.rmtree(brand_dir, ignore_errors=True)
            shutil.copytree(brand_bak, brand_dir)
            shutil.rmtree(brand_bak, ignore_errors=True)


if __name__ == "__main__":
    main()
