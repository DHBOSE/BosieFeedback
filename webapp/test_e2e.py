# -*- coding: utf-8 -*-
"""端到端测试：启动服务器 → 调 /api/generate → 校验 docx → 请求 PDF → 关闭服务器"""
import json
import subprocess
import sys
import time
import urllib.request
import urllib.error
import os

ROOT = os.path.dirname(os.path.abspath(__file__))
BASE = "http://127.0.0.1:7100"

OVERVIEW = """00:00-05:00 课程导入：物理学科初探，什么是物理，为什么要学物理
05:00-15:00 声音的产生：音叉实验，物体振动发声，振动停止发声停止
15:00-25:00 声音的传播：需要介质，真空不能传声，固体液体气体都能传声
25:00-35:00 声速：15℃空气中340m/s，固体中最快，气体中最慢
35:00-45:00 回声：声音遇到障碍物反射，回声测距 s=vt/2
45:00-50:00 声音的能量与应用：超声波、次声波
50:00-60:00 例题演练与作业布置"""

NOTES = """今天我们开始学物理第一章声现象。老师先用音叉做实验，把正在发声的音叉靠近乒乓球，乒乓球被弹开，说明发声的物体在振动。学生回答很积极，主动举手说吉他发声是琴弦振动。
然后讲声音的传播，老师问月球上能不能直接对话，学生一开始说能，后来明白真空不能传声，需要无线电。声音传播需要介质，固液气都可以，真空中不能传声。
声速部分：15℃时空气中声速是340m/s，要记住这个数。一般固体中最快、液体次之、气体最慢。学生在回声计算上卡了一下：小明对着山崖喊一声，2秒后听到回声，求距离。s=vt/2=340×2/2=340m，学生一开始忘了除以2，后来改过来了。答题时要写已知、求、解、答，单位不能丢。
声音能传递能量也能传递信息，超声波碎石、B超，次声波监测地震。
作业：完成《物理同步练习》第一章第1节第1-12题，明天交。注意答题规范，计算题要写公式、代数据、带单位。
最后聊了会儿最近流行的AI工具，这部分和课程无关。"""


def wait_server(timeout=15):
    for _ in range(timeout * 10):
        try:
            urllib.request.urlopen(BASE + "/", timeout=2)
            return True
        except Exception:
            time.sleep(0.1)
    return False


def main():
    proc = subprocess.Popen(
        [sys.executable, os.path.join(ROOT, "server.py"), "--host", "127.0.0.1", "--port", "7100"],
        stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    try:
        assert wait_server(), "服务器未启动"
        print("[1/6] 服务器已启动")

        # 供应商列表
        with urllib.request.urlopen(BASE + "/api/providers", timeout=5) as resp:
            prov = json.loads(resp.read().decode("utf-8"))
        names = [p["name"] for p in prov["providers"]]
        assert names == ["deepseek", "kimi", "bailian", "custom"], names
        ds = prov["providers"][0]
        # 真实 key 从本地 config.json 读取，断言接口响应中不出现明文（不硬编码 key）
        with open(os.path.join(ROOT, "config.json"), encoding="utf-8") as _f:
            _real_key = (json.load(_f).get("providers", {}).get("deepseek", {}) or {}).get("api_key", "")
        assert ds["has_key"] and "…" in ds["key_masked"], "key 未掩码!"
        assert not _real_key or _real_key not in json.dumps(prov), "key 泄漏!"
        print("[2/6] 供应商接口 OK（key 已掩码）:", names, "| 当前:", prov["current"])

        # 配置保存（给 custom 写模型列表 + 换默认再换回）
        cfg_body = json.dumps({
            "current": "deepseek",
            "updates": {"custom": {"base_url": "https://example.com/v1", "models": "m1, m2", "model": "m1"}}
        }).encode("utf-8")
        req = urllib.request.Request(BASE + "/api/config", data=cfg_body,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=5) as resp:
            prov2 = json.loads(resp.read().decode("utf-8"))
        custom = [p for p in prov2["providers"] if p["name"] == "custom"][0]
        assert custom["models"] == ["m1", "m2"] and custom["model"] == "m1"
        # 确认 deepseek key 未被误清
        ds2 = [p for p in prov2["providers"] if p["name"] == "deepseek"][0]
        assert ds2["has_key"], "key 被误清空!"
        print("[3/6] 配置保存 OK（custom.models=%s，deepseek key 保留）" % custom["models"])

        body = json.dumps({"overview": OVERVIEW, "notes": NOTES}).encode("utf-8")
        req = urllib.request.Request(BASE + "/api/generate", data=body,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=300) as resp:
            result = json.loads(resp.read().decode("utf-8"))
        job = result["job"]
        print("[4/8] 生成成功 job=%s filename=%s" % (job, result["filename"]))
        print("      meta:", json.dumps(result["data"]["meta"], ensure_ascii=False))
        print("      content 条数:", len(result["data"]["content"]),
              "| minutes 小节:", len(result["data"]["minutes"]),
              "| gains:", len(result["data"]["gains"]),
              "| mistakes:", len(result["data"]["mistakes"]))

        req = urllib.request.Request(BASE + "/api/file?job=%s&fmt=docx" % job)
        with urllib.request.urlopen(req, timeout=30) as resp:
            docx_bytes = resp.read()
        assert docx_bytes[:2] == b"PK", "docx 无效"
        print("[5/8] docx 有效，%d 字节" % len(docx_bytes))

        # /api/update：修改标题后重建 docx，并验证 pdf 缓存被作废
        data2 = result["data"]
        data2["meta"]["title"] = "声音是什么（修改测试）"
        data2["content"].append({"lead": "新增条目：", "text": "这是通过 /api/update 添加的一行。"})
        body = json.dumps({"job": job, "data": data2}).encode("utf-8")
        req = urllib.request.Request(BASE + "/api/update", data=body,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            upd = json.loads(resp.read().decode("utf-8"))
        assert upd.get("ok"), upd
        assert "修改测试" in upd["filename"], upd["filename"]
        print("[6/8] update 接口 OK，新文件名:", upd["filename"])

        import zipfile, io as _io
        req = urllib.request.Request(BASE + "/api/file?job=%s&fmt=docx&t=2" % job)
        with urllib.request.urlopen(req, timeout=30) as resp:
            docx2 = resp.read()
        assert docx2[:2] == b"PK"
        z = zipfile.ZipFile(_io.BytesIO(docx2))
        xml = z.read("word/document.xml").decode("utf-8")
        assert "修改测试" in xml and "api/update" in xml, "docx 未包含修改内容"
        pdf_stale = os.path.exists(os.path.join(ROOT, "runtime", job, "report.pdf"))
        assert not pdf_stale, "旧 PDF 未被作废"
        print("[7/8] docx 重建含修改内容，旧 PDF 已作废")

        req = urllib.request.Request(BASE + "/api/file?job=%s&fmt=pdf" % job)
        t0 = time.time()
        with urllib.request.urlopen(req, timeout=240) as resp:
            pdf_bytes = resp.read()
        assert pdf_bytes[:4] == b"%PDF", "pdf 无效"
        print("[8/8] pdf 有效，%d 字节，转换耗时 %.1fs" % (len(pdf_bytes), time.time() - t0))
        print("E2E PASS")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()


if __name__ == "__main__":
    main()
