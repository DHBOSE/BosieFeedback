# -*- coding: utf-8 -*-
"""批量打包验证：造两个 job（同文件名，测重名）→ /api/batch_zip → 校验 zip 内容 → 清理"""
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
JOBS = ["batchtest-000000-aaaaaa", "batchtest-000000-bbbbbb"]

DATA = {
    "meta": {"title": "批量测试", "subtitle": "—— 副标题", "subject": "初中物理",
             "form": "一对一", "content": "测试内容", "date": "2026 年 8 月 24 日"},
    "mainline": "主线", "content": [{"lead": "点", "text": "内容"}],
    "performance": {"pros": [{"lead": "优", "text": "好"}], "cons": [{"lead": "改", "text": "进"}]},
    "minutes": [{"title": "纪要", "points": ["要点"]}],
    "gains": [{"lead": "获", "text": "得"}], "homework": [{"lead": "作", "text": "业"}],
    "mistakes": [{"lead": "错", "text": "析"}], "suggestions": [{"lead": "建", "text": "议"}],
}


def post(route, body, timeout=120, raw=False):
    req = urllib.request.Request(BASE + route, data=json.dumps(body).encode("utf-8"),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        data = resp.read()
        return (data, resp.headers) if raw else json.loads(data.decode("utf-8"))


def main():
    for job in JOBS:
        d = os.path.join(ROOT, "runtime", job)
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "data.json"), "w", encoding="utf-8") as f:
            json.dump(DATA, f, ensure_ascii=False)

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

        # 1. 两个 job 各生成 docx
        for job in JOBS:
            d = post("/api/update", {"job": job, "data": DATA})
            assert d.get("ok"), d
        print("[1/3] 两个测试 job 的 docx 已生成（同文件名，测重名）")

        # 2. 打包 zip（不含 PDF，快速路径）
        raw, headers = post("/api/batch_zip", {"jobs": JOBS, "include_pdf": False}, raw=True)
        assert raw[:2] == b"PK", "zip 无效"
        assert "attachment" in headers.get("Content-Disposition", ""), "缺下载头"
        zpath = os.path.join(ROOT, "runtime", "_batch_test.zip")
        with open(zpath, "wb") as f:
            f.write(raw)
        z = zipfile.ZipFile(zpath)
        names = z.namelist()
        assert "初中物理-批量测试-课堂报告.docx" in names, names
        assert "初中物理-批量测试-课堂报告-2.docx" in names, "重名未加序号: %s" % names
        for n in names:
            assert z.read(n)[:2] == b"PK"
        z.close()
        os.remove(zpath)
        print("[2/3] zip OK：%s（重名自动加序号 ✓）" % names)

        # 3. 非法 job 拒绝
        try:
            post("/api/batch_zip", {"jobs": ["../../etc"]})
            raise SystemExit("非法 job 未被拒绝!")
        except urllib.error.HTTPError as e:
            assert e.code == 400, e.code
        print("[3/3] 非法 job 校验 OK（400）")
        print("BATCH TEST PASS")
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        for job in JOBS:
            shutil.rmtree(os.path.join(ROOT, "runtime", job), ignore_errors=True)
        shutil.rmtree(os.path.join(ROOT, "runtime", "_zips"), ignore_errors=True)


if __name__ == "__main__":
    main()
