# -*- coding: utf-8 -*-
"""docx → pdf：优先 Word COM（保真），失败时回退 LibreOffice。"""
import os
import shutil
import subprocess

POWERSHELL = (shutil.which("powershell")
              or r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe")

# 隐藏子进程控制台黑框（pythonw/无窗口环境下尤其重要）
_NO_WINDOW = {"creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0)} if os.name == "nt" else {}

PS_SCRIPT = r"""
param([string]$Src, [string]$Dst)
$ErrorActionPreference = "Stop"
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {
    $doc = $word.Documents.Open($Src, $false, $true)
    $doc.SaveAs2($Dst, 17)
    $doc.Close($false)
} finally {
    $word.Quit()
}
"""


def convert_with_word(src, dst):
    ps1 = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "runtime", "_convert.ps1")
    os.makedirs(os.path.dirname(ps1), exist_ok=True)
    with open(ps1, "w", encoding="utf-8") as f:
        f.write(PS_SCRIPT)
    subprocess.run(
        [POWERSHELL, "-NoProfile", "-ExecutionPolicy", "Bypass",
         "-File", ps1, "-Src", src, "-Dst", dst],
        check=True, capture_output=True, timeout=180, **_NO_WINDOW,
    )


def convert_with_libreoffice(src, dst):
    outdir = os.path.dirname(dst)
    subprocess.run(
        [r"C:\Program Files\LibreOffice\program\soffice.exe",
         "--headless", "--nologo", "--norestore",
         "--convert-to", "pdf", "--outdir", outdir, src],
        check=True, capture_output=True, timeout=180, **_NO_WINDOW,
    )
    produced = os.path.join(outdir, os.path.splitext(os.path.basename(src))[0] + ".pdf")
    if produced != dst and os.path.isfile(produced):
        os.replace(produced, dst)


def convert(src, dst):
    try:
        convert_with_word(src, dst)
    except Exception:
        convert_with_libreoffice(src, dst)
    if not os.path.isfile(dst):
        raise RuntimeError("PDF 转换失败：%s" % dst)
    return dst
