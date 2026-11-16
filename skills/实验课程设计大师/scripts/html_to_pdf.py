# -*- coding: utf-8 -*-
"""把 A4 HTML 教案转换为 PDF（调用本机 Edge / Chrome 的无头打印，零依赖）。

用法:
    python html_to_pdf.py <input.html> <output.pdf>

说明:
  * 优先使用 Microsoft Edge，其次 Google Chrome；均按 A4 无页边距打印
  * 生成的 PDF 尺寸为 A4（210x297mm），可直接送打印机
"""
import os
import subprocess
import sys
import tempfile

BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\Application\msedge.exe"),
    os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
]


def find_browser():
    for p in BROWSERS:
        if p and os.path.isfile(p):
            return p
    return None


def main():
    if len(sys.argv) != 3:
        print("用法: python html_to_pdf.py <input.html> <output.pdf>")
        sys.exit(1)

    src = os.path.abspath(sys.argv[1])
    dst = os.path.abspath(sys.argv[2])
    if not os.path.isfile(src):
        print("错误: 找不到输入文件", src)
        sys.exit(1)

    browser = find_browser()
    if not browser:
        print("未找到 Edge / Chrome，无法自动导出 PDF。"
              "请改用 HTML 内置的「一键生成 PDF」按钮手动另存。")
        sys.exit(2)

    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    if os.path.exists(dst):
        os.remove(dst)

    profile = os.path.join(tempfile.gettempdir(), "wb_pdf_profile")
    args = [
        browser,
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--no-first-run",
        "--disable-extensions",
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=6000",
        "--user-data-dir=" + profile,
        "--print-to-pdf=" + dst,
        "--no-pdf-header-footer",
        "file:///" + src.replace("\\", "/"),
    ]
    try:
        proc = subprocess.run(args, capture_output=True, timeout=120)
    except subprocess.TimeoutExpired:
        print("错误: 浏览器导出 PDF 超时")
        sys.exit(3)

    if not os.path.isfile(dst) or os.path.getsize(dst) == 0:
        err = (proc.stderr or b"").decode("utf-8", "ignore")[-400:]
        print("错误: PDF 生成失败\n" + err)
        sys.exit(4)

    print("OK:", dst)


if __name__ == "__main__":
    main()
