# -*- coding: utf-8 -*-
"""docx → pdf (설치된 MS Word 사용). 렌더 결과 육안 검수용.
사용법: python to_pdf.py <파일.docx> [출력.pdf]
"""
import sys
from pathlib import Path

import win32com.client


def convert(src, dst=None):
    src = Path(src).resolve()
    dst = Path(dst).resolve() if dst else src.with_suffix(".pdf")
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(str(src), False, True)
        doc.ExportAsFixedFormat(str(dst), 17)
        pages = doc.ComputeStatistics(2)
        doc.Close(0)
    finally:
        word.Quit()
    return dst, pages


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    out, n = convert(*sys.argv[1:3])
    print(f"OK {out} ({n}쪽)")
