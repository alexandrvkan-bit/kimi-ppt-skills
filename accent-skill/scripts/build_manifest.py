#!/usr/bin/env python3
"""Генерация manifest.json из library.pptx"""
import json
import os
import sys

from pptx import Presentation
from pptx.util import Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def extract_manifest(lib_path):
    prs = Presentation(lib_path)
    layouts = []
    for i, slide in enumerate(prs.slides):
        slots = []
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            # Ищем placeholder вида {{KEY}}
            import re
            for match in re.finditer(r"\{\{([A-Z0-9_]+)\}\}", text):
                key = match.group(1)
                font_size = Pt(12)
                if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                    font_size = shape.text_frame.paragraphs[0].runs[0].font.size or Pt(12)
                slots.append({
                    "key": key,
                    "label": text.replace(f"{{{{{key}}}}}", "").strip(),
                    "size": round(font_size.pt),
                    "capacity": estimate_capacity(shape, font_size),
                    "sample": text
                })
        layouts.append({
            "code": f"A{i+1:03d}",
            "num": i + 1,
            "name": f"Макет {i+1}",
            "layout_index": i,
            "slots": slots
        })
    return {"layouts": layouts}


def estimate_capacity(shape, font_size):
    """Оценка ёмкости поля в символах."""
    if not shape.width or not shape.height:
        return 100
    char_w = font_size.pt * 0.55 * 12700  # EMU
    char_h = font_size.pt * 1.2 * 12700
    cols = max(1, int(shape.width / char_w))
    rows = max(1, int(shape.height / char_h))
    return cols * rows


def main():
    lib_path = os.path.join(ROOT, "assets", "library.pptx")
    if len(sys.argv) > 1:
        lib_path = sys.argv[1]
    man = extract_manifest(lib_path)
    out_path = os.path.join(ROOT, "assets", "manifest.json")
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(man, f, ensure_ascii=False, indent=2)
    print(f"Manifest saved: {out_path} ({len(man['layouts'])} layouts)")


if __name__ == "__main__":
    main()
