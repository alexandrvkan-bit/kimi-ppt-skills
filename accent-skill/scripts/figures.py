#!/usr/bin/env python3
"""Схемы и диаграммы для презентаций."""
import os

from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def draw_bcg_matrix(slide, data, colors):
    """Рисует матрицу BCG."""
    from pptx.enum.shapes import MSO_SHAPE
    # Ось X — доля рынка, ось Y — темпы роста
    width = Inches(8)
    height = Inches(6)
    left = Inches(1)
    top = Inches(1.5)

    # Фон матрицы
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = colors.get("bg")
    bg.line.color.rgb = colors.get("rule")

    # Линии осей
    mid_x = left + width / 2
    mid_y = top + height / 2
    # ...


def draw_swot(slide, data, colors):
    """Рисует SWOT-матрицу 2x2."""
    from pptx.enum.shapes import MSO_SHAPE
    cell_w = Inches(4)
    cell_h = Inches(2.5)
    start_left = Inches(1)
    start_top = Inches(2)

    positions = [
        ("S", start_left, start_top),
        ("W", start_left + cell_w, start_top),
        ("O", start_left, start_top + cell_h),
        ("T", start_left + cell_w, start_top + cell_h),
    ]

    labels = {"S": "Strengths", "W": "Weaknesses", "O": "Opportunities", "T": "Threats"}
    fills = {"S": "E8F5E9", "W": "FFEBEE", "O": "E3F2FD", "T": "FFF3E0"}

    for key, l, t in positions:
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, cell_w, cell_h)
        shape.fill.solid()
        shape.fill.fore_color.rgb = colors.get("accent") if key == "S" else colors.get("bg")
        shape.line.color.rgb = colors.get("rule")
        if shape.has_text_frame:
            shape.text_frame.text = f"{labels[key]}\n{data.get(key, '')}"


def draw_timeline(slide, items, colors):
    """Рисует горизонтальный таймлайн."""
    from pptx.enum.shapes import MSO_SHAPE
    start = Inches(1)
    y = Inches(3)
    total_w = Inches(10)
    n = len(items)
    step = total_w / max(n - 1, 1)

    # Линия
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, start, y - Pt(2), total_w, Pt(4))
    line.fill.solid()
    line.fill.fore_color.rgb = colors.get("accent")

    for i, item in enumerate(items):
        x = start + i * step
        # Точка
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, x - Pt(8), y - Pt(8), Pt(16), Pt(16))
        dot.fill.solid()
        dot.fill.fore_color.rgb = colors.get("accent")
        # Текст
        tx = slide.shapes.add_textbox(x - Inches(1), y + Pt(12), Inches(2), Inches(1))
        tx.text_frame.text = item


def draw_pie(slide, data, colors, left, top, width, height):
    """Рисует круговую диаграмму (упрощённо — секторами)."""
    from pptx.enum.shapes import MSO_SHAPE
    import math
    total = sum(v for _, v in data)
    start_angle = 0
    cx = left + width / 2
    cy = top + height / 2
    r = min(width, height) / 2

    palette = [colors.get("accent"), colors.get("dark"), colors.get("muted"),
               colors.get("highlight"), colors.get("second")]

    for i, (label, value) in enumerate(data):
        angle = (value / total) * 360
        # Упрощённо: рисуем сектор как треугольник + дуга
        # В реальности нужна более сложная геометрия
        shape = slide.shapes.add_shape(MSO_SHAPE.PIE, left, top, width, height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = palette[i % len(palette)]
        shape.line.color.rgb = colors.get("bg")


def draw_bar_chart(slide, data, colors, left, top, width, height):
    """Рисует столбчатую диаграмму."""
    from pptx.enum.shapes import MSO_SHAPE
    n = len(data)
    bar_w = width / (n * 2)
    max_val = max(v for _, v in data) if data else 1

    for i, (label, value) in enumerate(data):
        bar_h = (value / max_val) * height
        x = left + i * (width / n) + bar_w / 2
        y = top + height - bar_h
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, bar_w, bar_h)
        bar.fill.solid()
        bar.fill.fore_color.rgb = colors.get("accent")
        bar.line.color.rgb = colors.get("bg")

        # Подпись
        tx = slide.shapes.add_textbox(x, top + height + Pt(4), bar_w, Pt(20))
        tx.text_frame.text = label
        tx.text_frame.paragraphs[0].font.size = Pt(10)

        # Значение
        val_tx = slide.shapes.add_textbox(x, y - Pt(16), bar_w, Pt(14))
        val_tx.text_frame.text = str(value)
        val_tx.text_frame.paragraphs[0].font.size = Pt(10)
