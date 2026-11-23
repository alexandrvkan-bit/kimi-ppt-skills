#!/usr/bin/env python3
"""Рисование фигур и схем на слайдах."""
import os

from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def draw_arrow(slide, start, end, color):
    """Рисует стрелку между двумя точками."""
    from pptx.enum.shapes import MSO_CONNECTOR
    connector = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT, start[0], start[1], end[0], end[1]
    )
    connector.line.color.rgb = color
    connector.line.width = Pt(2)
    return connector


def draw_box(slide, left, top, width, height, text, fill_color, text_color):
    """Рисует прямоугольник с текстом."""
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = text_color
    shape.line.width = Pt(1)
    if shape.has_text_frame:
        shape.text_frame.text = text
        for paragraph in shape.text_frame.paragraphs:
            paragraph.font.color.rgb = text_color
            paragraph.font.size = Pt(14)
    return shape


def draw_circle(slide, left, top, width, height, text, fill_color, text_color):
    """Рисует круг с текстом."""
    from pptx.enum.shapes import MSO_SHAPE
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.color.rgb = text_color
    shape.line.width = Pt(1)
    if shape.has_text_frame:
        shape.text_frame.text = text
        for paragraph in shape.text_frame.paragraphs:
            paragraph.font.color.rgb = text_color
            paragraph.font.size = Pt(14)
    return shape
