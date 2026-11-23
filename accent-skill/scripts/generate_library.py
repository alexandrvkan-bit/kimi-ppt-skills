#!/usr/bin/env python3
"""Generate minimal library.pptx for accent-skill."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    colors = {
        "white": RGBColor(0xFF, 0xFF, 0xFF),
        "blue": RGBColor(0x21, 0x96, 0xF3),
        "dark": RGBColor(0x21, 0x21, 0x21),
        "light_blue": RGBColor(0xE3, 0xF2, 0xFD),
        "gray": RGBColor(0x75, 0x75, 0x75),
    }

    def add_text(slide, left, top, width, height, text, size=14, bold=False, color=colors["dark"], align=PP_ALIGN.LEFT):
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        shape.fill.background()
        shape.line.fill.background()
        tf = shape.text_frame
        tf.text = text
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.runs[0] if p.runs else p.add_run()
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "Raleway"
        return shape

    def bar(slide, y):
        b = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(y), Inches(13.333), Inches(0.7))
        b.fill.solid()
        b.fill.fore_color.rgb = colors["blue"]
        b.line.fill.background()

    # Slide 1: Title
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(2), Inches(11), Inches(1.5), "{{TITLE}}", 44, True, colors["dark"], PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(3.5), Inches(11), Inches(1), "{{SUBTITLE}}", 20, False, colors["gray"], PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(5), Inches(11), Inches(0.5), "{{META}}", 15, False, colors["gray"], PP_ALIGN.CENTER)
    bar(s, 6.8)

    # Slide 2: Thesis
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(2.5), Inches(11), Inches(1.5), "{{THESIS}}", 40, True, colors["blue"], PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(4.2), Inches(11), Inches(0.8), "{{CAPTION}}", 16, False, colors["dark"], PP_ALIGN.CENTER)

    # Slide 3: List
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 28, True, colors["dark"])
    add_text(s, Inches(1.5), Inches(2), Inches(10), Inches(4), "{{ITEMS}}", 18, False, colors["dark"])

    # Slide 4: 3 Metrics
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 24, True, colors["dark"])
    for i in range(3):
        left = Inches(1 + i * 4)
        box = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.5), Inches(3.5), Inches(2.5))
        box.fill.solid()
        box.fill.fore_color.rgb = colors["light_blue"]
        box.line.fill.background()
        add_text(s, left, Inches(2.8), Inches(3.5), Inches(1), f"{{M{i+1}}}", 36, True, colors["blue"], PP_ALIGN.CENTER)
        add_text(s, left, Inches(3.8), Inches(3.5), Inches(0.6), f"{{L{i+1}}}", 14, False, colors["dark"], PP_ALIGN.CENTER)

    # Slide 5: Timeline
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 24, True, colors["dark"])
    bar_shape = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(3.5), Inches(11), Inches(0.1))
    bar_shape.fill.solid()
    bar_shape.fill.fore_color.rgb = colors["blue"]
    bar_shape.line.fill.background()
    add_text(s, Inches(1), Inches(4), Inches(11), Inches(2), "{{MILESTONES}}", 14, False, colors["dark"])

    # Slide 6: SWOT
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.5), Inches(11), Inches(0.6), "{{TITLE}}", 24, True, colors["dark"])
    quadrants = [
        ("S", colors["light_blue"], 1, 1.5),
        ("W", RGBColor(0xFF, 0xEB, 0xEE), 6.667, 1.5),
        ("O", RGBColor(0xE8, 0xF5, 0xE9), 1, 4.5),
        ("T", RGBColor(0xFF, 0xF3, 0xE0), 6.667, 4.5),
    ]
    for key, fill, left, top in quadrants:
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(left), Inches(top), Inches(5.667), Inches(2.5))
        box.fill.solid()
        box.fill.fore_color.rgb = fill
        box.line.color.rgb = colors["gray"]
        add_text(s, Inches(left + 0.2), Inches(top + 0.2), Inches(5), Inches(0.5), f"{{H{key}}}", 18, True, colors["dark"])
        add_text(s, Inches(left + 0.2), Inches(top + 0.8), Inches(5), Inches(1.5), f"{{{key}}}", 14, False, colors["dark"])

    # Slide 7: Table
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 24, True, colors["dark"])
    add_text(s, Inches(1), Inches(2), Inches(11), Inches(4), "{{TABLE}}", 14, False, colors["dark"])

    # Slide 8: Bar chart
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 24, True, colors["dark"])
    add_text(s, Inches(1), Inches(2), Inches(11), Inches(4.5), "{{CHART}}", 14, False, colors["dark"])

    # Slide 9: Org chart
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 24, True, colors["dark"])
    add_text(s, Inches(1), Inches(2), Inches(11), Inches(4.5), "{{ORG}}", 14, False, colors["dark"])

    # Slide 10: Status panel
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(0.8), Inches(11), Inches(0.8), "{{TITLE}}", 24, True, colors["dark"])
    add_text(s, Inches(1), Inches(2), Inches(11), Inches(4.5), "{{PROJECTS}}", 14, False, colors["dark"])

    # Slide 11: Conclusion
    s = prs.slides.add_slide(blank)
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = colors["white"]
    add_text(s, Inches(1), Inches(2.5), Inches(11), Inches(1.5), "{{CONCLUSION}}", 36, True, colors["blue"], PP_ALIGN.CENTER)
    add_text(s, Inches(1), Inches(4.2), Inches(11), Inches(0.8), "{{SUBTITLE}}", 18, False, colors["dark"], PP_ALIGN.CENTER)
    bar(s, 6.8)

    prs.save("assets/library.pptx")
    print("Generated assets/library.pptx with", len(prs.slides), "slides")

if __name__ == "__main__":
    main()
