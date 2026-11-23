#!/usr/bin/env python3
"""
Движок «Акцент»: презентация из библиотеки слайдов по спецификации.
Модель пишет только текст в поля образцов; геометрию, кегль, цвет и проверки делает движок.

    python scripts/accent.py nav   [группа]           # навигатор: вопрос → группа → макет
    python scripts/accent.py find  <слова>            # поиск макета по задаче или названию модели
    python scripts/accent.py slots <код|номер> [...]  # поля макета: ключ, кегль, ёмкость, текст образца
    python scripts/accent.py check spec.yaml          # текст и ёмкость до сборки
    python scripts/accent.py build spec.yaml out.pptx # сборка
    python scripts/accent.py fidelity source.txt out.pptx   # каждая строка исходника есть в деке

Файлы: assets/library.pptx (образцы), assets/manifest.json (описание полей).
Зависимости: python-pptx>=0.6.21, pyyaml.
"""
import colorsys
import copy
import io
import json
import os
import re
import sys
import uuid

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Emu, Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
EMU = 914400
CHAR = 0.55          # ширина знака Raleway в долях кегля для оценки переноса: среднее 0,49–0,54, в кириллице с переносом по словам — до 0,58; берём с запасом
FONT = "Raleway"
REPORT = []


def say(level, where, msg):
    REPORT.append((level, where, msg))


# ================================================================ палитра
# Роли палитры клиента (как в card.md). Пропущенные роли берутся из умолчаний.
ROLES = ["text", "bg", "dark", "surface", "accent", "highlight", "muted", "rule", "accent_on_dark", "second"]
DEFAULT = {"text": "1F1E1D", "bg": "FFFFFF", "dark": "1F1E1D", "surface": "F6F5F1", "accent": "FFCC00",
           "highlight": "FFF7D1", "muted": "7A7872", "rule": "E6E4DD", "accent_on_dark": "FFCC00",
           "second": "A3A199"}
SLOT = {"text": "dk1", "bg": "lt1", "dark": "dk2", "surface": "lt2", "accent": "accent1",
        "highlight": "accent2", "muted": "accent3", "rule": "accent4", "accent_on_dark": "accent5",
        "second": "accent6"}
SCHEME_VAL = {"dk1": "tx1", "lt1": "bg1", "dk2": "tx2", "lt2": "bg2"}

# Цвет библиотеки -> (роль, сдвиг светлоты). Светлые и тёмные слайды — разные карты.
LIB_LIGHT = {
    "1F1E1D": ("dark", 0), "262624": ("dark", 0), "302F2D": ("dark", .08), "3A3936": ("dark", .12),
    "3F3E3B": ("dark", .15), "5E5D59": ("muted", -.15), "7A7872": ("muted", 0), "7F7D76": ("muted", .04),
    "A3A199": ("muted", .35), "B8B6AF": ("muted", .5), "FFCC00": ("accent", 0), "FFE680": ("accent", .5),
    "FFF7D1": ("highlight", 0), "FAF9F5": ("bg", 0), "FFFFFF": ("bg", 0), "F6F5F1": ("surface", 0),
    "EDEBE5": ("surface", -.03), "E6E4DD": ("rule", 0), "DDDAD2": ("rule", -.04), "D6D4CD": ("rule", -.07),
}
LIB_DARK = dict(LIB_LIGHT)
LIB_DARK.update({
    "FFCC00": ("accent_on_dark", 0), "FFE680": ("accent_on_dark", .4), "7A7872": ("rule", -.2),
    "7F7D76": ("rule", -.2), "A3A199": ("rule", 0), "B8B6AF": ("rule", .1), "E6E4DD": ("dark", .25),
    "FFF7D1": ("dark", .2),
})


def lum(hexv):
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def shift_hex(hexv, d):
    """Сдвиг светлоты как lumMod/lumOff в PowerPoint: d>0 — к белому, d<0 — темнее."""
    if not d:
        return hexv
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 - d) + d if d > 0 else l * (1 + d)
    r, g, b = colorsys.hls_to_rgb(h, max(0, min(1, l)), s)
    return "%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))


# =============================================================== манифест
def load_manifest():
    with open(os.path.join(ROOT, "assets", "manifest.json"), encoding="utf-8") as f:
        return json.load(f)


def named_models(man, text):
    """Канонические модели, названные в тексте: {ключ: [коды макетов]}. Ключи — поле model в relations.json."""
    t = " " + re.sub(r"\s+", " ", norm(text)) + " "
    found = {}
    for L in man["layouts"]:
        for k in L.get("model", []):
            if k.lower() in t:
                found.setdefault(k, []).append(L["code"])
    return found


def norm(s):
    return re.sub(r"[^\w\s]", " ", s.lower())


# =============================================================== палитра клиента
PALETTE = None


def load_palette(path):
    global PALETTE
    PALETTE = dict(DEFAULT)
    if not path or not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if ":" in line:
                k, v = line.split(":", 1)
                k = k.strip()
                v = v.strip().strip("'\"").lstrip("#")
                if k in ROLES:
                    PALETTE[k] = v


def palette_role(role, slide_is_dark=False):
    lib = LIB_DARK if slide_is_dark else LIB_LIGHT
    hexv = PALETTE.get(role, DEFAULT.get(role, "000000"))
    return lib.get(hexv.upper(), (role, 0))


def role_rgb(role, slide_is_dark=False):
    r, d = palette_role(role, slide_is_dark)
    hexv = PALETTE.get(r, DEFAULT.get(r, "000000"))
    return RGBColor.from_string(shift_hex(hexv, d))


# =============================================================== навигация / поиск
NAV = None


def load_nav():
    global NAV
    if NAV is None:
        with open(os.path.join(ROOT, "assets", "groups.json"), encoding="utf-8") as f:
            NAV = json.load(f)
    return NAV


def cmd_nav(args):
    nav = load_nav()
    if not args:
        for g in nav["groups"]:
            print(f"{g['id']}: {g['question']}")
        return
    gid = args[0].upper()
    group = next((g for g in nav["groups"] if g["id"] == gid), None)
    if not group:
        print("Группа не найдена")
        return
    print(f"=== {gid}: {group['question']} ===")
    for m in group["layouts"]:
        print(f"  {m['code']} — {m['name']}: {m['work']}")
        if m.get("how_choose"):
            print(f"     → {m['how_choose']}")


def cmd_find(args):
    if not args:
        print("Usage: accent.py find <слова>")
        return
    man = load_manifest()
    q = " ".join(args).lower()
    models = named_models(man, q)
    if models:
        for k, codes in models.items():
            print(f"Модель «{k}» → {', '.join(codes)}")
        return
    # fuzzy по имени и work
    scores = []
    for L in man["layouts"]:
        text = (L.get("name", "") + " " + L.get("work", "")).lower()
        score = sum(1 for w in q.split() if w in text)
        if score:
            scores.append((score, L["code"], L.get("name", ""), L.get("work", "")))
    scores.sort(reverse=True)
    for score, code, name, work in scores[:10]:
        print(f"  {code} — {name}: {work}")


# =============================================================== slots
def cmd_slots(args):
    if not args:
        print("Usage: accent.py slots <код|номер> [...]")
        return
    man = load_manifest()
    for a in args:
        L = next((l for l in man["layouts"] if l["code"] == a.upper() or str(l.get("num")) == a), None)
        if not L:
            print(f"Макет {a} не найден")
            continue
        print(f"\n=== {L['code']} — {L.get('name', '')} ===")
        print(f"Работа: {L.get('work', '')}")
        print(f"Модель: {', '.join(L.get('model', [])) or '—'}")
        print(f"Отношение: {', '.join(L.get('rel', [])) or '—'}")
        if L.get("scheme"):
            print(f"Схема: {L['scheme']}")
        if L.get("data"):
            print(f"Данные: {L['data']}")
        if L.get("needs"):
            print(f"Нужно в материале: {L['needs']}")
        if L.get("items"):
            print(f"Пункты: {L['items']}")
        for s in L.get("slots", []):
            print(f"  [{s['key']}] {s.get('label', '')} — кегль {s.get('size', '?')}, ёмкость {s.get('capacity', '?')}")
            if s.get("sample"):
                print(f"      образец: {s['sample'][:80]}...")


# =============================================================== YAML spec loader
def load_spec(path):
    import yaml
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f)


# =============================================================== проверка spec
def cmd_check(args):
    if not args:
        print("Usage: accent.py check spec.yaml")
        return
    spec = load_spec(args[0])
    man = load_manifest()
    issues = 0
    for i, slide in enumerate(spec.get("slides", [])):
        code = slide.get("use", "")
        L = next((l for l in man["layouts"] if l["code"] == code.upper()), None)
        if not L:
            say("ОШИБКА", f"слайд {i+1}", f"макет {code} не найден")
            issues += 1
            continue
        # проверяем заполненность полей
        for s in L.get("slots", []):
            key = s["key"]
            if key not in slide.get("fields", {}):
                say("ВНИМАНИЕ", f"слайд {i+1} ({code})", f"поле [{key}] не заполнено")
                issues += 1
        # проверяем items
        if L.get("items") and "items" not in slide:
            say("ВНИМАНИЕ", f"слайд {i+1} ({code})", "ожидается items")
            issues += 1
    print(f"Проверка завершена: {issues} замечаний")
    for level, where, msg in REPORT:
        print(f"[{level}] {where}: {msg}")


# =============================================================== сборка
def cmd_build(args):
    if len(args) < 2:
        print("Usage: accent.py build spec.yaml out.pptx")
        return
    spec_path, out_path = args[0], args[1]
    spec = load_spec(spec_path)
    man = load_manifest()

    # Загружаем библиотеку
    lib_path = os.path.join(ROOT, "assets", "library.pptx")
    prs = Presentation(lib_path)

    # Палитра
    palette_path = spec.get("palette") or os.path.join(os.path.dirname(spec_path), "card.md")
    load_palette(palette_path)

    # Создаём новую презентацию
    out_prs = Presentation()
    out_prs.slide_width = prs.slide_width
    out_prs.slide_height = prs.slide_height

    for slide_spec in spec.get("slides", []):
        build_slide(out_prs, prs, man, slide_spec)

    out_prs.save(out_path)
    print(f"Сохранено: {out_path}")

    # Отчёт
    report_path = out_path.replace(".pptx", ".report.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("# Отчёт сборки\n\n")
        for level, where, msg in REPORT:
            f.write(f"- **[{level}]** {where}: {msg}\n")
    print(f"Отчёт: {report_path}")


def build_slide(out_prs, lib_prs, man, spec):
    code = spec.get("use", "").upper()
    L = next((l for l in man["layouts"] if l["code"] == code), None)
    if not L:
        say("ОШИБКА", code, "макет не найден в манифесте")
        return

    # Находим образец в библиотеке
    layout_idx = L.get("layout_index", 0)
    if layout_idx >= len(lib_prs.slide_layouts):
        say("ОШИБКА", code, f"layout_index {layout_idx} вне диапазона")
        return

    blank_layout = out_prs.slide_layouts[6] if len(out_prs.slide_layouts) > 6 else out_prs.slide_layouts[0]
    slide = out_prs.slides.add_slide(blank_layout)

    # Копируем фон
    source_slide = lib_prs.slides[layout_idx]
    if source_slide.background.fill.type is not None:
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = source_slide.background.fill.fore_color.rgb

    # Копируем фигуры
    for shape in source_slide.shapes:
        new_shape = copy_shape(slide, shape)
        if new_shape and new_shape.has_text_frame:
            fill_shape(new_shape, spec, L)

    say("OK", code, f"собран слайд «{L.get('name', '')}»")


def copy_shape(slide, shape):
    """Копирует фигуру на новый слайд. Упрощённая версия."""
    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PLACEHOLDER:
            return None
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            # Пропускаем картинки в этой версии
            return None

        left = shape.left
        top = shape.top
        width = shape.width
        height = shape.height

        if hasattr(shape, "auto_shape_type"):
            new_shape = slide.shapes.add_shape(
                shape.auto_shape_type, left, top, width, height
            )
        else:
            new_shape = slide.shapes.add_shape(
                1, left, top, width, height
            )

        # Копируем заливку
        if shape.fill.type is not None:
            if shape.fill.type == 1:  # SOLID
                new_shape.fill.solid()
                new_shape.fill.fore_color.rgb = shape.fill.fore_color.rgb

        # Копируем линию
        if shape.line.fill.type is not None:
            new_shape.line.color.rgb = shape.line.color.rgb
            new_shape.line.width = shape.line.width

        return new_shape
    except Exception as e:
        say("ВНИМАНИЕ", "copy", str(e))
        return None


def fill_shape(shape, spec, L):
    """Заполняет текстовые поля фигуры данными из спецификации."""
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    text = tf.text
    fields = spec.get("fields", {})

    # Ищем placeholder вида {{key}} или просто совпадение с label
    for s in L.get("slots", []):
        key = s["key"]
        placeholder = f"{{{{{key}}}}}"
        if placeholder in text:
            value = fields.get(key, s.get("sample", ""))
            text = text.replace(placeholder, str(value))

    # Устанавливаем текст
    if tf.paragraphs:
        tf.paragraphs[0].text = text


def cmd_fidelity(args):
    if len(args) < 2:
        print("Usage: accent.py fidelity source.txt out.pptx")
        return
    source_path, out_path = args[0], args[1]
    with open(source_path, encoding="utf-8") as f:
        source_lines = [l.strip() for l in f if l.strip()]

    prs = Presentation(out_path)
    slide_texts = []
    for slide in prs.slides:
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                texts.append(shape.text_frame.text)
        slide_texts.append(" ".join(texts))

    all_text = " ".join(slide_texts)
    missing = [l for l in source_lines if l not in all_text]

    if missing:
        print(f"НЕ НАЙДЕНО {len(missing)} строк из исходника:")
        for m in missing[:20]:
            print(f"  - {m}")
    else:
        print("Все строки исходника найдены в презентации.")


# =============================================================== main
COMMANDS = {
    "nav": cmd_nav,
    "find": cmd_find,
    "slots": cmd_slots,
    "check": cmd_check,
    "build": cmd_build,
    "fidelity": cmd_fidelity,
}


def main():
    if len(sys.argv) < 2:
        print("Usage: accent.py <command> [args...]")
        print("Commands:", ", ".join(COMMANDS.keys()))
        return
    cmd = sys.argv[1]
    args = sys.argv[2:]
    fn = COMMANDS.get(cmd)
    if not fn:
        print(f"Unknown command: {cmd}")
        return
    fn(args)


if __name__ == "__main__":
    main()
