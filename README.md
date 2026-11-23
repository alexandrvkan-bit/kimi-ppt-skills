# Акцент — TNF Презентация

Скилл для сборки презентаций из библиотеки макетов «Акцент».

## Быстрый старт

```bash
pip install python-pptx pyyaml pillow
python scripts/accent.py nav
python scripts/accent.py build spec.yaml out.pptx
```

## Файлы

- `SKILL.md` — полная инструкция по 4 ролям
- `brief.md` — дизайн-бриф (синий #2196F3, белый фон)
- `TNF_slide_content_v02.md` — исходные 54 блока
- `scripts/` — движок сборки
- `assets/` — конфиги макетов

## Дизайн

- Фон: `#FFFFFF`
- Акцент: `#2196F3`
- Заливка графиков: `#E3F2FD`
- Текст: `#000000` / `#212121`
- Шрифт: Raleway
- Flat design, без теней

## Сборка

1. Подготовь `spec.yaml` со списком слайдов
2. `python scripts/accent.py check spec.yaml`
3. `python scripts/accent.py build spec.yaml out.pptx`
