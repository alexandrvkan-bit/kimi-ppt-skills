# accent-skill

Система презентаций "Акцент" для Kimi Code. Корпоративный дизайн: синий (#2196F3) + белый, шрифт Raleway.

## Быстрый старт

```bash
# 1. Клонируй репозиторий
git clone https://github.com/alexandrvkan-bit/kimi-ppt-skills.git
cd kimi-ppt-skills/accent-skill

# 2. Установи зависимости
pip install python-pptx>=0.6.21 pyyaml Pillow

# 3. Сгенерируй библиотеку шаблонов
python scripts/generate_library.py

# 4. Собери презентацию
python scripts/accent.py build spec.yaml out.pptx
```

## Структура

- `SKILL.md` — полная инструкция по системе
- `brief.md` — дизайн-бриф с точными цветами
- `TNF_slide_content_v02.md` — контент 54 блоков ТНФ
- `scripts/accent.py` — движок сборки
- `scripts/generate_library.py` — генератор библиотеки шаблонов
- `assets/manifest.json` — описание 11 базовых макетов
- `assets/library.pptx` — шаблоны слайдов (генерируется скриптом)
- `assets/fonts/` — шрифт Raleway

## Макеты

| Код | Назначение |
|-----|------------|
| A001 | Титул |
| A002 | Тезис по центру |
| A003 | Заключение |
| B001 | Список пунктов |
| C001 | Таймлайн |
| D001 | 3 метрики |
| D004 | Столбчатая диаграмма |
| D007 | Таблица |
| M001 | SWOT |
| F001 | Оргструктура |
| P001 | Статус-панель |

## Сборка презентации ТНФ

1. Отредактируй `spec.yaml` — укажи макет и контент для каждого слайда
2. Запусти: `python scripts/accent.py build spec.yaml TNF.pptx`
3. Готово.
