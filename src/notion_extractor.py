"""Извлечение данных курса из Notion API.

Скрипт подключается к Notion API и извлекает структуру курса:
- Название, описание
- Модули и уроки
- Академические часы
- Свойства страниц (часы, формы контроля и т.д.)

Использование:
    python notion_extractor.py <notion_page_url>
    python notion_extractor.py <notion_page_url> --output output/course_data.json
"""

import json
import re
import sys
import argparse
from pathlib import Path

try:
    from notion_client import Client
    HAS_NOTION = True
except ImportError:
    HAS_NOTION = False


# Алиасы для поиска свойств Notion (регистронезависимый поиск)
HOURS_TOTAL_ALIASES = [
    "Часы", "Академические часы", "Количество часов", "Hours",
    "часы", "академические часы", "кол-во часов", "Кол-во часов",
    "Всего часов", "всего часов",
]
HOURS_LECTURE_ALIASES = [
    "Лекции", "Часы лекций", "Лекционные часы", "Hours lecture",
    "лекции", "часы лекций", "Лекц", "лекц",
]
HOURS_PRACTICE_ALIASES = [
    "Практика", "Часы практики", "Практические часы", "Hours practice",
    "практика", "часы практики", "Практ", "практ",
]
HOURS_CONTROL_ALIASES = [
    "Контроль", "Часы контроля", "Контрольные часы", "Hours control",
    "контроль", "часы контроля", "Сам. работа", "сам. работа",
]
CONTROL_TYPE_ALIASES = [
    "Форма контроля", "Тип контроля", "Control type",
    "форма контроля", "тип контроля", "Вид контроля", "вид контроля",
]
DESCRIPTION_ALIASES = [
    "Описание", "Описание урока", "Описание темы", "Description",
    "описание", "описание урока", "описание темы", "Аннотация", "аннотация",
]


def load_env(env_path=None):
    """Загрузить NOTION_API_KEY из .env файла."""
    if env_path is None:
        env_path = Path(__file__).parent.parent.parent / ".env"

    if not env_path.exists():
        return None

    for line in env_path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if line.startswith('#') or '=' not in line:
            continue
        key, value = line.split('=', 1)
        key = key.strip()
        value = value.strip()
        if key == 'NOTION_API_KEY' and value:
            return value
    return None


def extract_page_id(url_or_id):
    """Извлечь ID страницы Notion из URL или вернуть как есть."""
    # UUID формат
    uuid_pattern = r'([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})'
    match = re.search(uuid_pattern, url_or_id, re.IGNORECASE)
    if match:
        return match.group(1)

    # URL формат: https://www.notion.so/workspace/Page-Title-page_id
    # или https://www.notion.so/page_id
    url_pattern = r'notion\.so/.*?([0-9a-f]{32})'
    match = re.search(url_pattern, url_or_id, re.IGNORECASE)
    if match:
        hex_id = match.group(1)
        return f"{hex_id[:8]}-{hex_id[8:12]}-{hex_id[12:16]}-{hex_id[16:20]}-{hex_id[20:]}"

    return url_or_id


def find_property(props, aliases):
    """Найти свойство в Notion-объекте по алиасам."""
    for alias in aliases:
        if alias in props:
            prop = props[alias]
            ptype = prop.get('type')
            if ptype == 'number' and prop.get('number') is not None:
                return prop['number']
            elif ptype == 'rich_text':
                texts = prop.get('rich_text', [])
                return ''.join(t.get('plain_text', '') for t in texts)
            elif ptype == 'select':
                sel = prop.get('select')
                return sel.get('name', '') if sel else ''
            elif ptype == 'title':
                texts = prop.get('title', [])
                return ''.join(t.get('plain_text', '') for t in texts)
    return None


def get_block_text(block):
    """Извлечь текст из блока Notion."""
    btype = block.get('type', '')
    if btype in ('paragraph', 'heading_1', 'heading_2', 'heading_3',
                  'bulleted_list_item', 'numbered_list_item', 'toggle',
                  'quote', 'callout'):
        rich_text = block.get(btype, {}).get('rich_text', [])
        return ''.join(t.get('plain_text', '') for t in rich_text)
    return ''


def extract_course_data(page_id, api_key):
    """Извлечь полные данные курса из Notion."""
    client = Client(auth=api_key)

    # Получить главную страницу
    page = client.pages.retrieve(page_id)
    course_data = {
        'title': '',
        'url': page.get('url', ''),
        'description': '',
        'modules': [],
        'raw_properties': {},
    }

    # Извлечь название
    props = page.get('properties', {})
    for key, prop in props.items():
        if prop.get('type') == 'title':
            texts = prop.get('title', [])
            course_data['title'] = ''.join(t.get('plain_text', '') for t in texts)
            break

    # Извлечь свойства
    for key, prop in props.items():
        ptype = prop.get('type')
        if ptype == 'number':
            course_data['raw_properties'][key] = prop.get('number')
        elif ptype == 'rich_text':
            texts = prop.get('rich_text', [])
            course_data['raw_properties'][key] = ''.join(t.get('plain_text', '') for t in texts)
        elif ptype == 'select':
            sel = prop.get('select')
            course_data['raw_properties'][key] = sel.get('name', '') if sel else None
        elif ptype == 'multi_select':
            course_data['raw_properties'][key] = [s.get('name', '') for s in prop.get('multi_select', [])]

    # Получить дочерние блоки (модули и уроки)
    children = client.blocks.children.list(page_id)

    current_module = None
    module_number = 0

    for block in children.get('results', []):
        btype = block.get('type', '')
        text = get_block_text(block)

        if btype == 'heading_3' and text:
            # Новый модуль
            module_number += 1
            current_module = {
                'number': module_number,
                'title': text,
                'lessons': [],
                'description': '',
            }
            course_data['modules'].append(current_module)

        elif btype == 'heading_2' and text and not current_module:
            # Может быть вводный модуль
            module_number += 1
            current_module = {
                'number': module_number,
                'title': text,
                'lessons': [],
                'description': '',
            }
            course_data['modules'].append(current_module)

        # Дочерние страницы (уроки)
        if btype == 'child_page':
            child_title = block.get('child_page', {}).get('title', '')
            lesson = {
                'title': child_title,
                'id': block.get('id', ''),
                'hours_total': 2,
                'hours_lecture': 1,
                'hours_practice': 1,
                'hours_control': 0,
                'control_type': 'Практическая работа',
                'description': '',
            }

            # Попробовать извлечь свойства из дочерней страницы
            try:
                child_page = client.pages.retrieve(block['id'])
                child_props = child_page.get('properties', {})
                hours = find_property(child_props, HOURS_TOTAL_ALIASES)
                if hours:
                    lesson['hours_total'] = int(hours)
                hours_lec = find_property(child_props, HOURS_LECTURE_ALIASES)
                if hours_lec:
                    lesson['hours_lecture'] = int(hours_lec)
                hours_prac = find_property(child_props, HOURS_PRACTICE_ALIASES)
                if hours_prac:
                    lesson['hours_practice'] = int(hours_prac)
                hours_ctrl = find_property(child_props, HOURS_CONTROL_ALIASES)
                if hours_ctrl:
                    lesson['hours_control'] = int(hours_ctrl)
                ctrl_type = find_property(child_props, CONTROL_TYPE_ALIASES)
                if ctrl_type:
                    lesson['control_type'] = str(ctrl_type)
                desc = find_property(child_props, DESCRIPTION_ALIASES)
                if desc:
                    lesson['description'] = str(desc)
            except Exception as e:
                print(f"  ⚠ Не удалось извлечь свойства урока '{child_title}': {e}", file=sys.stderr)

            # Получить описание из содержимого дочерней страницы
            try:
                child_blocks = client.blocks.children.list(block['id'])
                desc_parts = []
                for cb in child_blocks.get('results', [])[:5]:  # Первые 5 блоков для описания
                    cb_text = get_block_text(cb)
                    if cb_text:
                        desc_parts.append(cb_text)
                if desc_parts and not lesson['description']:
                    lesson['description'] = '\n'.join(desc_parts)
            except Exception:
                pass

            if current_module:
                lesson_id = f"{current_module['number']}.{len(current_module['lessons']) + 1}"
                lesson['id'] = lesson_id
                current_module['lessons'].append(lesson)
            else:
                # Урок без модуля — создаём модуль "Вводный"
                module_number += 1
                current_module = {
                    'number': module_number,
                    'title': 'Вводный модуль',
                    'lessons': [],
                    'description': '',
                }
                lesson['id'] = f"{module_number}.1"
                current_module['lessons'].append(lesson)
                course_data['modules'].append(current_module)

    return course_data


def main():
    parser = argparse.ArgumentParser(description="Извлечение данных курса из Notion")
    parser.add_argument("page_url", help="URL или ID страницы Notion")
    parser.add_argument("--output", "-o", default=None, help="Путь к выходному JSON-файлу")
    parser.add_argument("--env", "-e", default=None, help="Путь к .env файлу")
    args = parser.parse_args()

    if not HAS_NOTION:
        print("ERROR: notion-client not installed. Run: pip install notion-client")
        sys.exit(1)

    api_key = load_env(args.env)
    if not api_key:
        print("ERROR: NOTION_API_KEY not found in .env file.")
        sys.exit(1)

    page_id = extract_page_id(args.page_url)
    print(f"Extracting from Notion: {args.page_url}")
    print(f"  Page ID: {page_id}")

    course_data = extract_course_data(page_id, api_key)

    # Stats
    total_lessons = sum(len(m['lessons']) for m in course_data['modules'])
    total_hours = sum(
        lesson['hours_total']
        for module in course_data['modules']
        for lesson in module['lessons']
    )
    print(f"\nExtracted:")
    print(f"  Title: {course_data['title']}")
    print(f"  Modules: {len(course_data['modules'])}")
    print(f"  Lessons: {total_lessons}")
    print(f"  Academic hours: {total_hours}")

    # Save JSON
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = Path(__file__).parent.parent / "output" / "course_data.json"

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(course_data, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f"  Data saved: {output_path}")

    return course_data


if __name__ == "__main__":
    main()