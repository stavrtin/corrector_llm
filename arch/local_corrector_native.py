import os
import time
import shutil
import logging
from datetime import datetime
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from difflib import SequenceMatcher
import requests

# --- НАСТРОЙКИ ---
API_URL = "http://127.0.0.1:1234/v1/chat/completions"  # LM Studio API
MODEL_NAME = "qwen2.5-7b-instruct"  # Ваша модель
REQUEST_DELAY = 1.0  # Задержка между запросами

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

SYSTEM_PROMPT = (
    "ТВОЯ ГЛАВНАЯ ЗАДАЧА - ИСПРАВЛЯТЬ ТОЛЬКО ОРФОГРАФИЧЕСКИЕ И ПУНКТУАЦИОННЫЕ ОШИБКИ.\n"
    "1. НЕ МЕНЯЙ СМЫСЛ, СТИЛЬ И СТРУКТУРУ ТЕКСТА\n"
    "2. НЕ ДОБАВЛЯЙ НОВЫЕ СЛОВА И ИНФОРМАЦИЮ\n"
    "3. СОХРАНЯЙ СПЕЦИАЛЬНОЕ ФОРМАТИРОВАНИЕ И ТАБУЛЯЦИИ\n"
    "4. НЕ ТРОГАЙ СОКРАЩЕНИЯ (л., экз., ак. час и т.д.)\n"
    "5. ВЕРНИ ТОЛЬКО ИСПРАВЛЕННЫЙ ТЕКСТ БЕЗ КОММЕНТАРИЕВ"
)


def send_to_ai(text: str) -> str:
    """Отправка текста в LM Studio"""
    try:
        response = requests.post(API_URL, json={
            "model": MODEL_NAME,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": text}
            ],
            "temperature": 0.1,
            "max_tokens": 4000
        }, timeout=60)
        response.raise_for_status()
        return response.json()["choices"][0]["message"]["content"]
    except Exception as e:
        logger.error(f"Ошибка LLM: {e}")
        return text


def is_special_paragraph(para) -> bool:
    """Пропускаем заголовки, таблицы и пустые абзацы"""
    if not para.text.strip():
        return True
    style_name = para.style.name or ""
    if any(x in style_name for x in ['Heading', 'Заголовок', 'Title', 'Table']):
        return True
    if para.runs and (para.runs[0].bold or (para.runs[0].font.size and para.runs[0].font.size > Pt(14))):
        return True
    return False


def create_track_change_xml(text, change_type="ins"):
    """
    Создает XML-элемент <w:ins> или <w:del> для нативного режима рецензирования.
    Это заменяет функцию comparator из Aspose.Words.
    """
    elem = OxmlElement(f'w:{change_type}')
    # Уникальный ID ревизии
    rev_id = str(hash((text, change_type, datetime.now())) % 100000)
    elem.set(qn('w:id'), rev_id)
    elem.set(qn('w:author'), "AI Корректор")
    elem.set(qn('w:date'), datetime.now().strftime('%Y-%m-%dT%H:%M:%SZ'))

    run = OxmlElement('w:r')
    t_elem = OxmlElement('w:t')
    t_elem.text = text
    run.append(t_elem)
    elem.append(run)

    return elem


def apply_native_track_changes(para, original_text, corrected_text):
    """
    Внедряет нативные правки Word (<w:ins>/<w:del>) прямо в параграф.
    Работает без Aspose.Words, используя только python-docx и lxml.
    """
    if original_text == corrected_text:
        return

    matcher = SequenceMatcher(None, original_text, corrected_text)

    # Очищаем параграф от старого контента, но сохраняем его свойства (стили, отступы)
    for run in para.runs:
        run.text = ""

    # Вставляем изменения непосредственно в XML параграфа (_p)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            # Совпадающий текст добавляем как обычный run
            new_run = para.add_run(original_text[i1:i2])

        elif tag in ('replace', 'delete'):
            # Удаляемый фрагмент оборачиваем в <w:del>
            del_xml = create_track_change_xml(original_text[i1:i2], change_type="del")
            para._p.append(del_xml)

        if tag in ('replace', 'insert'):
            # Вставляемый фрагмент оборачиваем в <w:ins>
            ins_xml = create_track_change_xml(corrected_text[j1:j2], change_type="ins")
            para._p.append(ins_xml)


def process_document(input_path: str, output_path: str) -> bool:
    """Основная функция обработки документа"""
    logger.info("Начало обработки...")
    try:
        shutil.copyfile(input_path, output_path)
        doc = Document(output_path)

        for i, para in enumerate(doc.paragraphs):
            if is_special_paragraph(para):
                continue

            original_text = para.text
            logger.info(f"Обработка параграфа {i + 1}: {original_text[:60]}...")

            corrected_text = send_to_ai(original_text)

            # Применяем нативные правки вместо простой замены
            apply_native_track_changes(para, original_text, corrected_text)

            time.sleep(REQUEST_DELAY)

        doc.save(output_path)
        logger.info(f"Документ сохранен: {output_path}")
        return True

    except Exception as e:
        logger.error(f"Критическая ошибка: {e}", exc_info=True)
        return False


if __name__ == "__main__":
    INPUT_FILE = "../ПРОГРАММА-1.docx"
    OUTPUT_FILE = "../ПРОГРАММА-1_РЕЦЕНЗИРОВАНИЕ.docx"

    print(" Запуск локального корректора с нативным режимом рецензирования...")
    print(f"📄 Входной файл: {INPUT_FILE}")
    print(f" Выходной файл: {OUTPUT_FILE}")

    if os.path.exists(INPUT_FILE):
        if process_document(INPUT_FILE, OUTPUT_FILE):
            print("\n✅ ГОТОВО! Откройте файл в MS Word.")
            print("   Перейдите во вкладку 'Рецензирование' -> 'Исправления'.")
            print("   Вы увидите правки от автора 'AI Корректор' с возможностью принять/отклонить.")
        else:
            print("\n❌ Ошибка при обработке.")
    else:
        print(f"\n❌ Файл '{INPUT_FILE}' не найден в текущей директории.")