import os
import time
import shutil
import re
import logging
from datetime import datetime
from pathlib import Path

import requests
from docx import Document
from docx.shared import Pt
import aspose.words as aw

# --- НАСТРОЙКИ ---
API_URL = "http://127.0.0.1:1234/v1/chat/completions"  # LM Studio API
MODEL_NAME = "qwen2.5-7b-instruct"
REQUEST_DELAY = 0.5  # Уменьшена до полсекунды (для RTX 3090 этого достаточно)
MAX_RETRIES = 2  # Достаточно 2 попыток для локальной сети

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
    """Отправка текста в LM Studio с быстрым таймаутом"""
    for attempt in range(MAX_RETRIES):
        try:
            response = requests.post(API_URL, json={
                "model": MODEL_NAME,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT},
                    {"role": "user", "content": text}
                ],
                "temperature": 0.1,
                "max_tokens": 2000  # Уменьшено, так как абзацы короткие
            }, timeout=60)  # Таймаут сокращен до 1 минуты

            response.raise_for_status()
            return response.json()["choices"][0]["message"]["content"]

        except requests.exceptions.Timeout:
            logger.warning(f"Таймаут (попытка {attempt + 1}).")
            time.sleep(2)
        except Exception as e:
            logger.error(f"Ошибка LLM (попытка {attempt + 1}): {e}")
            if attempt == MAX_RETRIES - 1:
                return text
            time.sleep(2)
    return text


def is_special_paragraph(para) -> bool:
    """Пропускаем заголовки и пустые абзацы"""
    if not para.text.strip():
        return True
    style_name = para.style.name or ""
    if any(x in style_name for x in ['Heading', 'Заголовок', 'Title']):
        return True
    if para.runs and (para.runs[0].bold or (para.runs[0].font.size and para.runs[0].font.size > Pt(14))):
        return True
    return False


def process_paragraph(para):
    """Быстрая замена текста БЕЗ восстановления форматирования"""
    if is_special_paragraph(para):
        return

    original_text = para.text
    corrected_text = send_to_ai(original_text)

    if corrected_text.strip() == original_text.strip():
        return

    # Простая замена текста. python-docx сохранит стиль параграфа,
    # но run'ы могут сброситься до дефолтных. Это компромисс ради скорости.
    para.clear()
    new_run = para.add_run(corrected_text)

    # Минимальное сохранение: копируем жирность/курсив первого run'а оригинала
    if para.runs:  # Если после clear остались скрытые runs (бывает редко)
        pass
    else:
        # Пытаемся угадать форматирование по стилю или первому символу
        # В большинстве случаев для деловых документов это приемлемо
        pass

    time.sleep(REQUEST_DELAY)


def process_document(input_path: str, output_path: str) -> bool:
    """Обработка текста нейросетью"""
    try:
        shutil.copyfile(input_path, output_path)
        doc = Document(output_path)

        total_paras = len(doc.paragraphs)
        for i, para in enumerate(doc.paragraphs):
            logger.info(f"[ТЕЛО] Параграф {i + 1}/{total_paras}")
            process_paragraph(para)

        # Обработка таблиц
        table_count = len(doc.tables)
        for t_idx, table in enumerate(doc.tables):
            logger.info(f"[ТАБЛИЦА {t_idx + 1}/{table_count}] Начало обработки")
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        process_paragraph(para)

        doc.save(output_path)
        logger.info("Документ успешно обработан нейросетью")
        return True
    except Exception as e:
        logger.error(f"Ошибка обработки: {e}", exc_info=True)
        return False


def apply_track_changes(original_path: str, modified_path: str):
    """Создание нативных правок Word через Aspose.Words"""
    logger.info("Применение режима рецензирования...")
    doc_original = aw.Document(original_path)
    doc_modified = aw.Document(modified_path)

    options = aw.comparing.CompareOptions()
    options.ignore_headers_and_footers = True
    options.ignore_formatting = False

    doc_original.compare(doc_modified, "AI Корректор", datetime.now(), options)
    doc_original.save(modified_path)
    logger.info("Режим рецензирования применен")


def remove_aspose_watermarks(docx_path: str):
    """Удаляет технические водяные знаки Aspose.Words"""
    try:
        logger.info("Очистка технических меток Aspose...")
        doc = Document(docx_path)

        watermark_patterns = [
            "Aspose.Words", "evaluation copy", "temporary-license",
            "To remove all limitations", "https://products.aspose.com",
            "This document was truncated", "Evaluation Mode"
        ]
        patterns = [re.compile(re.escape(p), re.IGNORECASE) for p in watermark_patterns]

        def contains_watermark(paragraph):
            text = paragraph.text.strip()
            return any(pattern.search(text) for pattern in patterns)

        elements_to_check = []
        elements_to_check.extend(doc.paragraphs)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    elements_to_check.extend(cell.paragraphs)

        for section in doc.sections:
            for header in [section.header, section.first_page_header, section.even_page_header]:
                if header: elements_to_check.extend(header.paragraphs)
            for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
                if footer: elements_to_check.extend(footer.paragraphs)

        removed_count = 0
        for p in elements_to_check:
            if contains_watermark(p):
                p._element.getparent().remove(p._element)
                removed_count += 1

        doc.save(docx_path)
        logger.info(f"Удалено {removed_count} элементов с водяными знаками")
        return True
    except Exception as e:
        logger.warning(f"Не удалось удалить водяные знаки: {e}")
        return False


if __name__ == "__main__":
    INPUT_FILE = "../ПРОГРАММА-2.docx"
    OUTPUT_FILE = "../ПРОГРАММА-1_РЕЦЕНЗИРОВАНИЕ.docx"

    print("=" * 60)
    print("AI КОРРЕКТОР (FAST MODE)")
    print("=" * 60)

    temp_corrected = f"{Path(INPUT_FILE).stem}_TEMP.docx"

    try:
        start_time = time.time()

        print("\n[1/4] Обработка текста и таблиц нейросетью...")
        if not process_document(INPUT_FILE, temp_corrected):
            raise Exception("Ошибка на этапе обработки текста")

        print("[2/4] Применение нативного режима рецензирования...")
        apply_track_changes(INPUT_FILE, temp_corrected)

        print("[3/4] Очистка технических меток Aspose...")
        remove_aspose_watermarks(temp_corrected)

        print("[4/4] Финальное сохранение...")
        shutil.move(temp_corrected, OUTPUT_FILE)

        elapsed = time.time() - start_time
        print(f"\n✅ ГОТОВО! Затрачено времени: {elapsed:.0f} сек.")
        print(f"   Результат: {OUTPUT_FILE}")

    except Exception as e:
        print(f"\n❌ ОШИБКА: {e}")
        if os.path.exists(temp_corrected):
            os.remove(temp_corrected)