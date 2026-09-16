import os
import time
import shutil
import tempfile
import logging
from datetime import datetime
from pathlib import Path

import requests
from docx import Document
from docx.shared import Pt
from difflib import SequenceMatcher
import aspose.words as aw

# --- НАСТРОЙКИ ---
API_URL = "http://127.0.0.1:1234/v1/chat/completions"  # LM Studio API
MODEL_NAME = "qwen2.5-7b-instruct"  # Ваша модель в LM Studio
REQUEST_DELAY = 1.0

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
    """Пропускаем заголовки, таблицы и пустые абзацы (как в оригинальном проекте)"""
    if not para.text.strip():
        return True
    style_name = para.style.name or ""
    if any(x in style_name for x in ['Heading', 'Заголовок', 'Title', 'Table']):
        return True
    # Проверка на жирный текст или крупный шрифт (признак заголовка)
    if para.runs and (para.runs[0].bold or (para.runs[0].font.size and para.runs[0].font.size > Pt(14))):
        return True
    return False


def process_document(input_path: str, output_path: str) -> bool:
    """Основная обработка: замена текста с сохранением форматирования"""
    try:
        shutil.copyfile(input_path, output_path)
        doc = Document(output_path)

        for i, para in enumerate(doc.paragraphs):
            if is_special_paragraph(para):
                continue

            original_runs = list(para.runs)
            corrected_text = send_to_ai(para.text)

            if corrected_text.strip() == para.text.strip():
                continue

            # Восстановление форматирования (логика из apply_formatting)
            para.clear()
            matcher = SequenceMatcher(None, para.text, corrected_text)

            # Упрощенное восстановление форматирования по словам
            words_orig = para.text.split()
            runs_map = {}
            pos = 0
            for run in original_runs:
                for char in run.text:
                    runs_map[pos] = run
                    pos += 1

            current_run = None
            current_text = ""

            for opcode in matcher.get_opcodes():
                tag, i1, i2, j1, j2 = opcode
                segment = corrected_text[j1:j2]

                if tag == 'equal':
                    ref_idx = i1 if i1 < len(runs_map) else len(runs_map) - 1
                    target_run = runs_map.get(ref_idx, original_runs[0])
                else:
                    target_run = original_runs[0] if original_runs else None

                if target_run and (current_run is None or
                                   (target_run.font.name != current_run.font.name or
                                    target_run.bold != current_run.bold)):
                    if current_run:
                        current_run.text = current_text
                    current_run = para.add_run("")
                    current_run.font.name = target_run.font.name
                    current_run.font.size = target_run.font.size
                    current_run.bold = target_run.bold
                    current_run.italic = target_run.italic
                    current_text = ""

                current_text += segment

            if current_run:
                current_run.text = current_text

            logger.info(f"Обработан параграф {i + 1}")
            time.sleep(REQUEST_DELAY)

        doc.save(output_path)
        return True
    except Exception as e:
        logger.error(f"Ошибка обработки: {e}", exc_info=True)
        return False


def apply_track_changes(original_path: str, modified_path: str):
    """Создание нативных правок Word через Aspose.Words (ядро вашего проекта)"""
    logger.info("Применение режима рецензирования...")
    doc_original = aw.Document(original_path)
    doc_modified = aw.Document(modified_path)

    options = aw.comparing.CompareOptions()
    options.ignore_headers_and_footers = True  # Не трогаем колонтитулы
    options.ignore_formatting = False  # Сравниваем форматирование

    # Сравнение создает теги <w:ins> и <w:del>
    doc_original.compare(doc_modified, "AI Корректор", datetime.now(), options)

    # Сохраняем результат поверх исправленного файла
    doc_original.save(modified_path)
    logger.info("Режим рецензирования успешно применен")


def clean_watermarks(docx_path: str):
    """Удаление водяных знаков Aspose (из вашего original кода)"""
    try:
        doc = Document(docx_path)
        patterns = ["Aspose.Words", "evaluation copy", "temporary-license"]

        def has_watermark(p):
            return any(pat in p.text for pat in patterns)

        # Удаляем из основного тела
        for p in [p for p in doc.paragraphs if has_watermark(p)]:
            p._element.getparent().remove(p._element)

        # Удаляем из таблиц
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in [p for p in cell.paragraphs if has_watermark(p)]:
                        p._element.getparent().remove(p._element)

        doc.save(docx_path)
    except Exception as e:
        logger.warning(f"Не удалось удалить водяные знаки: {e}")


if __name__ == "__main__":
    INPUT_FILE = "../ПРОГРАММА-1.docx"
    OUTPUT_FILE = "../ПРОГРАММА-1_РЕЦЕНЗИРОВАНИЕ.docx"

    # Создаем временный файл для промежуточных правок
    temp_corrected = tempfile.mktemp(suffix=".docx")

    print("1. Обработка текста нейросетью...")
    if process_document(INPUT_FILE, temp_corrected):
        print("2. Применение нативного режима рецензирования (Aspose)...")
        apply_track_changes(INPUT_FILE, temp_corrected)

        print("3. Очистка технических меток...")
        clean_watermarks(temp_corrected)

        # Финальное сохранение
        shutil.move(temp_corrected, OUTPUT_FILE)
        print(f"✅ Готово! Откройте '{OUTPUT_FILE}' в Word.")
        print("   Перейдите во вкладку 'Рецензирование' -> 'Исправления'.")
    else:
        print("❌ Ошибка при обработке документа")