import os
import sys
import time
import shutil
import re
import logging
import argparse
from datetime import datetime
from pathlib import Path
from functools import lru_cache

import requests
from docx import Document
from docx.shared import Pt
import aspose.words as aw

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# --- НАСТРОЙКИ ---
API_URL = "http://127.0.0.1:1234/v1/chat/completions"
MODEL_NAME = "qwen2.5-7b-instruct"
BATCH_SIZE = 6
MAX_WORKERS = 1
MAX_RETRIES = 3
TIMEOUT = 180
SEPARATOR = "<<<SEP>>>"

SYSTEM_PROMPT = (
    "Ты — корректор. Исправляй ТОЛЬКО орфографические и пунктуационные ошибки.\n"
    "СТРОГО ЗАПРЕЩЕНО: менять смысл, стиль, структуру, порядок слов, "
    "добавлять/удалять предложения.\n"
    "НЕ трогай сокращения (л., экз., ак. час, т.д., т.п.).\n"
    "Текст разделен маркером <<<SEP>>>. Верни текст в ТОМ ЖЕ формате: "
    "те же маркеры, тот же порядок фрагментов, ничего не пропуская.\n"
    "Верни ТОЛЬКО исправленный текст, без пояснений."
)


# ==================================================================
#                     ВЫБОР РЕЖИМА ОБРАБОТКИ
# ==================================================================
def ask_user_mode() -> bool:
    """True — пропускать таблицы, False — проверять всё."""
    print("\n" + "=" * 60)
    print("ВЫБОР РЕЖИМА ПРОВЕРКИ")
    print("=" * 60)
    print("В документе могут быть таблицы. Их проверка занимает")
    print("больше времени (иногда в 3–5 раз дольше), но даёт")
    print("более полную корректуру.\n")
    print("  1 — Проверять ВЕСЬ текст, включая таблицы (медленнее)")
    print("  2 — Проверять только основной текст, БЕЗ таблиц (быстрее)")
    print()
    while True:
        choice = input("Ваш выбор [1/2] (по умолчанию 2): ").strip() or "2"
        if choice == "1":
            print("→ Режим: полная проверка (с таблицами)\n")
            return False
        elif choice == "2":
            print("→ Режим: без таблиц\n")
            return True
        print("  Введите 1 или 2.")


def parse_cli_args():
    parser = argparse.ArgumentParser(description="AI-корректор Word-документов")
    parser.add_argument("input", nargs="?", default="ПРОГРАММА-2.docx")
    parser.add_argument("-o", "--output", default=None)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--skip-tables", action="store_true")
    group.add_argument("--check-tables", action="store_true")
    return parser.parse_args()


# ==================================================================
#                        ОБРАЩЕНИЕ К LLM
# ==================================================================
def send_to_ai_batch(texts: list[str], depth: int = 0) -> list[str]:
    if not texts:
        return []
    combined = f"\n{SEPARATOR}\n".join(texts)
    payload = {
        "model": MODEL_NAME,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": combined}
        ],
        "temperature": 0.0,
        "max_tokens": 1500,
    }
    for attempt in range(MAX_RETRIES):
        try:
            r = requests.post(API_URL, json=payload, timeout=TIMEOUT)
            r.raise_for_status()
            result = r.json()["choices"][0]["message"]["content"].strip()
            parts = re.split(rf"\s*{re.escape(SEPARATOR)}\s*", result)
            if len(parts) != len(texts):
                logger.warning(
                    f"Батч вернул {len(parts)} вместо {len(texts)}, дробим"
                )
                return _split_and_retry(texts, depth)
            return parts
        except requests.exceptions.Timeout:
            logger.warning(f"Таймаут батча (попытка {attempt + 1})")
            if len(texts) > 1 and attempt >= 1:
                return _split_and_retry(texts, depth)
            time.sleep(2)
        except Exception as e:
            logger.error(f"Ошибка LLM: {e}")
            time.sleep(2)
    if len(texts) > 1:
        return _split_and_retry(texts, depth)
    return texts


def _split_and_retry(texts: list[str], depth: int) -> list[str]:
    if depth > 3 or len(texts) == 1:
        return [send_to_ai_single(t) for t in texts]
    mid = len(texts) // 2
    left = send_to_ai_batch(texts[:mid], depth + 1)
    right = send_to_ai_batch(texts[mid:], depth + 1)
    return left + right


def send_to_ai_single(text: str) -> str:
    for attempt in range(MAX_RETRIES):
        try:
            r = requests.post(API_URL, json={
                "model": MODEL_NAME,
                "messages": [
                    {"role": "system", "content": SYSTEM_PROMPT.replace(SEPARATOR, "|")},
                    {"role": "user", "content": text}
                ],
                "temperature": 0.0,
                "max_tokens": 1000,
            }, timeout=TIMEOUT)
            r.raise_for_status()
            return r.json()["choices"][0]["message"]["content"].strip()
        except Exception:
            time.sleep(1)
    return text


# ==================================================================
#                          ФИЛЬТРЫ
# ==================================================================
def is_correctable(text: str) -> bool:
    if not text or len(text.strip()) < 5:
        return False
    if not re.search(r"[A-Za-zА-Яа-яЁё]", text):
        return False
    if re.fullmatch(r"[\d\s\W]+", text):
        return False
    return True


def is_heading(para) -> bool:
    style_name = (para.style.name or "") if para.style else ""
    if any(x in style_name for x in ('Heading', 'Заголовок', 'Title')):
        return True
    for run in para.runs:
        if run.bold:
            return True
        if run.font and run.font.size and run.font.size > Pt(14):
            return True
    return False


# ==================================================================
#                     ОБРАБОТКА ДОКУМЕНТА
# ==================================================================
def process_document(input_path: str, output_path: str,
                     skip_tables: bool = False) -> bool:
    """Пишет ИСПРАВЛЕННУЮ версию в output_path. Без трекинга."""
    try:
        shutil.copyfile(input_path, output_path)
        doc = Document(output_path)

        targets = []

        def collect(paras, scope=""):
            for p in paras:
                if is_heading(p):
                    continue
                txt = p.text
                if not is_correctable(txt):
                    continue
                targets.append((p, txt, scope))

        collect(doc.paragraphs, "body")

        total_tables = len(doc.tables)
        if skip_tables:
            logger.info(f"Таблицы ПРОПУЩЕНЫ (всего в документе: {total_tables})")
        else:
            for t_idx, table in enumerate(doc.tables):
                seen_cells = set()
                for row in table.rows:
                    for cell in row.cells:
                        if id(cell._element) in seen_cells:
                            continue
                        seen_cells.add(id(cell._element))
                        collect(cell.paragraphs, f"table{t_idx}")
            logger.info(f"Таблиц обработано: {total_tables}")

        logger.info(f"Параграфов к проверке: {len(targets)}")
        if not targets:
            doc.save(output_path)
            return True

        unique_texts = {}
        for _, txt, _ in targets:
            unique_texts.setdefault(txt, None)
        unique_list = list(unique_texts.keys())
        logger.info(f"Уникальных текстов: {len(unique_list)}")

        batches = [unique_list[i:i + BATCH_SIZE]
                   for i in range(0, len(unique_list), BATCH_SIZE)]
        logger.info(f"Батчей: {len(batches)}")

        results = {}
        for i, b in enumerate(batches):
            logger.info(f"  батч {i + 1}/{len(batches)} ({len(b)} эл.)")
            corrected = send_to_ai_batch(b)
            for orig, corr in zip(b, corrected):
                results[orig] = corr

        applied = 0
        for para, orig, _ in targets:
            corrected = results.get(orig, orig)
            if corrected.strip() == orig.strip():
                continue
            # сохраняем формат первого run
            first_run = para.runs[0] if para.runs else None
            bold = first_run.bold if first_run else None
            italic = first_run.italic if first_run else None
            underline = first_run.underline if first_run else None
            font_name = first_run.font.name if first_run else None
            font_size = first_run.font.size if first_run else None

            para.clear()
            new_run = para.add_run(corrected)
            if bold is not None:
                new_run.bold = bold
            if italic is not None:
                new_run.italic = italic
            if underline is not None:
                new_run.underline = underline
            if font_name:
                new_run.font.name = font_name
            if font_size:
                new_run.font.size = font_size
            applied += 1

        doc.save(output_path)
        logger.info(f"Применено правок: {applied}")
        return True
    except Exception as e:
        logger.error(f"Ошибка process_document: {e}", exc_info=True)
        return False


# ==================================================================
#                РЕЖИМ РЕЦЕНЗИРОВАНИЯ (через Aspose)
# ==================================================================
def apply_track_changes(original_path: str, modified_path: str) -> bool:
    """
    Aspose.Words сравнивает два документа и сохраняет результат
    в modified_path С ТРЕКИНГОМ. Word открывает без проблем.
    """
    try:
        logger.info("Aspose: загрузка документов...")
        doc_original = aw.Document(original_path)
        doc_modified = aw.Document(modified_path)

        options = aw.comparing.CompareOptions()
        options.ignore_headers_and_footers = True
        options.ignore_formatting = False
        options.ignore_comments = True
        options.ignore_tables = False       # таблицы тоже сравниваем, если их правили
        options.ignore_case_changes = False
        options.ignore_fields = True
        options.granularity = aw.comparing.Granularity.WORD_LEVEL

        logger.info("Aspose: сравнение...")
        doc_original.compare(
            doc_modified,
            "AI Корректор",
            datetime.now(),
            options
        )

        logger.info("Aspose: сохранение...")
        doc_original.save(modified_path)
        return True
    except Exception as e:
        logger.error(f"Aspose compare error: {e}", exc_info=True)
        return False


# ==================================================================
#               УДАЛЕНИЕ ВОДЯНОГО ЗНАКА ASPOSE
# ==================================================================
def remove_aspose_watermarks(docx_path: str) -> bool:
    """
    Aspose в trial-режиме добавляет текстовый/колонтитульный водяной знак.
    Убираем его. Работает и с <w:t>, и с колонтитулами.
    """
    try:
        logger.info("Очистка водяных знаков Aspose...")
        doc = Document(docx_path)

        watermark_patterns = [
            "Aspose.Words", "evaluation copy", "temporary-license",
            "To remove all limitations", "https://products.aspose.com",
            "This document was truncated", "Evaluation Mode",
            "Evaluation Only", "Created with Aspose",
        ]
        patterns = [re.compile(re.escape(p), re.IGNORECASE) for p in watermark_patterns]

        def has_wm(text: str) -> bool:
            t = (text or "").strip()
            return any(p.search(t) for p in patterns)

        removed = 0

        # --- Основное тело ---
        for p in list(doc.paragraphs):
            if has_wm(p.text):
                p._element.getparent().remove(p._element)
                removed += 1

        # --- Таблицы ---
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in list(cell.paragraphs):
                        if has_wm(p.text):
                            p._element.getparent().remove(p._element)
                            removed += 1

        # --- Колонтитулы ---
        for section in doc.sections:
            for hf in (section.header, section.footer,
                       section.first_page_header, section.first_page_footer,
                       section.even_page_header, section.even_page_footer):
                if hf is None:
                    continue
                for p in list(hf.paragraphs):
                    if has_wm(p.text):
                        p._element.getparent().remove(p._element)
                        removed += 1
                # Aspose иногда кладёт watermark в shape/картинку —
                # удаляем пустые параграфы с рисунками только если рядом есть watermark-текст
                # (не трогаем пользовательские картинки)

        doc.save(docx_path)
        logger.info(f"Удалено watermark-элементов: {removed}")
        return True
    except Exception as e:
        logger.warning(f"Не удалось удалить watermark: {e}")
        return False


# ==================================================================
#                              MAIN
# ==================================================================
if __name__ == "__main__":
    args = parse_cli_args()

    INPUT_FILE = args.input
    if args.output:
        OUTPUT_FILE = args.output
    else:
        OUTPUT_FILE = f"{Path(INPUT_FILE).stem}_РЕЦЕНЗИРОВАНИЕ.docx"

    # --- Режим таблиц ---
    if args.skip_tables:
        skip_tables = True
        print("→ Режим (CLI): БЕЗ таблиц")
    elif args.check_tables:
        skip_tables = False
        print("→ Режим (CLI): полная проверка")
    else:
        skip_tables = ask_user_mode()

    print("=" * 60)
    print("AI КОРРЕКТОР (Variant A: Aspose track changes)")
    print("=" * 60)
    print(f"Вход:   {INPUT_FILE}")
    print(f"Выход:  {OUTPUT_FILE}")
    print(f"Таблицы: {'НЕ проверяем' if skip_tables else 'проверяем'}")
    print("=" * 60)

    temp_corrected = f"{Path(INPUT_FILE).stem}_TEMP_CORRECTED.docx"

    try:
        t0 = time.time()

        # --- ШАГ 1: правим текст ---
        print("\n[1/3] Обработка текста нейросетью...")
        if not process_document(INPUT_FILE, temp_corrected,
                                skip_tables=skip_tables):
            raise Exception("Ошибка обработки текста")

        # --- ШАГ 2: накладываем трекинг Aspose ---
        print("\n[2/3] Наложение режима рецензирования (Aspose)...")
        if not apply_track_changes(INPUT_FILE, temp_corrected):
            logger.warning("Aspose не сработал — сохраняем без трекинга")
            shutil.copyfile(temp_corrected, OUTPUT_FILE)
        else:
            # --- ШАГ 3: чистим watermark ---
            print("\n[3/3] Очистка технических меток Aspose...")
            remove_aspose_watermarks(temp_corrected)
            shutil.move(temp_corrected, OUTPUT_FILE)

        # Подчищаем temp, если остался
        if os.path.exists(temp_corrected):
            os.remove(temp_corrected)

        print(f"\n✅ ГОТОВО за {time.time() - t0:.0f} сек.")
        print(f"   Результат: {OUTPUT_FILE}")
    except Exception as e:
        print(f"\n❌ ОШИБКА: {e}")
        if os.path.exists(temp_corrected):
            os.remove(temp_corrected)