from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from difflib import SequenceMatcher
import copy


def apply_native_track_changes(para, original_text, corrected_text):
    """
    Создает нативные правки Word (Track Changes) внутри параграфа.
    Удаляет старый текст и вставляет новый с тегами <w:ins>/<w:del>.
    """
    if original_text == corrected_text:
        return

    # Очищаем параграф от старого контента, но сохраняем его свойства
    for run in para.runs:
        run.text = ""

    # Создаем базовый run для копирования стилей
    base_run = para.add_run("")

    matcher = SequenceMatcher(None, original_text, corrected_text)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            # Совпадающий текст добавляем как есть
            new_run = para.add_run(original_text[i1:i2])
        elif tag in ('replace', 'delete'):
            # Удаляемый текст оборачиваем в <w:del>
            del_elem = OxmlElement('w:del')
            del_elem.set(qn('w:id'), str(hash(original_text[i1:i2]) % 100000))
            del_elem.set(qn('w:author'), "AI Корректор")
            del_elem.set(qn('w:date'), "2026-09-16T10:00:00Z")

            ins_elem = OxmlElement('w:r')
            t_elem = OxmlElement('w:t')
            t_elem.text = original_text[i1:i2]
            ins_elem.append(t_elem)
            del_elem.append(ins_elem)

            para._p.append(del_elem)

        if tag in ('replace', 'insert'):
            # Вставляемый текст оборачиваем в <w:ins>
            ins_elem = OxmlElement('w:ins')
            ins_elem.set(qn('w:id'), str(hash(corrected_text[j1:j2]) % 100000))
            ins_elem.set(qn('w:author'), "AI Корректор")
            ins_elem.set(qn('w:date'), "2026-09-16T10:00:00Z")

            r_elem = OxmlElement('w:r')
            t_elem = OxmlElement('w:t')
            t_elem.text = corrected_text[j1:j2]
            r_elem.append(t_elem)
            ins_elem.append(r_elem)

            para._p.append(ins_elem)

# Использование в вашем основном цикле:
# for para in doc.paragraphs:
#     if not is_heading(para) and para.text.strip():
#         corrected = send_to_ai(para.text)
#         apply_native_track_changes(para, para.text, corrected)