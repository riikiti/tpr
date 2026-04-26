"""
Конвертирует задания_защита_лаб1-2.md -> задания_защита_лаб1-2.pdf
Использует reportlab + Windows-шрифты (поддержка кириллицы).
"""

import re
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether
)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER

# ── Шрифты ──────────────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont('Arial',        'C:/Windows/Fonts/arial.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Bold',   'C:/Windows/Fonts/arialbd.ttf'))
pdfmetrics.registerFont(TTFont('Arial-Italic', 'C:/Windows/Fonts/ariali.ttf'))
pdfmetrics.registerFont(TTFont('Courier-New',  'C:/Windows/Fonts/cour.ttf'))
pdfmetrics.registerFont(TTFont('Courier-Bold', 'C:/Windows/Fonts/courbd.ttf'))

pdfmetrics.registerFontFamily(
    'Arial',
    normal='Arial',
    bold='Arial-Bold',
    italic='Arial-Italic',
    boldItalic='Arial-Bold',
)

# ── Стили ────────────────────────────────────────────────────────────────────
def make_styles():
    def p(name, **kw):
        defaults = dict(fontName='Arial', leading=14, spaceAfter=4,
                        spaceBefore=2, wordWrap='CJK')
        defaults.update(kw)
        return ParagraphStyle(name, **defaults)

    s = {}
    s['h1'] = p('h1', fontSize=16, fontName='Arial-Bold',
                leading=20, spaceBefore=14, spaceAfter=8,
                textColor=colors.HexColor('#1a1a6e'))
    s['h2'] = p('h2', fontSize=13, fontName='Arial-Bold',
                leading=17, spaceBefore=12, spaceAfter=6,
                textColor=colors.HexColor('#1a1a6e'))
    s['h3'] = p('h3', fontSize=11, fontName='Arial-Bold',
                leading=15, spaceBefore=8, spaceAfter=4,
                textColor=colors.HexColor('#333399'))
    s['h4'] = p('h4', fontSize=10, fontName='Arial-Bold',
                leading=14, spaceBefore=6, spaceAfter=3,
                textColor=colors.HexColor('#555555'))
    s['body']     = p('body', fontSize=9.5)
    s['bullet']   = p('bullet', fontSize=9.5, leftIndent=14, bulletIndent=0)
    s['blockquote'] = p('blockquote', fontSize=9, fontName='Arial-Italic',
                        leftIndent=16, textColor=colors.grey)
    s['table_header'] = p('th', fontSize=8.5, fontName='Arial-Bold', leading=11)
    s['table_cell']   = p('td', fontSize=8.5, fontName='Arial', leading=11)
    s['code_block']   = p('cb', fontSize=8, fontName='Courier-New',
                           leading=11, leftIndent=6)
    return s


STYLES = make_styles()

# ── Вспомогательные функции ──────────────────────────────────────────────────

def escape_xml(text):
    return text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

def inline_format(text):
    """Обрабатывает **bold**, *italic*, `code` внутри строки."""
    text = escape_xml(text)

    # 1. Защищаем inline-code до применения bold/italic (иначе * в формулах ломает теги)
    codes = []
    def save_code(m):
        codes.append(m.group(1))
        return f'\x00CODE{len(codes) - 1}\x00'
    text = re.sub(r'`([^`]+)`', save_code, text)

    # 2. bold+italic, bold, italic
    text = re.sub(r'\*\*\*(.+?)\*\*\*', r'<b><i>\1</i></b>', text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    # Italic: * не должна быть вплотную к букве/цифре снаружи (иначе p* ломает)
    text = re.sub(r'(?<!\w)\*(?!\s)([^*\n]+?)(?<!\s)\*(?!\w)', r'<i>\1</i>', text)

    # 3. Восстанавливаем code
    def restore_code(m):
        return f'<font name="Courier-New" size="8">{codes[int(m.group(1))]}</font>'
    text = re.sub(r'\x00CODE(\d+)\x00', restore_code, text)

    return text


# ── Таблицы ──────────────────────────────────────────────────────────────────

def parse_table(lines):
    """Принимает список строк таблицы Markdown, возвращает platypus Table."""
    rows_raw = []
    for line in lines:
        # Убираем обрамляющие |
        line = line.strip().strip('|')
        cells = [c.strip() for c in line.split('|')]
        rows_raw.append(cells)

    # Удаляем разделительную строку (содержит только ---)
    rows = [r for r in rows_raw
            if not all(re.match(r'^[-: ]+$', c) for c in r)]

    if not rows:
        return None

    # Выравниваем ширину строк
    max_cols = max(len(r) for r in rows)
    rows = [r + [''] * (max_cols - len(r)) for r in rows]

    # Формируем данные
    data = []
    for i, row in enumerate(rows):
        style = STYLES['table_header'] if i == 0 else STYLES['table_cell']
        data.append([Paragraph(inline_format(c), style) for c in row])

    # Ширина столбцов
    page_w = A4[0] - 3.5 * cm
    col_w = page_w / max_cols

    tbl = Table(data, colWidths=[col_w] * max_cols, repeatRows=1)
    tbl.setStyle(TableStyle([
        ('BACKGROUND',  (0, 0), (-1, 0),  colors.HexColor('#d0d8f0')),
        ('TEXTCOLOR',   (0, 0), (-1, 0),  colors.HexColor('#1a1a6e')),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1),
         [colors.white, colors.HexColor('#f5f5f5')]),
        ('GRID',        (0, 0), (-1, -1), 0.4, colors.HexColor('#aaaaaa')),
        ('VALIGN',      (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING',(0, 0), (-1, -1), 4),
        ('TOPPADDING',  (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING',(0,0), (-1, -1), 3),
    ]))
    return tbl


# ── Кодовые блоки ─────────────────────────────────────────────────────────────

def make_code_block(lines):
    """Рисует блок кода с серым фоном."""
    text = '\n'.join(lines)
    # Escape XML
    text = escape_xml(text)
    # Сохраняем пробелы через неразрывные пробелы
    text = text.replace(' ', '\u00a0')
    para = Paragraph(text.replace('\n', '<br/>'), STYLES['code_block'])
    tbl = Table([[para]],
                colWidths=[A4[0] - 3.5 * cm])
    tbl.setStyle(TableStyle([
        ('BACKGROUND',   (0, 0), (-1, -1), colors.HexColor('#f0f0f0')),
        ('BOX',          (0, 0), (-1, -1), 0.5, colors.HexColor('#cccccc')),
        ('LEFTPADDING',  (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING',   (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING',(0, 0), (-1, -1), 5),
    ]))
    return tbl


# ── Основной парсер ──────────────────────────────────────────────────────────

def md_to_story(md_text):
    story = []
    lines = md_text.splitlines()
    i = 0
    n = len(lines)

    while i < n:
        line = lines[i]
        stripped = line.strip()

        # ── Пустая строка ────────────────────────────────────────────────────
        if stripped == '':
            story.append(Spacer(1, 4))
            i += 1
            continue

        # ── Горизонтальная линия ─────────────────────────────────────────────
        if re.match(r'^---+$', stripped):
            story.append(Spacer(1, 4))
            story.append(HRFlowable(width='100%', thickness=0.8,
                                     color=colors.HexColor('#aaaaaa')))
            story.append(Spacer(1, 4))
            i += 1
            continue

        # ── Заголовки ────────────────────────────────────────────────────────
        m = re.match(r'^(#{1,4})\s+(.*)', stripped)
        if m:
            level = len(m.group(1))
            text  = inline_format(m.group(2))
            key   = ['h1', 'h2', 'h3', 'h4'][min(level - 1, 3)]
            story.append(Paragraph(text, STYLES[key]))
            i += 1
            continue

        # ── Блок кода ────────────────────────────────────────────────────────
        if stripped.startswith('```'):
            i += 1
            code_lines = []
            while i < n and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # пропускаем закрывающий ```
            if code_lines:
                story.append(Spacer(1, 3))
                story.append(make_code_block(code_lines))
                story.append(Spacer(1, 3))
            continue

        # ── Таблица ──────────────────────────────────────────────────────────
        if stripped.startswith('|'):
            tbl_lines = []
            while i < n and lines[i].strip().startswith('|'):
                tbl_lines.append(lines[i])
                i += 1
            tbl = parse_table(tbl_lines)
            if tbl:
                story.append(Spacer(1, 3))
                story.append(tbl)
                story.append(Spacer(1, 3))
            continue

        # ── Цитата (blockquote) ──────────────────────────────────────────────
        if stripped.startswith('>'):
            text = inline_format(stripped.lstrip('> '))
            story.append(Paragraph(text, STYLES['blockquote']))
            i += 1
            continue

        # ── Маркированный список ─────────────────────────────────────────────
        m = re.match(r'^[-*]\s+(.*)', stripped)
        if m:
            text = inline_format(m.group(1))
            story.append(Paragraph('• ' + text, STYLES['bullet']))
            i += 1
            continue

        # ── Нумерованный список ──────────────────────────────────────────────
        m = re.match(r'^(\d+)\.\s+(.*)', stripped)
        if m:
            text = inline_format(m.group(2))
            story.append(Paragraph(m.group(1) + '. ' + text, STYLES['bullet']))
            i += 1
            continue

        # ── Обычный абзац ────────────────────────────────────────────────────
        text = inline_format(stripped)
        story.append(Paragraph(text, STYLES['body']))
        i += 1

    return story


# ── Генерация PDF ────────────────────────────────────────────────────────────

INPUT  = r"C:\Users\Ruslan\Проекты\тпр\примеры для защиты\1-2\задания_защита_лаб1-2.md"
OUTPUT = r"C:\Users\Ruslan\Проекты\тпр\примеры для защиты\1-2\задания_защита_лаб1-2.pdf"

with open(INPUT, encoding='utf-8') as f:
    md_text = f.read()

doc = SimpleDocTemplate(
    OUTPUT,
    pagesize=A4,
    leftMargin=2.0 * cm,
    rightMargin=1.5 * cm,
    topMargin=1.8 * cm,
    bottomMargin=1.8 * cm,
    title='Задания для защиты Лаб 1–2',
    author='ТПР',
)

story = md_to_story(md_text)
doc.build(story)
print(f"PDF создан: {OUTPUT}")
