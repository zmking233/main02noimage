import re
import sys
import fitz  # PyMuPDF
from docx import Document




def reduce_repeated_chars(text):
    def replacer(m):
        return m.group(2)
    pattern = r'(([\u4e00-\u9fa5])(\s*\2){2,})'
    return re.sub(pattern, replacer, text)

def fuzzy_kw_regex(kw):
    return r'\s*'.join([fr'{ch}(?:\s*{ch})*' for ch in kw])

def fuzzy_find_section(text, start_keywords, end_keywords):
    start_pattern = '|'.join(fuzzy_kw_regex(kw) for kw in start_keywords)
    end_pattern = '|'.join(fuzzy_kw_regex(kw) for kw in end_keywords)
    regex = re.compile(fr'({start_pattern})(.*?)({end_pattern})', re.DOTALL)
    m = regex.search(text)
    return m.group(2).strip() if m else ""

def clean_text(text):
    text = re.sub(r'第?\s*\d+\s*页', '', text)
    text = re.sub(r'\d+\s*/\s*\d+\s*页', '', text)
    watermark_pattern = r'(人\s*民\s*法\s*院\s*案\s*例\s*库){1,}'
    text = re.sub(watermark_pattern, '', text)
    text = reduce_repeated_chars(text)
    lines = text.splitlines()
    lines = [line.rstrip() for line in lines]
    return '\n'.join(lines)

def process_pdf_text_lines(raw_text):
    lines = raw_text.splitlines()
    new_lines = []
    buffer = ""
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if buffer and not buffer.endswith(('。', '！', '？', '：', '.', '!', '?', ':', '；')):
            buffer += line
        else:
            if buffer:
                new_lines.append(buffer)
            buffer = line
    if buffer:
        new_lines.append(buffer)
    final_lines = []
    for i, line in enumerate(new_lines):
        final_lines.append(line)
        hanzi_count = len([ch for ch in line if '\u4e00' <= ch <= '\u9fa5'])
        if hanzi_count < 20 and line.endswith('。'):
            final_lines.append('')
    return '\n\n'.join([l for l in final_lines if l or (i > 0 and final_lines[i-1])])

def parse_fields(text):
    m_num = re.search(r'(\d{4}(?:-\d+){3,4})', text)
    case_number = m_num.group(1) if m_num else "未知编号"
    m_name = re.search(fr'{re.escape(case_number)}\s*\n?(.*?案)', text, re.DOTALL)
    case_name = m_name.group(1).strip() if m_name else "未知案件名称"
    m_desc = re.search(fr'案(.*?){fuzzy_kw_regex("关键词")}', text, re.DOTALL)
    case_desc = m_desc.group(1).strip() if m_desc else ""
    key_word = fuzzy_find_section(text, ["关键词"], ["基本案情"])
    case_text = fuzzy_find_section(text, ["基本案情"], ["裁判理由"])
    trial_process = fuzzy_find_section(text, ["裁判理由"], ["裁判要旨"])
    trial_abbr = fuzzy_find_section(text, ["裁判要旨"], ["关联索引"])
    m_index = re.search(fr'{fuzzy_kw_regex("关联索引")}(.*)', text, re.DOTALL)
    relevant_index = m_index.group(1).strip() if m_index else ""
    return {
        "case_number": case_number,
        "case_name": case_name,
        "case_desc": case_desc,
        "key_word": key_word,
        "case_text": case_text,
        "trial_process": trial_process,
        "trial_abbr": trial_abbr,
        "relevant_index": relevant_index
    }

def styled_paragraph(text, color, size=16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=0, background="#ffffff"):
    font_weight = 'bold' if bold else 'normal'
    style = (
        f"color:{color};"
        f"font-size:{size}px;"
        f"text-align:{align};"
        f"line-height:{line_height};"
        f"margin-top:{margin_top}px;"
        f"margin-bottom:{margin_bottom}px;"
        f"margin-left:{indent_px}px;"
        f"margin-right:{indent_px}px;"
        f"background-color:{background};"
        f"font-family:'Helvetica Neue', Helvetica, 'Hiragino Sans GB', 'Microsoft YaHei', Arial, sans-serif;"
        f"font-weight:{font_weight};"
    )
    return f'<p style="{style}">{text}</p>'

def render_multiline_section(text, color, size=16, bold=False):
    paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
    return [
        styled_paragraph(para, color, size, bold, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=32)
        for para in paragraphs
    ]

def generate_wechat_html(data):
    parts = []
    parts.append('<meta charset="UTF-8">') # 添加编码声明

    parts.append(styled_paragraph(
        f"本期推送人民法院案例库编号为{data['case_number']}的参考案例",
        "#5e5e5e", 16, bold=False, align='justify', line_height=1.6, indent_px=16,  margin_top=0, margin_bottom=8
    ))

    parts.append(styled_paragraph(
        f'<span style="background-color:#5287b7;color:#ffffff;">延伸阅读</span>',
        "#5e5e5e", 16, bold=False, align='justify', line_height=1.6, indent_px=16, margin_top=0, margin_bottom=8, background="#ffffff"
    ))

    parts.append(styled_paragraph(
        data['case_name'], "#5287b7", 18, bold=True, align='left', line_height=2, indent_px=16, margin_top=32, margin_bottom=0
    ))

    parts.append(styled_paragraph(
        data['case_desc'], "#424242", 18, bold=True, align='left', line_height=2, indent_px=16, margin_top=0, margin_bottom=32
    ))

    parts.append(styled_paragraph('【关键词】', "#5287b7", 16, bold=True,  margin_bottom=32))
    parts.extend(render_multiline_section(data['key_word'], "#5e5e5e"))

    parts.append(styled_paragraph('【基本案情】', "#5287b7", 16, bold=True,  margin_bottom=32))
    parts.extend(render_multiline_section(data['case_text'], "#5e5e5e"))

    parts.append(styled_paragraph('【裁判理由】', "#5287b7", 16, bold=True,  margin_bottom=32))
    parts.extend(render_multiline_section(data['trial_process'], "#5e5e5e"))

    parts.append(styled_paragraph('【裁判要旨】', "#5287b7", 16, bold=True,  margin_bottom=32))
    parts.extend(render_multiline_section(data['trial_abbr'], "#5e5e5e"))

    parts.append(styled_paragraph('【关联索引】', "#5287b7", 16, bold=True,  margin_bottom=32))

    relevant_index = data['relevant_index']
    if relevant_index:
        count = [0]
        def replace_bracket(match):
            count[0] += 1
            return '<br/>《' if count[0] > 1 else '《'
        relevant_index = re.sub(r'《', replace_bracket, relevant_index)
        relevant_index = re.sub(r'(一审)', r'<br/><br/>\1', relevant_index)
        relevant_index = re.sub(r'(二审)', r'<br/>\1', relevant_index)
        relevant_index = re.sub(r'(其他审理程序)', r'<br/>\1', relevant_index)
        relevant_index = re.sub(r'(再审)', r'<br/>\1', relevant_index)
        relevant_index = re.sub(r'(本案例文本已于)', r'<br/><br/>\1', relevant_index)

        paragraphs = [p for p in relevant_index.split('\n\n') if p.strip()]
        for para in paragraphs:
            parts.append(styled_paragraph(
                para, "#5e5e5e", 16, bold=False, align='justify', line_height=2, indent_px=16, margin_top=0, margin_bottom=32
            ))

    parts.append(styled_paragraph(
        '编辑团队：薛政  黄琳娜  初相钰  赵绮', "#5e5e5e", 16, bold=False, align='center', line_height=2, indent_px=16, margin_top=0, margin_bottom=32
    ))

    return "\n".join(parts)

def extract_text_from_pdf(path):
    doc = fitz.open(path)
    full_text = [page.get_text() for page in doc]
    return process_pdf_text_lines("\n".join(full_text))

def extract_text_from_docx(path):
    doc = Document(path)
    lines = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return '\n\n'.join(lines)

def main():
    if len(sys.argv) < 2:
        print("用法: python script.py 文件路径")
        return
    
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))

    path = sys.argv[1]
    base_name = os.path.splitext(os.path.basename(path))[0]

    text_dir = os.path.join(base_path, 'text')
    output_dir = os.path.join(base_path, 'output')
    os.makedirs(text_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)

    if path.endswith(".pdf"):
        raw_text = extract_text_from_pdf(path)
    else:
        print("仅支持PDF或DOCX文件")
        return

    cleaned = clean_text(raw_text)
    parsed = parse_fields(cleaned)

    text_file = os.path.join(text_dir, f"{base_name}-文本.txt")
    with open(text_file, "w", encoding="utf-8") as f:
        f.write(cleaned)

    html_output = generate_wechat_html(parsed)
    html_file = os.path.join(output_dir, f"{base_name}-公众号格式.html")
    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_output)

    print(f"✅ 提取完成，生成文件：\n - {text_file}\n - {html_file}")

import sys
import os
import subprocess
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QListWidget, QLabel, QFileDialog, QPushButton, QSpacerItem, QSizePolicy
)
from PyQt5.QtCore import Qt

def get_base_path():
    if getattr(sys, 'frozen', False): # 如果是打包运行的
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

class DropWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("自动排版工具")
        self.setAcceptDrops(True)
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        self.label = QLabel("拖入案例库PDF文件（可多选）", self)
        self.label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.label)
        self.listWidget = QListWidget(self)
        layout.addWidget(self.listWidget)
        self.btn_select = QPushButton("手动选择文件", self)
        layout.addWidget(self.btn_select)
        self.btn_select.clicked.connect(self.open_file_dialog)
        base_path = get_base_path()
        self.output_dir = os.path.join(base_path, "output")
        os.makedirs(self.output_dir, exist_ok=True)
        self.listWidget.itemDoubleClicked.connect(self.open_file)

        # 添加弹性空间和右下角标签
        layout.addSpacerItem(QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.author_label = QLabel("By LeClaire", self)
        self.author_label.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        layout.addWidget(self.author_label)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        self.process_files(files)

    def open_file_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择PDF文件", "", "PDF(*.pdf)")
        if files:
            self.process_files(files)

    def process_files(self, files):
        for path in files:
            base_name = os.path.splitext(os.path.basename(path))[0]
            if path.endswith(".pdf"):
                raw_text = extract_text_from_pdf(path)
            else:
                continue
            cleaned = clean_text(raw_text)
            parsed = parse_fields(cleaned)
            html_output = generate_wechat_html(parsed)
            html_file = os.path.join(self.output_dir, f"{base_name}-公众号格式.html")
            with open(html_file, "w", encoding="utf-8") as f:
                f.write(html_output)
            self.listWidget.addItem(html_file)

    def open_file(self, item):
        path = item.text()
        if sys.platform.startswith('win'):
            os.startfile(path)
        elif sys.platform.startswith('darwin'):
            subprocess.call(['open', path])
        else:
            subprocess.call(['xdg-open', path])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = DropWidget()
    w.show()
    sys.exit(app.exec_())
