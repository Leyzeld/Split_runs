import re
from docx import Document
from docx.text.run import Run
from docx.text.hyperlink import Hyperlink
from docx.oxml import parse_xml
from tqdm import tqdm
import time

def copy_styles(new_run, run):
    if not '<w:hyperlink' in run._element.xml:
        new_run.bold = run.bold
        new_run.italic = run.italic
        new_run.underline = run.underline
        new_run.font.size = run.font.size
        new_run.font.name = run.font.name
        new_run.font.color.rgb = run.font.color.rgb
        new_run.font.highlight_color = run.font.highlight_color
        new_run.font.strike = run.font.strike
        new_run.font.shadow = run.font.shadow
        new_run.font.outline = run.font.outline
    else:
        run_copy_xml = parse_xml(run._element.xml)
        hyperlink = Hyperlink(run_copy_xml, run._parent)
        return hyperlink._element

def castom_deepcopy(runs):
    copied_runs = []
    for run in runs:
        if '<w:hyperlink' in run._element.xml:
            run_copy_xml = parse_xml(run._element.xml)
            new_run = Hyperlink(run_copy_xml, run._parent)
        else:
            run_copy_xml = parse_xml(run._element.xml)
            new_run = Run(run_copy_xml, run._parent)
        copy_styles(new_run, run)
        copied_runs.append(new_run)
    return copied_runs

def split_words_into_runs(doc_path, output_path='out.docx'):
    doc = Document(doc_path)

    for paragraph in tqdm(doc.paragraphs, desc='paragraphs'):
        hyperlinks = paragraph.hyperlinks
        old_runs = castom_deepcopy(paragraph.iter_inner_content())
        full_text = ''.join(run.text for run in old_runs)
        parts = re.findall(r'\S+|\s+', full_text)

        paragraph.clear()

        run_index = 0
        char_index = 0
        for part in parts:
            while run_index < len(old_runs) and char_index >= len(old_runs[run_index].text):
                char_index -= len(old_runs[run_index].text)
                run_index += 1
            if run_index < len(old_runs):
                if '<w:hyperlink' in old_runs[run_index]._element.xml:
                    element = copy_styles(old_runs[run_index], old_runs[run_index])
                    if hyperlinks:
                        if element.rId == hyperlinks[0]._element.rId:
                            paragraph._element.append(element)
                            del hyperlinks[0]
                else:
                    new_run = paragraph.add_run(part)
                    copy_styles(new_run, old_runs[run_index])
            char_index += len(part)

            for index in range(0, len(old_runs)):
                if 'footnoteReference' in old_runs[index]._element.xml:
                    pre_footnote_text = old_runs[max(index-1, 0)].text
                    pre_footnote_text_list = pre_footnote_text.split()
                    if pre_footnote_text_list:
                        if pre_footnote_text_list[-1] in part:
                            paragraph._element.append(old_runs[index].element)
                    else:
                        paragraph._element.append(old_runs[index].element)
        for drawing in old_runs:
            if 'w:drawing' in drawing._element.xml:
                paragraph._element.append(drawing.element)
    doc.save(output_path)

start_time = time.time()
split_words_into_runs('in.docx', 'out.docx')
end_time = time.time()

print(f'Time: {end_time - start_time:.6f}')
