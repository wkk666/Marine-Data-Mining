"""仅检查本次文档转换产物；不导入或运行任何实验代码。"""
import hashlib
import json
import re
from pathlib import Path
from zipfile import ZipFile

import pdfplumber
from docx import Document
from docx.oxml.ns import qn
from lxml import etree

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
DOCX = ROOT / '实验报告_审阅版.docx'
PDF = ROOT / '实验报告_审阅版.pdf'
SOURCE = ROOT / '实验报告_完整初稿.md'
MODEL = ROOT.parent / 'runs/selector_cal_20261111_140827_086716/online_model.json'
META = ROOT / '审阅版预览/Word分页.json'
EXPECTED_SOURCE_SHA = '7b3314ba7fb942dd534f3ab560b9978eff0c430fa70f09d1e5577ab6161ffa69'


def sha(p):
    return hashlib.sha256(p.read_bytes()).hexdigest()


def main():
    doc = Document(DOCX)
    with ZipFile(DOCX) as z:
        root = etree.fromstring(z.read('word/document.xml'))
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
          'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math'}
    text = ''.join(root.xpath('//w:t/text() | //m:t/text()', namespaces=ns))
    norm = re.sub(r'\s+', '', text)
    checks = {}
    checks['original_markdown_unchanged'] = sha(SOURCE) == EXPECTED_SOURCE_SHA
    checks['online_task_cost_boundary'] = all(x in norm for x in [
        '冻结校准参数已经存在时', '一次在线任务', '不在每个验证任务中重复计入', '不据此评价整个项目从零建设的总收益'])
    checks['separate_appendix_references'] = '完整校准参数见附录A，离线校准费用见附录E.1' in norm
    checks['K1000_qualification'] = '在K=1000时，准确比例下原始模型最低成本方法是r6' in norm
    checks['key_results_retained'] = all(x in norm for x in ['16/18', '14/18', '−1.8359', '+0.3621', '24.4056', '6.9551%', '本验证集未触发新增档位，未验证其被选中后的收益'])
    checks['course_extensions_retained'] = '3.6基于课程的延伸设计' in norm
    checks['repository_placeholder'] = '代码、数据与AI协作记录公开仓库：提交前补齐' in norm
    checks['no_markdown_blocks'] = all(x not in text for x in ['```', '![', '**'])
    model = json.loads(MODEL.read_text(encoding='utf-8'))
    expected = [str(v) for method in model['methods'].values() for v in method.values()]
    expected += [str(model[k]) for k in ['f_min', 'f_max', 'sample_size', 'sample_seed', 'common_geometry_ms']]
    checks['full_precision_model_parameters'] = all(x in text for x in expected)
    checks['editable_math'] = len(root.xpath('//m:oMath', namespaces=ns)) == 28
    checks['display_math'] = len(root.xpath('//m:oMathPara', namespaces=ns)) == 9
    checks['editable_pseudocode'] = len([p for p in doc.paragraphs if p.style.name == 'Pseudocode']) >= 15
    checks['seven_figures'] = len(doc.inline_shapes) == 7
    checks['all_table_headers_repeat'] = all(t.rows[0]._tr.find(qn('w:trPr')).find(qn('w:tblHeader')) is not None for t in doc.tables)
    checks['complete_18_scenario_tables'] = sum(len(t.rows) == 19 for t in doc.tables) == 4
    refs = [p.text for p in doc.paragraphs if p.style.name == 'Reference']
    checks['six_numbered_references'] = all(any(p.startswith(f'[{i}]') for p in refs) for i in range(1, 7))
    meta = json.loads(META.read_text(encoding='utf-8-sig'))
    checks['preview_matches_docx'] = meta['docx_sha256'] == sha(DOCX)
    checks['preview_matches_pdf'] = meta['pdf_sha256'] == sha(PDF)
    checks['captions_same_page'] = all(i['caption_on_same_page'] for i in meta['image_pages'])
    checks['continuous_page_numbers'] = all(p['printed_page_at_start'] == p['physical_page'] - 1 for p in meta['pages'][1:])
    bounds = []
    with pdfplumber.open(PDF) as pdf:
        checks['pdf_page_count'] = len(pdf.pages) == meta['physical_pages']
        for i, p in enumerate(pdf.pages, 1):
            chars = [c for c in p.chars if c.get('text', '').strip()]
            overflow = [c for c in chars if c['x0'] < 65 or c['x1'] > p.width - 65 or c['top'] < 15 or c['bottom'] > p.height - 15]
            bounds.append({'physical_page': i, 'width_pt': float(p.width), 'height_pt': float(p.height), 'char_count': len(chars), 'outside_safe_bounds': ''.join(c['text'] for c in overflow)})
    checks['no_blank_pages'] = all(p['char_count'] > 20 for p in bounds)
    checks['no_text_outside_safe_bounds'] = all(not p['outside_safe_bounds'] for p in bounds)
    out = {'checks': checks, 'all_passed': all(checks.values()), 'page_bounds': bounds,
           'physical_pages': meta['physical_pages'], 'manual_visual_inspection': '逐页打开140dpi PNG；公式、图注、数字、表头、页码和页边距均复查。',
           'source_sha256': sha(SOURCE), 'docx_sha256': sha(DOCX), 'pdf_sha256': sha(PDF), 'frozen_model_sha256': sha(MODEL),
           'scope': '只检查内容转换与排版，不重新生成数据、拟合或执行查询计时。'}
    (HERE / '审阅版检查.json').write_text(json.dumps(out, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    files = [DOCX, PDF, ROOT / '审阅版说明.md', SOURCE, MODEL,
             HERE / 'build_review_docx.py', HERE / 'export_word_pdf.ps1', HERE / 'render_review.py',
             HERE / 'check_review.py', HERE / '内容修订建议.json', HERE / '构建记录.json',
             HERE / '审阅版检查.json', META]
    manifest = {'schema': 1, 'master': DOCX.name, 'preview_only': PDF.name,
                'files': [{'path_relative_to_materials': Path(__import__('os').path.relpath(p, ROOT)).as_posix(), 'sha256': sha(p)} for p in files]}
    (HERE / '交付指纹.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    print(json.dumps({'all_passed': out['all_passed'], 'failed': [k for k,v in checks.items() if not v], 'pages': out['physical_pages']}, ensure_ascii=False))
    if not out['all_passed']:
        raise SystemExit(1)


if __name__ == '__main__':
    main()
