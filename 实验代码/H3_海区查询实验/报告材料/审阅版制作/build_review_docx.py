"""本次Word审阅版构建；原Markdown不改。后续以DOCX为主稿，不自动同步全文。"""
import argparse
import hashlib
import json
import re
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_SECTION_START, WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SOURCE = ROOT / '实验报告_完整初稿.md'
TARGET = ROOT / '实验报告_审阅版.docx'


def el(tag, **attrs):
    e=OxmlElement(tag)
    for k,v in attrs.items(): e.set(qn(k),str(v))
    return e


def font(style, size, east='宋体', latin='Times New Roman', bold=False):
    style.font.name=latin;style.font.size=Pt(size);style.font.bold=bold
    style.font.color.rgb=RGBColor(0,0,0)
    style.font.italic=False
    style.font.underline=False
    rf=style.element.get_or_add_rPr().get_or_add_rFonts()
    rf.set(qn('w:eastAsia'),east);rf.set(qn('w:ascii'),latin);rf.set(qn('w:hAnsi'),latin)


def config_doc(doc):
    styles=doc.styles
    font(styles['Normal'],11)
    pf=styles['Normal'].paragraph_format
    pf.line_spacing=1.4;pf.space_after=Pt(4);pf.first_line_indent=Pt(22)
    pf.widow_control=True
    for name,size in [('Title',20),('Subtitle',13),('Heading 1',15),('Heading 2',12.5),('Heading 3',11.5)]:
        font(styles[name],size,'黑体',bold=name!='Subtitle')
        p=styles[name].paragraph_format;p.first_line_indent=Pt(0);p.space_before=Pt(12);p.space_after=Pt(6)
        p.keep_with_next=True;p.keep_together=True
    for name,size,latin in [('Caption',9.5,'Times New Roman'),('Equation',11,'Cambria Math'),
                            ('Pseudocode',10,'Consolas'),('Reference',10,'Times New Roman'),
                            ('Table Text',10.5,'Times New Roman'),('Material Path',9.5,'Times New Roman')]:
        if name not in styles: styles.add_style(name,WD_STYLE_TYPE.PARAGRAPH)
        font(styles[name],size,latin=latin)
        p=styles[name].paragraph_format;p.first_line_indent=Pt(0);p.line_spacing=1.15
        p.space_after=Pt(3);p.widow_control=True
    styles['Caption'].paragraph_format.keep_together=True
    styles['Equation'].paragraph_format.space_before=Pt(3)
    styles['Equation'].paragraph_format.space_after=Pt(5)
    styles['Pseudocode'].paragraph_format.left_indent=Mm(5)
    styles['Pseudocode'].paragraph_format.line_spacing=1.05
    styles['Reference'].paragraph_format.left_indent=Mm(6)
    styles['Reference'].paragraph_format.first_line_indent=Mm(-6)
    styles['Table Text'].paragraph_format.space_after=Pt(0)
    styles['Table Text'].paragraph_format.line_spacing=1.08
    font(styles['List Bullet'],9.5)
    styles['List Bullet'].paragraph_format.line_spacing=1.15
    styles['List Bullet'].paragraph_format.space_after=Pt(4)
    for name in ['Header','Footer']:
        font(styles[name],9);styles[name].paragraph_format.first_line_indent=Pt(0)
    doc.settings.element.append(el('w:updateFields',**{'w:val':'true'}))
    mathpr=el('m:mathPr');mathpr.append(el('m:mathFont',**{'m:val':'Cambria Math'}));doc.settings.element.append(mathpr)


def section_format(sec,landscape=False):
    sec.page_width=Mm(297 if landscape else 210);sec.page_height=Mm(210 if landscape else 297)
    sec.orientation=WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    sec.left_margin=Mm(25);sec.right_margin=Mm(25)
    sec.top_margin=Mm(20 if landscape else 24);sec.bottom_margin=Mm(18 if landscape else 23)
    sec.header_distance=Mm(9);sec.footer_distance=Mm(10)


def clean(s): return s.replace('**','').replace('`','')


def hyperlink(p,label,url):
    h=el('w:hyperlink',**{'r:id':p.part.relate_to(url,RT.HYPERLINK,is_external=True)})
    r=el('w:r');pr=el('w:rPr');pr.append(el('w:color',**{'w:val':'155A8A'}));pr.append(el('w:u',**{'w:val':'single'}));r.append(pr)
    t=el('w:t');t.text=label;r.append(t);h.append(r);p._p.append(h)


def inline(p,s,paths=False):
    pattern=r'(\*\*.+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))'
    for part in re.split(pattern,s):
        if not part: continue
        if part.startswith('**') and part.endswith('**'): p.add_run(part[2:-2]).bold=True
        elif part.startswith('`') and part.endswith('`'):
            value=part[1:-1]
            if re.fullmatch(r'(P_m|a_m|b_m|C_m|q_m|f_used|f_min|f_max)',value):
                inline_vars(p,value)
            else:
                r=p.add_run(value);r.font.name='Consolas';r.font.size=Pt(9.5)
        elif re.fullmatch(r'\[[^\]]+\]\([^)]+\)',part):
            label,url=re.match(r'\[([^\]]+)\]\(([^)]+)\)',part).groups()
            if url.startswith(('http://','https://')): hyperlink(p,label,url)
            else:
                # 未发布的项目内材料是可追溯路径，不能伪装成网页。
                p.add_run(f'{label}（仓库内路径：{url}）')
        else: inline_vars(p,part)


def inline_vars(p,s):
    pattern=r'(P̂_m|Ĉ_m|Q_m,j|P_m|P_s|a_m|b_m|C_m|q_m|n_b|n_s|f_used|f_min|f_max)'
    for part in re.split(pattern,s):
        if not part:continue
        if re.fullmatch(pattern,part):
            m=el('m:oMath')
            if part in ['P̂_m','Ĉ_m']:node=hat('P' if part.startswith('P') else 'C','m')
            elif part=='q_m':node=hat('q','m')
            else:
                a,b=part.split('_',1);node=sub(a,b)
            m.append(node);p._p.append(m)
        else:p.add_run(part)


def mr(t,normal=False):
    r=el('m:r')
    if normal:
        pr=el('m:rPr');pr.append(el('m:sty',**{'m:val':'p'}));r.append(pr)
    x=el('m:t');x.text=str(t);r.append(x);return r


def sub(t,s):
    e=el('m:sSub');a=el('m:e');a.append(mr(t));b=el('m:sub');b.append(mr(s));e.extend([a,b]);return e


def sup(base,s):
    e=el('m:sSup');a=el('m:e');a.append(base if not isinstance(base,str) else mr(base));b=el('m:sup');b.append(mr(s,True));e.extend([a,b]);return e


def hat(t,s=None):
    a=el('m:acc');pr=el('m:accPr');pr.append(el('m:chr',**{'m:val':'̂'}));a.append(pr);b=el('m:e');b.append(mr(t));a.append(b)
    if s is None:return a
    e=el('m:sSub');b=el('m:e');b.append(a);c=el('m:sub');c.append(mr(s));e.extend([b,c]);return e


def frac(num,den):
    e=el('m:f');a=el('m:num');b=el('m:den')
    for v in num:a.append(v)
    for v in den:b.append(v)
    e.extend([a,b]);return e


def summ(method):
    e=el('m:nary');pr=el('m:naryPr');pr.append(el('m:chr',**{'m:val':'∑'}));pr.append(el('m:limLoc',**{'m:val':'undOvr'}));e.append(pr)
    a=el('m:sub');a.append(mr('j=1'));b=el('m:sup');b.append(mr('K'));c=el('m:e');c.append(sub('Q',method+',j'));e.extend([a,b,c]);return e


def equation(doc,parts):
    p=doc.add_paragraph(style='Equation');p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    m=el('m:oMathPara');eq=el('m:oMath')
    for r in parts:eq.append(r)
    m.append(eq);p._p.append(m);return p


def equation_for(doc,s):
    if s.startswith('f原始'):
        equation(doc,[sub('f','raw'),mr('='),frac([sub('n','b')],[sub('n','s')]),mr(',  '),sub('n','s'),mr('=200')])
        return '其中n_b为满足r7边界格及矩形条件的样本记录数，n_s为样本量。'
    if s.startswith('预计单次'):
        equation(doc,[hat('q','m'),mr('(f)='),sub('a','m'),mr('+'),sub('b','m'),mr('f')]).paragraph_format.keep_with_next=True;return None
    if s.startswith('预计方法'):
        equation(doc,[hat('C','m'),mr('='),hat('P','m'),mr('+K'),hat('q','m'),mr('(f)')]);return '帽号表示校准得到的预测值；m表示候选方法。'
    if s.startswith('固定方法总'):
        equation(doc,[sub('T','m'),mr('(K)=G+'),sub('P','m'),mr('+'),summ('m')]);return None
    if s.startswith('选择策略总'):
        equation(doc,[sub('T','sel'),mr('(K)=G+S+'),sup(sub('P','s'),'rem'),mr('+'),summ('s')]);return '在实际成本式中，G为共同几何准备，S为完整选择耗时，P_m为专属准备实际耗时，P_s的rem上标表示所选方法剩余准备，Q_m,j为第j次正常查询耗时；s表示选中方法。'
    if s.startswith('f = min'):
        equation(doc,[mr('f='),mr('min',True),mr('('),sub('f','max'),mr(', '),mr('max',True),mr('('),sub('f','min'),mr(', '),sub('f','raw'),mr('))')]);return None
    if s.startswith('净节省（中位数差）'):
        equation(doc,[sub('Δ','med'),mr('='),mr('med',True),mr('('),sub('T','c,r'),mr(')−'),mr('med',True),mr('('),sub('T','s,r'),mr(')')]);return 'med表示对七轮取中位数，c和s分别为对照与策略，r为计时轮次。'
    if s.startswith('每轮净节省'):
        equation(doc,[sub('d','r'),mr('='),sub('T','c,r'),mr('−'),sub('T','s,r')]);return None
    if s.startswith('净节省（同轮差中位数）'):
        equation(doc,[sub('Δ','pair'),mr('='),mr('med',True),mr('('),sub('d','r'),mr(')')]);return None
    return False


def table(doc,rows,kind):
    # 分面只改变附录展示，不丢场景或原始精度。
    groups=[list(range(len(rows[0])))]
    if kind=='D':groups=[[0,1,2,3],[0,1,4,5]]
    for index,cols in enumerate(groups):
        if index:
            doc.add_paragraph('附录 D 同轮差值统计 续表',style='Heading 2').paragraph_format.page_break_before=True
        if kind=='D':doc.add_paragraph('同轮差值中位数 [Q25, Q75]，单位ms；'+('S2/KO与S4/S2' if index==0 else 'S4/固定D与S4/固定r7'),style='Caption')
        n=len(cols)
        t=doc.add_table(rows=1,cols=n);t.autofit=False
        if kind=='B':widths=[26,15]+[29.42]*7
        elif kind=='C':widths=[26,15]+[51.5]*4
        elif kind=='D':widths=[26,15,103,103]
        elif kind=='Aparam':widths=[22,23,40,40,35]
        elif kind=='Ainterval':widths=[13,83,32,32]
        else:widths=[27,43,30,30,30]
        for c,w in zip(t.columns,widths):c.width=Mm(w)
        allrows=[rows[0]]+rows[1:]
        for i,row in enumerate(allrows):
            cells=t.rows[0].cells if i==0 else t.add_row().cells
            trpr=t.rows[i]._tr.get_or_add_trPr();trpr.append(el('w:cantSplit'))
            if i==0:trpr.append(el('w:tblHeader',**{'w:val':'true'}))
            for j,orig in enumerate(cols):
                cell=cells[j];cell.width=Mm(widths[j]);cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
                text=clean(row[orig])
                if i==0:
                    if kind=='Aparam':text=['方法','准备P / ms','a / ms·次⁻¹','b / ms·次⁻¹','K-only均值 / ms'][orig]
                    elif kind=='Ainterval':text=['K','截断后比例范围','样本二选一','样本四选一'][orig]
                    elif kind=='B':text=['数据','K','固定D','固定r6','固定r7','固定r8','KO（选择）','S2（选择）','S4（选择）'][orig]
                p=cell.paragraphs[0];p.style='Table Text';p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.keep_with_next=False
                r=p.add_run(text);r.bold=i==0
                if kind in ['B','C','D']:r.font.size=Pt(10.5)
                pr=cell._tc.get_or_add_tcPr();borders=el('w:tcBorders')
                for edge in ['top','bottom','left','right']:borders.append(el('w:'+edge,**{'w:val':'single','w:sz':'4','w:color':'D9D9D9'}))
                pr.append(borders);mar=el('w:tcMar')
                for edge,value in [('top',35),('bottom',35),('left',65),('right',65)]:mar.append(el('w:'+edge,**{'w:w':value,'w:type':'dxa'}))
                pr.append(mar)
                if i==0:pr.append(el('w:shd',**{'w:fill':'E7EBEE'}))
                elif i%2==0:pr.append(el('w:shd',**{'w:fill':'F7F8F9'}))
                if i>0 and kind in ['B','C','D','Aparam'] and orig>0:pr.append(el('w:noWrap'))
        p=doc.add_paragraph(style='Table Text');p.paragraph_format.line_spacing=Pt(1);p.paragraph_format.space_after=Pt(2);p.add_run().font.size=Pt(1)


def parse_blocks(s):
    lines=s.splitlines();i=0;out=[]
    while i<len(lines):
        line=lines[i].strip()
        if not line or line.startswith('<!--'):i+=1;continue
        if line.startswith('```'):
            code=[];i+=1
            while i<len(lines) and not lines[i].startswith('```'):code.append(lines[i]);i+=1
            out.append(('code',code));i+=1;continue
        if line.startswith('|'):
            rows=[]
            while i<len(lines) and lines[i].strip().startswith('|'):
                cells=[v.strip() for v in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r'[:\- ]+',c) for c in cells):rows.append(cells)
                i+=1
            out.append(('table',rows));continue
        if line.startswith('#'):
            out.append(('heading',(len(line)-len(line.lstrip('#')),line.lstrip('#').strip())));i+=1;continue
        if line.startswith('!['):
            m=re.match(r'!\[([^\]]*)\]\(([^)]+)\)',line);out.append(('image',m.groups()));i+=1;continue
        if line.startswith('`') and line.endswith('`'):out.append(('equation',line[1:-1]));i+=1;continue
        if line.startswith('- '):out.append(('paragraph',line));i+=1;continue
        text=line;i+=1
        while i<len(lines) and lines[i].strip() and not lines[i].startswith(('#','|','![','```','- ')):
            text+=' '+lines[i].strip();i+=1
        out.append(('paragraph',text))
    return out


def prepare_content():
    source=SOURCE.read_text(encoding='utf-8');s=source
    changes=json.loads((HERE/'内容修订建议.json').read_text(encoding='utf-8'))
    for item in changes['mandatory']+changes['optional_compressions']:
        assert s.count(item['old'])==1,item['location'];s=s.replace(item['old'],item['new'])
    # 学术内容与复现路径保留，删掉作为协作交付说明的重复句。
    s=s.replace('本报告图片采用[正文版目录](图稿/正文版/)，','本报告使用[正文图稿](图稿/正文版/)，')
    s=s.replace('图3连续分为四张，每张保留三个K下的六组数据及126个原始差值，合计504点。','图3a置于正文，图3b—3d见附录F；四图分别保留三个K下六组数据的126个原始差值，合计504点。')
    s=s.replace('当前200点单档特征策略提高了事后固定方法命中数','当前200点单档特征策略提高了事后固定方法命中数')
    s=s.replace('其中P_m为该方法尚需支付的专属准备费用，a_m、b_m为固定海区和规模下的两个查询参数。','其中P̂_m为该方法专属准备费用的校准估计，a_m、b_m为固定海区和规模下的两个查询参数。')
    s=s.replace('C_m不含各候选共有的共同几何准备','预测成本Ĉ_m不含各候选共有的共同几何准备')
    s=s.replace('这一结果尚不能评价新增档位被选中后的净收益。','')
    s=re.sub(r'\[项目内转换课件\]\([^\)]+\)。','课程内部资料，仓库内路径见附录E。',s)
    s=s.replace('访问日期2026-11-11；实验运行版本为2.2.0。','访问日期2026-11-11。')
    s=s.replace('[SIGSPATIAL 2026官方程序](https://sigspatial2026.sigspatial.org/detailed-program.html)列为Demo D-661；相关DOI：','SIGSPATIAL 2026演示工作。相关DOI：')
    s=s.replace('SIGSPATIAL 2026演示工作，本文引用arXiv v2作者稿。','本文引用arXiv v2作者稿。')
    s=s.replace('代码、数据、冻结协议和统计记录的项目内相对位置见附录。','仓库内材料路径见附录E。')
    s=s.replace('以`P_m + K × 校准单次均值`预测','以专属准备成本的校准估计加上K倍校准单次均值预测')
    # 原正文的外部文献核对表述统一由建议JSON转换；不写入第二份全文。
    blocks=parse_blocks(s);moved=[];result=[];i=0
    skip_d=False
    while i<len(blocks):
        typ,val=blocks[i]
        if typ=='heading' and val[1].startswith('D.1 '):skip_d=True
        if typ=='heading' and val[1].startswith('附录 E'):skip_d=False
        if skip_d:i+=1;continue
        if typ=='image' and any(v in val[0] for v in ['图3b','图3c','图3d']):
            moved.append(blocks[i])
            if i+1<len(blocks) and blocks[i+1][0]=='paragraph' and blocks[i+1][1].startswith('**图3'):
                moved.append(blocks[i+1]);i+=1
        else:
            if typ=='paragraph' and ('图3b' in val or '图3c' in val or '图3d' in val) and not val.startswith('**图'):
                val=val.replace('（图3b）','（见附录F图3b）').replace('图3c展示','附录F图3c展示').replace('图3d按','附录F图3d按')
            result.append((typ,val))
        i+=1
    return result,moved,source


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--overwrite',action='store_true');args=parser.parse_args()
    if TARGET.exists() and not args.overwrite:raise SystemExit('Word现为主稿；要重建请显式使用--overwrite，避免覆盖人工编辑。')
    blocks,moved,source=prepare_content();doc=Document();config_doc(doc);section_format(doc.sections[0])
    cover=doc.sections[0];cover.header.is_linked_to_previous=False;cover.footer.is_linked_to_previous=False
    title=blocks[0][1][1];p=doc.add_paragraph(title,style='Title');p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(80);p.paragraph_format.space_after=Pt(28)
    for text in ['海洋数据挖掘课程报告','Word审阅版','姓名：待填写    学号：待填写','任课教师：待填写','提交日期：待填写','审阅版日期：2026年11月12日']:
        p=doc.add_paragraph(text,style='Subtitle' if text in ['海洋数据挖掘课程报告','Word审阅版'] else 'Normal');p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.space_after=Pt(12)
    p=doc.add_paragraph('代码、数据与AI协作记录公开仓库：提交前补齐');p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.space_before=Pt(28)
    sec=doc.add_section(WD_SECTION_START.NEW_PAGE);section_format(sec)
    sec.header.is_linked_to_previous=False;sec.footer.is_linked_to_previous=False
    p=sec.header.paragraphs[0];p.text='海洋数据挖掘课程报告  H3分辨率成本与选择评估';p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p=sec.footer.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    fld=el('w:fldSimple',**{'w:instr':'PAGE'});p._p.append(fld)
    num=el('w:pgNumType',**{'w:start':'1'});sec._sectPr.append(num)
    start=next(i for i,b in enumerate(blocks) if b[0]=='heading' and b[1][1]=='摘要')
    appendix='';heading='';captions=[];tables=[];figures=[]
    for typ,val in blocks[start:]:
        if typ=='heading':
            level,text=val;heading=text
            if text.startswith('附录 B'):
                sec=doc.add_section(WD_SECTION_START.NEW_PAGE);section_format(sec,True)
            elif text.startswith('附录 E'):
                sec=doc.add_section(WD_SECTION_START.NEW_PAGE);section_format(sec)
            if text.startswith('附录 '):appendix=text[3:4]
            p=doc.add_paragraph(text,style='Heading 1' if level==2 else 'Heading 2')
            if text in ['参考文献'] or text.startswith(('附录 A','附录 C','附录 D')):p.paragraph_format.page_break_before=True
            if text.startswith('附录 E'):
                p=doc.add_paragraph(style='Material Path')
                p.add_run('课程内部资料的仓库内路径：../../PPT_Markdown/第3讲_课时2_渔场真的东移了吗_修订版/第3讲_课时2_渔场真的东移了吗_修订版.md。以下路径均以报告材料目录为相对起点，尚不是已发布的网页地址。')
            continue
        if typ=='image':
            label,path=val;p=doc.add_paragraph();p.paragraph_format.first_line_indent=Pt(0);p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next=True;p.paragraph_format.keep_together=True;p.paragraph_format.space_before=Pt(5)
            p.add_run().add_picture(str(ROOT/path),width=Mm(160));figures.append(label);continue
        if typ=='paragraph':
            if re.match(r'^\*\*图[1-4][a-d]?\*\*',val):
                p=doc.add_paragraph(style='Caption');inline(p,val);captions.append(clean(val).split('　')[0]);continue
            if val.startswith('来源：'):
                # 详细字段和完整路径已汇于E，避免附录表后重复长路径。
                val='数据来源及字段对应见附录E；原始CSV保留完整重复记录和详细统计。'
            if appendix=='A' and val.startswith('K=1按既定捷径'):
                val='执行第3.5节的冻结规则：K=1直接查询，其他K先定共同基准B；新增档严格低于0.98倍B成本才切换。等式点不切换，新增同价优先r6，浮点等式容差为1e-12 ms。原始最低者仅作解释。'
            if appendix=='A' and val.startswith('不保守时的四方法'):
                val='原始模型最低者的转折为：K=100时f=0.18766315370056594，从r7转为direct，等式点direct；K=1000时f=0.4956989843023951之后从r7转为r6，等式点r7。全部区间及边界记录见附录E校准目录的decision_intervals.csv与JSON。'
            style='Reference' if heading=='参考文献' else ('Material Path' if appendix=='E' else 'Normal')
            if val.startswith('- '):style='List Bullet';val=val[2:]
            p=doc.add_paragraph(style=style)
            if appendix=='A':p.paragraph_format.line_spacing=1.2;p.paragraph_format.space_after=Pt(3)
            if appendix=='E' and val.startswith('实际运行环境见'):p.paragraph_format.keep_together=True
            if val.startswith('对每种方法m'):p.paragraph_format.keep_with_next=True
            if val.startswith(('**第一','**第二','**第三')):p.paragraph_format.keep_together=True
            inline(p,val,paths=appendix=='E');continue
        if typ=='equation':
            extra=equation_for(doc,val)
            if extra is False:
                p=doc.add_paragraph(style='Material Path');p.add_run(val)
            elif extra:
                p=doc.add_paragraph();inline(p,extra)
            continue
        if typ=='code':
            code=val
            for j,line in enumerate(code):
                p=doc.add_paragraph(style='Pseudocode');p.add_run(line);p.paragraph_format.keep_with_next=j<len(code)-1
            continue
        if typ=='table':
            kind='Aparam' if appendix=='A' and 'prepare_ms' in val[0][1] else 'Ainterval' if appendix=='A' else appendix
            table(doc,val,kind);tables.append({'kind':kind,'rows':len(val)-1});continue
    doc.add_paragraph('附录 F 策略比较分图',style='Heading 1').paragraph_format.page_break_before=True
    doc.add_paragraph('本附录承接正文图3a，保留其余三种比较的全部18场景、七轮原始差值、中位数与IQR。各图仍使用同轮对照减策略口径，与附录C的两个中位数差分开。')
    for i in range(0,len(moved),2):
        (_, (label,path)),(_,caption)=moved[i:i+2]
        p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.first_line_indent=Pt(0);p.paragraph_format.keep_with_next=True
        if i:p.paragraph_format.page_break_before=True
        p.add_run().add_picture(str(ROOT/path),width=Mm(160));figures.append(label)
        p=doc.add_paragraph(style='Caption');inline(p,caption);captions.append(clean(caption).split('　')[0])
    props=doc.core_properties;props.title=title;props.subject='海洋数据挖掘课程报告审阅版';props.author='待填写';props.comments='Word为后续主稿；原Markdown为保留来源。'
    # 不继承模板Title的蓝色装饰线；横向和后续节继续页码。
    for root in (doc.element,doc.styles.element):
        for border in list(root.iter(qn('w:pBdr'))):border.getparent().remove(border)
    for sec in doc.sections[2:]:
        for pg in sec._sectPr.findall(qn('w:pgNumType')):sec._sectPr.remove(pg)
    doc.save(TARGET)
    assert SOURCE.read_text(encoding='utf-8')==source
    record={'target':TARGET.name,'source_sha256':hashlib.sha256(SOURCE.read_bytes()).hexdigest(),'source_unchanged':True,
            'figures':figures,'captions':captions,'tables':tables,'native_equation_count':len(doc.element.xpath('//m:oMath')),
            'layout':'A4 portrait 160mm text width; appendices B-D landscape','fonts':'宋体11pt / Times New Roman; 黑体 headings; Cambria Math; table10.5pt; caption9.5pt',
            'future_master':'DOCX; do not rebuild over later manual edits without explicit overwrite'}
    (HERE/'构建记录.json').write_text(json.dumps(record,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print(json.dumps(record,ensure_ascii=False))


if __name__=='__main__':main()
