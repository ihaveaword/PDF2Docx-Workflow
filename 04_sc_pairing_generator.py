import os
import re
from pdf2image import convert_from_path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH

# === 配置区域 ===
SC_SOURCE_FOLDER = 'SC_Source'      # 软件著作权PDF文件夹
OUTPUT_SC_FOLDER = 'Output_Sc'      # 输出文件夹

# 确保输出目录存在
os.makedirs(OUTPUT_SC_FOLDER, exist_ok=True)

def setup_page_layout(doc):
    """设置页面为 A4 横向，并应用自定义页边距"""
    section = doc.sections[0]
    
    # 1. 设置方向为横向 (Landscape)
    section.orientation = WD_ORIENT.LANDSCAPE
    # 必须手动交换宽高才能真正生效
    section.page_width, section.page_height = section.page_height, section.page_width
    
    # 2. 设置页边距
    section.top_margin = Cm(0.6)
    section.bottom_margin = Cm(0.6)
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)

def extract_sort_id(filename):
    """
    从文件名中提取序号
    例如从 "授权专利-计算机软件著作权1-xxx.pdf" 中提取出 1
    """
    match = re.search(r'(\d+)', filename)
    if match:
        return int(match.group(1))
    return 9999

def extract_sc_name(filename):
    """
    从文件名中提取软著名称（去掉序号前缀和.pdf后缀）
    例如: "授权专利-计算机软件著作权1-一种视觉间隙测量软件V1.0-2021SR0433578.pdf"
    返回: "一种视觉间隙测量软件V1.0-2021SR0433578"
    """
    # 去掉 .pdf 后缀
    name_without_ext = filename.replace('.pdf', '')
    # 只保留最后一个 "-" 后面的内容作为名称
    # 匹配模式: 授权专利-计算机软件著作权数字-名称
    match = re.search(r'著作权\d+-(.+)$', name_without_ext)
    if match:
        return match.group(1)
    return name_without_ext

def add_title_line(doc, text):
    """添加标题行（等线字体，五号，加粗）"""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = '等线'
    run._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '等线')
    run.font.size = Pt(10.5)  # 五号字
    run.bold = True

def add_empty_line(doc):
    """添加空行（设置字体避免默认明朝）"""
    empty_p = doc.add_paragraph()
    empty_run = empty_p.add_run()
    empty_run.font.name = '等线'
    empty_run._element.rPr.rFonts.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}eastAsia', '等线')

def add_pdf_image(cell, pdf_path, seq_id):
    """将PDF转换为图片并插入到表格单元格"""
    try:
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # PDF 转图片
        images = convert_from_path(pdf_path, first_page=1, last_page=1)
        temp_img = f"temp_sc_{seq_id}.jpg"
        images[0].save(temp_img, 'JPEG')
        
        # 插入图片，宽度 11cm
        run = p.add_run()
        run.add_picture(temp_img, width=Cm(11))
        
        # 清理临时图片
        if os.path.exists(temp_img):
            os.remove(temp_img)
            
    except Exception as e:
        cell.text = f"PDF读取失败: {e}"
        print(f"   ❌ PDF错误: {e}")

def process_sc_pairs(max_count=None):
    """处理软件著作权文件，两两配对生成文档"""
    # 获取所有PDF文件
    files = [f for f in os.listdir(SC_SOURCE_FOLDER) if f.lower().endswith('.pdf')]
    
    # 按序号排序
    files.sort(key=extract_sort_id)
    
    # 限制处理数量（测试用）
    if max_count:
        files = files[:max_count]
    
    total_to_process = len(files)
    print(f"🚀 开始处理 {total_to_process} 个软件著作权文件，两两配对...")
    
    # 两两配对处理
    pair_count = 0
    for i in range(0, total_to_process, 2):
        pair_count += 1
        
        # 获取当前配对的两个文件
        file1 = files[i]
        file2 = files[i + 1] if i + 1 < total_to_process else None
        
        pdf1_path = os.path.join(SC_SOURCE_FOLDER, file1)
        name1 = extract_sc_name(file1)
        
        # 构建输出文件名
        if file2:
            pdf2_path = os.path.join(SC_SOURCE_FOLDER, file2)
            name2 = extract_sc_name(file2)
            output_filename = f"附件 {name1}+{name2}.docx"
            print(f"[{pair_count}] 配对: {name1} + {name2}")
        else:
            # 只剩一个文件，单独成文档
            output_filename = f"附件 {name1}.docx"
            print(f"[{pair_count}] 单个: {name1}")
        
        # 创建 Word 文档
        doc = Document()
        setup_page_layout(doc)
        
        # 添加第一个软著的标题
        add_title_line(doc, f"附件 {name1}")
        
        # 如果有第二个软著，添加第二个标题（紧接上一行，不加空行）
        if file2:
            add_title_line(doc, f"附件 {name2}")
        
        # 添加一个空行后再放表格
        add_empty_line(doc)
        
        # 创建表格：1行2列（左右并排）
        if file2:
            # 两个软著：1行2列，每个单元格一个PDF
            table = doc.add_table(rows=1, cols=2)
            
            # 左列：第一个软著的PDF
            cell_left = table.cell(0, 0)
            add_pdf_image(cell_left, pdf1_path, f"pair{pair_count}_1")
            
            # 右列：第二个软著的PDF
            cell_right = table.cell(0, 1)
            add_pdf_image(cell_right, pdf2_path, f"pair{pair_count}_2")
        else:
            # 只有一个软著：1行2列，左边放PDF，右边空白
            table = doc.add_table(rows=1, cols=2)
            
            cell_left = table.cell(0, 0)
            add_pdf_image(cell_left, pdf1_path, f"pair{pair_count}_1")
        
        # 设置表格自适应
        table.autofit = True
        
        # 保存文档
        output_path = os.path.join(OUTPUT_SC_FOLDER, output_filename)
        doc.save(output_path)
    
    print(f"\n🎉 处理完毕！共生成 {pair_count} 个文档。请检查 {OUTPUT_SC_FOLDER} 文件夹。")

if __name__ == '__main__':
    # 在这里修改处理数量（测试用）
    # None = 处理全部，2 = 只处理前2个文件
    process_sc_pairs(max_count=None)
