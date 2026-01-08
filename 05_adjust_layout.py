import os
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT

# === 配置区域 ===
SOURCE_FOLDER = 'Output_Docs'       # 源文件夹
OUTPUT_FOLDER = 'Output_Docs2'      # 输出文件夹

# 确保输出目录存在
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def adjust_document_layout(input_path, output_path):
    """
    调整Word文档的布局：
    1. 修改页边距：上1cm，下0.8cm，左右1.27cm
    2. 修改图片尺寸：高18cm，宽12.7cm（不锁定纵横比）
    3. 删除标题后的空行
    """
    try:
        # 打开文档
        doc = Document(input_path)
        
        # === 1. 调整页边距 ===
        for section in doc.sections:
            section.top_margin = Cm(1.0)      # 上边距 1cm
            section.bottom_margin = Cm(0.8)   # 下边距 0.8cm
            section.left_margin = Cm(1.27)    # 左边距 1.27cm
            section.right_margin = Cm(1.27)   # 右边距 1.27cm
        
        # === 2. 删除空段落（标题后的空行）===
        # 从后向前遍历，避免索引问题
        paragraphs_to_remove = []
        for i, para in enumerate(doc.paragraphs):
            # 如果段落为空（没有文本且没有图片）
            if len(para.text.strip()) == 0 and len(para.runs) == 0:
                paragraphs_to_remove.append(para)
        
        # 删除空段落
        for para in paragraphs_to_remove:
            p_element = para._element
            p_element.getparent().remove(p_element)
        
        # === 3. 调整所有图片的尺寸 ===
        for shape in doc.inline_shapes:
            # 设置图片尺寸：高18cm，宽12.7cm
            shape.height = Cm(18.0)
            shape.width = Cm(12.7)
        
        # 保存文档
        doc.save(output_path)
        return True
        
    except Exception as e:
        print(f"   ❌ 处理失败: {e}")
        return False

def process_documents(max_count=None):
    """处理文档，调整布局"""
    # 获取所有 .docx 文件
    files = [f for f in os.listdir(SOURCE_FOLDER) if f.endswith('.docx')]
    files.sort()
    
    # 限制处理数量（测试用）
    if max_count:
        files = files[:max_count]
    
    total_to_process = len(files)
    print(f"🚀 开始处理 {total_to_process} 个Word文档...")
    
    success_count = 0
    for index, filename in enumerate(files):
        input_path = os.path.join(SOURCE_FOLDER, filename)
        output_path = os.path.join(OUTPUT_FOLDER, filename)
        
        print(f"[{index+1}/{total_to_process}] 处理: {filename}")
        
        if adjust_document_layout(input_path, output_path):
            success_count += 1
    
    print(f"\n🎉 处理完毕！成功 {success_count}/{total_to_process} 个文档。")
    print(f"   输出目录: {OUTPUT_FOLDER}")

if __name__ == '__main__':
    # 测试模式：只处理1个文件
    # 修改为 None 可处理全部文件
    process_documents(max_count=None)  # 处理全部文件
