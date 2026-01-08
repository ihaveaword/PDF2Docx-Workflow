import os
import re
from docx import Document
from docxcompose.composer import Composer

# === 配置区域 ===
SOURCE_FOLDER = 'Output_Docs2'      # 源文件夹
OUTPUT_FILE = '专利合并文档.docx'    # 输出文件名

def extract_sort_id(filename):
    """从文件名中提取序号用于排序"""
    match = re.search(r'^(\d+)\.', filename)
    if match:
        return int(match.group(1))
    return 9999

def extract_patent_type(filename):
    """
    从文件名中提取专利类型
    返回类型序号用于排序：
    1 = 发明专利
    2 = 实用新型专利
    3 = 外观设计专利
    4 = PCT专利
    """
    if '发明专利' in filename:
        return 1
    elif '实用新型' in filename:
        return 2
    elif '外观设计' in filename:
        return 3
    elif 'PCT' in filename or 'pct' in filename.lower():
        return 4
    else:
        return 9999  # 其他类型放在最后

def custom_sort_key(filename):
    """
    自定义排序键：先按专利类型，再按序号
    """
    patent_type = extract_patent_type(filename)
    sort_id = extract_sort_id(filename)
    return (patent_type, sort_id)

def merge_documents(max_count=None):
    """合并多个Word文档"""
    # 获取所有文档
    files = [f for f in os.listdir(SOURCE_FOLDER) if f.endswith('.docx')]
    
    # 按专利类型和序号排序
    files.sort(key=custom_sort_key)
    
    # 限制处理数量（测试用）
    if max_count:
        files = files[:max_count]
    
    total_files = len(files)
    print(f"🚀 准备合并 {total_files} 个文档...")
    
    if not files:
        print("❌ 没有找到文档！")
        return
    
    # 使用第一个文档作为基础
    first_file_path = os.path.join(SOURCE_FOLDER, files[0])
    master_doc = Document(first_file_path)
    composer = Composer(master_doc)
    print(f"[1/{total_files}] 基础文档: {files[0]}")
    
    # 合并剩余文档
    for index, filename in enumerate(files[1:], start=2):
        file_path = os.path.join(SOURCE_FOLDER, filename)
        print(f"[{index}/{total_files}] 添加: {filename}")
        
        try:
            # 读取要合并的文档
            doc_to_append = Document(file_path)
            
            # 使用 docxcompose 的 append 方法合并文档
            # 这会自动处理图片关系和其他资源
            composer.append(doc_to_append)
            
        except Exception as e:
            print(f"   ❌ 处理失败: {e}")
            continue
    
    # 保存合并后的文档
    composer.save(OUTPUT_FILE)
    print(f"\n🎉 合并完成！输出文件: {OUTPUT_FILE}")

if __name__ == '__main__':
    # 测试模式：合并前2个文件
    # 修改为 None 可合并全部文件
    merge_documents(max_count=None)  # 合并全部文件
