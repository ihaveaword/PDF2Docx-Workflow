# 文件名: 01_parse_list.py
import re
import pandas as pd
import os

INPUT_FILE = 'raw_list.txt'
OUTPUT_EXCEL = 'task_list.xlsx'

def parse_text_to_excel():
    if not os.path.exists(INPUT_FILE):
        print(f"❌ 错误：找不到 {INPUT_FILE}，请确认文件名")
        return

    with open(INPUT_FILE, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    data = []
    current_category = "未知分类"
    
    # 你的文本里有些是用横杠，有些可能是空格，这里统一处理
    # 核心逻辑：提取 "序号"、"标题"、"专利号"
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # 识别大分类 (比如 "授权专利-发明专利")
        # 逻辑：如果一行包含 "专利" 或 "著作权"，且不以 "数字." 开头，则它是分类行
        is_item = bool(re.match(r'^(\d+)\.', line))
        if not is_item and ("专利" in line or "著作权" in line):
            # 去掉开头可能存在的特殊符号（如 , o, [], 0 等）
            current_category = re.sub(r'^[^\w\u4e00-\u9fa5]+', '', line).strip()
            continue

        # 跳过软著
        if "软件著作权" in current_category:
            continue

        # === 核心正则匹配 ===
        # 匹配模式：数字. 作者 - 标题 - 专利号
        # 例子：9.陈永伟... -设备管理方法... -发明专利ZL...
        # 考虑到分隔符混乱，我们用 split 配合正则
        
        # 1. 提取开头的序号
        id_match = re.match(r'^(\d+)\.', line)
        if not id_match:
            continue
        
        seq_id = id_match.group(1) # 提取出 "9"
        
        # 2. 提取标题
        # 逻辑：标题通常在第一个和第二个“-”之间，或者在人名之后
        parts = re.split(r'[-—]', line) # 按横杠切割
        
        if len(parts) >= 3:
            # parts[0]: "9.陈永伟..."
            # parts[1]: "设备管理方法..." (通常是标题)
            # parts[2]: "发明专利ZL..." (包含专利号)
            
            title = parts[1].strip()
            patent_part = parts[2].strip()
            
            # 提取纯净的专利号
            p_match = re.search(r'ZL[\d\.\sX]+', patent_part + (parts[3] if len(parts)>3 else ""))
            patent_no = p_match.group(0).replace(" ", "") if p_match else "未识别"

            data.append({
                "Sort_ID": int(seq_id), # 排序用的ID
                "Category": current_category,
                "Title": title,
                "Patent_No": patent_no,
                "Raw_Line": line
            })

    # 生成 Excel
    df = pd.DataFrame(data)
    df.to_excel(OUTPUT_EXCEL, index=False)
    print(f"✅ 步骤1完成！已提取 {len(df)} 条数据到 '{OUTPUT_EXCEL}'")
    print("请打开 Excel 检查一下 Title 列是否正确，如果不正确可以手动微调。")

if __name__ == '__main__':
    parse_text_to_excel()