# 文件名: 03_generate_links.py
import os
import re
from urllib.parse import quote

OUTPUT_FOLDER = 'Output_Docs'  # 从这个文件夹读取生成的文件
HTML_OUTPUT = '03_专利快速查询.html'

def extract_patent_info(filename):
    """
    从文件名提取信息
    例如: "1. 授权专利-发明专利1-基于核电厂真实工况与仿真系统的故障诊断方法及系统-ZL201310699915.1.docx"
    返回: {
        'seq_id': 1,
        'clean_name': '授权专利-发明专利1-基于核电厂真实工况与仿真系统的故障诊断方法及系统-ZL201310699915.1',
        'patent_no': 'ZL201310699915.1'
    }
    """
    # 去掉 .docx 后缀
    name_without_ext = filename.replace('.docx', '')
    
    # 提取序号（文件名开头的数字）
    seq_match = re.match(r'^(\d+)\.\s*(.+)$', name_without_ext)
    if seq_match:
        seq_id = int(seq_match.group(1))
        clean_name = seq_match.group(2)  # 去掉序号后的文件名
    else:
        seq_id = 0
        clean_name = name_without_ext
    
    # 提取专利号 (ZL开头的)
    patent_match = re.search(r'(ZL[\d\.X]+)', clean_name)
    patent_no = patent_match.group(1) if patent_match else ''
    
    return {
        'seq_id': seq_id,
        'clean_name': clean_name,
        'patent_no': patent_no,
        'original_filename': filename
    }

def generate_search_url(clean_name, patent_no):
    """
    生成专利查询链接
    使用 Uyanip（领导推荐）- 直接搜索专利号
    """
    # 清理专利号：去掉 ZL 前缀，保留小数点
    clean_patent_no = patent_no.replace('ZL', '').strip()
    
    # 使用 Uyanip 搜索页面（领导推荐）
    # 格式: https://www.uyanip.com/search/keyword?q=专利号
    url = f"https://www.uyanip.com/search/keyword?q={quote(clean_patent_no)}"
    
    return url

def generate_html(max_count=None):
    """生成专利查询辅助HTML页面"""
    if not os.path.exists(OUTPUT_FOLDER):
        print(f"❌ 错误：找不到文件夹 {OUTPUT_FOLDER}")
        return
    
    # 获取所有 .docx 文件
    files = [f for f in os.listdir(OUTPUT_FOLDER) if f.endswith('.docx')]
    # 按序号数字排序（而不是字母排序）
    files.sort(key=lambda f: extract_patent_info(f)['seq_id'])
    
    # 限制处理数量（测试模式）
    if max_count:
        files = files[:max_count]
    
    print(f"🔍 扫描到 {len(files)} 个文件，开始生成链接...")
    
    # HTML 头部
    html_content = """
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>专利快速查询助手</title>
        <style>
            body {{ 
                font-family: '微软雅黑', sans-serif; 
                padding: 20px; 
                background: #f5f5f5;
            }}
            h2 {{ color: #333; }}
            .container {{ max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; }}
            .item {{ 
                margin-bottom: 15px; 
                padding: 15px; 
                border: 1px solid #ddd; 
                border-radius: 5px; 
                background: #fafafa;
                transition: background 0.2s;
                display: flex;
                align-items: center;
                flex-wrap: wrap;
            }}
            .item:hover {{ background: #f0f0f0; }}
            .seq {{ 
                display: inline-block;
                width: 40px;
                font-weight: bold; 
                color: #007bff; 
                font-size: 16px;
            }}
            .patent-no {{ 
                display: inline-block;
                background: #007bff;
                color: white;
                padding: 5px 10px;
                border-radius: 3px;
                margin: 0 10px;
                font-size: 14px;
                font-weight: bold;
            }}
            .copy-btn {{ 
                background: #28a745; 
                color: white;
                border: none;
                padding: 6px 12px; 
                border-radius: 4px; 
                font-size: 14px;
                cursor: pointer;
                display: inline-flex;
                align-items: center;
                gap: 5px;
                transition: background 0.2s;
            }}
            .copy-btn:hover {{ background: #218838; }}
            .copy-btn.copied {{ background: #6c757d; }}
            .title {{ 
                color: #666; 
                font-size: 14px; 
                margin-top: 8px;
                width: 100%;
            }}
            .link-btn {{
                background: #17a2b8;
                color: white;
                text-decoration: none;
                padding: 6px 12px;
                border-radius: 4px;
                font-size: 14px;
                margin-left: 10px;
                display: inline-flex;
                align-items: center;
                gap: 5px;
            }}
            .link-btn:hover {{ background: #138496; }}
        </style>
    </head>
    <body>
        <div class="container">
            <h2>📋 专利快速查询助手</h2>
            <p>共 <strong>{count}</strong> 个专利，点击"📋 复制"按钮复制专利号，然后手动访问 <a href="https://www.uyanip.com/search/keyword" target="_blank">Uyanip</a> 查询。</p>
            <hr>
    """.format(count=len(files))
    
    # 生成每个专利项
    for filename in files:
        info = extract_patent_info(filename)
        # 纯数字专利号（去掉 ZL）
        clean_patent_no = info['patent_no'].replace('ZL', '').strip()
        # Uyanip 首页（简化版）
        search_url = "https://www.uyanip.com/"
        
        html_content += f"""
        <div class="item">
            <span class="seq">{info['seq_id']}</span>
            <span class="patent-no">{info['patent_no']}</span>
            <button class="copy-btn" onclick="copyPatentNo('{clean_patent_no}', this)">📋 复制</button>
            <a href="{search_url}" target="_blank" class="link-btn">🔍 查询</a>
            <div class="title">{info['clean_name']}</div>
        </div>
        """
    
    # HTML 尾部 + JavaScript
    html_content += """
        </div>
        <script>
            function copyPatentNo(patentNo, button) {{
                // 复制到剪贴板
                navigator.clipboard.writeText(patentNo).then(function() {{
                    // 修改按钮状态
                    const originalText = button.innerHTML;
                    button.innerHTML = '✅ 已复制';
                    button.classList.add('copied');
                    
                    // 2秒后恢复
                    setTimeout(function() {{
                        button.innerHTML = originalText;
                        button.classList.remove('copied');
                    }}, 2000);
                }}).catch(function(err) {{
                    alert('复制失败：' + err);
                }});
            }}
        </script>
    </body>
    </html>
    """
    
    # 保存文件
    with open(HTML_OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print(f"✅ 成功生成：{HTML_OUTPUT}")
    print(f"   包含 {len(files)} 个专利的查询链接")

if __name__ == '__main__':
    # 在这里修改处理数量（测试用）
    # None = 处理全部，2 = 只处理前2个
    generate_html(max_count=None)