import streamlit as st
import os
import tempfile
from pdf2docx import Converter
from docx import Document
from io import BytesIO
from openai import OpenAI
from zhipuai import ZhipuAI
import concurrent.futures

# 设置页面标题和布局
st.set_page_config(page_title="PDF翻译工具", layout="wide")
st.title("PDF 转 Word 并翻译工具")

# API配置
DOUBAO_API_KEY = "ad8016f6-641c-494f-8f2d-1425d8dd2e08"  # 请替换为你的实际API密钥
DOUBAO_API_URL = "https://ark.cn-beijing.volces.com/api/v3"  # 假设的API地址，请替换为实际地址
ZHIPU_API_KEY = "22077bfa73e94c518393f921d00d1d48.7PIvlmDe1ynfEtot"  # 请替换为你的实际API密钥

def pdf_to_word(pdf_path, word_path):
    """将PDF转换为Word文档"""
    cv = Converter(pdf_path)
    cv.convert(word_path, start=0, end=None)
    cv.close()

def extract_text_from_word(word_path):
    """从Word文档中提取文本（按段落）"""
    doc = Document(word_path)
    paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
    return doc, paragraphs, doc.tables

def doubao_translate_text(text, target_lang='zh'):
    """使用豆包API翻译文本"""
    client = OpenAI(
        api_key=DOUBAO_API_KEY,
        base_url=DOUBAO_API_URL,
    )
    completion = client.chat.completions.create(
        model='doubao-1-5-lite-32k-250115',
        messages=[
            {"role": "system", "content": """# 角色
你是一名军事专家，按照规则将用户上传的文本内容翻译成中文，需要地道表达。
# 规则
1. 中文版的内容需要符合中国的法律，尤其是对于台海问题，香港问题等涉及到意识形态的问题。
2. 只输出中文翻译文本，不需要添加注释以及自己的思考过程。"""},
            {"role": "user", "content": f"{text}"},
        ],
    )
    try:
        if completion.choices[0].message.content:
            return completion.choices[0].message.content
        else:
            return text
    except Exception as e:
        print(f'豆包翻译出错: {e}')
        return text

def zhipu_translate_text(text):
    """使用智谱API翻译文本"""
    client = ZhipuAI(api_key=ZHIPU_API_KEY)
    response = client.chat.completions.create(
        model="glm-4-flash",
        temperature=0.8,
        max_tokens=4095,
        messages=[
            {"role": "system", "content": """# 角色
你是一名军事专家，按照规则将用户上传的文本内容翻译成中文，需要地道表达。
# 规则
1. 中文版的内容需要符合中国的法律，尤其是对于台海问题，香港问题等涉及到意识形态的问题。
2. 只输出中文翻译文本，不需要添加注释以及自己的思考过程。"""},
            {"role": "user", "content": f"{text}"}
        ],
    )
    if response and response.choices and len(response.choices) > 0:
        return response.choices[0].message.content
    else:
        return text

def process_cell(cell, target_lang):
    """处理单个单元格的翻译"""
    try:
        if cell.text.strip():
            # 使用智谱翻译表格内容
            translated_text = zhipu_translate_text(cell.text)
            if '很抱歉，您提供的信息似乎不完整。请上传需要翻译的文本内容，我将根据您的要求进行翻译。' not in translated_text:
                if len(cell.paragraphs) > 0:
                    cell.paragraphs[0].text = translated_text
                    # 清除多余段落
                    for para in cell.paragraphs[1:]:
                        para.clear()
                else:
                    cell.text = translated_text
    except Exception as e:
        print(f"处理单元格失败: {e}")

def process_table(table, target_lang):
    """处理单个表格的所有单元格"""
    # 收集所有单元格
    cells = []
    for row in table.rows:
        for cell in row.cells:
            cells.append(cell)
    
    # 使用线程池处理单元格
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = [executor.submit(process_cell, cell, target_lang) for cell in cells]
        for future in concurrent.futures.as_completed(futures):
            future.result()

def translate_word_document(doc, paragraphs, source_lang, target_lang):
    """翻译Word文档中的文本并保留格式"""
    st.info("开始翻译文档...")
    progress_bar = st.progress(0)
    total_paragraphs = len([p for p in doc.paragraphs if p.text.strip()])
    processed = 0
    
    for paragraph in doc.paragraphs:
        if paragraph.text.strip():
            # 使用豆包翻译文档内容
            translated_text = doubao_translate_text(paragraph.text, target_lang)
            paragraph.text = translated_text
            processed += 1
            progress_bar.progress(processed / total_paragraphs)
    
    # 翻译表格
    if doc.tables:
        st.info("开始翻译表格...")
        table_progress = st.progress(0)
        total_tables = len(doc.tables)
        
        with concurrent.futures.ThreadPoolExecutor() as executor:
            futures = [executor.submit(process_table, table, target_lang) 
                      for table in doc.tables]
            for i, _ in enumerate(concurrent.futures.as_completed(futures)):
                table_progress.progress((i + 1) / total_tables)
        
        table_progress.empty()
    
    progress_bar.empty()
    st.success("翻译完成!")
    return doc

def save_word_document(doc, output_path):
    """保存Word文档"""
    doc.save(output_path)

def main():
    # 文件上传区域
    uploaded_file = st.file_uploader("上传PDF文件", type=["pdf"])
    
    # 语言选择
    col1, col2 = st.columns(2)
    with col1:
        source_lang = st.selectbox("源语言", ["en", "zh"], index=0)
    with col2:
        target_lang = st.selectbox("目标语言", ["zh", "en"], index=0)
    
    if uploaded_file is not None:
        # 创建临时文件
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(uploaded_file.getvalue())
            pdf_path = tmp_pdf.name
        
        # 转换为Word
        word_path = pdf_path.replace(".pdf", ".docx")
        pdf_to_word(pdf_path, word_path)
        
        # 提取文本
        doc, paragraphs, tables = extract_text_from_word(word_path)
        
        # 显示原始文本预览
        with st.expander("原始文本预览"):
            st.write("\n\n".join(paragraphs[:10]))  # 只显示前10段
        
        # 翻译按钮
        if st.button("开始翻译"):
            translated_doc = translate_word_document(doc, paragraphs, source_lang, target_lang)
            
            # 保存翻译后的文档
            translated_word_path = word_path.replace(".docx", f"_translated_{target_lang}.docx")
            save_word_document(translated_doc, translated_word_path)
            
            # 提供下载
            with open(translated_word_path, "rb") as f:
                bytes_data = f.read()
            
            st.download_button(
                label="下载翻译后的Word文档",
                data=BytesIO(bytes_data),
                file_name=os.path.basename(translated_word_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            # 清理临时文件
            os.unlink(pdf_path)
            os.unlink(word_path)
            os.unlink(translated_word_path)

if __name__ == "__main__":
    main()