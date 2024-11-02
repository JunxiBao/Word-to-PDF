import pdfkit
from docx import Document
import os
from docx.shared import Inches

def save_images_from_word(doc, output_dir):
    """提取 Word 文档中的图像并保存到指定目录"""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    image_index = 0
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            image_index += 1
            image_name = f'image_{image_index}.png'
            image_path = os.path.join(output_dir, image_name)
            with open(image_path, "wb") as img_file:
                img_file.write(rel.target_part.blob)
            yield image_name, os.path.abspath(image_path)  # 返回绝对路径

def word_to_pdf(word_file, output_pdf):
    # 打开 Word 文件
    doc = Document(word_file)
    
    # 创建临时图像保存目录
    image_dir = 'temp_images'
    
    # 临时生成一个HTML文件
    temp_html = 'temp_file.html'
    
    # 把 Word 内容写入 HTML 文件，包含 UTF-8 编码和字体设置
    with open(temp_html, 'w', encoding='utf-8') as f:
        f.write('<html><head>')
        f.write('<meta charset="utf-8">')
        f.write('<style> body { font-family: "SimSun", "Noto Sans CJK", sans-serif; } </style>')
        f.write('</head><body>\n')
        
        # 写入文本内容
        for para in doc.paragraphs:
            f.write(f"<p>{para.text}</p>\n")
        
        # 写入图像内容
        for image_name, image_path in save_images_from_word(doc, image_dir):
            f.write(f'<p><img src="{image_path}" style="max-width:600px;"></p>\n')
        
        f.write('</body></html>')
    
    # 设置 wkhtmltopdf 选项，确保正确处理文件协议
    options = {
        'enable-local-file-access': None,  # 允许访问本地文件
    }

    # 使用 pdfkit 把 HTML 转换为 PDF，并传入选项
    pdfkit.from_file(temp_html, output_pdf, options=options)
    
    # 删除临时的 HTML 文件和图像目录
    os.remove(temp_html)
    for img_file in os.listdir(image_dir):
        os.remove(os.path.join(image_dir, img_file))
    os.rmdir(image_dir)

# 示例用法
word_file = '神经冲动的产生和传导课后习题.docx'
output_pdf = 'output_pdf_file.pdf'
word_to_pdf(word_file, output_pdf)
print(f'{output_pdf} 转换成功！')
