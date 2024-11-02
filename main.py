import os
from docx2pdf import convert

def convert_to_pdf(word_file):
    #将Word文档转换为PDF
    convert(word_file)

def main():
    #输入要转换的Word文档路径
    word_file = input("请输入要转换为PDF的Word文档路径：")
    
    if os.path.exists(word_file):
        convert_to_pdf(word_file)
        print("转换成功！")
    else:
        print("文件不存在！")

if __name__ == "__main__":
    main()