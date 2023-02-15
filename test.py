import os
import docx
import pandas as pd

# 读取docx文件并提取前五行
def extract_data_from_docx(path):
    doc = docx.Document(path)
    data = []
    for i in range(5):
        data.append(doc.paragraphs[i].text)
    return data

# 新建excel表格
def create_excel(path, data):
    df = pd.DataFrame(data)
    df.to_excel(path, index=False, header=False)

# 存储数据到表格中
def save_data_to_excel(path, data):
    df = pd.read_excel(path, header=None)
    df = df.append(data, ignore_index=True)
    df.to_excel(path, index=False, header=False)
    
# 去除文本中的*号，并以原文件名输出
def remove_star(path, data):
    for i in range(len(data)):
        data[i] = data[i].replace('*', '')
    name = os.path.basename(path)
    name = name.replace('.docx', '.txt')
    with open(name, 'w') as f:
        for i in range(len(data)):
            f.write(data[i] + '\n')

# 批量读取docx文件，并进行上述处理
def process_docx_files(path):
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith('.docx'):
                data = extract_data_from_docx(os.path.join(root, file))
                remove_star(os.path.join(root, file), data)
                if os.path.exists('data.xlsx'):
                    save_data_to_excel('data.xlsx', data)
                else:
                    create_excel('data.xlsx', data)