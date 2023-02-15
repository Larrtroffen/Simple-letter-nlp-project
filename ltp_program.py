import torch.cuda
from ltp import LTP

ltp = LTP("../LTP/base")

# 检测显卡cuda是否可用
if torch.cuda.is_available():
    ltp("cuda")
    
# 将导入的文本分词
def segment(text):
    output = ltp.pipeline([text])
    return output[0]
    
 
# 输出结果到excel表格中
import pandas as pd
def save_data_to_excel(path, data):
    df = pd.read_excel(path, header=None)
    df = df.append(data, ignore_index=True)
    df.to_excel(path, index=False, header=False)
