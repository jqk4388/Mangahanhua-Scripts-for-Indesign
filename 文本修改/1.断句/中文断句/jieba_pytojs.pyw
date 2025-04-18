import sys
import jieba
import os

# 默认输入和输出文件路径
default_input_file = os.path.join(os.getenv('LOCALAPPDATA'), 'Temp', 'jieba_temp_input.txt')
default_output_file = os.path.join(os.getenv('LOCALAPPDATA'), 'Temp', 'jieba_temp_output.txt')

try:
    # 检查参数数量
    if len(sys.argv) == 3:
        input_file = sys.argv[1]
        output_file = sys.argv[2]
    else:
        # print("参数错误，使用默认路径")
        input_file = default_input_file
        output_file = default_output_file

    # 读取输入文件内容
    with open(input_file, 'r', encoding='utf-8') as f:
        text = f.read().strip()
    
    # 使用 jieba 分词
    seg_list = jieba.cut(text)
    
    # 用 \r 连接分词结果
    result = '\r'.join(seg_list)
    
    # 写入输出文件
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(result)
        
    # print(f"分词完成，结果已写入 {output_file}")
except FileNotFoundError as e:
    # print(f"文件未找到: {str(e)}")
    sys.exit(1)
except Exception as e:
    # print(f"Error: {str(e)}")
    sys.exit(1)