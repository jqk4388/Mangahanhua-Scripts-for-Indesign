#author: 几千块
#version: 1.1.1
#description: 基于jieba的文本分词工具，词性标注、合并词语、自动分句等功能。
import os
import sys
import jieba.posseg as pseg
import argparse
from tkinter import messagebox
import tkinter as tk

# 句末标点列表
sentence_endings_lianyong = {"…", "。", "！", "!", "？", "?", "，", "；", "：", "、", "—", "”", "）", "】", "｝", "」", "』", "》", ")"}
sentence_endings = {"…", "。", "！", "!", "？", "?", "，", "；", "：", "、", "—", "”"} #分句用标点
# 根据系统类型获取临时目录
if sys.platform.startswith('win'):
    # Windows 系统：使用 LOCALAPPDATA\Temp（兼容原逻辑）
    temp_dir = os.path.join(os.getenv('LOCALAPPDATA', 'C:\\Temp'), 'Temp')
else:
    # macOS/Linux 系统：使用系统默认临时目录（TMPDIR 或 /tmp）
    temp_dir = os.getenv('TMPDIR', '/tmp')

# 生成输入/输出文件路径（保持文件名不变）
default_input_file = os.path.join(temp_dir, 'jieba_temp_input.txt')
default_output_file = os.path.join(temp_dir, 'jieba_temp_output.txt')
# 调试用
# default_input_file = "M:\\汉化\\PS_PNG\\断句测试.txt"
# default_output_file = "M:\\汉化\\PS_PNG\\断句测试_jieba_output.txt"

#词性合并
def process_sentence(word_list):
    """对单句进行词性合并"""
    units = []
    i = 0
    n = len(word_list)
    while i < n:
        word, pos = word_list[i]
        
        # 处理空格，将其转换为断句标识
        if word.isspace():
            if units:
                units[-1] += "\\r"
            else:
                units.append("\\r")
            i += 1
            continue

        # 处理动词+数词+动词结构
        if pos.startswith('v') and i + 2 < n:
            m_word, m_pos = word_list[i+1]
            v2_word, v2_pos = word_list[i+2]
            if m_pos == 'm' and v2_pos.startswith('v'):
                units.append(word + m_word + v2_word)
                i += 3
                continue

        # 处理专有名词
        if pos.startswith(('nr', 'ns', 'nt', 'nz')):
            units.append(word)
            i += 1
            continue
        
        # 处理习用语/成语
        if pos in ('l', 'i'):
            units.append(word)
            i += 1
            continue
        
        # 处理代词+动词+名词结构
        if pos == 'r' and i + 2 < n:
            v_word, v_pos = word_list[i+1]
            n_word, n_pos = word_list[i+2]
            if v_pos.startswith('v') and n_pos.startswith('n'):
                units.append(word + v_word + n_word)
                i += 3
                continue
        
        # 处理代词+量词+名词结构：这扇窗户
        if pos == 'r' and i + 2 < n:
            v_word, v_pos = word_list[i+1]
            n_word, n_pos = word_list[i+2]
            if v_pos.startswith('q') and n_pos.startswith('n'):
                units.append(word + v_word + n_word)
                i += 3
                continue
        
        # 处理名词+动词短语+的结构
        if pos.startswith('n'):
            j = i + 1
            while j < n and word_list[j][1] != 'uj':
                j += 1
            if j < n:
                units.append(''.join([w for w, p in word_list[i:j+1]]))
                i = j + 1
                continue
        
        # 处理名词加方位词
        if pos == 'n' and i + 1 < n and word_list[i+1][1] == 'f':
            units.append(word + word_list[i+1][0])
            i += 2
            continue
                
        # 处理副词+动词+了结构：突然响了
        if pos == 'ad' and i + 2 < n:
            v_word, v_pos = word_list[i+1]
            n_word, n_pos = word_list[i+2]
            if v_pos.startswith('v') and n_pos.startswith('ul'):
                units.append(word + v_word + n_word)
                i += 3
                continue
        
        # 处理副词修饰结构
        if pos in ('d', 'ad') and i + 1 < n:
            next_word, next_pos = word_list[i+1]
            if next_pos in ('v', 'a'):
                units.append(word + next_word)
                i += 2
                continue
        
        # 处理动词+助词结构
        if pos.startswith('v') and i + 1 < n:
            next_word, next_pos = word_list[i+1]
            if next_pos in ('uz', 'ug', 'ud', 'ul','uj','uv'): #着得过了的
                units.append(word + next_word)
                i += 2
                continue
        
        # 处理副词+助词结构
        if pos.startswith('d') and i + 1 < n:
            next_word, next_pos = word_list[i+1]
            if next_pos in ('uz', 'ug', 'ud', 'ul','uj'): 
                units.append(word + next_word)
                i += 2
                continue
        
        # 处理介词短语
        if pos == 'p' and i + 1 < n:
            units.append(word + word_list[i+1][0])
            i += 2
            continue
        
        # 处理数词+量词
        if pos == 'm' and i + 1 < n and word_list[i+1][1] == 'q':
            units.append(word + word_list[i+1][0])
            i += 2
            continue
        
        # 处理代词+动词
        if pos == 'r' and i + 1 < n and word_list[i+1][1].startswith('v'):
            units.append(word + word_list[i+1][0])
            i += 2
            continue
        
        # 处理代词+助词
        if pos == 'r' and i + 1 < n and word_list[i+1][1] in ('uj', 'ul'):
            units.append(word + word_list[i+1][0])
            i += 2
            continue
        
        # 处理语气词
        if pos in ('y', 'zg') and units:
            units[-1] += word
            i += 1
            continue

        # 处理好+形容词+啊结构
        if pos.startswith('a') and i + 2 < n:
            m_word, m_pos = word_list[i+1]
            v2_word, v2_pos = word_list[i+2]
            if m_pos == 'a' and v2_pos.startswith('y'):
                units.append(word + m_word + v2_word)
                i += 3
                continue

        
        # 处理标点符号
        if pos == 'x':
            if word in sentence_endings_lianyong:
                if units:
                    units[-1] += word
                else:
                    units.append(word)
            else:
                # 定义左右括号对应关系
                brackets = {'（': '）', '『': '』', '《': '》', '「': '」', '【': '】', '(': ')', '[': ']', '{': '}', '“': '”', '‘': '’', '"': '"'}
                if word in brackets or word in brackets.values():  # 处理括号包裹内容
                    if word in brackets:  # 左括号
                        j = i + 1
                        while j < n and word_list[j][0] != brackets[word]:
                            j += 1
                        if j < n:
                            units.append(''.join([w for w, p in word_list[i:j+1]]))
                            i = j + 1
                            continue
                if units:
                    units[-1] += word
                else:
                    units.append(word)
            i += 1
            continue
        
        # 处理普通词语
        units.append(word)
        i += 1
    
    # 合并连续句末标点
    cleaned = []
    for unit in units:
        if cleaned and unit[-1] in sentence_endings and cleaned[-1][-1] in sentence_endings:
            cleaned[-1] = cleaned[-1][:-1] + unit[-1]
        else:
            cleaned.append(unit)
    return [u for u in cleaned if u]  # 过滤空单元

def merge_words(word_list):
    units = []
    sentence = []
    for word, pos in word_list:
        # 处理标点符号
        if pos == 'x' and word in sentence_endings:
            limit = 6 # 字数阈值降低
            if sentence:
                units.extend(process_sentence(sentence))  # 调用函数处理当前句子
                sentence = []
            if units:
                units[-1] += word
            else:
                units.append(word)
        else:
            limit = 11 
            sentence.append((word, pos))
    
    # 处理最后一句话
    if sentence:
        units.extend(process_sentence(sentence))
    
    return units,limit

def split_into_lines(units, limit):
    """按句末标点分句，并根据字数规则分行"""
    sentences = []
    sentence = []
    
    # 按句末标点分句
    for unit in units:
        sentence.append(unit)
        if unit[-1] in sentence_endings:
            sentences.append(sentence)
            sentence = []
    if sentence:  # 处理最后未结束的句子
        sentences.append(sentence)
    
    # 按字数规则分行
    result = []
    for sentence in sentences:
        # 短句处理
        total_length = sum(len(unit) for unit in sentence)
        if len(sentences) == 1 and total_length <= 6:  # 只有一句且字数不超过阈值
            sentence_result = [sentence[0], '\\r'.join(sentence[1:])]
            result.append(''.join(sentence_result))
            break
        
        # 普通分行逻辑
        line = []
        line_length = 0
        sentence_result = []
        for unit in sentence:
            unit_length = len(unit)
            if line_length + unit_length > limit:  # 超过11字换行
                if line == []:
                    line = [unit]
                    line_length = unit_length
                else:
                    sentence_result.append(''.join(line))
                    line = [unit]
                    line_length = unit_length

            else:
                line.append(unit)
                line_length += unit_length
        if line:  # 添加最后一行
            sentence_result.append(''.join(line))
        
        # 调整行数和字数分布
        total_lines = len(sentence_result)
        if total_lines > 6:  # 限制最多6行
            while len(sentence_result) > 6:
                # 找出最短的行及其索引
                shortest_idx = min(range(len(sentence_result)), key=lambda i: len(sentence_result[i]))
                shortest_len = len(sentence_result[shortest_idx])
                
                # 计算与前后行合并的长度
                prev_merged_len = float('inf')
                next_merged_len = float('inf')
                
                if shortest_idx > 0:
                    prev_merged_len = len(sentence_result[shortest_idx - 1]) + shortest_len
                if shortest_idx < len(sentence_result) - 1:
                    next_merged_len = shortest_len + len(sentence_result[shortest_idx + 1])
                
                # 选择较短的合并方案
                if prev_merged_len <= next_merged_len and shortest_idx > 0:
                    sentence_result[shortest_idx - 1] += sentence_result[shortest_idx]
                    sentence_result.pop(shortest_idx)
                else:
                    sentence_result[shortest_idx] += sentence_result[shortest_idx + 1]
                    sentence_result.pop(shortest_idx + 1)
        elif total_lines > 1:
            avg_length = sum(len(l) for l in sentence_result[1:-1]) // max(1, total_lines - 2)
            for i in range(1, total_lines - 1):
                if len(sentence_result[i]) < avg_length - 4:
                    sentence_result[i] += sentence_result[i + 1]
                    sentence_result.pop(i + 1)
                    break
                elif len(sentence_result[i]) > avg_length + 4:
                    overflow = sentence_result[i][avg_length:]
                    sentence_result[i] = sentence_result[i][:avg_length]
                    sentence_result.insert(i + 1, overflow)
                    break
        
        # 合并句中行并用 \r 连接
        result.append('\\r'.join(sentence_result))
    
    # 用 \r 连接每句
    result = '\\r'.join(result)
    
    # 最后检查字数
    total_length = len(result.replace('\\r', ''))
    if 5 <= total_length <= 14 and len(sentence_result) > 2:
        parts = result.split('\\r')
        result = '\\r'.join(parts[:-1]) + parts[-1]  # 保留1个 \r
    elif 15 <= total_length <= 24 and len(sentence_result) > 4:
        parts = result.split('\\r')
        result = '\\r'.join(parts[:-2]) + parts[-2] + parts[-1]  # 保留2个 \r
    
    return result

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', help='输入txt文件路径')
    parser.add_argument('--output', help='输出txt文件路径')
    args = parser.parse_args()
    
    input_path = args.input or default_input_file
    output_path = args.output or default_output_file
    
    
    # 读取输入文件
    with open(input_path, 'r', encoding='utf-8') as f:
        content = f.readlines()
    
    processed_content = []
    for line in content:
        line = line.strip()
        if not line:
            processed_content.append('')
            continue
        
        # 分词处理
        words = pseg.cut(line)
        word_list = [(w.word, w.flag) for w in words]
        
        # 合并短语和标点
        units,limit = merge_words(word_list)
        
        # 断句处理
        if len(line) > 4:
            lines = split_into_lines(units, limit)
            textbox = ''.join(lines)
        else:
            textbox = line
        processed_content.append(textbox)
    
    # 写入输出文件
    with open(output_path, 'w', encoding='utf-8') as f:
        processed_content = [line.replace('\\r\\r', '\\r') for line in processed_content]
        f.write('\n'.join(processed_content))
    
    # 创建并显示5秒后自动关闭的消息框
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    root.after(5000, root.destroy)  # 5秒后销毁窗口
    messagebox.showinfo("结巴断句", "断句已完成！")
    root.mainloop()

if __name__ == '__main__':
    main()