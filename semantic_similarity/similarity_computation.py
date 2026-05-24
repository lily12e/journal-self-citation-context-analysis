"""
最佳句子嵌入方法：首末层平均 + whitening (匹配)
基于实验结果选择的最优方案
"""

import openpyxl
import torch
import numpy as np
from transformers import BertTokenizer, BertModel
from scipy.spatial.distance import cosine
from all_utils import load_whiten, transform_and_normalize
import sys
import re
import time
from typing import List, Tuple

# 初始化分词器和模型
BERT_PATH = r'D:\2Tools\Model\bert-base-uncased'
tokenizer = BertTokenizer.from_pretrained(BERT_PATH, clean_up_tokenization_spaces=True)
model = BertModel.from_pretrained(BERT_PATH)

# 加载whitening参数
kernel, bias = load_whiten('../bert-base-uncased-first_last_avg-whiten(NLI).pkl')


def is_chinese(text):
    """检查文本是否为中文"""
    return re.match("^[\u4e00-\u9fff]+$", text) is not None


def chunk_text(text: str, max_length: int = 450) -> List[str]:
    """
    将长文本分块处理，避免超过BERT的最大长度限制
    预留一些token给[CLS]和[SEP]，所以使用450而不是512
    """
    if not text or not isinstance(text, str):
        return [""]

    # 先用句号分割
    sentences = text.split('. ')
    if len(sentences) == 1:
        # 如果没有句号，按空格分割
        words = text.split()
        chunks = []
        current_chunk = []
        current_length = 0

        for word in words:
            word_tokens = len(tokenizer.tokenize(word))
            if current_length + word_tokens > max_length and current_chunk:
                chunks.append(' '.join(current_chunk))
                current_chunk = [word]
                current_length = word_tokens
            else:
                current_chunk.append(word)
                current_length += word_tokens

        if current_chunk:
            chunks.append(' '.join(current_chunk))

        return chunks if chunks else [text[:1000]]  # 最后的保险措施

    # 按句子组合成块
    chunks = []
    current_chunk = []
    current_length = 0

    for sentence in sentences:
        sentence_tokens = len(tokenizer.tokenize(sentence))
        if current_length + sentence_tokens > max_length and current_chunk:
            chunks.append('. '.join(current_chunk) + '.')
            current_chunk = [sentence]
            current_length = sentence_tokens
        else:
            current_chunk.append(sentence)
            current_length += sentence_tokens

    if current_chunk:
        chunks.append('. '.join(current_chunk) + ('.' if not current_chunk[-1].endswith('.') else ''))

    return chunks if chunks else [text]


def get_first_last_avg_embedding(text: str) -> np.ndarray:
    """
    使用首末层平均方法获取句子嵌入
    如果文本过长，会进行分块处理并平均
    """
    if not text or not isinstance(text, str):
        return np.zeros(768)  # BERT-base的隐藏层维度

    chunks = chunk_text(text)
    chunk_embeddings = []

    for chunk in chunks:
        if not chunk.strip():
            continue

        try:
            inputs = tokenizer(
                chunk,
                return_tensors='pt',
                truncation=True,
                padding=True,
                max_length=512
            )

            with torch.no_grad():
                outputs = model(**inputs, output_hidden_states=True)

            # 获取第一层和最后一层的CLS token
            first_layer_cls = outputs.hidden_states[1][:, 0, :]  # 第一层CLS
            last_layer_cls = outputs.last_hidden_state[:, 0, :]  # 最后一层CLS

            # 计算平均
            avg_embedding = (first_layer_cls + last_layer_cls) / 2
            chunk_embeddings.append(avg_embedding)

        except Exception as e:
            print(f"处理文本块时出错: {e}")
            print(f"文本块内容: {chunk[:100]}...")
            continue

    if not chunk_embeddings:
        return np.zeros(768)

    # 如果有多个块，计算平均
    if len(chunk_embeddings) > 1:
        final_embedding = torch.mean(torch.stack(chunk_embeddings), dim=0)
    else:
        final_embedding = chunk_embeddings[0]

    return final_embedding


def calculate_similarity_first_last_avg_whitening(text1: str, text2: str) -> float:
    """
    使用首末层平均 + whitening方法计算两个文本的语义相似度
    """
    try:
        # 获取两个文本的嵌入
        embedding1 = get_first_last_avg_embedding(text1)
        embedding2 = get_first_last_avg_embedding(text2)

        # 转换为torch tensor并添加batch维度
        embeddings = torch.stack([
            torch.tensor(embedding1).unsqueeze(0),
            torch.tensor(embedding2).unsqueeze(0)
        ])

        # 应用whitening变换
        embeddings = transform_and_normalize(embeddings, kernel=kernel, bias=bias)

        # 计算余弦相似度
        emb1 = embeddings[0].flatten().detach().cpu().numpy()
        emb2 = embeddings[1].flatten().detach().cpu().numpy()

        similarity = 1 - cosine(emb1, emb2)
        return float(similarity)

    except Exception as e:
        print(f"计算相似度时出错: {e}")
        return 0.0


def process_excel_file(file_path: str, citation_col: str = 'G', abstract_col: str = 'C', output_col: str = 'O'):
    """
    处理Excel文件，计算指定列之间的语义相似度

    Args:
        file_path: Excel文件路径
        citation_col: 引文内容列（默认I列 - citation content filtered）
        abstract_col: 摘要列（默认J列 - Self-citing Article Abstract）
        output_col: 输出列（默认K列）
    """
    print("=" * 80)
    print("使用最佳方法计算语义相似度：首末层平均 + whitening")
    print("=" * 80)
    print(f"文件路径: {file_path}")
    print(f"引文内容列: {citation_col}")
    print(f"摘要列: {abstract_col}")
    print(f"输出列: {output_col}")

    try:
        # 加载Excel文件
        wb = openpyxl.load_workbook(file_path)
        ws = wb.active
        print(f"成功加载Excel文件，共{ws.max_row}行数据")

        # 统计信息
        total_rows = 0
        processed_rows = 0
        chinese_rows = 0
        error_rows = 0
        similarities = []

        start_time = time.time()

        # 处理每一行（跳过标题行）
        for row_num in range(2, ws.max_row + 1):
            total_rows += 1

            # 获取单元格值
            citation_cell = ws[f'{citation_col}{row_num}']
            abstract_cell = ws[f'{abstract_col}{row_num}']
            output_cell = ws[f'{output_col}{row_num}']

            citation_text = citation_cell.value
            abstract_text = abstract_cell.value

            # 转换为字符串
            citation_text = str(citation_text) if citation_text is not None else ''
            abstract_text = str(abstract_text) if abstract_text is not None else ''

            # 检查是否为空
            if not citation_text.strip() or not abstract_text.strip():
                output_cell.value = 0.0
                continue

            # 检查是否为中文
            if is_chinese(citation_text):
                output_cell.value = citation_text  # 保持原样
                chinese_rows += 1
                continue

            try:
                # 计算相似度
                similarity = calculate_similarity_first_last_avg_whitening(
                    abstract_text, citation_text
                )
                output_cell.value = similarity
                similarities.append(similarity)
                processed_rows += 1

                # 每100行显示一次进度
                if processed_rows % 100 == 0:
                    elapsed_time = time.time() - start_time
                    print(f"已处理 {processed_rows} 行，当前相似度: {similarity:.4f}，"
                          f"用时: {elapsed_time:.1f}s")

            except Exception as e:
                print(f"处理第{row_num}行时出错: {e}")
                output_cell.value = 0.0
                error_rows += 1

        # 保存文件
        wb.save(file_path)

        # 计算统计信息
        total_time = time.time() - start_time

        print("\n" + "=" * 80)
        print("处理完成！")
        print("=" * 80)
        print(f"总行数: {total_rows}")
        print(f"成功处理: {processed_rows}")
        print(f"中文行数: {chinese_rows}")
        print(f"错误行数: {error_rows}")
        print(f"总用时: {total_time:.1f}秒")
        print(f"平均每行用时: {total_time / max(total_rows, 1):.2f}秒")

        if similarities:
            similarities = np.array(similarities)
            print(f"\n相似度统计:")
            print(f"  均值: {np.mean(similarities):.4f}")
            print(f"  标准差: {np.std(similarities):.4f}")
            print(f"  最小值: {np.min(similarities):.4f}")
            print(f"  最大值: {np.max(similarities):.4f}")
            print(f"  中位数: {np.median(similarities):.4f}")

        print(f"\n文件已保存: {file_path}")

    except Exception as e:
        print(f"处理文件时出现错误: {e}")
        import traceback
        traceback.print_exc()


def main():
    """主函数"""
    # Excel文件路径
   XLSX_PATH = "./input/your_annotation_file.xlsx"

    # 列设置（根据你的需求调整）
    abstract_col = 'D'  # Self-citing Article Abstract 列
    citation_col = 'H'  # citation content (cleaned) 列
    output_col = 'O'  # 输出相似度的列

    print("句子嵌入语义相似度计算器")
    print("方法：首末层平均 + Whitening")
    print("基于实验验证的最佳方案")
    print()

    # 处理Excel文件
    process_excel_file(file_path, citation_col, abstract_col, output_col)


if __name__ == "__main__":
    main()