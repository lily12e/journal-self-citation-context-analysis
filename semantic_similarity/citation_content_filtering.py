import pandas as pd
from openai import OpenAI
from tqdm import tqdm
import os
from datetime import datetime
from dotenv import load_dotenv

# 1️⃣ 初始化 OpenAI 客户端
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")
client = OpenAI(api_key=api_key)

# 2️⃣ 读取 excel 文件
df = pd.read_excel("./input/your_annotation_file.xlsx")
# 8️⃣ 设置输出目录
output_dir = "./output"
os.makedirs(output_dir, exist_ok=True)

# ✨ 新增（紧跟在读表后）
df["Citation Content(±3)"] = df["Citation Content(±3)"].fillna("").astype(str)
# 3️⃣ 定义全局 system prompt（指令信息放这里，避免重复输入）
system_prompt = """
【Role】
You are an expert in citation analysis and academic text processing with extensive experience in bibliometrics research. Your specialty is extracting precise citation contexts while maintaining strict adherence to linguistic inclusion rules.

【Task Overview】
Extract citation content/context for a given reference from academic text, following specific linguistic rules for including surrounding sentences.

【Key Concept】
Citation content = the citation sentence itself + any qualifying surrounding sentences based on linguistic markers.

【Extraction Rules - Follow These Steps】
STEP 1: Locate and always include the citation sentence containing the target reference mark.

STEP 2: Analyze preceding sentences for inclusion:
- Check if the citation sentence contains ANY of these linguistic elements:
  • Conjunctive adverbs: Use your complete internal knowledge of discourse connectives 
    (including but not limited to: however, therefore, moreover, furthermore, nevertheless, 
    consequently, similarly, in contrast, in addition, additionally, likewise, thus, hence, etc.)
  • Demonstrative pronouns: this, that, these, those, such
  • Third-person pronouns: it, they, them, their, its, theirs, he, she, him, her, his, hers
  • First-person plural pronouns when referring to research: we, our, us
- IF the citation sentence contains any of the above → include the immediately preceding sentence
- IF that preceding sentence ALSO contains any of the above elements → include ITS preceding sentence too

STEP 3: Analyze following sentences:
- Check the immediately following sentence for:
  • Demonstrative pronouns that refer back to the citation content
  • Third-person pronouns that refer back to the citation content  
  • Conjunctive adverbs that continue the discourse thread
  • Any pronoun or discourse marker that creates clear cohesive links back to the citation

- IF present → include the following sentence

【Strict Constraints】
Never rewrite, paraphrase, summarize, or alter the original wording.
Output must be verbatim text from the input.
Each block of output = one citation content (citation sentence + qualifying surrounding sentences).
"""

# 4️⃣ 生成动态的 user prompt
def build_user_prompt(self_cited_article, citation_context):
    return f"""Self-cited article title and authors: {self_cited_article}
Citation context (given text segment): {citation_context}
"""

# 5️⃣ 统计变量
total_input_tokens = 0
total_output_tokens = 0

# 6️⃣ GPT 调用函数（替换整个函数体）
def call_gpt(user_prompt, fallback_text, retries=2):
    global total_input_tokens, total_output_tokens
    for i in range(retries):
        try:
            resp = client.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=400,
            )
            content = (resp.choices[0].message.content or "").strip()
            usage = getattr(resp, "usage", None)
            total_input_tokens += (getattr(usage, "prompt_tokens", 0) or 0)
            total_output_tokens += (getattr(usage, "completion_tokens", 0) or 0)
            return content if content else fallback_text
        except Exception as e:
            if i == retries - 1:
                print("API Error:", e)
                with open("gpt_error_log.txt", "a", encoding="utf-8") as log:
                    log.write("=" * 60 + "\n")
                    log.write(f"❌ GPT 调用失败时间：{datetime.now()}\n")
                    log.write(f"Prompt 内容：\n{user_prompt}\n")
                    log.write(f"错误信息：{e}\n")
                    log.write("=" * 60 + "\n\n")
                return fallback_text


# 7️⃣ 生成 user prompt 并调用 GPT 生成 filtered 内容
tqdm.pandas()
df["user_prompt"] = df.progress_apply(
    lambda row: build_user_prompt(row["Self-cited Article Title"], row["Citation Content(±3)"]),
    axis=1
)
df["citation content (filtered)"] = df.progress_apply(
    lambda row: call_gpt(row["user_prompt"], row["Citation Content(±3)"]),
    axis=1
)



# 9️⃣ 生成当前时间戳
timestamp = datetime.now().strftime("%Y%m%d_%H%M")
output_path = os.path.join(output_dir, f"citations-filtered-{timestamp}.xlsx")

# 🔟 移除 "user_prompt" 列，避免冗余信息输出
df.drop(columns=["user_prompt"], inplace=True)

# 将 "citation content (filtered)" 插入到 "Citation Content(±3)" 列后
citation_index = df.columns.get_loc("Citation Content(±3)")
filtered_series = df["citation content (filtered)"]
df.drop(columns=["citation content (filtered)"], inplace=True)
df.insert(citation_index + 1, "citation content (filtered)", filtered_series)

# 保存为 Excel 文件（.xlsx）
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Filtered Results")
print(f"✅ 已生成清洗结果：{output_path}")

# 🔥 打印 Token 使用与费用
total_tokens = total_input_tokens + total_output_tokens
input_cost = total_input_tokens / 1000 * 0.01  # $0.01 per 1K input tokens
output_cost = total_output_tokens / 1000 * 0.03  # $0.03 per 1K output tokens
total_cost = input_cost + output_cost
# 🔥 打印 Token 使用与费用（按 3.5 的价格！）
# total_tokens = total_input_tokens + total_output_tokens
# input_cost = total_input_tokens / 1000 * 0.0015  # $0.0015 per 1K input tokens (3.5)
# output_cost = total_output_tokens / 1000 * 0.002  # $0.002 per 1K output tokens (3.5)
# total_cost = input_cost + output_cost


print(f"\n📊 Token 使用统计：")
print(f"  - 输入 tokens：{total_input_tokens}")
print(f"  - 输出 tokens：{total_output_tokens}")
print(f"  - 总 tokens：{total_tokens}")

print(f"\n💰 费用估算：")
print(f"  - 输入成本：${input_cost:.4f}")
print(f"  - 输出成本：${output_cost:.4f}")
print(f"  - 总成本：${total_cost:.4f}")
