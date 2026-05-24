import os
import json
import random
import pandas as pd
from datetime import datetime

# ========= 配置 =========
INPUT_XLSX = "./training_data/your_training_set.xlsx"
OUTPUT_DIR = "./training_data/"
RANDOM_SEED  = 42
VAL_RATIO    = 0

# 混合提示：长版占比（建议 0.2~0.4）
USE_LONG_RATE = 0.5

# ========= 允许值 =========
FN_OPTS = {
    "Foundation","Inspiration","Extension","Application","Elaborated Citation",
    "Comparison","Similarity","Affirmation","Related work","Simple mention",
    "Comparison between Related Work","Future work","Further reading","Historical background",
}
DEPTH_OPTS = {"Deep citation","Moderate citation","Shallow citation"}

# ========= system 提示 =========
SYSTEM_LONG = r"""
# Role
You are an expert in scientometrics and citation analysis with extensive experience in citation annotation.

# Task 
Your task is to use the provided information, together with your training knowledge, to annotate the most appropriate citation function and citation depth for each cited work in the citing article.

# Citation Function (Definition)
The citation function refers to the role or purpose of a citation within the citing text. In other words, it addresses the questions: Why did the author cite this work? What role does the citation play in the argument or the overall research structure?

Citation Function Options:
- Foundation
- Inspiration
- Extension
- Application
- Elaborated Citation
- Comparison
- Similarity
- Affirmation
- Related work
- Simple mention
- Comparison between Related Work
- Future work
- Further reading
- Historical background

# Citation Function Category Definitions
Foundation: The citing work takes the cited work as the starting point of its research. The cited work provides the basis and premise of the citing work. Generally, the citing work is regarded as a continuation of the cited work.
Inspiration: The citing work refers to the cited work to gain conceptual insights or ideas, such as motivating research questions or formulating hypotheses. The cited work functions as a source of intellectual stimulation, not as something directly applied in the research process.
Extension: The citing work identifies the limitations of the cited work and then extends it.
Application: The citing work employs the cited work in a practical or functional way within its own study. The cited work is not only mentioned but is used to carry out part of the research (e.g., by applying its approaches, frameworks, or resources).
Elaborated Citation: The citing work provides a lengthy description of the cited work, but the reference does not play a role in research design, analysis, or theoretical support. Used only when no functional role applies. 
Comparison: The citing work compares its own methods, results, or conclusions with those of the cited work, thereby underscoring the features of the cited study.
Similarity: The citing work and the cited work are similar, for example in research findings, research methods, research design, or research perspectives, but does not use them to support arguments or construct its research framework.
Affirmation: The citing work supports its own research by drawing on the conclusions of the cited work, for example by citing its empirical findings or theoretical arguments.
Related work: The citing work introduces cited studies, in order to provide background literature and show the academic context.
Simple mention: The main feature is that the citation is only mentioned in passing and does not substantively contribute to the research. Such citations are brief in length and lack detailed explanation, discussion, or application.
Comparison between Related Work: The citing work compares two or more cited studies with each other, without involving its own study.
Future work: The citing work refers to the cited work in relation to possible directions for further research.
Further reading: The cited work is provided as a reference for readers to consult for additional information.
Historical background: The citing work uses the cited work to provide historical, policy, or societal context. Its function is to situate the research within a broader environment, not to support arguments, methods, or findings. Historical Background typically appears in the Introduction section.

#Distinction
Application and related work: Application emphasizes “use” → the cited work’s contribution is directly integrated into the research process.Related Work emphasizes “introduction” → the cited work contextualizes the study but does not directly drive it.
Simple mention and historical background: Historical Background provides contextual information (history, policy, social phenomena) that frames the study’s environment. 
Inspiration and extension: if it sparks → Inspiration; If it builds after limitations → Extension.

# Citation Depth (Definition)
Citation depth aims to reveal the extent to which cited knowledge is utilized and assimilated within the citing document. It reflects the degree of linguistic engagement with the cited content and the functional value attributed to it by the citing author(s).

Citation Depth Options:
- Deep citation
- Moderate citation
- Shallow citation

# Citation Depth Category Definitions, Examples, and Notes

Deep Citation: A citation that arises from the need for research innovation. It usually refers to literature that is inspirational for the study or supportive in a way that highlights the value of the new research. Such citations are of critical importance to the citing work. 
Example: “Taylor and Mott [1] recognized at an early stage the important role of the Coulomb force in internal conversion. Without considering relativistic effects, multiple internal conversion would mainly arise from the Coulomb force. In contrast, Rose [6] argued that the Coulomb force does not play a role in internal conversion. Due to gauge invariance, longitudinal photons make no contribution. Tralli and Goerizel [3] maintained that, due to selection rules, longitudinal photons cannot affect electric multipole or magnetic multipole internal conversion. Therefore, it is necessary to discuss: (1) whether the Coulomb force should be considered in the calculation of internal conversion coefficients, and (2) if so, whether the Coulomb force would make an important contribution.” 
Note: In this case, the citing work raises two research questions on the basis of the divergent viewpoints presented in the cited literature and develops its study from there. These references can to some extent be regarded as the origin of the research problem, and thus are considered deep citations.

Moderate Citation: A citation that arises from the needs of research and argumentation. The citing work usually presents the content of the cited work objectively and clearly, or applies its data, methods, or theories, but does not develop, extend, or innovate on this basis.
Example: “According to Woodward et al. [11], since the refractive index of the solution differs from that of pure carbon tetrachloride, the refraction of the two also differs when the scattered light emerges from the scattering cell. Therefore, the volume of the scattering liquid ‘seen’ by the spectrometer is also different, and thus the recorded scattered light is considered to be proportional to n².”
Note: In this case, the citing work provides a concise explanation of the cited study’s results and applies them in its own research, but does not further explore or develop these results. The cited work thus serves as essential and organizational material required for the research process, and this type of reference is regarded as a moderate citation.

Shallow Citation: A citation that arises during the writing process mainly for narrative or stylistic purposes, or when the author cites a work secondhand. Such citations largely remain at the level of thematic generalization or mere mention, and are not essential to the article.
Example: “For a long time, many researchers in speech acoustics have worked on automatic speech recognition [1–15]. Because the length of pronunciation poses great difficulties for recognition, many approaches have been used in the past to address this problem, such as end-cutting [3], zero-padding [3], linear stretching [3,10], and logarithmic compression of the time axis [8].”
Note: In this case, the citing work merely lists several references without providing detailed discussion, explanation, or application of their contents. Although such citations may offer readers additional material to understand the background, their practical utility from an evaluative perspective is low. This type of reference is therefore regarded as a shallow citation.

No extra text. Output strictly as JSON.
"""

SYSTEM_SHORT = r"""
You are an expert in scientometrics and citation annotation.
Task: Given (Self-cited Article Index/Title, Citation Location, Citation Content),
output ONLY a JSON with keys "Final Citation Function" and "Final Citation Depth".
Options:
- Final Citation Function ∈ {Foundation, Inspiration, Extension, Application, Elaborated Citation,
  Comparison, Similarity, Affirmation, Related work, Simple mention,
  Comparison between Related Work, Future work, Further reading, Historical background}
- Final Citation Depth ∈ {Deep citation, Moderate citation, Shallow citation}
No extra text. Output strictly as JSON.
"""

def norm_str(x: str) -> str:
    return "" if x is None else str(x).strip()

def valid_or_raise(label: str, opts: set, kind: str):
    if label not in opts:
        raise ValueError(f"[{kind}] 非法标签：{label} ；允许值：{sorted(list(opts))}")

def build_user_prompt(row) -> str:
    idx = norm_str(row.get("Self-cited Article Index", ""))  # APA 可为空字符串
    title = norm_str(row.get("Self-cited Article Title", ""))
    content = norm_str(row.get("Citation Content", ""))
    location = norm_str(row.get("citation location", ""))
    return (
        f"Self-cited Article Index: {idx}\n"
        f"Self-cited Article Title: {title}\n"
        f"Citation Location: {location}\n"
        f"Citation Content: {content}"
    )

def build_assistant_json(function_en: str, depth_en: str) -> str:
    return json.dumps({
        "Final Citation Function": function_en,
        "Final Citation Depth": depth_en,
    }, ensure_ascii=False)

# ——关键改动：按行号稳定选择 system（可复现）——
def pick_system_for_row(row_idx: int) -> str:
    return SYSTEM_LONG if (hash((row_idx, RANDOM_SEED)) % 1000)/1000.0 < USE_LONG_RATE else SYSTEM_SHORT

def to_messages(row_idx: int, row, fn_en: str, depth_en: str) -> dict:
    return {
        "messages": [
            {"role": "system", "content": pick_system_for_row(row_idx)},
            {"role": "user", "content": build_user_prompt(row)},
            {"role": "assistant", "content": build_assistant_json(fn_en, depth_en)},
        ]
    }

def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    df = pd.read_excel(INPUT_XLSX, engine="openpyxl")

    # 必要列检查
    must_cols_in = [
        "Self-cited Article Index",
        "Self-cited Article Title",
        "Citation Content",
        "citation location",
        "citation function",
        "citation depth",
    ]
    miss = [c for c in must_cols_in if c not in df.columns]
    if miss:
        raise RuntimeError(f"Excel 缺少必要列：{miss}")

    examples, bad_rows = [], []
    for i, row in df.iterrows():
        fn_en = norm_str(row.get("citation function", ""))
        dp_en = norm_str(row.get("citation depth", ""))

        try:
            valid_or_raise(fn_en, FN_OPTS, "Function")
            valid_or_raise(dp_en, DEPTH_OPTS, "Depth")
        except Exception as e:
            bad_rows.append((i, str(e), fn_en, dp_en))
            continue

        examples.append(to_messages(i, row, fn_en, dp_en))  # 传入 i 以选择 system

    if bad_rows:
        print("⚠️ 发现非法标签样本（已跳过）：")
        for r in bad_rows[:10]:
            print("  行号/错误/Function/Depth =>", r)
        print(f"总计 {len(bad_rows)} 条。")

    # 切分 80/20（可复现）
    random.seed(RANDOM_SEED)
    random.shuffle(examples)
    n_total = len(examples)
    n_val = max(1, int(round(n_total * VAL_RATIO)))
    valid_set = examples[:n_val]
    train_set = examples[n_val:]

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    train_path = os.path.join(OUTPUT_DIR, f"train_{ts}.jsonl")
    valid_path = os.path.join(OUTPUT_DIR, f"valid_{ts}.jsonl")

    with open(train_path, "w", encoding="utf-8") as f:
        for ex in train_set:
            f.write(json.dumps(ex, ensure_ascii=False) + "\n")
    with open(valid_path, "w", encoding="utf-8") as f:
        for ex in valid_set:
            f.write(json.dumps(ex, ensure_ascii=False) + "\n")

    print("✅ 转换完成：")
    print(f"  总样本：{n_total}，训练：{len(train_set)}，验证：{len(valid_set)}")
    print(f"  训练集：{train_path}")
    print(f"  验证集：{valid_path}")

if __name__ == "__main__":
    main()
