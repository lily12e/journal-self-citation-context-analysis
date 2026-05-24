# train_finetune.py
# 功能：上传训练/验证 JSONL → 创建微调任务 → 轮询状态 → 打印新模型名
# 依赖：pip install --upgrade openai
# 环境：设置好环境变量 OPENAI_API_KEY

import os
# Set your OPENAI_API_KEY as an environment variable before running
import time
from openai import OpenAI
from typing import Optional

# ========= 必填：本地文件路径 =========
TRAIN_PATH = "./training_data/batch3_train_53.jsonl"
VALID_PATH = "./training_data/test_set_84.jsonl"

# ========= 可调参数 =========
BASE_MODEL   = "ft:gpt-4o-2024-08-06:personal:citation-func-depth-v1:CGmPftcm"
JOB_SUFFIX   = "citation-func-depth-v3"   # 微调任务后缀，便于识别
N_EPOCHS     = 1                          # 2-4 起步较常见
LR_MULT      = 0.5                        # 初始学习率倍率
BATCH_SIZE   = "auto"                     # 让平台自动选择
POLL_SECS    = 15                         # 轮询间隔（秒）

def assert_file(path: str):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"未找到文件：{path}")

def upload_file(client: OpenAI, path: str, purpose: str = "fine-tune") -> str:
    print(f"⬆️ 正在上传：{path}")
    with open(path, "rb") as f:
        fobj = client.files.create(file=f, purpose=purpose)
    print(f"   → file.id = {fobj.id}")
    return fobj.id

def create_job(client: OpenAI,
               train_file_id: str,
               valid_file_id: Optional[str] = None):
    print("🚀 正在创建微调任务...")
    job = client.fine_tuning.jobs.create(
        model=BASE_MODEL,
        training_file=train_file_id,
        validation_file=valid_file_id,
        suffix=JOB_SUFFIX,
        hyperparameters={
            "batch_size": BATCH_SIZE,
            "learning_rate_multiplier": LR_MULT,
            "n_epochs": N_EPOCHS
        }
    )
    print(f"   → job.id = {job.id}")
    return job

def poll_job(client: OpenAI, job_id: str, interval_sec: int = 15):
    """轮询直到完成；返回最终的 fine_tuned_model 或抛错"""
    printed_event_ids = set()
    while True:
        info = client.fine_tuning.jobs.retrieve(job_id)
        status = info.status
        if status in ("succeeded", "failed", "cancelled"):
            # 打印最后一波事件
            events = client.fine_tuning.jobs.list_events(job_id, limit=50)
            for e in reversed(events.data):
                if e.id not in printed_event_ids:
                    print(f"[{e.created_at}] {e.type} | {e.message}")
                    printed_event_ids.add(e.id)

            print(f"✅ 任务结束，状态：{status}")
            if status == "succeeded":
                print(f"🆕 fine_tuned_model = {info.fine_tuned_model}")
                return info.fine_tuned_model
            else:
                raise RuntimeError(f"微调失败或被取消，status={status}")
        else:
            # 打印新增事件
            events = client.fine_tuning.jobs.list_events(job_id, limit=20)
            for e in reversed(events.data):
                if e.id not in printed_event_ids:
                    print(f"[{e.created_at}] {e.type} | {e.message}")
                    printed_event_ids.add(e.id)
            time.sleep(interval_sec)

def main():
    # 0) API Key 检查
    if not os.getenv("OPENAI_API_KEY"):
        raise EnvironmentError("未检测到 OPENAI_API_KEY，请先设置环境变量。")

    # 1) 路径检查
    assert_file(TRAIN_PATH)
    assert_file(VALID_PATH)

    client = OpenAI()

    # 2) 上传文件
    train_file_id = upload_file(client, TRAIN_PATH)
    valid_file_id = upload_file(client, VALID_PATH)

    # 3) 创建微调任务
    job = create_job(client, train_file_id, valid_file_id)

    # 4) 轮询直到完成，返回新模型名
    model_name = poll_job(client, job.id, interval_sec=POLL_SECS)

    # 5) 可选：保存模型名到本地 txt，便于后续推理脚本读取
    out_txt = os.path.join(os.path.dirname(TRAIN_PATH), "fine_tuned_model.txt")
    with open(out_txt, "w", encoding="utf-8") as f:
        f.write(model_name or "")
    print(f"📄 模型名已写入：{out_txt}")

if __name__ == "__main__":
    main()
