import subprocess
import sys
import pandas as pd
import numpy as np
from transformers import T5Tokenizer, T5ForConditionalGeneration, Trainer, TrainingArguments
from datasets import Dataset
import torch
from rouge_score import rouge_scorer
import os

# Hàm cài đặt các thư viện cần thiết
def install_dependencies():
    try:
        print("Đang cài đặt các thư viện cần thiết...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "transformers", "datasets", "pandas", "torch", "openpyxl", "rouge-score", "numpy"])
        print("Cài đặt hoàn tất!")
    except subprocess.CalledProcessError as e:
        print(f"Lỗi khi cài đặt thư viện: {e}")
        sys.exit(1)

import pandas as pd
import numpy as np
from transformers import T5Tokenizer, T5ForConditionalGeneration, Trainer, TrainingArguments
from datasets import Dataset
import torch
from rouge_score import rouge_scorer
import os

# 1. Đọc và tiền xử lý dữ liệu
def load_data(file_path):
    # Đọc file Excel, chỉ định dòng đầu tiên là tiêu đề
    df = pd.read_excel(file_path, header=0)  # header=0 nghĩa là dòng 0 là tiêu đề
    # Lấy dữ liệu từ cột 2 (văn bản gốc) và cột 3 (tóm tắt), bỏ qua tiêu đề
    data = {
        "input_text": df.iloc[:, 1].tolist(),  # Cột 2
        "target_text": df.iloc[:, 2].tolist()  # Cột 3
    }
    return Dataset.from_dict(data)

# 2. Chuẩn bị dữ liệu cho mô hình T5
def preprocess_data(dataset, tokenizer, max_input_length=1500, max_target_length=128):
    def tokenize_function(examples):
        inputs = ["summarize: " + doc for doc in examples["input_text"]]
        model_inputs = tokenizer(inputs, max_length=max_input_length, truncation=True, padding="max_length")
        
        with tokenizer.as_target_tokenizer():
            labels = tokenizer(examples["target_text"], max_length=max_target_length, truncation=True, padding="max_length")
        
        model_inputs["labels"] = labels["input_ids"]
        return model_inputs

    tokenized_dataset = dataset.map(tokenize_function, batched=True)
    return tokenized_dataset

# 3. Fine-tune mô hình T5
def fine_tune_t5(dataset):
    model_name = "t5-base"
    tokenizer = T5Tokenizer.from_pretrained(model_name)
    model = T5ForConditionalGeneration.from_pretrained(model_name)

    tokenized_dataset = preprocess_data(dataset, tokenizer, max_input_length=1500)

    train_test_split = tokenized_dataset.train_test_split(test_size=0.2)
    train_dataset = train_test_split["train"]
    eval_dataset = train_test_split["test"]

    training_args = TrainingArguments(
        output_dir="./t5_finetuned",
        evaluation_strategy="epoch",
        learning_rate=5e-5,
        per_device_train_batch_size=1,
        per_device_eval_batch_size=1,
        gradient_accumulation_steps=4,
        num_train_epochs=3,
        weight_decay=0.01,
        save_strategy="epoch",
        load_best_model_at_end=True,
        metric_for_best_model="eval_loss",
        logging_dir="./logs",
        logging_steps=10,
    )

    trainer = Trainer(
        model=model,
        args=training_args,
        train_dataset=train_dataset,
        eval_dataset=eval_dataset,
    )

    trainer.train()
    model.save_pretrained("./t5_finetuned_final")
    tokenizer.save_pretrained("./t5_finetuned_final")
    return model, tokenizer

# 4. Đánh giá mô hình bằng ROUGE
def evaluate_model(model, tokenizer, dataset):
    model.eval()
    predictions = []
    references = dataset["target_text"]

    for example in dataset["input_text"]:
        inputs = tokenizer("summarize: " + example, return_tensors="pt", max_length=1500, truncation=True)
        inputs = {k: v.to(model.device) for k, v in inputs.items()}
        outputs = model.generate(**inputs, max_length=128, num_beams=4, early_stopping=True)
        pred = tokenizer.decode(outputs[0], skip_special_tokens=True)
        predictions.append(pred)

    scorer = rouge_scorer.RougeScorer(['rouge1', 'rouge2', 'rougeL'], use_stemmer=True)
    rouge_scores = {"rouge1": [], "rouge2": [], "rougeL": []}

    for ref, pred in zip(references, predictions):
        scores = scorer.score(ref, pred)
        for key in rouge_scores.keys():
            rouge_scores[key].append(scores[key].fmeasure)

    avg_rouge = {key: np.mean(scores) for key, scores in rouge_scores.items()}
    return avg_rouge

# 5. Chạy chương trình
def main():
    file_path = "dataset.xlsx"
    dataset = load_data(file_path)
    model, tokenizer = fine_tune_t5(dataset)
    rouge_scores = evaluate_model(model, tokenizer, dataset)
    print("ROUGE Scores:", rouge_scores)

if __name__ == "__main__":
    main()