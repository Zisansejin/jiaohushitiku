import requests
import os
from dotenv import load_dotenv
load_dotenv()

class MultiAIClient:
    def __init__(self, model_type: str, api_key: str):
        self.model_type = model_type
        self.api_key = api_key
        self.base_url = ""
        self.model_name = ""
        self.headers = {"Content-Type": "application/json"}
        self.headers["Authorization"] = f"Bearer {self.api_key}"
        if model_type == "deepseek":
            self.base_url = "https://api.deepseek.com/v1/chat/completions"
            self.model_name = "deepseek-chat"
        elif model_type == "openai":
            self.base_url = "https://api.openai.com/v1/chat/completions"
            self.model_name = "gpt-3.5-turbo"
        elif model_type == "tongyi":
            self.base_url = "https://dashscope.aliyuncs.com/api/v1/services/aigc/text-generation/generation"
            self.model_name = "qwen-turbo"
        elif model_type == "ernie":
            self.base_url = "https://qianfan.baidubce.com/v2/chat/completions"
            self.model_name = "ernie-3.5"

    def chat_completion(self, prompt: str, temperature=0.3):
        sys_prompt = "你是资深临床药学教学专家，严格按照药学规培大纲、处方审核规范出题，答案解析专业严谨，分级清晰，格式规整。"
        payload = {
            "model": self.model_name,
            "messages": [
                {"role": "system", "content": sys_prompt},
                {"role": "user", "content": prompt}
            ],
            "temperature": temperature,
            "max_tokens": 4096
        }
        resp = requests.post(self.base_url, headers=self.headers, json=payload, timeout=60)
        resp.raise_for_status()
        data = resp.json()
        return data["choices"][0]["message"]["content"]