import os
import json
import threading
import tempfile
import time
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import platform
import base64
from concurrent.futures import ThreadPoolExecutor, as_completed

# 跨平台临时文件路径和默认文件名
TMPDIR = tempfile.gettempdir()
DEFAULT_INPUT = os.path.join(TMPDIR, "LLM_input.txt")
DEFAULT_OUTPUT = os.path.join(TMPDIR, "LLM_output.txt")

# 默认API URLs
DEFAULT_APIS = {
    "Ollama": "http://localhost:11434/api/generate",
    "OpenAI": "https://api.openai.com/v1/chat/completions",
    "Doubao": "https://ark.cn-beijing.volces.com/api/v3/chat/completions",
    "DeepSeek": "https://api.deepseek.com/chat/completions",
    "Qianwen": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
    "Baidu": "https://qianfan.baidubce.com/v2/chat/completions",
    "Tencent": "https://api.hunyuan.cloud.tencent.com/v1/chat/completions",
    "Zhipu": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
    "Gemini": "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
}

# 默认模型列表
DEFAULT_MODELS = {
    "Ollama": ["deepseek-v3.2:cloud"],
    "OpenAI": ["gpt-4o", "gpt-3.5-turbo"],
    "Doubao": ["doubao-seed-1-6-251015"],
    "DeepSeek": ["deepseek-chat"],
    "Qianwen": ["qwen-plus", "qwen-max"],
    "Baidu": ["ernie-x1.1-preview"],
    "Tencent": ["hunyuan-2.0-instruct-20251111"],
    "Zhipu": ["glm-4.7"],
    "Gemini": ["gemini-3-flash-preview"]
}

# 系统提示词
DEFAULT_SYSTEM_PROMPT = (
    "You are a sentence-splitting assistant. For the given line of raw Chinese text, "
    "split it into natural sentences with \\r. Consider sentence-final punctuation, morphemes, "
    "and avoid breaking words or meaningful units. Place commas appropriately at the "
    "end of clauses when natural. Keep sentences balanced in length where possible. "
    "Short fragments should be kept as-is (do not force splits). Do not change the original punctuation marks."
    "Output only the finalized split text for each requested line, preserving original order. Do not add extra"
    " commentary. Return plain text." 
    "Example Input:"
    "Line 57: 听说我家附近的寺庙有驱鬼的业务，于是我特地前来拜访。"
    "Line 58: 你知道马是从什么时候开始出现在人类历史中的吗?"
    "Example Output:"
    "听说\\r我家附近的寺庙\\r有驱鬼的业务，\\r于是我特地\\r前来拜访。"
    "你知道\\r马是从什么时候\\r开始出现在人类历史中的吗？"
)

class App:
    def __init__(self, root):
        self.root = root
        root.title("大语言模型断句器")

        # 变量
        self.input_var = tk.StringVar(value=DEFAULT_INPUT)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        self.api_type_var = tk.StringVar(value="Ollama")
        self.api_var = tk.StringVar(value=DEFAULT_APIS["Ollama"])
        self.api_key_var = tk.StringVar(value="")
        self.model_var = tk.StringVar(value="deepseek-v3.2:cloud")
        self.task_size_var = tk.IntVar(value=20)
        self.max_chars_var = tk.IntVar(value=4000)
        self.think_mode_var = tk.BooleanVar(value=False)
        self.prompt_var = tk.StringVar(value=DEFAULT_SYSTEM_PROMPT)
        self.api_keys = {}  # 各平台的API Key

        # 配置
        self.config_file = self.get_config_path()

        # UI
        self.setup_ui()

        self.load_config()

        # 状态
        self._stop_flag = threading.Event()
        self.worker_thread = None
        self.output_lock = threading.Lock()
        self.output_buffer = None
        self._write_ptr = 0
        self.max_retries = 10
        self.failure_counts = {}
        self.total_tokens = 0
        self.start_time = None

    def setup_ui(self):
        # 文件路径
        frame_files = ttk.Frame(self.root)
        frame_files.pack(fill="x", padx=8, pady=6)

        ttk.Label(frame_files, text="输入文件:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_files, textvariable=self.input_var, width=60).grid(row=0, column=1, sticky="we")
        ttk.Button(frame_files, text="浏览", command=self.browse_input).grid(row=0, column=2, padx=4)

        ttk.Label(frame_files, text="输出文件:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame_files, textvariable=self.output_var, width=60).grid(row=1, column=1, sticky="we")
        ttk.Button(frame_files, text="浏览", command=self.browse_output).grid(row=1, column=2, padx=4)

        # 参数
        frame_params = ttk.Frame(self.root)
        frame_params.pack(fill="x", padx=8, pady=6)

        ttk.Label(frame_params, text="API类型:").grid(row=0, column=0, sticky="w")
        self.api_type_cb = ttk.Combobox(frame_params, textvariable=self.api_type_var, values=list(DEFAULT_APIS.keys()), state="readonly", width=15)
        self.api_type_cb.grid(row=0, column=1, sticky="w", padx=4)
        self.api_type_cb.bind("<<ComboboxSelected>>", self.on_api_type_changed)

        ttk.Label(frame_params, text="API地址:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame_params, textvariable=self.api_var, width=50).grid(row=1, column=1, sticky="we", padx=4)

        ttk.Label(frame_params, text="API Key:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame_params, textvariable=self.api_key_var, width=50, show="*").grid(row=2, column=1, sticky="we", padx=4)

        ttk.Label(frame_params, text="模型名:").grid(row=3, column=0, sticky="w")
        self.model_cb = ttk.Combobox(frame_params, textvariable=self.model_var, values=DEFAULT_MODELS["OpenAI"], width=25)
        self.model_cb.grid(row=3, column=1, sticky="w", padx=4)
        ttk.Button(frame_params, text="加载模型", command=self.load_models).grid(row=3, column=2, padx=4)

        ttk.Label(frame_params, text="每任务行数:").grid(row=4, column=0, sticky="w")
        ttk.Spinbox(frame_params, from_=1, to=200, increment=1, textvariable=self.task_size_var, width=8).grid(row=4, column=1, sticky="w", padx=4)

        ttk.Label(frame_params, text="发送字符阈值:").grid(row=4, column=2, sticky="w")
        ttk.Spinbox(frame_params, from_=500, to=20000, increment=500, textvariable=self.max_chars_var, width=10).grid(row=4, column=3, sticky="w", padx=4)

        ttk.Checkbutton(frame_params, text="开启思考（慢）", variable=self.think_mode_var).grid(row=5, column=0, columnspan=2, sticky="w")

        # 高级设置
        self.advanced_shown = False
        ac_frame = ttk.Frame(self.root)
        ac_frame.pack(fill="x", padx=8, pady=2)
        self.adv_button = ttk.Button(ac_frame, text="显示高级设置", command=self.toggle_advanced)
        self.adv_button.pack(side="left")
        self.adv_panel = ttk.Frame(self.root)

        ttk.Label(self.adv_panel, text="系统提示词:").grid(row=0, column=0, sticky="nw")
        ttk.Entry(self.adv_panel, textvariable=self.prompt_var, width=80).grid(row=0, column=1, sticky="we", padx=4)

        root.protocol("WM_DELETE_WINDOW", self.on_closing)

        # 控制
        frame_ctrl = ttk.Frame(self.root)
        frame_ctrl.pack(fill="x", padx=8, pady=6)
        self.progress = ttk.Progressbar(frame_ctrl, length=400, mode="determinate")
        self.progress.grid(row=0, column=0, padx=4, pady=4)
        self.status_label = ttk.Label(frame_ctrl, text="未开始")
        self.status_label.grid(row=0, column=1, padx=8)
        self.start_btn = ttk.Button(frame_ctrl, text="开始处理", command=self.toggle_start)
        self.start_btn.grid(row=0, column=2, padx=4)

    def browse_input(self):
        p = filedialog.askopenfilename(title="选择输入 txt 文件", filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if p:
            self.input_var.set(p)

    def browse_output(self):
        p = filedialog.asksaveasfilename(title="选择输出文件", defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if p:
            self.output_var.set(p)

    def toggle_advanced(self):
        if not self.advanced_shown:
            self.adv_panel.pack(fill="x", padx=8, pady=4)
            self.adv_button.config(text="隐藏高级设置")
            self.advanced_shown = True
        else:
            self.adv_panel.forget()
            self.adv_button.config(text="显示高级设置")
            self.advanced_shown = False

    def on_api_type_changed(self, event=None):
        api_type = self.api_type_var.get()
        self.api_var.set(DEFAULT_APIS.get(api_type, ""))
        self.model_cb['values'] = DEFAULT_MODELS.get(api_type, [])
        self.model_var.set(DEFAULT_MODELS.get(api_type, [""])[0] if DEFAULT_MODELS.get(api_type) else "")
        self.api_key_var.set(self.api_keys.get(api_type, ""))

    def load_models(self):
        api_type = self.api_type_var.get()
        api_url = self.api_var.get().strip()
        api_key = self.api_key_var.get().strip()
        
        # Ollama 不需要 API Key
        if api_type != "Ollama" and (not api_url or not api_key):
            messagebox.showerror("错误", "请先填写API地址和Key")
            return
        elif api_type == "Ollama" and not api_url:
            messagebox.showerror("错误", "请先填写API地址")
            return

        print(f"加载模型列表: {api_type}")
        models = self.fetch_models(api_type, api_url, api_key)
        if models:
            self.model_cb['values'] = models
            self.model_var.set(models[0])
            print(f"加载成功: {len(models)} 个模型")
            # messagebox.showinfo("成功", f"加载了 {len(models)} 个模型")
        else:
            print("加载模型列表失败")
            messagebox.showerror("错误", "无法加载模型列表")

    def fetch_models(self, api_type, api_url, api_key):
        try:
            if api_type == "Ollama":
                tags_url = api_url.replace("/api/generate", "/api/tags")
                resp = requests.get(tags_url, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["name"] for m in data.get("models", [])]
            elif api_type in ["OpenAI", "Doubao", "DeepSeek", "Qianwen", "Zhipu","Baidu"]:
                models_url = api_url.replace("/chat/completions", "/models")
                headers = {"Authorization": f"Bearer {api_key}"}
                resp = requests.get(models_url, headers=headers, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["id"] for m in data.get("data", [])]
            elif api_type == "Tencent":
                return DEFAULT_MODELS.get("Tencent", [])
            elif api_type == "Gemini":
                resp = requests.get(f"{api_url}/models", headers={"x-goog-api-key": api_key}, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    return [m["name"].split("/")[-1] for m in data.get("models", [])]
        except Exception as e:
            print(f"获取模型列表失败: {e}")
        return []

    def toggle_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self._stop_flag.set()
            self.start_btn.config(text="开始处理")
            self.status_label.config(text="停止中...")
        else:
            self._stop_flag.clear()
            self.start_btn.config(text="停止")
            self.worker_thread = threading.Thread(target=self.run_processing, daemon=True)
            self.worker_thread.start()

    def run_processing(self):
        self.start_time = time.time()
        self.total_tokens = 0
        input_path = self.input_var.get()
        output_path = self.output_var.get()
        task_size = self.task_size_var.get()

        print(f"开始处理输入文件: {input_path}")
        print(f"输出文件: {output_path}")
        print(f"每任务行数: {task_size}")

        self.failure_counts = {}  # 初始化失败计数

        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"输入文件不存在: {input_path}")
            self.start_btn.config(text="开始处理")
            return

        with open(input_path, "r", encoding="utf-8") as f:
            lines = [ln.rstrip("\n") for ln in f]

        total = len(lines)
        print(f"总行数: {total}")
        processed_count = 0
        if os.path.exists(output_path):
            with open(output_path, "r", encoding="utf-8") as f:
                out_lines = [ln.rstrip("\n") for ln in f]
                processed_count = len(out_lines)
            print(f"检测到已有输出: {processed_count} 行")
            if processed_count < total:
                if not messagebox.askyesno("续传", f"检测到输出文件已有 {processed_count} 行结果，是否从该行继续处理？"):
                    processed_count = 0
                    open(output_path, "w", encoding="utf-8").close()
                    print("覆盖输出文件，从头开始")
            else:
                restart = messagebox.askyesno("完成", "输出文件已完成所有行。是否重新开始并覆盖已有输出？")
                if restart:
                    processed_count = 0
                    open(output_path, "w", encoding="utf-8").close()
                    print("重新开始，覆盖输出")
                else:
                    messagebox.showinfo("完成", "已取消处理。")
                    self.start_btn.config(text="开始处理")
                    return

        self.output_buffer = lines[:]
        if processed_count > 0:
            for i in range(min(processed_count, total)):
                self.output_buffer[i] = out_lines[i]

        # 过滤短行，直接设置为原文（已经是）
        short_lines = [i for i in range(total) if len(lines[i].strip()) < 5]
        print(f"短行数: {len(short_lines)}")

        # 收集待处理的行（长行）
        pending_indices = [i for i in range(processed_count, total) if len(lines[i].strip()) >= 5]
        tasks = self.group_indices_into_tasks(pending_indices, task_size)
        print(f"初始任务数: {len(tasks)}（过滤后长行: {len(pending_indices)}）")

        self.progress["maximum"] = total
        self.progress["value"] = processed_count

        completed = processed_count
        # 先处理所有初始任务
        all_failed = []
        with ThreadPoolExecutor(max_workers=5) as ex:
            futures = [ex.submit(self.process_task, task, lines) for task in tasks]
            for fut in as_completed(futures):
                if self._stop_flag.is_set():
                    break
                try:
                    success_indices, failed_indices = fut.result()
                    completed += len(success_indices)
                    all_failed.extend(failed_indices)
                    self._flush_buffer_to_file(output_path, lines)
                    self.progress["value"] = completed+len(short_lines)
                    self.status_label.config(text=f"已完成 {completed+len(short_lines)}/{total}")
                    print(f"任务完成: 成功 {len(success_indices)} 行，失败 {len(failed_indices)} 行")
                except Exception as e:
                    print("任务异常:", e)
        # 重试失败的行
        retry_count = 0
        while all_failed and not self._stop_flag.is_set() and retry_count < self.max_retries:
            retry_count += 1
            print(f"第 {retry_count} 轮重试，开始处理 {len(all_failed)} 个失败行")
            retry_tasks = self.group_indices_into_tasks(all_failed, task_size)
            all_failed = []
            with ThreadPoolExecutor(max_workers=5) as ex:
                futures = [ex.submit(self.process_task, task, lines) for task in retry_tasks]
                for fut in as_completed(futures):
                    if self._stop_flag.is_set():
                        break
                    try:
                        success_indices, failed_indices = fut.result()
                        completed += len(success_indices)
                        all_failed.extend(failed_indices)
                        self._flush_buffer_to_file(output_path, lines)
                        self.progress["value"] = completed
                        self.status_label.config(text=f"已完成 {completed}/{total}")
                        print(f"重试任务完成: 成功 {len(success_indices)} 行，失败 {len(failed_indices)} 行")
                    except Exception as e:
                        print("重试任务异常:", e)
            if all_failed:
                print(f"重试后仍有 {len(all_failed)} 行失败")
        if self._stop_flag.is_set():
            print("处理被停止")
            messagebox.showinfo("已停止", f"已停止，当前已完成 {completed}/{len(pending_indices)} 行。")
        elif completed >= len(pending_indices):
            elapsed_time = time.time() - self.start_time
            print(f"处理完成，用时 {elapsed_time:.2f} 秒，总共消耗 {self.total_tokens} tokens")
            messagebox.showinfo("完成", "断句已完成")
        else:
            elapsed_time = time.time() - self.start_time
            print(f"处理未完成，用时 {elapsed_time:.2f} 秒，总共消耗 {self.total_tokens} tokens，剩余 {len(pending_indices) - completed} 行")
            messagebox.showwarning("未完成", f"处理结束，但仍有 {len(pending_indices) - completed} 行未成功。")
        self.start_btn.config(text="开始处理")

    def process_task(self, indices, lines):
        print(f"处理任务: 行 {indices[0]} 到 {indices[-1]} (共 {len(indices)} 行)")
        
        # 所有indices都是长行，直接处理
        prompt = self.build_prompt(indices, lines)
        print("发送的提示内容:")
        print(prompt)
        print(f"发送API请求，提示长度: {len(prompt)} 字符")
        response, tokens = self.call_api(prompt)
        self.total_tokens += tokens
        if not response:
            print("API调用失败，所有行失败")
            return [], indices  # 所有失败

        print("AI返回的结果:")
        print(response)
        parts = self.split_response(response, len(indices))
        print(f"收到响应，分割成 {len(parts)} 部分")
        success = []
        failed = []
        for idx, part in zip(indices, parts):
            is_valid, reason = self.validate_line(lines[idx], part)
            if is_valid:
                self._set_buffer_line(idx, part)
                success.append(idx)
                print(f"\033[92m行 {idx} 验证成功: {part}\033[0m")
            else:
                self.failure_counts[idx] = self.failure_counts.get(idx, 0) + 1
                if self.failure_counts[idx] > 3:
                    self._set_buffer_line(idx, lines[idx])  # 使用原文
                    success.append(idx)
                    print(f"行 {idx} 失败次数超过3，使用原文")
                else:
                    failed.append(idx)
                    print(f"\033[91m行 {idx} 验证失败: {part}||原因: {reason}\033[0m")
        print(f"任务结果: 成功 {len(success)} 行，失败 {len(failed)} 行")
        return success, failed

    def build_prompt(self, indices, lines):
        prompt = self.prompt_var.get() + "\n\n"
        for idx in indices:
            prompt += f"Line {idx}: {lines[idx]}\n"
        prompt += "\nOutput the split text for each line, separated by '----'."
        return prompt

    def call_api(self, prompt):
        api_type = self.api_type_var.get()
        api_url = self.api_var.get()
        api_key = self.api_key_var.get()
        model = self.model_var.get()

        print(f"调用API: {api_type}, 模型: {model}, URL: {api_url}")
        try:
            if api_type == "Ollama":
                payload = {
                    "model": model,
                    "prompt": prompt,
                    "stream": False,
                    "keep_alive": 60
                }
                resp = requests.post(api_url, json=payload, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                response = data.get("response", "")
                tokens = data.get("eval_count", 0)  # Ollama返回eval_count作为tokens
                return response, tokens
            elif api_type in ["OpenAI", "Doubao", "DeepSeek", "Qianwen", "Zhipu"]:
                headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                payload = {
                    "model": model,
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.0
                }
                resp = requests.post(api_url, json=payload, headers=headers, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                response = data["choices"][0]["message"]["content"]
                tokens = data.get("usage", {}).get("total_tokens", 0)
                return response, tokens
            elif api_type == "Baidu":
                headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                payload = {
                    "model": model,
                    "messages": [{"role": "user", "content": prompt}],
                }
                resp = requests.post(api_url, json=payload, headers=headers, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                response = data["choices"][0]["message"]["content"]
                tokens = data.get("usage", {}).get("total_tokens", 0)
                return response, tokens
            elif api_type == "Tencent":
                headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
                payload = {
                    "model": model,
                    "messages": [{"role": "user", "content": prompt}],
                    "enable_enhancement": True
                }
                resp = requests.post(api_url, json=payload, headers=headers, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                response = data["choices"][0]["message"]["content"]
                tokens = data.get("usage", {}).get("total_tokens", 0)
                return response, tokens
            elif api_type == "Gemini":
                headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
                payload = {
                    "contents": [
                        {
                            "parts": [
                                {
                                    "text": prompt
                                }
                            ]
                        }
                    ]
                }
                resp = requests.post(api_url, json=payload, headers=headers, timeout=120)
                resp.raise_for_status()
                data = resp.json()
                response = data.get("contents", [{}])[0].get("parts", [{}])[0].get("text", "")
                tokens = data.get("usageMetadata", {}).get("totalTokenCount", 0)
                return response, tokens
        except Exception as e:
            print(f"API调用失败: {e}")
            return "", 0

    def split_response(self, response, num_parts):
        if "----" in response:
            parts = [p.strip() for p in response.split("----") if p.strip()]
        else:
            parts = [p.strip() for p in response.split("\n\n") if p.strip()]
        if len(parts) < num_parts:
            parts += [""] * (num_parts - len(parts))
        return parts[:num_parts]

    def validate_line(self, original, processed):
        original = original.strip()
        processed_clean = processed.replace('\\r', '').strip()
        if len(original) > 13 and original == processed: 
            return False, "断句后内容与原文相同，原文："+original
        if original != processed_clean: 
            return False, "断句后内容与原文不符，原文："+original
        return True, ""

    def group_indices_into_tasks(self, indices, task_size):
        if not indices:
            return []
        indices = sorted(set(indices))
        tasks = []
        i = 0
        while i < len(indices):
            tasks.append(indices[i:i+task_size])
            i += task_size
        return tasks

    def _set_buffer_line(self, index, text):
        cleaned = text.strip().replace("\r\n", "\n").replace("\r", "")
        with self.output_lock:
            if self.output_buffer and 0 <= index < len(self.output_buffer):
                self.output_buffer[index] = cleaned

    def _flush_buffer_to_file(self, output_path, lines=None):
        with self.output_lock:
            if self.output_buffer is None:
                return
            with open(output_path, "w", encoding="utf-8") as wf:
                for i in range(len(self.output_buffer)):
                    text = self.output_buffer[i]
                    if text is None:
                        if lines and i < len(lines):
                            text = lines[i]
                        else:
                            text = ""
                    wf.write(text + "\n")

    def get_config_path(self):
        if platform.system() == 'Windows':
            appdata = os.environ.get('APPDATA')
            config_dir = os.path.join(appdata, 'Mangahanhua')
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, 'config.json')
        elif platform.system() == 'Darwin':  # Mac
            app_support = os.path.expanduser('~/Library/Application Support')
            config_dir = os.path.join(app_support, 'Mangahanhua')
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, 'config.json')
        else:  # Linux or others
            config_dir = os.path.expanduser('~/.config/Mangahanhua')
            os.makedirs(config_dir, exist_ok=True)
            return os.path.join(config_dir, 'config.json')

    def load_config(self):
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    encrypted = f.read()
                decrypted = base64.b64decode(encrypted).decode('utf-8')
                data = json.loads(decrypted)
                self.api_type_var.set(data.get('api_type', 'Ollama'))
                self.api_var.set(data.get('api_url', DEFAULT_APIS['Ollama']))
                self.api_keys = data.get('api_keys', {})
                self.model_var.set(data.get('model', 'deepseek-v3.2:cloud'))
                self.task_size_var.set(data.get('task_size', 20))
                self.max_chars_var.set(data.get('max_chars', 4000))
                self.think_mode_var.set(data.get('think_mode', False))
                self.prompt_var.set(data.get('prompt', DEFAULT_SYSTEM_PROMPT))
                self.on_api_type_changed()  # 更新模型列表和API Key
                print("配置加载成功")
            except Exception as e:
                print(f"加载配置失败: {e}")

    def save_config(self):
        # 更新当前平台的API Key
        self.api_keys[self.api_type_var.get()] = self.api_key_var.get()
        data = {
            'api_type': self.api_type_var.get(),
            'api_url': self.api_var.get(),
            'api_keys': self.api_keys,
            'model': self.model_var.get(),
            'task_size': self.task_size_var.get(),
            'max_chars': self.max_chars_var.get(),
            'think_mode': self.think_mode_var.get(),
            'prompt': self.prompt_var.get()
        }
        try:
            json_str = json.dumps(data, ensure_ascii=False)
            encrypted = base64.b64encode(json_str.encode('utf-8')).decode('utf-8')
            with open(self.config_file, 'w', encoding='utf-8') as f:
                f.write(encrypted)
            print("配置保存成功")
        except Exception as e:
            print(f"保存配置失败: {e}")

    def on_closing(self):
        self.save_config()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()