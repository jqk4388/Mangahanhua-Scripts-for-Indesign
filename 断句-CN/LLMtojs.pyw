import os
import json
import threading
import tempfile
import time
import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import re

# 新增：跨平台临时文件路径和默认文件名
TMPDIR = tempfile.gettempdir()
DEFAULT_INPUT = os.path.join(TMPDIR, "LLM_input.txt")
DEFAULT_OUTPUT = os.path.join(TMPDIR, "LLM_output.txt")
STATE_FILE = os.path.join(TMPDIR, "mangahanhua_state.json")

# 调试用
# DEFAULT_INPUT = "M:\\汉化\\PS_PNG\\断句测试.txt"
# DEFAULT_OUTPUT = "M:\\汉化\\PS_PNG\\断句测试_LLM_output.txt"

# 默认 Ollama API 与系统提示词（英文）
DEFAULT_OLLAMA_API = "http://localhost:11434/api/generate"
DEFAULT_OPENAI_API = "https://api.openai.com/v1/chat/completions"
DEFAULT_SYSTEM_PROMPT = (
    "You are a sentence-splitting assistant. For the given line of raw Chinese text, "
    "split it into natural sentences with \\r. Consider sentence-final punctuation, morphemes, "
    "and avoid breaking words or meaningful units. Place commas appropriately at the "
    "end of clauses when natural. Keep sentences balanced in length where possible. "
    "Short fragments should be kept as-is (do not force splits). If the original word"
    " order is scrambled, correct it to be natural and fluent. Output only the finalized"
    " split text for each requested line, preserving original order. Do not add extra"
    " commentary. Return plain text." 
    "Example Input:"
    "听说我家附近的寺庙有驱鬼的业务，于是我特地前来拜访。"
    "Example Output:"
    "听说\\r我家附近的寺庙\\r有驱鬼的业务，\\r于是我特地\\r前来拜访。"
)

# UI and control state
class App:
    def __init__(self, root):
        self.root = root
        root.title("大语言模型断句器")
        # ...existing code...
        # top frame for file paths
        frame_files = ttk.Frame(root)
        frame_files.pack(fill="x", padx=8, pady=6)

        ttk.Label(frame_files, text="输入文件:").grid(row=0, column=0, sticky="w")
        self.input_var = tk.StringVar(value=DEFAULT_INPUT)
        self.input_entry = ttk.Entry(frame_files, textvariable=self.input_var, width=60)
        self.input_entry.grid(row=0, column=1, sticky="we")
        ttk.Button(frame_files, text="浏览", command=self.browse_input).grid(row=0, column=2, padx=4)

        ttk.Label(frame_files, text="输出文件:").grid(row=1, column=0, sticky="w")
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        self.output_entry = ttk.Entry(frame_files, textvariable=self.output_var, width=60)
        self.output_entry.grid(row=1, column=1, sticky="we")
        ttk.Button(frame_files, text="浏览", command=self.browse_output).grid(row=1, column=2, padx=4)

        # model and send chars
        frame_params = ttk.Frame(root)
        frame_params.pack(fill="x", padx=8, pady=6)
        ttk.Label(frame_params, text="模型名:").grid(row=0, column=0, sticky="w")
        self.model_var = tk.StringVar(value="qwen3")
        ttk.Entry(frame_params, textvariable=self.model_var, width=20).grid(row=0, column=1, sticky="w", padx=4)

        ttk.Label(frame_params, text="发送字符阈值:").grid(row=0, column=2, sticky="w")
        self.max_chars_var = tk.IntVar(value=4000)
        ttk.Spinbox(frame_params, from_=500, to=20000, increment=500, textvariable=self.max_chars_var, width=10).grid(row=0, column=3, sticky="w", padx=4)

        # 替换为 think_mode_var（用于 get_model_response 中判断是否添加 /no_think）
        self.think_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame_params, text="开启思考（慢）", variable=self.think_mode_var).grid(row=0, column=4, padx=8)

        # advanced hidden settings
        self.advanced_shown = False
        ac_frame = ttk.Frame(root)
        ac_frame.pack(fill="x", padx=8, pady=2)
        self.adv_button = ttk.Button(ac_frame, text="显示高级设置", command=self.toggle_advanced)
        self.adv_button.pack(side="left")
        self.adv_panel = ttk.Frame(root)
        # advanced widgets (initially hidden)
        ttk.Label(self.adv_panel, text="系统提示词:").grid(row=0, column=0, sticky="nw")
        # 使用 StringVar 以便 get_model_response 访问 .get()
        self.prompt_var = tk.StringVar(value=DEFAULT_SYSTEM_PROMPT)
        ttk.Entry(self.adv_panel, textvariable=self.prompt_var, width=80).grid(row=0, column=1, sticky="we", padx=4)

        ttk.Label(self.adv_panel, text="API 类型:").grid(row=1, column=0, sticky="w")
        self.api_type_var = tk.StringVar(value="local")  # local 或 openai
        ttk.Combobox(self.adv_panel, textvariable=self.api_type_var, values=["local", "openai"], width=10, state="readonly").grid(row=1, column=1, sticky="w", padx=4)

        ttk.Label(self.adv_panel, text="Ollama / API 地址:").grid(row=2, column=0, sticky="w")
        # api_var 可用于本地或 OpenAI 的自定义地址
        self.api_var = tk.StringVar(value=DEFAULT_OLLAMA_API)
        ttk.Entry(self.adv_panel, textvariable=self.api_var, width=60).grid(row=2, column=1, sticky="w", padx=4)

        ttk.Label(self.adv_panel, text="API Key (OpenAI):").grid(row=3, column=0, sticky="w")
        self.api_key_var = tk.StringVar(value="")
        ttk.Entry(self.adv_panel, textvariable=self.api_key_var, width=60, show="*").grid(row=3, column=1, sticky="w", padx=4)

        # progress and control
        frame_ctrl = ttk.Frame(root)
        frame_ctrl.pack(fill="x", padx=8, pady=6)
        self.progress = ttk.Progressbar(frame_ctrl, length=400, mode="determinate")
        self.progress.grid(row=0, column=0, padx=4, pady=4)
        self.status_label = ttk.Label(frame_ctrl, text="未开始")
        self.status_label.grid(row=0, column=1, padx=8)
        self.start_btn = ttk.Button(frame_ctrl, text="开始处理", command=self.toggle_start)
        self.start_btn.grid(row=0, column=2, padx=4)

        # settings
        self.batch_size = 1  # 默认每次处理 1 行
        self._stop_flag = threading.Event()
        self.worker_thread = None

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

    def toggle_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            # stop
            self._stop_flag.set()
            self.start_btn.config(text="开始处理")
            self.status_label.config(text="停止中...")
        else:
            # start
            self._stop_flag.clear()
            self.start_btn.config(text="停止")
            self.worker_thread = threading.Thread(target=self.run_processing, daemon=True)
            self.worker_thread.start()

    # 修改：统一的 API 调用函数，支持 local (ollama-like) 与 openai 两种模式
    def get_model_response(self, text):
        try:
            # 根据思考模式决定是否添加/no_think
            suffix = "" if self.think_mode_var.get() else "/no_think"
            api_type = self.api_type_var.get() if hasattr(self, "api_type_var") else "local"
            # OpenAI 在线调用
            if api_type == "openai":
                openai_url = self.api_var.get() or DEFAULT_OPENAI_API
                headers = {"Authorization": f"Bearer {self.api_key_var.get()}", "Content-Type": "application/json"}
                payload = {
                    "model": self.model_var.get(),
                    "messages": [
                        {"role": "system", "content": self.prompt_var.get()},
                        {"role": "user", "content": text}
                    ],
                    "temperature": 0.0
                }
                resp = requests.post(openai_url, json=payload, headers=headers, timeout=120)
                resp.raise_for_status()
                result = resp.json()
                # extract text
                try:
                    return result["choices"][0]["message"]["content"]
                except Exception:
                    return str(result)
            else:
                # local / ollama-like
                api_url = self.api_var.get() or DEFAULT_OLLAMA_API
                resp = requests.post(
                    api_url,
                    json={
                        "model": self.model_var.get(),
                        "prompt": self.prompt_var.get() + text + suffix,
                        "stream": False
                    },
                    timeout=120
                )
                resp.raise_for_status()
                result = resp.json()
                # try common shapes: {'response': '...'} or Ollama-like {'choices':[{'message':{'content':{'text':...}}}]}
                if isinstance(result, dict):
                    if "response" in result:
                        return result["response"]
                    if "choices" in result:
                        # reuse existing extractor helper if present
                        try:
                            return self._extract_text_from_ollama_response(result)
                        except Exception:
                            return str(result)
                return str(result)
        except Exception as e:
            messagebox.showerror("错误", f"API调用失败: {str(e)}")
            return text

    def run_processing(self):
        input_path = self.input_var.get() or DEFAULT_INPUT
        output_path = self.output_var.get() or DEFAULT_OUTPUT
        api_url = self.api_var.get() if self.advanced_shown else DEFAULT_OLLAMA_API
        system_prompt = self.prompt_var.get().strip() if self.advanced_shown else DEFAULT_SYSTEM_PROMPT
        model_name = self.model_var.get().strip() or "qwen3"
        max_chars = int(self.max_chars_var.get())

        # read input
        if not os.path.exists(input_path):
            messagebox.showerror("错误", f"输入文件不存在: {input_path}")
            self.start_btn.config(text="开始处理")
            return

        with open(input_path, "r", encoding="utf-8") as f:
            lines = [ln.rstrip("\n") for ln in f]

        total = len(lines)
        # check existing output to decide resume
        processed_count = 0
        if os.path.exists(output_path):
            with open(output_path, "r", encoding="utf-8") as f:
                out_lines = [ln.rstrip("\n") for ln in f]
                processed_count = len(out_lines)
            if processed_count < total:
                if not messagebox.askyesno("续传", f"检测到输出文件已有 {processed_count} 行结果，是否从该行继续处理？"):
                    # overwrite
                    processed_count = 0
                    open(output_path, "w", encoding="utf-8").close()
            else:
                # 如果输出已包含所有行，询问用户是否重新开始（覆盖已有输出并从头处理）
                restart = messagebox.askyesno("完成", "输出文件已完成所有行。是否重新开始并覆盖已有输出？")
                if restart:
                    processed_count = 0
                    # 清空输出文件，准备从头开始
                    open(output_path, "w", encoding="utf-8").close()
                    # 继续执行后续处理（不 return），让后面的逻辑从头开始处理
                else:
                    messagebox.showinfo("完成", "已取消处理。")
                    self.start_btn.config(text="开始处理")
                    return

        self.progress["maximum"] = total
        self.progress["value"] = processed_count

        # processing loop by batch
        i = processed_count
        while i < total and not self._stop_flag.is_set():
            # prepare batch indices
            batch_end = min(i + self.batch_size, total)
            batch_indices = list(range(i, batch_end))
            # prepare combined user message: for each target line include its context (3 before/after)
            sub_prompts = []
            for idx in batch_indices:
                context_start = max(0, idx - 3)
                context_end = min(total, idx + 4)  # slice end exclusive
                context_lines = lines[context_start:context_end]
                # mark which is the target line inside context
                relative_idx = idx - context_start
                prompt = f"Context (lines {context_start}-{context_end-1}):\n" \
                         + "\n".join(f"{j + context_start}: {ln}" for j, ln in enumerate(context_lines)) \
                         + f"\n\nOnly process line index {idx} (relative {relative_idx}). Input line:\n{lines[idx]}\n\n" \
                         + "Output the final split sentence(s) for that single line (do not output context or commentary)."
                sub_prompts.append(prompt)

            combined_user_message = "\n\n----\n\n".join(sub_prompts)
            # if combined too long, fall back to per-line requests
            if len(combined_user_message) > max_chars:
                for sp_idx, idx in enumerate(batch_indices):
                    if self._stop_flag.is_set():
                        break
                    single_prompt = sub_prompts[sp_idx]
                    try:
                        text = self.get_model_response(single_prompt)
                        text = re.sub(r'<think>\n\n</think>\n\n', '', text)
                        # 如果用户开启思考（think），删除 <think>...</think> 块
                        if self.think_mode_var.get():
                            text = re.sub(r'<think>[\s\S]*?</think>\s*\n*', '', text)
                    except Exception as e:
                        messagebox.showerror("请求失败", f"请求失败: {e}")
                        self._stop_flag.set()
                        break
                    # write single line result
                    self._append_output_line(output_path, text)
                    i += 1
                    self.progress["value"] = i
                    self.status_label.config(text=f"已处理 {i}/{total}")
            else:
                try:
                    text = self.get_model_response(combined_user_message)
                    text = re.sub(r'<think>\n\n</think>\n\n', '', text)
                    if self.think_mode_var.get():
                        text = re.sub(r'<think>[\s\S]*?</think>\s*\n*', '', text)
                except Exception as e:
                    messagebox.showerror("请求失败", f"请求失败: {e}")
                    self._stop_flag.set()
                    break
                # The response contains concatenated outputs for each sub_prompt. We split by the separator we used.
                splitted = self._split_response_into_parts(text, len(batch_indices))
                for part in splitted:
                    self._append_output_line(output_path, part)
                    i += 1
                    self.progress["value"] = i
                    self.status_label.config(text=f"已处理 {i}/{total}")
            # small delay to avoid overwhelming
            time.sleep(0.2)

        # ended loop
        if self._stop_flag.is_set():
            messagebox.showinfo("已停止", f"已停止，当前已处理 {i}/{total} 行。下次运行会询问是否续传。")
            self.start_btn.config(text="开始处理")
            return
        # finished
        self.progress["value"] = total
        self.status_label.config(text=f"已完成 {total}/{total}")
        messagebox.showinfo("完成", "断句已完成")
        self.start_btn.config(text="开始处理")

    def _append_output_line(self, output_path, text):
        # normalize text: ensure single-line per input line
        cleaned = text.strip().replace("\r\n", "\n")
        # append as single line (if multiple sentences include appropriate separators)
        with open(output_path, "a", encoding="utf-8") as wf:
            wf.write(cleaned + "\n")

    def _extract_text_from_ollama_response(self, data):
        # best-effort extraction for typical Ollama chat response structure
        try:
            choices = data.get("choices")
            if choices and len(choices) > 0:
                msg = choices[0].get("message", {})
                content = msg.get("content")
                if isinstance(content, dict):
                    # content may be {"type":"text","text":"..."}
                    return content.get("text", "") or content.get("text", "")
                elif isinstance(content, str):
                    return content
            # fallback to raw text
            return str(data)
        except Exception:
            return str(data)

    def _split_response_into_parts(self, text, expect_parts):
        # try to split the combined response into expect_parts pieces using the separator or heuristics
        # If user used our "----" separator, model may preserve it. Try split by that first.
        if "----" in text:
            parts = [p.strip() for p in text.split("----") if p.strip()]
        else:
            # fallback: split by double newlines into expect_parts (best-effort)
            parts = [p.strip() for p in text.split("\n\n") if p.strip()]
        # If we have fewer parts than expected, pad by repeating last; if more, truncate
        if len(parts) < expect_parts:
            parts += [""] * (expect_parts - len(parts))
        if len(parts) > expect_parts:
            parts = parts[:expect_parts]
        return parts

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()