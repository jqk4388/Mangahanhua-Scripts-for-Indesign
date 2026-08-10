# -*- coding: utf-8 -*-
"""LLMtojs.pyw - 漫画台词智能断句工具（重构版）

功能：
  - 调用云端大模型或本地小模型，对中文漫画台词进行智能断句
  - 导入/导出文件格式保持不变（每行一条台词，<BR> 为段内分隔符）
  - 结构化 JSON 输出，减少 token 消耗，提高一致性
  - 全/半角标点自动校正，避免无谓失败
  - 与人工断句示例自动比对，量化评估断句质量

被 大语言模型断句.jsx 通过以下方式调用：
  pythonw LLMtojs.pyw <input.txt> <output.txt>
也可独立运行弹出 GUI。
"""

import os
import sys
import json
import re
import time
import base64
import platform
import tempfile
import threading
import difflib
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


# ============================================================================
# 常量与默认配置
# ============================================================================

TMPDIR = tempfile.gettempdir()
DEFAULT_INPUT = os.path.join(TMPDIR, "LLM_input.txt")
DEFAULT_OUTPUT = os.path.join(TMPDIR, "LLM_output.txt")

# 默认 API 地址
DEFAULT_APIS = {
    "LM Studio": "http://localhost:1234/v1/chat/completions",
    "Ollama": "http://localhost:11434/api/generate",
    "OpenAI": "https://api.openai.com/v1/chat/completions",
    "Doubao": "https://ark.cn-beijing.volces.com/api/v3/chat/completions",
    "DeepSeek": "https://api.deepseek.com/chat/completions",
    "Qwen": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
    "Baidu": "https://qianfan.baidubce.com/v2/chat/completions",
    "Tencent": "https://api.hunyuan.cloud.tencent.com/v1/chat/completions",
    "Zhipu": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
    "Gemini": "https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent",
    "MiniMax": "https://api.minimaxi.com/v1/chat/completions",
    "OpenRouter": "https://openrouter.ai/api/v1/chat/completions",
    "Anthropic": "https://api.anthropic.com/v1",
    "天翼云": "https://wishub-x6.ctyun.cn/v1",
    "免费模型": "https://api.kilo.ai/api/gateway/chat/completions",
    "MiMo": "https://api.xiaomimimo.com/v1/chat/completions",
    "移动云": "https://zhenze-huhehaote.cmecloud.cn/v1/chat/completions",
    "移动云(Coding)": "https://zhenze-huhehaote.cmecloud.cn/api/coding/v1/chat/completions",
    "联通云": "https://aigw-gzgy2.cucloud.cn:8443/v1/chat/completions",
}

# 默认模型列表
DEFAULT_MODELS = {
    "Ollama": ["deepseek-v4-flash", "deepseek-v4-pro"],
    "OpenAI": ["gpt-4.1-2025-04-14", "gpt-5.4-nano-2026-03-17", "gpt-5.5"],
    "Doubao": ["doubao-seed-2-0-mini-260428", "doubao-seed-evolving",
               "doubao-seed-2-1-turbo-260628", "deepseek-v4-flash-ga-260731"],
    "DeepSeek": ["deepseek-v4-flash", "deepseek-v4-pro"],
    "Qwen": ["qwen3.6-flash", "deepseek-v4-flash-260731", "MiniMax/MiniMax-M2.7",
             "glm-5.1", "qwen3.7-max"],
    "Baidu": ["ernie-5.0", "deepseek-v4-flash", "ernie-5.1"],
    "Tencent": ["hy3-preview", "deepseek-v4-flash", "glm-5.1", "glm-5-turbo"],
    "Zhipu": ["glm-5-turbo", "glm-4.7-flash", "glm-4.6v-flash"],
    "Gemini": ["gemini-3.1-flash-lite", "gemini-3.5-flash"],
    "LM Studio": ["local-model"],
    "MiniMax": ["MiniMax-M2.7", "MiniMax-M3"],
    "OpenRouter": ["openrouter/free"],
    "Anthropic": ["claude-haiku-4-5", "claude-sonnet-4-6", "claude-opus-4-7"],
    "天翼云": ["DeepSeek-V4-Flash"],
    "免费模型": [
        "kilo-auto/free", "mimo-v2.5-free", "north-mini-code-free",
        "nemotron-3-ultra-free", "deepseek-v4-flash-free", "openrouter/free",
    ],
    "MiMo": ["mimo-v2.5-pro", "mimo-v2-flash"],
    "移动云": ["deepseek-v4-flash"],
    "移动云(Coding)": ["deepseek-v4-flash"],
    "联通云": ["DeepSeek-V4-Flash", "MiniMax-M2.5"],
}

# 免费模型对应的 API 地址
FREE_MODEL_API_URLS = {
    "kilo-auto/free": "https://api.kilo.ai/api/gateway/chat/completions",
    "mimo-v2.5-free": "https://opencode.ai/zen/v1/chat/completions",
    "north-mini-code-free": "https://opencode.ai/zen/v1/chat/completions",
    "nemotron-3-ultra-free": "https://opencode.ai/zen/v1/chat/completions",
    "deepseek-v4-flash-free": "https://opencode.ai/zen/v1/chat/completions",
    "openrouter/free": "https://openrouter.ai/api/v1/chat/completions",
}
DEFAULT_OPENROUTER_FREE_KEY = ""

# 思考模式 extra body 示例
DEFAULT_EXTRA_BODIES = [
    '{"think": false}',
    '{"think": true}',
    '{"enable_thinking": false}',
    '{"enable_thinking": true}',
    '{"thinking": {"type": "disabled"}}',
    '{"thinking": {"type": "enabled"}}',
    '{"reasoning_effort": "low"}',
    '{"output_config": {"effort": "low"}}',
]

# 短行阈值：字符数低于此值的行直接保留原文，不送入模型
SHORT_LINE_THRESHOLD = 3
# 原文与结果完全相同时判为未断句的最小原文长度
MIN_SPLITTABLE_LEN = 8
# 单行最大重试次数（超过后回退原文）
MAX_LINE_FAILURES = 5

# 脚本版本号：每次更新提示词或配置字段时递增
SCRIPT_VERSION = "2.1.0"


# ============================================================================
# 提示词模板
# ============================================================================

# 系统提示词（纯规则，不含示例）：
SYSTEM_PROMPT = """你是中文漫画台词断句助手。把每行台词断成几个自然片段，片段之间用 <BR> 分隔。

规则：
1. 在逗号后、分句处、语气转换处断句；每段约 4-10 字，长短均衡。
2. 绝对不改原文标点（? ! — 「」 …… 等原样保留），只插入 <BR>。
3. 不拆散人名、地名、成语等固定词。
4. 短句（5 字以下）无需断句，原样返回。
5. 必须断句：超过 10 字且含逗号/分句的行，至少插入一个 <BR>。
6. 一般在"的"字后断句，除非整句太短。"""

# 预设示例对话（多轮 few-shot，用作系统提示词后的上下文）
# 每条为 (用户原文列表, 助手断句结果列表)
FEWSHOT_EXAMPLES = [
    (
        [
            "这就是传说中的石像鬼吧，我还是头一回见。",
            "直子!多半是那家人自己收起来了吧——",
            "……自从那东西来到家里之后，不好的事情就一直没断过。",
            "古董商带过来的藏品当中正好有这个，我一眼看中了它，就买下来了。",
            "嗯?",
        ],
        [
            "这就是<BR>传说中的<BR>石像鬼吧，<BR>我还是头一回见。",
            "直子!<BR>多半是那家人<BR>自己收起来了吧——",
            "……自从那东西<BR>来到家里之后，<BR>不好的事情<BR>就一直没断过。",
            "古董商带过来的<BR>藏品当中正好有这个，<BR>我一眼看中了它，<BR>就买下来了。",
            "嗯?",
        ],
    ),
]

# 输出格式说明（追加到系统提示词之后）
OUTPUT_FORMAT_JSON = (
    "\n\n现在对下面给出的每行台词断句，直接输出 JSON，不要输出思考过程、解释或 markdown 标记。\n"
    '格式严格为：{"r": ["第1行结果", "第2行结果", "..."]}\n'
    "要点：r 数组长度必须等于输入行数；每个元素只含原文与 <BR>，不得增删任何字符。"
)

OUTPUT_FORMAT_TEXT = (
    "\n\n现在对下面给出的每行台词断句，每行输出一个结果，结果之间用 ---- 分隔。\n"
    "不要输出行号、思考过程或任何解释文字。"
)


# ============================================================================
# 工具函数
# ============================================================================

# 全角→半角 标点映射（用于规范化比较）
_PUNCT_FULL_TO_HALF = {
    "？": "?", "！": "!", "。": ".", "，": ",", "、": ",",
    "：": ":", "；": ";", "（": "(", "）": ")",
    "「": '"', "」": '"', "『": '"', "』": '"',
    "【": "[", "】": "]", "《": "<", "》": ">",
    "～": "~", "“": '"', "”": '"', "‘": "'", "’": "'",
}


def normalize_punct(text):
    """全角标点归一化为半角，用于比较（不修改原文）。"""
    for fw, hw in _PUNCT_FULL_TO_HALF.items():
        text = text.replace(fw, hw)
    # 破折号 / 连字符统一
    text = text.replace("—", "-").replace("–", "-")
    # 省略号统一为半角点
    text = text.replace("…", ".")
    return text


def normalize_for_compare(text):
    """去除 <BR>、空白、全/半角差异后的归一化文本，仅用于比较。"""
    text = (text or "").replace("\r", "").replace("\\r", "")
    text = re.sub(r"<BR>", "", text, flags=re.IGNORECASE)
    text = normalize_punct(text)
    text = re.sub(r"[\s\u3000]+", "", text)
    # 连续点（……/.../…）归一为单点，避免省略号长度差异影响比较
    text = re.sub(r"\.{2,}", ".", text)
    return text


def extract_br_positions(text):
    """返回 <BR> 在去除标记后的文本中出现的位置（字符索引，从 1 开始计数表示"在第 N 个字符之后"）。

    例如 "ab<BR>cd" -> [2]  表示在第 2 个字符之后插入 <BR>。
    """
    positions = []
    text_pos = 0
    idx = 0
    while idx < len(text):
        if text[idx:idx + 4].upper() == "<BR>":
            positions.append(text_pos)
            idx += 4
        else:
            text_pos += 1
            idx += 1
    return positions


def apply_br_positions(original, positions):
    """在 original 的指定位置插入 <BR>。positions 为"在第 N 个字符之后"的列表（0 表示开头）。"""
    pos_set = set(positions)
    result = []
    if 0 in pos_set:
        result.append("<BR>")
    for i, char in enumerate(original):
        result.append(char)
        if (i + 1) in pos_set:
            result.append("<BR>")
    return "".join(result)


def rebuild_with_original_punct(original, processed):
    """用 processed 中的 <BR> 位置，结合 original 的原始标点重建结果。

    若 processed 去除 <BR> 后与 original 在归一化后一致，
    则用 original 的标点重建，自动修正模型可能的全/半角、省略号、破折号等替换。
    使用序列对齐处理标点长度差异（如 …… 2 字 → ... 3 字）。
    返回 (重建结果, 是否成功)。
    """
    original = original or ""
    processed = processed or ""
    processed_clean = processed.replace("<BR>", "")
    if normalize_for_compare(original) != normalize_for_compare(processed_clean):
        return processed, False

    # 用序列对齐把 processed_clean 的字符位置映射回 original 位置
    sm = difflib.SequenceMatcher(None, original, processed_clean, autojunk=False)
    proc_to_orig = {}  # processed_clean 索引 -> original 索引
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(j2 - j1):
                proc_to_orig[j1 + k] = i1 + k
        elif tag == "replace":
            n = min(i2 - i1, j2 - j1)
            for k in range(n):
                proc_to_orig[j1 + k] = i1 + k
            # processed 多出的字符映射到 original 区间末尾
            for k in range(n, j2 - j1):
                proc_to_orig[j1 + k] = i2 - 1 if i2 > i1 else max(i1 - 1, 0)

    br_boundaries = extract_br_positions(processed)  # "在第 N 个真实字符之后"
    orig_boundaries = set()
    for b in br_boundaries:
        if b == 0:
            orig_boundaries.add(0)
        elif (b - 1) in proc_to_orig:
            orig_boundaries.add(proc_to_orig[b - 1] + 1)
    return apply_br_positions(original, orig_boundaries), True


# ============================================================================
# 提示词构建与响应解析
# ============================================================================

class PromptBuilder:
    """构建发送给模型的提示词（多轮消息格式）。"""

    def __init__(self, system_prompt, output_format="json"):
        self.system_prompt = system_prompt
        self.output_format = output_format

    # ---------- 内部工具 ----------
    @staticmethod
    def _format_fewshot_user(example_inputs, output_format):
        """构造 few-shot 的用户消息：输出格式说明 + 示例输入行。"""
        lines = []
        lines.append(OUTPUT_FORMAT_JSON if output_format == "json" else OUTPUT_FORMAT_TEXT)
        lines.append("\n\n待断句的台词：")
        lines.extend(example_inputs)
        return "\n".join(lines).strip()

    @staticmethod
    def _format_fewshot_assistant(example_outputs, output_format):
        """构造 few-shot 的助手消息：与用户消息对应的格式化回答。"""
        if output_format == "json":
            safe = [json.dumps(s, ensure_ascii=False) for s in example_outputs]
            return '{"r": [' + ", ".join(safe) + "]}"
        else:
            return "\n----\n".join(example_outputs)

    @staticmethod
    def messages_to_prompt(messages):
        """将多轮 messages 折叠成单字符串 prompt（兼容不支持 messages 的接口）。"""
        if not messages:
            return ""
        buf = []
        for m in messages:
            role = m.get("role", "user")
            content = m.get("content", "") or ""
            if role == "system":
                buf.append("【系统】\n" + content)
            elif role == "user":
                buf.append("【用户】\n" + content)
            elif role == "assistant":
                buf.append("【助手】\n" + content)
            else:
                buf.append(f"【{role}】\n" + content)
        return "\n\n".join(buf)

    # ---------- 主接口 ----------
    def build(self, indices, lines, failure_info=None):
        """构建多轮消息列表 messages。

        返回: list[dict]，每项形如 {"role": "system|user|assistant", "content": "..."}

        failure_info: dict，索引 idx -> {"reason": 失败原因, "ai_output": AI上次返回}
                      仅重试时传入，用于在提示词中附加上次失败上下文。
        """
        messages = []
        # 1. 系统提示词：纯规则
        messages.append({"role": "system", "content": self.system_prompt})

        # 2. 注入预设 few-shot 对话（用户提问 → 助手正确回答）
        fmt = self.output_format
        for example_inputs, example_outputs in FEWSHOT_EXAMPLES:
            user_msg = self._format_fewshot_user(example_inputs, fmt)
            assistant_msg = self._format_fewshot_assistant(example_outputs, fmt)
            messages.append({"role": "user", "content": user_msg})
            messages.append({"role": "assistant", "content": assistant_msg})

        # 3. 构造当前真实用户消息
        user_parts = []
        user_parts.append(OUTPUT_FORMAT_JSON if fmt == "json" else OUTPUT_FORMAT_TEXT)

        # 重试模式：附加上轮失败上下文
        has_retry = failure_info and any(idx in failure_info for idx in indices)
        if has_retry:
            user_parts.append("\n\n【重要】以下台词之前断句失败，请针对失败原因修正结果：")
            for i, idx in enumerate(indices):
                info = failure_info.get(idx) if failure_info else None
                if info:
                    user_parts.append(
                        f"第 {i + 1} 条「{lines[idx]}」"
                        f" - 上次错误原因：{info.get('reason', '未知')}"
                        f"；上次AI返回：「{info.get('ai_output', '')}」"
                    )
            user_parts.append("请重新给出正确断句，避免再次出现同类错误。\n")

        user_parts.append("\n\n待断句的台词：")
        for idx in indices:
            user_parts.append(f"{lines[idx]}")

        messages.append({"role": "user", "content": "\n".join(user_parts)})
        return messages


class ResponseParser:
    """解析模型返回的结构化响应，兼容 JSON / 文本分隔两种格式。"""

    @staticmethod
    def parse(response, expected_count):
        """返回长度恰为 expected_count 的结果列表（不足补空，多余截断）。"""
        if response is None:
            response = ""
        elif not isinstance(response, str):
            response = str(response)
        response = response.strip()
        # 部分模型会把 < > 转义成 &lt; &gt;，先还原
        response = response.replace("&lt;", "<").replace("&gt;", ">")

        parts = ResponseParser._try_json(response, expected_count)
        if parts is None:
            parts = ResponseParser._try_text(response, expected_count)
        # 对齐长度
        if len(parts) < expected_count:
            parts += [""] * (expected_count - len(parts))
        return parts[:expected_count]

    @staticmethod
    def _try_json(response, expected_count):
        """尝试从响应中提取 JSON {"r": [...]} 或纯数组。"""
        # 去除可能的 markdown 代码块标记
        text = response.strip()
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```\s*$", "", text)
        # 定位首个 { 或 [ 到末尾
        start = -1
        for i, ch in enumerate(text):
            if ch in "{[":
                start = i
                break
        if start < 0:
            return None
        candidate = text[start:]
        try:
            data = json.loads(candidate)
        except json.JSONDecodeError:
            # 尝试截取到匹配的闭合括号
            candidate = ResponseParser._extract_balanced(text, start)
            if not candidate:
                # JSON 被截断（未闭合）：用正则提取数组内已完整的字符串元素
                partial = ResponseParser._extract_partial_array(text, start)
                if partial:
                    return partial
                return None
            try:
                data = json.loads(candidate)
            except json.JSONDecodeError:
                partial = ResponseParser._extract_partial_array(text, start)
                if partial:
                    return partial
                return None
        if isinstance(data, dict):
            for key in ("r", "results", "data", "output"):
                if isinstance(data.get(key), list):
                    return [str(x) for x in data[key]]
            # 取第一个列表值
            for v in data.values():
                if isinstance(v, list):
                    return [str(x) for x in v]
        if isinstance(data, list):
            return [str(x) for x in data]
        return None

    @staticmethod
    def _extract_balanced(text, start):
        """从 start 处的 { 或 [ 开始，提取到匹配的闭合括号。"""
        open_ch = text[start]
        close_ch = "}" if open_ch == "{" else "]"
        depth = 0
        in_str = False
        escape = False
        for i in range(start, len(text)):
            ch = text[i]
            if in_str:
                if escape:
                    escape = False
                elif ch == "\\":
                    escape = True
                elif ch == '"':
                    in_str = False
            else:
                if ch == '"':
                    in_str = True
                elif ch == open_ch:
                    depth += 1
                elif ch == close_ch:
                    depth -= 1
                    if depth == 0:
                        return text[start:i + 1]
        return None

    @staticmethod
    def _extract_partial_array(text, start):
        """JSON 被截断时，用正则提取数组内已完整的字符串元素。

        定位 {"r": [ 之后，逐个提取完整闭合的 "..." 字符串（处理转义）。
        """
        seg = text[start:]
        # 找到数组开始：[ （兼容 {"r":[ 与 { "r" : [ 等）
        m = re.search(r'"r"\s*:\s*\[', seg)
        if m:
            arr_start = m.end()
            seg = seg[arr_start:]
        else:
            # 纯数组 [ 开头
            m2 = re.match(r'\s*\[', seg)
            if m2:
                seg = seg[m2.end():]
            else:
                return None
        # 提取完整字符串元素："..." （支持 \" 转义，不含未闭合的末尾残串）
        items = []
        i = 0
        n = len(seg)
        while i < n:
            # 跳过逗号空白
            while i < n and seg[i] in " ,\t\r\n":
                i += 1
            if i >= n or seg[i] != '"':
                break
            # 解析一个完整字符串
            i += 1
            buf = []
            closed = False
            while i < n:
                c = seg[i]
                if c == "\\" and i + 1 < n:
                    buf.append(seg[i:i + 2])
                    i += 2
                    continue
                if c == '"':
                    closed = True
                    i += 1
                    break
                buf.append(c)
                i += 1
            if closed:
                raw = "".join(buf)
                # 反转义常见序列
                raw = raw.replace('\\"', '"').replace("\\\\", "\\")
                items.append(raw)
            else:
                break
        return items if items else None

    @staticmethod
    def _try_text(response, expected_count):
        """文本分隔解析：优先 ---- 分隔，其次双换行，最后单换行。"""
        if "----" in response:
            parts = [p.strip() for p in response.split("----")]
        elif "\n\n" in response:
            parts = [p.strip() for p in response.split("\n\n")]
        else:
            parts = [p.strip() for p in response.split("\n")]
        return [p for p in parts if p]


# ============================================================================
# 验证器
# ============================================================================

class Validator:
    """校验单行断句结果是否合法，并自动修正标点。"""

    @staticmethod
    def validate(original, processed):
        """返回 (修正后结果, 是否通过, 原因)。"""
        original = (original or "").strip()
        processed_clean = (processed or "").replace("\r", "").replace("\\r", "").strip()
        # 去除首尾无效的 <BR>（漫画台词中行首/行尾断句无意义）
        processed_clean = re.sub(r"^(<BR>\s*)+", "", processed_clean, flags=re.IGNORECASE)
        processed_clean = re.sub(r"(\s*<BR>)+$", "", processed_clean, flags=re.IGNORECASE)

        # 1. 尝试用原文标点重建（自动修正全/半角问题）
        rebuilt, ok = rebuild_with_original_punct(original, processed_clean)
        if ok:
            # 重建成功，内容一致。再检查是否真的断句了
            if len(original) > MIN_SPLITTABLE_LEN and "<BR>" not in rebuilt:
                return rebuilt, False, "断句后内容与原文相同（未断句）"
            return rebuilt, True, ""
        return processed_clean, False, "断句后内容与原文不符"


# ============================================================================
# 评估器：与人工断句示例对比
# ============================================================================

class Evaluator:
    """将 AI 断句结果与人工参考断句进行量化对比。"""

    @staticmethod
    def evaluate(ai_lines, ref_lines):
        """返回评估结果字典。

        指标：
          - line_count: (ai 数量, ref 数量, 是否一致)
          - exact_match_rate: 归一化后完全一致的行占比
          - text_match_rate: 仅文本一致（忽略 <BR> 位置）的行占比
          - br_count_f1: <BR> 数量的 F1
          - br_position_f1: <BR> 精确位置的 F1
          - overall: 综合得分（0-1）
          - details: 每行的对比详情
        """
        ai_lines = [l or "" for l in ai_lines]
        ref_lines = [l or "" for l in ref_lines]
        n = max(len(ai_lines), len(ref_lines))

        exact_match = 0
        text_match = 0
        br_tp = br_fp = br_fn = 0
        details = []

        for i in range(n):
            ai = ai_lines[i] if i < len(ai_lines) else ""
            ref = ref_lines[i] if i < len(ref_lines) else ""

            ai_norm = normalize_for_compare(ai)
            ref_norm = normalize_for_compare(ref)
            text_eq = (ai_norm == ref_norm)
            if text_eq:
                text_match += 1

            ai_pos = set(extract_br_positions(ai))
            ref_pos = set(extract_br_positions(ref))
            pos_eq = (ai_pos == ref_pos)
            if text_eq and pos_eq:
                exact_match += 1

            tp = len(ai_pos & ref_pos)
            fp = len(ai_pos - ref_pos)
            fn = len(ref_pos - ai_pos)
            br_tp += tp
            br_fp += fp
            br_fn += fn

            details.append({
                "line": i + 1,
                "ai": ai,
                "ref": ref,
                "text_match": text_eq,
                "br_match": pos_eq,
                "ai_br_count": len(ai_pos),
                "ref_br_count": len(ref_pos),
                "br_correct": tp,
            })

        compare_count = n
        exact_rate = exact_match / compare_count if compare_count else 0
        text_rate = text_match / compare_count if compare_count else 0
        br_count_precision = br_tp / (br_tp + br_fp) if (br_tp + br_fp) else 1.0
        br_count_recall = br_tp / (br_tp + br_fn) if (br_tp + br_fn) else 1.0
        br_count_f1 = (2 * br_count_precision * br_count_recall /
                       (br_count_precision + br_count_recall)
                       if (br_count_precision + br_count_recall) else 0)
        # 位置 F1 与数量 F1 在此实现下等价（tp 基于"位置完全相同"集合）
        br_pos_f1 = br_count_f1

        overall = 0.5 * exact_rate + 0.3 * text_rate + 0.2 * br_pos_f1

        return {
            "line_count": (len(ai_lines), len(ref_lines), len(ai_lines) == len(ref_lines)),
            "exact_match_rate": exact_rate,
            "text_match_rate": text_rate,
            "br_count_f1": br_count_f1,
            "br_position_f1": br_pos_f1,
            "overall": overall,
            "details": details,
        }

    @staticmethod
    def format_report(result):
        """将评估结果格式化为可读文本。"""
        ai_n, ref_n, count_ok = result["line_count"]
        lines = [
            "===== 断句质量评估报告 =====",
            f"行数对比 : AI {ai_n} 行 / 参考 {ref_n} 行 {'（一致）' if count_ok else '（不一致）'}",
            f"精确匹配 : {result['exact_match_rate'] * 100:.1f}%  （文本与 <BR> 位置完全一致）",
            f"文本匹配 : {result['text_match_rate'] * 100:.1f}%  （仅文本一致，忽略 <BR> 位置）",
            f"<BR> 位置 F1 : {result['br_position_f1'] * 100:.1f}%",
            f"综合得分 : {result['overall'] * 100:.1f}%",
            "",
            "===== 差异明细（仅列出有差异的行）=====",
        ]
        diff_lines = [d for d in result["details"] if not (d["text_match"] and d["br_match"])]
        if not diff_lines:
            lines.append("（无差异，所有行完全一致）")
        else:
            for d in diff_lines:
                lines.append(f"[第 {d['line']} 行] AI: {d['ai']}")
                lines.append(f"          参考: {d['ref']}")
                if not d["text_match"]:
                    lines.append("          ⚠ 文本内容不一致")
                elif not d["br_match"]:
                    lines.append(
                        f"          ⚠ <BR> 位置不同 (AI {d['ai_br_count']} 处 / 参考 {d['ref_br_count']} 处, 命中 {d['br_correct']})"
                    )
                lines.append("")
        return "\n".join(lines)


# ============================================================================
# LLM 提供商
# ============================================================================

class LLMProvider:
    """LLM 提供商基类。子类实现 chat / list_models。"""

    def __init__(self, api_url, api_key, model, extra_params, think_mode):
        self.api_url = api_url
        self.api_key = api_key
        self.model = model
        self.extra_params = extra_params if isinstance(extra_params, dict) else {}
        self.think_mode = think_mode

    def chat(self, messages):
        """返回 (response_text, token_count)。

        messages: list[dict]，多轮对话，形如:
            [{"role": "system", "content": "..."},
             {"role": "user", "content": "..."},
             {"role": "assistant", "content": "..."},
             {"role": "user", "content": "..."}]
        """
        raise NotImplementedError

    def list_models(self):
        """返回模型列表。"""
        return []

    @staticmethod
    def extract_content(data):
        """从 OpenAI 兼容返回结构中提取文本内容。"""
        if isinstance(data, str):
            return data
        if not isinstance(data, dict):
            return ""
        if isinstance(data.get("choices"), list) and data["choices"]:
            choice = data["choices"][0]
            if isinstance(choice, dict):
                message = choice.get("message")
                if isinstance(message, dict):
                    content = message.get("content")
                    if isinstance(content, str):
                        return content
                    if isinstance(content, list):
                        parts = []
                        for item in content:
                            if isinstance(item, dict):
                                text = item.get("text") or item.get("content")
                                if isinstance(text, str):
                                    parts.append(text)
                        return "".join(parts)
                text = choice.get("text")
                if isinstance(text, str):
                    return text
        for key in ("response", "text", "completion", "output"):
            value = data.get(key)
            if isinstance(value, str):
                return value
            if isinstance(value, list):
                parts = []
                for item in value:
                    if isinstance(item, dict):
                        text = item.get("text") or item.get("content")
                        if isinstance(text, str):
                            parts.append(text)
                    elif isinstance(item, str):
                        parts.append(item)
                if parts:
                    return "".join(parts)
        return ""

    @staticmethod
    def extract_tokens(data):
        """从返回中提取 token 用量。"""
        if not isinstance(data, dict):
            return 0
        usage = data.get("usage")
        if isinstance(usage, dict):
            return usage.get("total_tokens", 0) or 0
        return data.get("eval_count", 0) or 0

    @staticmethod
    def estimate_chars(messages_or_prompt):
        """估算输入的字符数（messages 列表或单字符串都支持）。"""
        if isinstance(messages_or_prompt, str):
            return len(messages_or_prompt)
        if isinstance(messages_or_prompt, list):
            total = 0
            for m in messages_or_prompt:
                if isinstance(m, dict):
                    c = m.get("content", "")
                    if isinstance(c, str):
                        total += len(c)
                    elif isinstance(c, list):
                        for item in c:
                            if isinstance(item, dict):
                                t = item.get("text") or item.get("content") or ""
                                total += len(t)
            return total
        return len(str(messages_or_prompt or ""))

    @staticmethod
    def get_timeout(prompt_chars, api_type="", think_mode=False):
        # 思考模式下模型先思考再输出，本地小模型 (~100 token/s) 需要更长等待
        if think_mode:
            if api_type == "Doubao":
                return 600 if prompt_chars > 2000 else 360
            return 600 if prompt_chars > 4000 else 360
        if api_type == "Doubao":
            return 300 if prompt_chars > 2000 else 180
        return 300 if prompt_chars > 4000 else 180


class OpenAICompatProvider(LLMProvider):
    """OpenAI 兼容接口（OpenAI / Doubao / DeepSeek / Qwen / Zhipu / Baidu / Tencent /
    LM Studio / MiniMax / OpenRouter / 移动云 / 联通云 等）。"""

    # 需要 thinking 字段的供应商
    THINKING_PROVIDERS = {"Doubao", "DeepSeek"}

    def chat(self, messages):
        headers = {"Authorization": f"Bearer {self.api_key}",
                   "Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": messages,
            "stream": False,
            # 给足输出空间，避免本地小模型默认 max_tokens 过小导致 JSON 被截断
            "max_tokens": 8192,
        }
        if self.api_url_type in self.THINKING_PROVIDERS:
            payload["thinking"] = {"type": "enabled" if self.think_mode else "disabled"}
        elif self.api_url_type in ("OpenAI", "Qwen", "Zhipu", "LM Studio",
                                    "MiniMax", "OpenRouter", "移动云",
                                    "移动云(Coding)", "联通云"):
            payload["enable_thinking"] = self.think_mode
        payload.update(self.extra_params)

        timeout = self.get_timeout(self.estimate_chars(messages), self.api_url_type, self.think_mode)

        # 豆包单独处理超时重试
        if self.api_url_type == "Doubao":
            return self._chat_doubao(payload, headers, timeout)

        endpoint = self._resolve_chat_endpoint()
        try:
            resp = requests.post(endpoint, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            return self.extract_content(data), self.extract_tokens(data)
        except Exception as e:
            return "", 0

    def _chat_doubao(self, payload, headers, timeout):
        last_error = None
        for attempt in range(3):
            try:
                resp = requests.post(self.api_url, json=payload,
                                     headers=headers, timeout=timeout)
                resp.raise_for_status()
                data = resp.json()
                return self.extract_content(data), self.extract_tokens(data)
            except requests.exceptions.Timeout as exc:
                last_error = exc
                if attempt < 2:
                    time.sleep(2 + attempt)
                    continue
                return "", 0
            except Exception:
                return "", 0
        return "", 0

    def _resolve_chat_endpoint(self):
        """将用户填写的地址归一化为 chat/completions 端点。

        兼容三种写法：
          1) 含 /chat/completions：原样使用
          2) 以 /v1 结尾：追加 /chat/completions
          3) 仅根地址（如 http://192.168.1.77:1234）：追加 /v1/chat/completions
        """
        url = self.api_url
        if "/chat/completions" in url:
            return url
        if url.rstrip("/").endswith("/v1"):
            return url.rstrip("/") + "/chat/completions"
        return url.rstrip("/") + "/v1/chat/completions"

    def list_models(self):
        """获取模型列表。兼容完整 chat/completions 地址与仅根地址两种写法。"""
        url = self.api_url
        # 1) 若地址含 /chat/completions，替换为 /models
        if "/chat/completions" in url:
            models_url = url.replace("/chat/completions", "/models")
        # 2) 若地址以 /v1 结尾（如天翼云根地址），直接追加 /models
        elif url.rstrip("/").endswith("/v1"):
            models_url = url.rstrip("/") + "/models"
        # 3) 仅根地址（如 http://192.168.1.77:1234），追加 /v1/models
        else:
            models_url = url.rstrip("/") + "/v1/models"
        headers = {}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
        try:
            resp = requests.get(models_url, headers=headers, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                return [m["id"] for m in data.get("data", []) if isinstance(m, dict) and m.get("id")]
        except Exception:
            pass
        return []

    # 通过工厂注入 api_type
    api_url_type = ""


class BaiduProvider(OpenAICompatProvider):
    def chat(self, messages):
        headers = {"Authorization": f"Bearer {self.api_key}",
                   "Content-Type": "application/json"}
        payload = {"model": self.model, "messages": messages}
        payload.update(self.extra_params)
        endpoint = self._resolve_chat_endpoint()
        try:
            resp = requests.post(endpoint, json=payload, headers=headers, timeout=120)
            resp.raise_for_status()
            data = resp.json()
            return data["choices"][0]["message"]["content"], self.extract_tokens(data)
        except Exception:
            return "", 0


class TencentProvider(OpenAICompatProvider):
    def chat(self, messages):
        headers = {"Authorization": f"Bearer {self.api_key}",
                   "Content-Type": "application/json"}
        payload = {"model": self.model,
                   "messages": messages,
                   "enable_enhancement": True}
        payload.update(self.extra_params)
        endpoint = self._resolve_chat_endpoint()
        try:
            resp = requests.post(endpoint, json=payload, headers=headers, timeout=120)
            resp.raise_for_status()
            data = resp.json()
            return data["choices"][0]["message"]["content"], self.extract_tokens(data)
        except Exception:
            return "", 0


class OllamaProvider(LLMProvider):
    def chat(self, messages):
        # Ollama /api/generate 接口只接收单字符串 prompt，
        # 若 URL 指向 /api/chat 则直接传 messages，否则折叠成 prompt
        is_chat_endpoint = "/api/chat" in self.api_url
        headers = {}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"

        if is_chat_endpoint:
            payload = {
                "model": self.model,
                "messages": messages,
                "stream": False,
                "think": self.think_mode,
                "keep_alive": 60,
            }
        else:
            prompt = PromptBuilder.messages_to_prompt(messages)
            payload = {
                "model": self.model,
                "prompt": prompt,
                "stream": False,
                "think": self.think_mode,
                "keep_alive": 60,
            }
        payload.update(self.extra_params)
        timeout = self.get_timeout(self.estimate_chars(messages), think_mode=self.think_mode)
        try:
            resp = requests.post(self.api_url, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            response = ""
            if isinstance(data, dict):
                response = data.get("response") or data.get("text") or ""
                if not response and data.get("choices"):
                    try:
                        response = data["choices"][0]["message"]["content"]
                    except Exception:
                        pass
                if not response and data.get("message"):
                    msg = data["message"]
                    if isinstance(msg, dict):
                        response = msg.get("content", "") or ""
            elif isinstance(data, str):
                response = data
            return response or "", self.extract_tokens(data)
        except Exception:
            return "", 0

    def list_models(self):
        try:
            tags_url = self.api_url.replace("/api/generate", "/api/tags")
            headers = {}
            if self.api_key:
                headers["Authorization"] = f"Bearer {self.api_key}"
            resp = requests.get(tags_url, timeout=5, headers=headers)
            if resp.status_code == 200:
                data = resp.json()
                models = []
                if isinstance(data, dict):
                    src = data.get("models") or data.get("tags") or []
                    models = [m.get("name") if isinstance(m, dict) else m for m in src]
                elif isinstance(data, list):
                    models = [m.get("name") if isinstance(m, dict) else m for m in data]
                return models
        except Exception:
            pass
        return []


class GeminiProvider(LLMProvider):
    @staticmethod
    def _messages_to_gemini_contents(messages):
        """将通用 messages 转为 Gemini contents 格式。
        Gemini 用 role=user/model，且必须交替 user/model，system 角色需内嵌到用户消息。
        """
        contents = []
        system_buf = []
        for m in messages:
            role = m.get("role", "user")
            text = m.get("content", "") or ""
            if role == "system":
                system_buf.append(text)
                continue
            gemini_role = "user" if role in ("user",) else "model"
            parts_text = text
            if system_buf:
                # 将累计的 system 指令拼入第一个用户消息前
                parts_text = "\n\n".join(system_buf) + "\n\n" + parts_text
                system_buf = []
            contents.append({"role": gemini_role, "parts": [{"text": parts_text}]})
        # 若最后还剩 system_buf（即无后续用户消息），追加一个占位
        if system_buf and not contents:
            contents.append({"role": "user", "parts": [{"text": "\n\n".join(system_buf)}]})
        # 确保首条为 user
        if contents and contents[0].get("role") != "user":
            contents.insert(0, {"role": "user", "parts": [{"text": "(请开始)"}]})
        return contents

    def chat(self, messages):
        headers = {"x-goog-api-key": self.api_key, "Content-Type": "application/json"}
        contents = self._messages_to_gemini_contents(messages)
        payload = {"contents": contents}
        payload.update(self.extra_params)
        timeout = self.get_timeout(self.estimate_chars(messages), think_mode=self.think_mode)
        try:
            resp = requests.post(self.api_url, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            # Gemini 返回结构: candidates[0].content.parts[0].text
            response = ""
            cands = data.get("candidates") or []
            if cands:
                cand = cands[0] if isinstance(cands[0], dict) else {}
                content = cand.get("content") or {}
                parts = content.get("parts") or []
                if parts and isinstance(parts[0], dict):
                    response = parts[0].get("text", "") or ""
            # 兼容旧结构
            if not response:
                response = (data.get("contents", [{}])[0].get("parts", [{}])[0]
                            .get("text", ""))
            tokens = data.get("usageMetadata", {}).get("totalTokenCount", 0)
            return response, tokens
        except Exception:
            return "", 0

    def list_models(self):
        try:
            url = "https://generativelanguage.googleapis.com/v1beta/models"
            resp = requests.get(url, params={"key": self.api_key}, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                return [m["name"].split("/")[-1] for m in data.get("models", [])]
        except Exception:
            pass
        return []


class AnthropicProvider(LLMProvider):
    def chat(self, messages):
        headers = {"x-api-key": self.api_key, "Content-Type": "application/json"}
        version = self.extra_params.get("anthropic-version", "2023-06-01")
        headers["anthropic-version"] = version
        endpoint = self.api_url
        if not (endpoint.rstrip("/").endswith("/messages") or
                endpoint.rstrip("/").endswith("/messages/")):
            endpoint = endpoint.rstrip("/") + "/messages"

        # Anthropic messages 接口需要 system 单独参数
        system_text = ""
        chat_messages = []
        for m in messages:
            role = m.get("role", "user")
            content = m.get("content", "") or ""
            if role == "system":
                if system_text:
                    system_text += "\n\n" + content
                else:
                    system_text = content
            else:
                # Anthropic 只接受 user/assistant 角色
                ar = "assistant" if role == "assistant" else "user"
                chat_messages.append({"role": ar, "content": content})

        payload = {
            "model": self.model,
            "max_tokens": self.extra_params.get("max_tokens", 4096),
            "messages": chat_messages,
        }
        if system_text:
            payload["system"] = system_text
        for k, v in self.extra_params.items():
            if k not in payload and k != "anthropic-version":
                payload[k] = v
        timeout = self.get_timeout(self.estimate_chars(messages), think_mode=self.think_mode)
        try:
            resp = requests.post(endpoint, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            response = ""
            if isinstance(data, dict):
                response = (data.get("completion") or data.get("response")
                            or data.get("text") or "")
                if not response and data.get("choices"):
                    try:
                        response = data["choices"][0]["message"]["content"]
                    except Exception:
                        pass
                # 新版 Anthropic 返回 content 数组
                if not response and isinstance(data.get("content"), list):
                    parts = []
                    for blk in data["content"]:
                        if isinstance(blk, dict) and blk.get("type") == "text":
                            parts.append(blk.get("text", ""))
                    response = "".join(parts)
            return response, self.extract_tokens(data)
        except Exception:
            return "", 0

    def list_models(self):
        try:
            url = self.api_url
            if "/messages" in url:
                url = url.replace("/messages", "/models")
            else:
                url = url.rstrip("/") + "/models"
            headers = {"x-api-key": self.api_key,
                       "anthropic-version": self.extra_params.get("anthropic-version", "2023-06-01")}
            params = {}
            if self.extra_params.get("limit"):
                params["limit"] = int(self.extra_params["limit"])
            resp = requests.get(url, headers=headers, timeout=5, params=params or None)
            if resp.status_code == 200:
                data = resp.json()
                candidates = []
                if isinstance(data, dict):
                    candidates = (data.get("data") or data.get("models")
                                 or (data.get("models", {}).get("data") if isinstance(data.get("models"), dict) else [])
                                 or [])
                elif isinstance(data, list):
                    candidates = data
                models = []
                for m in candidates:
                    if isinstance(m, dict):
                        mid = (m.get("id") or m.get("model") or m.get("name")
                               or m.get("displayName") or m.get("display_name"))
                        if mid:
                            models.append(mid)
                    else:
                        models.append(m)
                return models
        except Exception:
            pass
        return []


class MiMoProvider(LLMProvider):
    def chat(self, messages):
        headers = {"api-key": self.api_key, "Content-Type": "application/json"}
        payload = {
            "model": self.model,
            "messages": messages,
            "max_completion_tokens": self.extra_params.get("max_completion_tokens", 4096),
            "temperature": self.extra_params.get("temperature", 1.0),
            "top_p": self.extra_params.get("top_p", 0.95),
            "stream": False,
        }
        payload.update(self.extra_params)
        endpoint = self.api_url
        if not endpoint.endswith("/chat/completions"):
            endpoint = endpoint.rstrip("/") + "/chat/completions"
        timeout = self.get_timeout(self.estimate_chars(messages), think_mode=self.think_mode)
        try:
            resp = requests.post(endpoint, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            return self.extract_content(data), self.extract_tokens(data)
        except Exception:
            return "", 0


class TianyiProvider(LLMProvider):
    """天翼云 Wishub X6。"""

    def chat(self, messages):
        headers = {"Authorization": f"Bearer {self.api_key}",
                   "Content-Type": "application/json",
                   "User-Agent": "PostmanRuntime-ApipostRuntime/1.1.0"}
        endpoint = self.api_url
        if not endpoint.endswith("/chat/completions"):
            endpoint = endpoint.rstrip("/") + "/chat/completions"
        payload = {"model": self.model,
                   "messages": messages,
                   "stream": False}
        payload.update(self.extra_params)
        timeout = self.get_timeout(self.estimate_chars(messages), think_mode=self.think_mode)
        try:
            resp = requests.post(endpoint, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            response = ""
            if isinstance(data, dict):
                try:
                    response = data["choices"][0]["message"]["content"]
                except Exception:
                    response = data.get("response") or data.get("text") or ""
            return response, self.extract_tokens(data)
        except Exception:
            return "", 0

    def list_models(self):
        try:
            headers = {"Authorization": f"Bearer {self.api_key}",
                        "Content-Type": "application/json",
                        "User-Agent": "PostmanRuntime-ApipostRuntime/1.1.0"}
            if "/v1" in self.api_url:
                base = self.api_url.split("/v1")[0] + "/v1"
            else:
                base = self.api_url.rstrip("/")
            url = base.rstrip("/") + "/models"
            resp = requests.get(url, headers=headers, timeout=5)
            if resp.status_code == 200:
                data = resp.json()
                return [m.get("id") for m in data.get("data", [])]
        except Exception:
            pass
        return []


class FreeModelProvider(LLMProvider):
    """免费模型聚合（kilo / opencode / openrouter）。"""

    def chat(self, messages):
        headers = {"Content-Type": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
            headers["x-api-key"] = self.api_key
        payload = {"model": self.model,
                   "messages": messages,
                   "stream": False}
        payload.update(self.extra_params)
        timeout = self.get_timeout(self.estimate_chars(messages), think_mode=self.think_mode)
        try:
            resp = requests.post(self.api_url, json=payload, headers=headers, timeout=timeout)
            resp.raise_for_status()
            data = resp.json()
            return self.extract_content(data), self.extract_tokens(data)
        except Exception:
            return "", 0


# 供应商注册表：api_type -> Provider 类
PROVIDER_REGISTRY = {}


def _register_providers():
    """注册各 api_type 对应的 Provider。"""
    openai_compat_types = ["OpenAI", "Qwen", "Zhipu", "LM Studio", "MiniMax",
                          "OpenRouter", "移动云", "移动云(Coding)", "联通云",
                          "Doubao", "DeepSeek"]
    for t in openai_compat_types:
        PROVIDER_REGISTRY[t] = OpenAICompatProvider
    PROVIDER_REGISTRY["Baidu"] = BaiduProvider
    PROVIDER_REGISTRY["Tencent"] = TencentProvider
    PROVIDER_REGISTRY["Ollama"] = OllamaProvider
    PROVIDER_REGISTRY["Gemini"] = GeminiProvider
    PROVIDER_REGISTRY["Anthropic"] = AnthropicProvider
    PROVIDER_REGISTRY["MiMo"] = MiMoProvider
    PROVIDER_REGISTRY["天翼云"] = TianyiProvider
    PROVIDER_REGISTRY["免费模型"] = FreeModelProvider


_register_providers()


def create_provider(api_type, api_url, api_key, model, extra_params, think_mode):
    """工厂方法：根据 api_type 创建对应的 Provider 实例。"""
    cls = PROVIDER_REGISTRY.get(api_type, OpenAICompatProvider)
    provider = cls(api_url, api_key, model, extra_params, think_mode)
    provider.api_url_type = api_type
    return provider


# ============================================================================
# 核心处理器
# ============================================================================

class SplitProcessor:
    """断句处理核心：任务分组、并行调用、重试、缓冲写出。"""

    def __init__(self, app, log=None):
        self.app = app
        self.log = log or (lambda msg: None)
        self._stop_flag = threading.Event()
        self.output_lock = threading.Lock()
        self.output_buffer = None
        self.failure_counts = {}
        self.last_failure_info = {}
        self.all_fallback = []
        self.total_tokens = 0
        self.start_time = None

    def request_stop(self):
        self._stop_flag.set()

    @staticmethod
    def group_indices_into_tasks(indices, task_size, lines, max_chars):
        """将待处理行索引分组为任务，兼顾行数与字符阈值。"""
        if not indices:
            return []
        indices = sorted(set(indices))
        tasks = []
        current = []
        current_chars = 0
        for idx in indices:
            line_text = lines[idx] if lines and idx < len(lines) else ""
            line_chars = len(line_text) + 20
            if current and (len(current) >= task_size or
                            (max_chars and current_chars + line_chars > max_chars)):
                tasks.append(current)
                current = []
                current_chars = 0
            current.append(idx)
            current_chars += line_chars
        if current:
            tasks.append(current)
        return tasks

    def run(self, input_path, output_path, task_size, max_chars,
            think_mode, prompt_builder, progress_cb=None, done_cb=None):
        """主处理流程。"""
        self._stop_flag.clear()
        self.start_time = time.time()
        self.total_tokens = 0
        self.failure_counts = {}
        self.last_failure_info = {}
        self.all_fallback = []

        if not os.path.exists(input_path):
            self.log(f"错误：输入文件不存在: {input_path}")
            return

        with open(input_path, "r", encoding="utf-8") as f:
            lines = [ln.rstrip("\n") for ln in f]
        total = len(lines)
        self.log(f"开始处理：输入 {input_path}")
        self.log(f"输出文件: {output_path}")
        self.log(f"总行数: {total}，每任务行数: {task_size}，字符阈值: {max_chars}")

        # 续传检测
        processed_count = 0
        out_lines = []
        if os.path.exists(output_path):
            with open(output_path, "r", encoding="utf-8") as f:
                out_lines = [ln.rstrip("\n") for ln in f]
            processed_count = len(out_lines)
            if processed_count < total:
                if not messagebox.askyesno("续传", f"检测到输出文件已有 {processed_count} 行结果，是否从该行继续处理？"):
                    processed_count = 0
                    open(output_path, "w", encoding="utf-8").close()
            elif processed_count >= total:
                if messagebox.askyesno("完成", "输出文件已完成所有行。是否重新开始并覆盖已有输出？"):
                    processed_count = 0
                    open(output_path, "w", encoding="utf-8").close()
                else:
                    self.log("已取消处理。")
                    if done_cb:
                        done_cb()
                    return

        self.output_buffer = lines[:]
        if processed_count > 0:
            for i in range(min(processed_count, total)):
                self.output_buffer[i] = out_lines[i]

        short_lines = [i for i in range(total) if len(lines[i].strip()) < SHORT_LINE_THRESHOLD]
        pending = [i for i in range(processed_count, total)
                   if len(lines[i].strip()) >= SHORT_LINE_THRESHOLD]
        tasks = self.group_indices_into_tasks(pending, task_size, lines, max_chars)
        self.log(f"短行 {len(short_lines)} 行跳过；待处理 {len(pending)} 行，初始任务 {len(tasks)} 个")

        if progress_cb:
            progress_cb(0, total, short_lines)
        completed = processed_count + len(short_lines) if processed_count == 0 else processed_count

        all_failed, first_success, first_failed = self._run_tasks(
            tasks, lines, prompt_builder, output_path,
            total, completed, short_lines, progress_cb)
        first_fallback_count = len(self.all_fallback)

        # 重试失败行
        retry_round = 0
        while all_failed and not self._stop_flag.is_set() and retry_round < self.app.max_retries:
            retry_round += 1
            self.log(f"第 {retry_round} 轮重试，剩余 {len(all_failed)} 行")
            retry_tasks = self.group_indices_into_tasks(all_failed, task_size, lines, max_chars)
            all_failed, _, _ = self._run_tasks(
                retry_tasks, lines, prompt_builder, output_path,
                total, 0, [], progress_cb, is_retry=True)

        elapsed = time.time() - self.start_time
        model_name = self.app.model_var.get() or "未指定模型"
        first_fail_total = first_failed + first_fallback_count
        first_total = first_success + first_fail_total
        first_rate = (first_success / first_total * 100) if first_total > 0 else 100.0
        final_fallback = len(self.all_fallback)
        final_success = total - len(all_failed) - final_fallback
        final_failed = len(all_failed) + final_fallback
        final_rate = (final_success / total * 100) if total > 0 else 0.0
        summary = (f"[{model_name}] 统计: 总行数 {total} | "
                   f"首轮: 成功 {first_success}/失败 {first_fail_total}, 成功率 {first_rate:.1f}% | "
                   f"最终: 成功 {final_success}/失败 {final_failed}, 成功率 {final_rate:.1f}%")
        if self._stop_flag.is_set():
            self.log(f"已停止，用时 {elapsed:.1f}s，消耗 {self.total_tokens} tokens")
        elif not all_failed:
            self.log(f"处理完成，用时 {elapsed:.1f}s，消耗 {self.total_tokens} tokens")
        else:
            self.log(f"处理结束，仍有 {len(all_failed)} 行未成功，用时 {elapsed:.1f}s")
        self.log(summary)
        if done_cb:
            done_cb()

    def _run_tasks(self, tasks, lines, prompt_builder, output_path,
                   total, initial_completed, short_lines, progress_cb, is_retry=False):
        """并行执行一批任务，返回 (失败行索引列表, 本轮成功数, 本轮失败数)。"""
        all_failed = []
        round_success = 0
        round_failed = 0
        completed = initial_completed
        max_workers = 5
        with ThreadPoolExecutor(max_workers=max_workers) as ex:
            futures = {ex.submit(self._process_task, task, lines, prompt_builder, is_retry): task
                        for task in tasks}
            for fut in as_completed(futures):
                if self._stop_flag.is_set():
                    break
                try:
                    success, failed, fallback = fut.result()
                    round_success += len(success)
                    round_failed += len(failed)
                    completed += len(success) + len(fallback)
                    all_failed.extend(failed)
                    self.all_fallback.extend(fallback)
                    self._flush_buffer(output_path, lines)
                    if progress_cb:
                        progress_cb(completed, total, short_lines)
                except Exception as e:
                    self.log(f"任务异常: {e}")
        return all_failed, round_success, round_failed

    def _process_task(self, indices, lines, prompt_builder, is_retry=False):
        """处理单个任务（一组行）。返回 (success_indices, failed_indices, fallback_indices)。"""
        failure_info = self.last_failure_info if is_retry else None
        messages = prompt_builder.build(indices, lines, failure_info=failure_info)
        response, tokens = self._call_api(messages)
        self.total_tokens += tokens

        # 空响应处理：可能是本地小模型思考模式导致 content 为空
        if not response or not response.strip():
            # 若开启思考模式，尝试关闭后重试一次
            if self.app.think_mode_var.get():
                self.log(f"行 {indices[0]}-{indices[-1]}: 响应为空，关闭思考模式重试")
                response, tokens = self._call_api(messages, force_no_think=True)
                self.total_tokens += tokens
            if not response or not response.strip():
                self.log(f"行 {indices[0]}-{indices[-1]}: API 返回空响应")
                for idx in indices:
                    self.last_failure_info[idx] = {"reason": "API 返回空响应", "ai_output": ""}
                return [], indices[:], []

        parts = ResponseParser.parse(response, len(indices))
        success, failed, fallback = [], [], []
        for idx, part in zip(indices, parts):
            original = lines[idx]
            result, ok, reason = Validator.validate(original, part)
            if ok:
                self._set_buffer(idx, result)
                success.append(idx)
                self.last_failure_info.pop(idx, None)
                # 详细日志：断句前 → 断句后
                if result != original:
                    self.log(f"行 {idx} ✓: {original}  →  {result}")
                else:
                    self.log(f"行 {idx} ✓: (短句原样) {original}")
            else:
                self.failure_counts[idx] = self.failure_counts.get(idx, 0) + 1
                self.last_failure_info[idx] = {"reason": reason, "ai_output": part}
                if self.failure_counts[idx] > MAX_LINE_FAILURES:
                    self._set_buffer(idx, lines[idx])
                    fallback.append(idx)
                    self.log(f"行 {idx}: 失败超限，回退原文")
                else:
                    failed.append(idx)
                    self.log(f"行 {idx} ✗ {reason}")
                    self.log(f"        原文: {original}")
                    self.log(f"        AI返回: {part}")
        self.log(f"任务完成 行 {indices[0]}-{indices[-1]}: 成功 {len(success)}，失败 {len(failed)}，回退 {len(fallback)}")
        return success, failed, fallback

    def _call_api(self, messages, force_no_think=False):
        api_type = self.app.api_type_var.get()
        api_url = self.app.api_var.get()
        api_key = self.app.api_key_var.get()
        model = self.app.model_var.get()
        try:
            extra_params = json.loads(self.app.extra_params_var.get().strip() or "{}")
        except json.JSONDecodeError:
            extra_params = {}
        think_mode = self.app.think_mode_var.get() and not force_no_think
        provider = create_provider(api_type, api_url, api_key, model, extra_params, think_mode)
        return provider.chat(messages)

    def _set_buffer(self, index, text):
        cleaned = (text or "").strip().replace("\r\n", "\n").replace("\r", "")
        with self.output_lock:
            if self.output_buffer and 0 <= index < len(self.output_buffer):
                self.output_buffer[index] = cleaned

    def _flush_buffer(self, output_path, lines):
        with self.output_lock:
            if self.output_buffer is None:
                return
            with open(output_path, "w", encoding="utf-8") as wf:
                for i in range(len(self.output_buffer)):
                    text = self.output_buffer[i]
                    if text is None:
                        text = lines[i] if lines and i < len(lines) else ""
                    wf.write(text + "\n")


# ============================================================================
# GUI 应用
# ============================================================================

class App:
    """断句工具主界面。"""

    def __init__(self, root, input_path=None, output_path=None):
        self.root = root
        root.title(f"漫画台词智能断句工具 v{SCRIPT_VERSION}")
        self._version_checked = False

        # 变量
        self.input_var = tk.StringVar(value=input_path or DEFAULT_INPUT)
        self.output_var = tk.StringVar(value=output_path or DEFAULT_OUTPUT)
        self.reference_var = tk.StringVar(value="")
        self.api_type_var = tk.StringVar(value="Ollama")
        self.api_var = tk.StringVar(value=DEFAULT_APIS["Ollama"])
        self.api_key_var = tk.StringVar(value="")
        self.ollama_mode_var = tk.StringVar(value="Local")
        self.model_var = tk.StringVar(value="deepseek-v3.2:cloud")
        self.task_size_var = tk.IntVar(value=40)
        self.max_chars_var = tk.IntVar(value=3500)
        self.think_mode_var = tk.BooleanVar(value=False)
        self.output_format_var = tk.StringVar(value="json")  # json / text
        self.prompt_var = tk.StringVar(value=SYSTEM_PROMPT)
        self.extra_params_var = tk.StringVar(value='{"think": false}')
        self.extra_body_var = tk.StringVar(value=DEFAULT_EXTRA_BODIES[0])
        self.api_keys = {}
        self.max_retries = 10

        self.config_file = self._get_config_path()
        self.processor = SplitProcessor(self, log=self._log)
        self.worker_thread = None

        self._setup_ui()
        self.load_config()

    def _setup_ui(self):
        # 文件路径
        frame_files = ttk.LabelFrame(self.root, text="文件")
        frame_files.pack(fill="x", padx=8, pady=4)

        ttk.Label(frame_files, text="输入文件:").grid(row=0, column=0, sticky="w")
        ttk.Entry(frame_files, textvariable=self.input_var, width=55).grid(row=0, column=1, sticky="we", padx=2)
        ttk.Button(frame_files, text="浏览", command=self._browse_input).grid(row=0, column=2, padx=4)

        ttk.Label(frame_files, text="输出文件:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame_files, textvariable=self.output_var, width=55).grid(row=1, column=1, sticky="we", padx=2)
        ttk.Button(frame_files, text="浏览", command=self._browse_output).grid(row=1, column=2, padx=4)

        ttk.Label(frame_files, text="参考文件:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame_files, textvariable=self.reference_var, width=55).grid(row=2, column=1, sticky="we", padx=2)
        ttk.Button(frame_files, text="浏览", command=self._browse_reference).grid(row=2, column=2, padx=4)

        frame_files.columnconfigure(1, weight=1)

        # 参数
        frame_params = ttk.LabelFrame(self.root, text="模型与参数")
        frame_params.pack(fill="x", padx=8, pady=4)

        ttk.Label(frame_params, text="API类型:").grid(row=0, column=0, sticky="w")
        self.api_type_cb = ttk.Combobox(frame_params, textvariable=self.api_type_var,
                                        values=list(DEFAULT_APIS.keys()), state="readonly", width=14)
        self.api_type_cb.grid(row=0, column=1, sticky="w", padx=4)
        self.api_type_cb.bind("<<ComboboxSelected>>", self._on_api_type_changed)

        ttk.Label(frame_params, text="Ollama模式:").grid(row=0, column=2, sticky="e")
        self.ollama_mode_cb = ttk.Combobox(frame_params, textvariable=self.ollama_mode_var,
                                           values=["Local", "Online"], state="readonly", width=10)
        self.ollama_mode_cb.grid(row=0, column=3, sticky="w", padx=4)
        self.ollama_mode_cb.bind("<<ComboboxSelected>>", self._on_ollama_mode_changed)

        ttk.Label(frame_params, text="API地址:").grid(row=1, column=0, sticky="w")
        ttk.Entry(frame_params, textvariable=self.api_var, width=45).grid(row=1, column=1, columnspan=3, sticky="we", padx=4)

        ttk.Label(frame_params, text="API Key:").grid(row=2, column=0, sticky="w")
        ttk.Entry(frame_params, textvariable=self.api_key_var, width=45, show="*").grid(row=2, column=1, columnspan=3, sticky="we", padx=4)

        ttk.Label(frame_params, text="模型名:").grid(row=3, column=0, sticky="w")
        self.model_cb = ttk.Combobox(frame_params, textvariable=self.model_var,
                                     values=DEFAULT_MODELS.get("OpenAI", []), width=22)
        self.model_cb.grid(row=3, column=1, sticky="w", padx=4)
        self.model_cb.bind("<<ComboboxSelected>>", self._on_model_changed)
        ttk.Button(frame_params, text="加载模型", command=self._load_models).grid(row=3, column=3, padx=4)

        ttk.Label(frame_params, text="每任务行数:").grid(row=4, column=0, sticky="w")
        ttk.Spinbox(frame_params, from_=1, to=200, increment=1,
                    textvariable=self.task_size_var, width=8).grid(row=4, column=1, sticky="w", padx=4)

        ttk.Label(frame_params, text="字符阈值:").grid(row=4, column=2, sticky="e")
        ttk.Spinbox(frame_params, from_=500, to=20000, increment=500,
                    textvariable=self.max_chars_var, width=10).grid(row=4, column=3, sticky="w", padx=4)

        ttk.Label(frame_params, text="输出格式:").grid(row=5, column=0, sticky="w")
        ttk.Radiobutton(frame_params, text="JSON (省token, 推荐)", value="json",
                        variable=self.output_format_var).grid(row=5, column=1, sticky="w")
        ttk.Radiobutton(frame_params, text="纯文本", value="text",
                        variable=self.output_format_var).grid(row=5, column=2, columnspan=2, sticky="w")

        ttk.Checkbutton(frame_params, text="开启思考（慢，本地小模型不建议）",
                        variable=self.think_mode_var).grid(row=6, column=0, columnspan=2, sticky="w")

        # 高级设置
        self.advanced_shown = False
        ac_frame = ttk.Frame(self.root)
        ac_frame.pack(fill="x", padx=8, pady=2)
        self.adv_button = ttk.Button(ac_frame, text="显示高级设置", command=self._toggle_advanced)
        self.adv_button.pack(side="left")
        self.adv_panel = ttk.Frame(self.root)

        ttk.Label(self.adv_panel, text="系统提示词:").grid(row=0, column=0, sticky="nw")
        self.prompt_text = tk.Text(self.adv_panel, width=80, height=8, wrap="word")
        self.prompt_text.grid(row=0, column=1, sticky="we", padx=4)
        self.prompt_text.insert("1.0", SYSTEM_PROMPT)

        ttk.Label(self.adv_panel, text="额外参数(JSON):").grid(row=1, column=0, sticky="nw")
        ttk.Entry(self.adv_panel, textvariable=self.extra_params_var, width=80).grid(row=1, column=1, sticky="we", padx=4)

        ttk.Label(self.adv_panel, text="思考模式body:").grid(row=2, column=0, sticky="nw")
        self.extra_body_cb = ttk.Combobox(self.adv_panel, textvariable=self.extra_body_var,
                                          values=DEFAULT_EXTRA_BODIES, width=70, state="readonly")
        self.extra_body_cb.grid(row=2, column=1, sticky="we", padx=4)
        ttk.Button(self.adv_panel, text="应用", command=self._apply_extra_body).grid(row=2, column=2, padx=4)

        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)

        # 控制与进度
        frame_ctrl = ttk.Frame(self.root)
        frame_ctrl.pack(fill="x", padx=8, pady=4)
        self.progress = ttk.Progressbar(frame_ctrl, length=380, mode="determinate")
        self.progress.grid(row=0, column=0, padx=4)
        self.status_label = ttk.Label(frame_ctrl, text="未开始")
        self.status_label.grid(row=0, column=1, padx=8)
        self.start_btn = ttk.Button(frame_ctrl, text="开始处理", command=self._toggle_start)
        self.start_btn.grid(row=0, column=2, padx=4)
        self.eval_btn = ttk.Button(frame_ctrl, text="对比评估", command=self._evaluate)
        self.eval_btn.grid(row=0, column=3, padx=4)

        # 日志面板
        frame_log = ttk.LabelFrame(self.root, text="日志")
        frame_log.pack(fill="both", expand=True, padx=8, pady=4)
        self.log_text = tk.Text(frame_log, height=10, wrap="word", state="disabled")
        self.log_text.pack(fill="both", expand=True, side="left")
        log_scroll = ttk.Scrollbar(frame_log, command=self.log_text.yview)
        log_scroll.pack(side="right", fill="y")
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.tag_configure("success", foreground="#2e7d32")
        self.log_text.tag_configure("failure", foreground="#c62828")
        self._last_log_tag = None

    # ---------- 文件浏览 ----------
    def _browse_input(self):
        p = filedialog.askopenfilename(title="选择输入 txt",
                                       filetypes=[("Text", "*.txt"), ("All", "*.*")])
        if p:
            self.input_var.set(p)

    def _browse_output(self):
        p = filedialog.asksaveasfilename(title="选择输出文件", defaultextension=".txt",
                                        filetypes=[("Text", "*.txt"), ("All", "*.*")])
        if p:
            self.output_var.set(p)

    def _browse_reference(self):
        p = filedialog.askopenfilename(title="选择人工断句参考 txt",
                                       filetypes=[("Text", "*.txt"), ("All", "*.*")])
        if p:
            self.reference_var.set(p)

    # ---------- 高级设置 ----------
    def _apply_extra_body(self):
        body = self.extra_body_var.get().strip()
        if body:
            self.extra_params_var.set(body)

    def _toggle_advanced(self):
        if not self.advanced_shown:
            self.adv_panel.pack(fill="x", padx=8, pady=4)
            self.adv_button.config(text="隐藏高级设置")
            self.advanced_shown = True
        else:
            self.adv_panel.forget()
            self.adv_button.config(text="显示高级设置")
            self.advanced_shown = False

    # ---------- API 切换 ----------
    def _on_model_changed(self, event=None):
        if self.api_type_var.get() == "免费模型":
            model = self.model_var.get()
            url = FREE_MODEL_API_URLS.get(model)
            if url:
                self.api_var.set(url)
            if model == "openrouter/free":
                self.api_key_var.set(DEFAULT_OPENROUTER_FREE_KEY)
            elif not self.api_key_var.get().strip():
                self.api_key_var.set("")

    def _on_api_type_changed(self, event=None):
        api_type = self.api_type_var.get()
        if api_type == "Ollama":
            mode = self.ollama_mode_var.get()
            self.api_var.set("https://ollama.com/api/generate" if mode == "Online"
                             else DEFAULT_APIS.get("Ollama", ""))
            self.ollama_mode_cb.grid()
            self.api_key_var.set(self.api_keys.get(f"Ollama_{mode}", ""))
        elif api_type == "免费模型":
            self.api_var.set(DEFAULT_APIS.get(api_type, ""))
            self.model_cb["values"] = DEFAULT_MODELS.get(api_type, [])
            self.model_var.set(DEFAULT_MODELS.get(api_type, [""])[0] if DEFAULT_MODELS.get(api_type) else "")
            self.model_cb.config(state="readonly")
            if self.model_var.get() == "openrouter/free":
                self.api_key_var.set(DEFAULT_OPENROUTER_FREE_KEY)
            else:
                self.api_key_var.set("")
            try:
                self.ollama_mode_cb.grid_forget()
            except Exception:
                pass
        else:
            self.api_var.set(DEFAULT_APIS.get(api_type, ""))
            try:
                self.ollama_mode_cb.grid_forget()
            except Exception:
                pass
            self.api_key_var.set(self.api_keys.get(api_type, ""))
            self.model_cb.config(state="normal")
        if api_type != "免费模型":
            self.model_cb["values"] = DEFAULT_MODELS.get(api_type, [])
            self.model_var.set(DEFAULT_MODELS.get(api_type, [""])[0] if DEFAULT_MODELS.get(api_type) else "")

    def _on_ollama_mode_changed(self, event=None):
        if self.api_type_var.get() == "Ollama":
            mode = self.ollama_mode_var.get()
            self.api_var.set("https://ollama.com/api/generate" if mode == "Online"
                             else DEFAULT_APIS.get("Ollama", ""))
            self.api_key_var.set(self.api_keys.get(f"Ollama_{mode}", ""))

    def _load_models(self):
        api_type = self.api_type_var.get()
        api_url = self.api_var.get().strip()
        api_key = self.api_key_var.get().strip()

        if api_type == "免费模型":
            self.model_cb["values"] = DEFAULT_MODELS["免费模型"]
            self.model_var.set(DEFAULT_MODELS["免费模型"][0])
            self.api_var.set(FREE_MODEL_API_URLS[self.model_var.get()])
            self._log("免费模型已加载")
            return

        if api_type not in ("Ollama", "LM Studio") and (not api_url or not api_key):
            messagebox.showerror("错误", "请先填写API地址和Key")
            return
        if api_type == "Ollama":
            if not api_url:
                messagebox.showerror("错误", "请先填写API地址")
                return
            if self.ollama_mode_var.get() == "Online" and not api_key:
                messagebox.showerror("错误", "使用 Ollama 在线服务需要填写 API Key")
                return

        try:
            extra_params = json.loads(self.extra_params_var.get().strip() or "{}")
        except json.JSONDecodeError:
            extra_params = {}
        provider = create_provider(api_type, api_url, api_key, self.model_var.get(),
                                   extra_params, self.think_mode_var.get())
        models = provider.list_models()
        if models:
            self.model_cb["values"] = models
            self.model_var.set(models[0])
            self._log(f"加载成功: {len(models)} 个模型")
        else:
            self._log("加载模型列表失败")
            messagebox.showerror("错误", "无法加载模型列表")

    # ---------- 处理控制 ----------
    def _toggle_start(self):
        if self.worker_thread and self.worker_thread.is_alive():
            self.processor.request_stop()
            self.start_btn.config(text="开始处理")
            self.status_label.config(text="停止中...")
        else:
            self.worker_thread = threading.Thread(target=self._run_processing, daemon=True)
            self.worker_thread.start()

    def _run_processing(self):
        self.start_btn.config(text="停止")
        input_path = self.input_var.get()
        output_path = self.output_var.get()

        def progress_cb(done, total, short_lines):
            self.progress["maximum"] = total
            self.progress["value"] = done
            self.status_label.config(text=f"已完成 {done}/{total}")

        def done_cb():
            self.start_btn.config(text="开始处理")
            if not self.processor._stop_flag.is_set():
                messagebox.showinfo("完成", "断句已完成")

        prompt = self.prompt_text.get("1.0", "end-1c") or SYSTEM_PROMPT
        # 保存用户编辑的提示词
        self.prompt_var.set(prompt)
        prompt_builder = PromptBuilder(prompt, self.output_format_var.get())

        self.processor.run(
            input_path=input_path,
            output_path=output_path,
            task_size=self.task_size_var.get(),
            max_chars=self.max_chars_var.get(),
            think_mode=self.think_mode_var.get(),
            prompt_builder=prompt_builder,
            progress_cb=progress_cb,
            done_cb=done_cb,
        )

    # ---------- 对比评估 ----------
    def _evaluate(self):
        ref_path = self.reference_var.get().strip()
        out_path = self.output_var.get().strip()
        if not ref_path:
            messagebox.showwarning("提示", "请先选择人工断句参考文件")
            return
        if not out_path or not os.path.exists(out_path):
            messagebox.showwarning("提示", "输出文件不存在，请先运行断句")
            return
        try:
            with open(ref_path, "r", encoding="utf-8") as f:
                ref_lines = [ln.rstrip("\n") for ln in f]
            with open(out_path, "r", encoding="utf-8") as f:
                ai_lines = [ln.rstrip("\n") for ln in f]
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {e}")
            return

        result = Evaluator.evaluate(ai_lines, ref_lines)
        report = Evaluator.format_report(result)
        self._log(report)
        # 弹窗显示
        win = tk.Toplevel(self.root)
        win.title("断句质量评估报告")
        win.geometry("780x560")
        txt = tk.Text(win, wrap="word")
        txt.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(win, command=txt.yview)
        sb.pack(side="right", fill="y")
        txt.configure(yscrollcommand=sb.set)
        txt.insert("1.0", report)
        txt.config(state="disabled")

    # ---------- 日志 ----------
    def _log(self, msg):
        def append():
            text = str(msg)
            tag = ""
            if "✓" in text:
                tag = "success"
                self._last_log_tag = None
            elif "✗" in text:
                tag = "failure"
                self._last_log_tag = "failure"
            elif text.startswith("        原文:") or text.startswith("        AI返回:"):
                tag = self._last_log_tag or ""
            else:
                self._last_log_tag = None
            self.log_text.config(state="normal")
            if tag:
                self.log_text.insert("end", text + "\n", tag)
            else:
                self.log_text.insert("end", text + "\n")
            self.log_text.see("end")
            self.log_text.config(state="disabled")
        try:
            self.root.after(0, append)
        except Exception:
            pass

    # ---------- 配置 ----------
    def _get_config_path(self):
        system = platform.system()
        if system == "Windows":
            d = os.path.join(os.environ.get("APPDATA", ""), "Mangahanhua")
        elif system == "Darwin":
            d = os.path.join(os.path.expanduser("~/Library/Application Support"), "Mangahanhua")
        else:
            d = os.path.expanduser("~/.config/Mangahanhua")
        os.makedirs(d, exist_ok=True)
        return os.path.join(d, "config.json")

    @staticmethod
    def _parse_version(v):
        """把 'x.y.z' 版本号转成可比的整数元组，无法解析返回 (0,)。"""
        try:
            return tuple(int(x) for x in str(v).strip().split("."))
        except Exception:
            return (0,)

    def _prompt_keys_override(self, data):
        """用脚本默认值覆盖提示词相关字段，保留模型/API相关。"""
        self.prompt_var.set(SYSTEM_PROMPT)
        self.prompt_text.delete("1.0", "end")
        self.prompt_text.insert("1.0", SYSTEM_PROMPT)
        self.output_format_var.set("json")
        self.task_size_var.set(40)
        self.max_chars_var.set(3500)
        self.extra_params_var.set('{"think": false}')
        self.extra_body_var.set(DEFAULT_EXTRA_BODIES[0])
        # 其他字段从 data 保留（模型/API/Key 等）

    def load_config(self):
        if not os.path.exists(self.config_file):
            self._version_checked = True
            return
        try:
            with open(self.config_file, "r", encoding="utf-8") as f:
                encrypted = f.read()
            data = json.loads(base64.b64decode(encrypted).decode("utf-8"))

            saved_version = data.get("version")
            current_parsed = self._parse_version(SCRIPT_VERSION)
            saved_parsed = self._parse_version(saved_version) if saved_version else (0,)
            needs_upgrade = (not saved_version) or (saved_parsed < current_parsed)

            # —— 先从配置加载所有字段 ——
            self.api_type_var.set(data.get("api_type", "Ollama"))
            self.api_var.set(data.get("api_url", DEFAULT_APIS["Ollama"]))
            self.api_keys = data.get("api_keys", {})
            self.ollama_mode_var.set(data.get("ollama_mode", "Local"))
            self.model_var.set(data.get("model", "deepseek-v3.2:cloud"))
            self.task_size_var.set(data.get("task_size", 8))
            self.max_chars_var.set(data.get("max_chars", 3500))
            self.think_mode_var.set(data.get("think_mode", False))
            self.output_format_var.set(data.get("output_format", "json"))
            self.reference_var.set(data.get("reference", ""))
            self.extra_params_var.set(data.get("extra_params", '{"think": false}'))
            self.extra_body_var.set(data.get("extra_body", DEFAULT_EXTRA_BODIES[0]))
            saved_prompt = data.get("prompt", SYSTEM_PROMPT)
            self.prompt_var.set(saved_prompt)
            self.prompt_text.delete("1.0", "end")
            self.prompt_text.insert("1.0", saved_prompt)
            self._on_api_type_changed()
            self._version_checked = True

            # —— 版本升级判断：提示词字段用脚本新值覆盖，保留模型/API设置 ——
            if needs_upgrade:
                old_ver = saved_version if saved_version else "（无版本号）"
                msg = (f"检测到提示词模板已更新：\n"
                       f"  配置版本: {old_ver}\n"
                       f"  当前脚本版本: {SCRIPT_VERSION}\n\n"
                       f"是否用脚本内的新提示词覆盖配置？\n"
                       f"（模型、API地址、密钥等设置会保留）")
                if messagebox.askyesno("提示词升级", msg):
                    self._prompt_keys_override(data)
                    self._log(f"提示词模板已升级到 v{SCRIPT_VERSION}")
                else:
                    self._log(f"用户保留旧提示词（配置 v{old_ver}）")
            else:
                self._log("配置加载成功")
        except Exception as e:
            self._log(f"加载配置失败: {e}")

    def save_config(self):
        if self.api_type_var.get() == "Ollama":
            self.api_keys[f"Ollama_{self.ollama_mode_var.get()}"] = self.api_key_var.get()
        else:
            self.api_keys[self.api_type_var.get()] = self.api_key_var.get()
        current_prompt = self.prompt_text.get("1.0", "end-1c")
        data = {
            "version": SCRIPT_VERSION,
            "api_type": self.api_type_var.get(),
            "api_url": self.api_var.get(),
            "api_keys": self.api_keys,
            "ollama_mode": self.ollama_mode_var.get(),
            "model": self.model_var.get(),
            "task_size": self.task_size_var.get(),
            "max_chars": self.max_chars_var.get(),
            "think_mode": self.think_mode_var.get(),
            "output_format": self.output_format_var.get(),
            "reference": self.reference_var.get(),
            "prompt": current_prompt,
            "extra_params": self.extra_params_var.get(),
            "extra_body": self.extra_body_var.get(),
        }
        try:
            encrypted = base64.b64encode(
                json.dumps(data, ensure_ascii=False).encode("utf-8")
            ).decode("utf-8")
            with open(self.config_file, "w", encoding="utf-8") as f:
                f.write(encrypted)
        except Exception as e:
            self._log(f"保存配置失败: {e}")

    def _on_closing(self):
        self.save_config()
        self.root.destroy()


# ============================================================================
# 入口
# ============================================================================

def main():
    # 接受 JSX 传入的输入/输出文件路径参数
    args = sys.argv[1:]
    input_path = args[0] if len(args) > 0 else DEFAULT_INPUT
    output_path = args[1] if len(args) > 1 else DEFAULT_OUTPUT

    root = tk.Tk()
    App(root, input_path=input_path, output_path=output_path)
    root.mainloop()


if __name__ == "__main__":
    main()
