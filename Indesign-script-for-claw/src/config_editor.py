# -*- coding: utf-8 -*-
"""
Manga Layout Config Editor
用于编辑 manga_layout_config.json 的图形界面
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os

class ConfigEditor:
    def __init__(self, root, config_path):
        self.root = root
        self.root.title("Manga Layout Config Editor")
        self.config_path = config_path
        self.config = {}
        self.widgets = {}
        self.current_row = 0
        self.current_col = 0
        
        # 创建主框架和滚动区域
        self.main_frame = ttk.Frame(root, padding="3")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建Canvas和Scrollbar
        self.canvas = tk.Canvas(self.main_frame)
        self.scrollbar = ttk.Scrollbar(self.main_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # 绑定鼠标滚轮
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
        self.canvas.pack(side="left", fill=tk.BOTH, expand=True)
        self.scrollbar.pack(side="right", fill=tk.Y)
        
        # 底部按钮
        self.button_frame = ttk.Frame(root, padding="3")
        self.button_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        ttk.Button(self.button_frame, text="保存", command=self.save_config, width=8).pack(side=tk.RIGHT, padx=2)
        ttk.Button(self.button_frame, text="另存为...", command=self.save_as, width=8).pack(side=tk.RIGHT, padx=2)
        ttk.Button(self.button_frame, text="重新加载", command=self.reload_config, width=8).pack(side=tk.RIGHT, padx=2)
        
        # 加载配置
        self.load_config()
        self.create_widgets()
    
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def load_config(self):
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
        except FileNotFoundError:
            messagebox.showerror("错误", f"配置文件未找到: {self.config_path}")
            self.config = {}
        except json.JSONDecodeError as e:
            messagebox.showerror("错误", f"JSON解析错误: {e}")
            self.config = {}
    
    def save_config(self, path=None):
        self.collect_values()
        save_path = path or self.config_path
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"配置已保存到: {save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"保存失败: {e}")
    
    def save_as(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialfile=os.path.basename(self.config_path)
        )
        if path:
            self.save_config(path)
    
    def reload_config(self):
        self.load_config()
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.widgets = {}
        self.current_row = 0
        self.current_col = 0
        self.create_widgets()
    
    def add_group(self, parent, title, key, fields, sub_groups=None):
        """添加一个配置组，使用两列布局"""
        frame = ttk.LabelFrame(parent, text=title, padding="3")
        frame.pack(fill=tk.X, padx=2, pady=2)
        
        # 两列容器
        col_frame = ttk.Frame(frame)
        col_frame.pack(fill=tk.X)
        
        col0 = ttk.Frame(col_frame)
        col0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        col1 = ttk.Frame(col_frame)
        col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 分配字段到两列
        mid = (len(fields) + 1) // 2
        for i, field in enumerate(fields):
            parent_col = col0 if i < mid else col1
            self.create_field(parent_col, field, key)
        
        # 子分组
        if sub_groups:
            for sub_title, sub_key, sub_fields in sub_groups:
                sub_frame = ttk.LabelFrame(frame, text=sub_title, padding="2")
                sub_frame.pack(fill=tk.X, pady=2)
                sub_col_frame = ttk.Frame(sub_frame)
                sub_col_frame.pack(fill=tk.X)
                sub_col0 = ttk.Frame(sub_col_frame)
                sub_col0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                sub_col1 = ttk.Frame(sub_col_frame)
                sub_col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                sub_mid = (len(sub_fields) + 1) // 2
                for i, f in enumerate(sub_fields):
                    self.create_field(sub_col0 if i < sub_mid else sub_col1, f, sub_key)
    
    def create_field(self, parent, field_info, parent_key):
        """创建单个字段"""
        if isinstance(field_info, tuple):
            key, label = field_info
        else:
            key = field_info
            label = field_info
        
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=1)
        
        # 获取值
        value = self.config
        full_key = parent_key
        for k in parent_key.split('.'):
            value = value.get(k, {})
        value = value.get(key, "") if isinstance(value, dict) else ""
        
        # 标签
        ttk.Label(row, text=label, width=18, anchor='w', font=('', 8)).pack(side=tk.LEFT)
        
        # 根据类型创建控件
        widget_key = f"{parent_key}.{key}"
        
        if isinstance(value, bool):
            var = tk.StringVar(value=str(value))
            cb = ttk.Combobox(row, textvariable=var, values=["true", "false"], width=10, state="readonly")
            cb.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.widgets[widget_key] = ('bool', var)
        elif isinstance(value, (int, float)):
            var = tk.StringVar(value=str(value))
            entry = ttk.Entry(row, textvariable=var, width=12)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.widgets[widget_key] = ('number', var)
        elif isinstance(value, list):
            var = tk.StringVar(value=json.dumps(value, ensure_ascii=False))
            entry = ttk.Entry(row, textvariable=var, width=12)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            self.widgets[widget_key] = ('list', var)
        else:
            var = tk.StringVar(value=str(value) if value is not None else "")
            entry = ttk.Entry(row, textvariable=var, width=12)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            if 'Path' in key or 'path' in key:
                ttk.Button(row, text="...", width=2,
                    command=lambda v=var: self.browse_file(v)).pack(side=tk.LEFT, padx=1)
            self.widgets[widget_key] = ('string', var)
    
    def create_widgets(self):
        # 基本信息
        self.add_group(self.scrollable_frame, "基本信息", "", 
            [("version", "版本"), ("description", "描述")])
        
        # templateDocument
        self.add_group(self.scrollable_frame, "模板文档", "templateDocument", 
            [("enabled", "启用"), ("indtPath", "模板路径"), ("description", "说明")])
        
        # documentSettings
        self.add_group(self.scrollable_frame, "文档设置", "documentSettings",
            [("pageWidth", "页面宽度"), ("pageHeight", "页面高度"), ("unit", "单位"),
             ("pageBinding", "装订方向"), ("startPage", "起始页"), ("totalPages", "总页数"),
             ("bleedTop", "出血上"), ("bleedBottom", "出血下"), 
             ("bleedInside", "出血内"), ("bleedOutside", "出血外")],
            sub_groups=[("margin", "documentSettings.margin", 
                        [("top", "上"), ("bottom", "下"), ("inside", "内"), ("outside", "外")])])
        
        # layers
        self.create_layers_group()
        
        # styleImport
        self.add_group(self.scrollable_frame, "样式导入", "styleImport", 
            [("enabled", "启用"), ("styletemplatePath", "样式模板路径"), 
             ("importParagraphStyles", "段落样式"), ("importCharacterStyles", "字符样式"),
             ("importObjectStyles", "对象样式"), ("importCompositeFonts", "复合字体"),
             ("description", "说明")])
        
        # imageImport
        self.add_group(self.scrollable_frame, "图片导入", "imageImport",
            [("enabled", "启用"), ("artFolderPath", "图片文件夹路径"),
             ("fileExtensions", "文件扩展名"), ("placeByFileName", "按文件名放置"),
             ("scaleFactor", "缩放比例"), ("anchorPoint", "锚点"),
             ("createMissingFrames", "创建缺失框"), ("fitOptions", "适应选项"),
             ("description", "说明")])
        
        # textImport
        self.create_text_import()
        
        # fontMapping
        self.add_group(self.scrollable_frame, "字体映射", "fontMapping",
            [("enabled", "启用"), ("mapconfig", "配置路径"), 
             ("baseFontSize", "基准字号"), ("description", "说明")])
        
        # segmentation
        self.add_group(self.scrollable_frame, "断句设置", "segmentation",
            [("enabled", "启用"), ("range", "范围"), ("pythonScriptPath", "脚本路径"),
             ("waitForUser", "等待用户"), ("description", "说明")])
        
        # ghostFrameAlignment
        self.add_group(self.scrollable_frame, "辅助方框对齐", "ghostFrameAlignment",
            [("enabled", "启用"), ("jsonPath", "JSON路径"), ("autoFindJson", "自动查找JSON"),
             ("jsonFileNamePattern", "JSON文件名模式"), ("createGhostFrames", "创建辅助框"),
             ("alignTextFrames", "对齐文本框"), ("alignRange", "对齐范围"),
             ("hideGhostLayerAfterAlign", "对齐后隐藏"), ("ghostLayerName", "辅助图层名"),
             ("ghostColorName", "辅助颜色名"), ("ghostStrokeWeight", "描边粗细"),
             ("description", "说明")])
        
        # output
        self.add_group(self.scrollable_frame, "输出设置", "output",
            [("saveDocument", "保存文档"), ("savePath", "保存路径"),
             ("saveAsIdml", "存为IDML"), ("closeAfterSave", "保存后关闭"),
             ("logPath", "日志路径"), ("logEnabled", "启用日志")])
        
        # delays
        self.add_group(self.scrollable_frame, "延迟设置 (ms)", "delays",
            [("afterDocumentCreate", "文档创建后"), ("afterLayerCreate", "图层创建后"),
             ("afterStyleImport", "样式导入后"), ("afterImagePlace", "图片放置后"),
             ("afterTextImport", "文本导入后"), ("afterStyleMatching", "样式匹配后"),
             ("afterSegmentation", "断句后")])
        
        # errorHandling
        self.add_group(self.scrollable_frame, "错误处理", "errorHandling",
            [("stopOnError", "出错停止"), ("logErrors", "记录错误"), ("showAlerts", "显示警告")])
    
    def create_layers_group(self):
        """创建图层设置分组"""
        frame = ttk.LabelFrame(self.scrollable_frame, text="图层设置", padding="3")
        frame.pack(fill=tk.X, padx=2, pady=2)
        
        layers = self.config.get("layers", {})
        for layer_key, layer_config in layers.items():
            layer_frame = ttk.LabelFrame(frame, text=layer_key, padding="2")
            layer_frame.pack(fill=tk.X, pady=1)
            
            col_frame = ttk.Frame(layer_frame)
            col_frame.pack(fill=tk.X)
            col0 = ttk.Frame(col_frame)
            col0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            col1 = ttk.Frame(col_frame)
            col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            
            fields = [("name", "名称"), ("visible", "可见"), ("locked", "锁定"), ("color", "颜色")]
            for i, (key, label) in enumerate(fields):
                self.create_field(col0 if i < 2 else col1, (key, label), f"layers.{layer_key}")
    
    def create_text_import(self):
        """创建文本导入分组"""
        frame = ttk.LabelFrame(self.scrollable_frame, text="文本导入", padding="3")
        frame.pack(fill=tk.X, padx=2, pady=2)
        
        # 基本字段
        col_frame = ttk.Frame(frame)
        col_frame.pack(fill=tk.X)
        col0 = ttk.Frame(col_frame)
        col0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        col1 = ttk.Frame(col_frame)
        col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        fields = [("enabled", "启用"), ("lpTxtPath", "LP文本路径"), ("description", "说明"),
                  ("singleLineMode", "单行模式"), ("multiLineMode", "多行模式"),
                  ("pageOffset", "页偏移"), ("matchByNumber", "按数字匹配"), ("fromStartToEnd", "从头到尾")]
        mid = (len(fields) + 1) // 2
        for i, f in enumerate(fields):
            self.create_field(col0 if i < mid else col1, f, "textImport")
        
        # textFrameSize
        size_frame = ttk.LabelFrame(frame, text="textFrameSize", padding="2")
        size_frame.pack(fill=tk.X, pady=2)
        size_col = ttk.Frame(size_frame)
        size_col.pack(fill=tk.X)
        ttk.Label(size_col, text="宽度", width=18, anchor='w', font=('', 8)).pack(side=tk.LEFT)
        var_w = tk.StringVar(value=str(self.config.get("textImport", {}).get("textFrameSize", {}).get("width", 10)))
        ttk.Entry(size_col, textvariable=var_w, width=10).pack(side=tk.LEFT)
        self.widgets["textImport.textFrameSize.width"] = ('number', var_w)
        
        size_col2 = ttk.Frame(size_frame)
        size_col2.pack(fill=tk.X)
        ttk.Label(size_col2, text="高度", width=18, anchor='w', font=('', 8)).pack(side=tk.LEFT)
        var_h = tk.StringVar(value=str(self.config.get("textImport", {}).get("textFrameSize", {}).get("height", 25)))
        ttk.Entry(size_col2, textvariable=var_h, width=10).pack(side=tk.LEFT)
        self.widgets["textImport.textFrameSize.height"] = ('number', var_h)
        
        # replacements
        repl_frame = ttk.LabelFrame(frame, text="replacements (替换)", padding="2")
        repl_frame.pack(fill=tk.X, pady=2)
        
        replacements = self.config.get("textImport", {}).get("replacements", {})
        self.replacements_frame = repl_frame
        self.replacement_vars = []
        
        for orig, new in replacements.items():
            self.add_replacement_row(repl_frame, orig, new)
        
        ttk.Button(repl_frame, text="+ 添加", command=self.add_new_replacement).pack(anchor='w')
        
        # styleRules
        rules_frame = ttk.LabelFrame(frame, text="styleRules (样式规则)", padding="2")
        rules_frame.pack(fill=tk.X, pady=2)
        
        style_rules = self.config.get("textImport", {}).get("styleRules", [])
        self.style_rules_frame = rules_frame
        self.style_rule_vars = []
        
        for i, rule in enumerate(style_rules):
            self.add_style_rule_row(rules_frame, rule)
        
        ttk.Button(rules_frame, text="+ 添加规则", command=self.add_new_style_rule).pack(anchor='w')
    
    def add_replacement_row(self, parent, orig="", new=""):
        """添加替换行"""
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=1)
        
        var_orig = tk.StringVar(value=orig)
        var_new = tk.StringVar(value=new)
        
        ttk.Entry(row, textvariable=var_orig, width=6, font=('', 8)).pack(side=tk.LEFT, padx=1)
        ttk.Label(row, text="→", font=('', 8)).pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=var_new, width=6, font=('', 8)).pack(side=tk.LEFT, padx=1)
        ttk.Button(row, text="×", width=2, 
                   command=lambda r=row: self.remove_replacement_row(r)).pack(side=tk.LEFT)
        
        self.replacement_vars.append((row, var_orig, var_new))
    
    def add_new_replacement(self):
        self.add_replacement_row(self.replacements_frame, "", "")
    
    def remove_replacement_row(self, row):
        self.replacement_vars = [t for t in self.replacement_vars if t[0] != row]
        row.destroy()
    
    def add_style_rule_row(self, parent, rule):
        """添加样式规则行"""
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=1)
        
        var_enabled = tk.StringVar(value=str(rule.get("enabled", True)))
        var_match = tk.StringVar(value=rule.get("match", ""))
        var_style = tk.StringVar(value=rule.get("style", ""))
        
        ttk.Combobox(row, textvariable=var_enabled, values=["true", "false"], width=5, state="readonly").pack(side=tk.LEFT)
        ttk.Label(row, text="匹配:", font=('', 8)).pack(side=tk.LEFT, padx=1)
        ttk.Entry(row, textvariable=var_match, width=8, font=('', 8)).pack(side=tk.LEFT)
        ttk.Label(row, text="样式:", font=('', 8)).pack(side=tk.LEFT, padx=1)
        ttk.Entry(row, textvariable=var_style, width=16, font=('', 8)).pack(side=tk.LEFT)
        ttk.Button(row, text="×", width=2, 
                   command=lambda r=row: self.remove_style_rule_row(r)).pack(side=tk.LEFT, padx=1)
        
        self.style_rule_vars.append((row, var_enabled, var_match, var_style))
    
    def add_new_style_rule(self):
        self.add_style_rule_row(self.style_rules_frame, {"match": "", "style": "", "enabled": True})
    
    def remove_style_rule_row(self, row):
        self.style_rule_vars = [t for t in self.style_rule_vars if t[0] != row]
        row.destroy()
    
    def browse_file(self, var):
        path = filedialog.askopenfilename()
        if path:
            var.set(path)
    
    def browse_folder(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)
    
    def collect_values(self):
        """收集所有控件的值并更新config"""
        for widget_key, (widget_type, var) in self.widgets.items():
            keys = widget_key.split('.')
            value = var.get()
            
            if widget_type == 'bool':
                value = value == "true"
            elif widget_type == 'number':
                try:
                    value = float(value) if '.' in value else int(value)
                except ValueError:
                    value = 0
            elif widget_type == 'list':
                try:
                    value = json.loads(value)
                except:
                    value = []
            
            self.set_nested_value(keys, value)
        
        # 收集replacements
        replacements = {}
        for row, var_orig, var_new in self.replacement_vars:
            orig = var_orig.get()
            new = var_new.get()
            if orig:
                replacements[orig] = new
        self.config["textImport"]["replacements"] = replacements
        
        # 收集styleRules
        style_rules = []
        for row, var_enabled, var_match, var_style in self.style_rule_vars:
            style_rules.append({
                "match": var_match.get(),
                "style": var_style.get(),
                "enabled": var_enabled.get() == "true"
            })
        self.config["textImport"]["styleRules"] = style_rules
    
    def set_nested_value(self, keys, value):
        """设置嵌套字典的值"""
        d = self.config
        for key in keys[:-1]:
            if key not in d:
                d[key] = {}
            d = d[key]
        d[keys[-1]] = value


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    jsonname = "manga_layout_config - template.json"
    config_path = os.path.join(script_dir, jsonname)
    if not os.path.exists(config_path):
        jsonname = "manga_layout_config.json"
        config_path = os.path.join(script_dir, jsonname)
        if not os.path.exists(config_path):
            messagebox.showerror("错误", f"未找到配置文件: {jsonname} 或 manga_layout_config - template.json")
        return
    
    root = tk.Tk()
    root.geometry("650x750")
    root.minsize(500, 400)
    
    # 设置紧凑样式
    style = ttk.Style()
    style.configure('TLabel', padding=1)
    style.configure('TEntry', padding=1)
    style.configure('TLabelframe', padding=2)
    style.configure('TLabelframe.Label', font=('', 8, 'bold'))
    
    app = ConfigEditor(root, config_path)
    root.mainloop()


if __name__ == "__main__":
    main()