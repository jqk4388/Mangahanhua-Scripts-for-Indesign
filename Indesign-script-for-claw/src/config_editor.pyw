# -*- coding: utf-8 -*-
"""
Manga Layout Config Editor & CLI
用于编辑 manga_layout_config.json 的图形界面，同时提供命令行接口供AI和脚本调用。

用法:
  GUI模式:
    python config_editor.pyw                          # 使用默认配置文件启动GUI
    python config_editor.pyw --config path/to/config.json  # 指定配置文件启动GUI

  CLI模式:
    python config_editor.pyw get <key>                # 获取配置值
    python config_editor.pyw set <key> <value>        # 设置配置值
    python config_editor.pyw list [--flat]            # 列出所有配置项
    python config_editor.pyw save                     # 保存配置到文件
    python config_editor.pyw schema                   # 输出配置schema说明
    python config_editor.pyw run                      # 执行manga_layout脚本
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import os
import subprocess
import copy
import argparse
import sys


CONFIG_SCHEMA = {
    "version": {"type": "string", "desc": "配置版本号", "example": "1.0.0"},
    "description": {"type": "string", "desc": "配置描述", "example": "Manga Layout Automation Configuration"},
    "templateDocument": {
        "desc": "InDesign模板文档设置",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用模板文档功能", "example": "true"},
            "indtPath": {"type": "string", "desc": "InDesign模板文件(.indt)的绝对路径", "example": "M:/path/to/template.indt"},
            "description": {"type": "string", "desc": "功能说明", "example": "填入Indesign模板文件indt路径"}
        }
    },
    "documentSettings": {
        "desc": "文档页面尺寸与装订设置",
        "children": {
            "pageWidth": {"type": "number", "desc": "页面宽度(单位由unit字段决定)", "example": "128"},
            "pageHeight": {"type": "number", "desc": "页面高度(单位由unit字段决定)", "example": "182"},
            "unit": {"type": "string", "desc": "尺寸单位: mm / cm / in / pt / px / c / ag", "example": "mm"},
            "pageBinding": {"type": "string", "desc": "装订方向: rightToLeft(右绑/日式漫画) / leftToRight(左绑/西式)", "example": "rightToLeft"},
            "startPage": {"type": "number", "desc": "起始页码", "example": "1"},
            "totalPages": {"type": "number", "desc": "总页数(0=自动)", "example": "0"},
            "bleedTop": {"type": "number", "desc": "顶部出血量", "example": "3"},
            "bleedBottom": {"type": "number", "desc": "底部出血量", "example": "3"},
            "bleedInside": {"type": "number", "desc": "内侧出血量", "example": "5"},
            "bleedOutside": {"type": "number", "desc": "外侧出血量", "example": "3"},
            "margin": {
                "desc": "页边距设置",
                "children": {
                    "top": {"type": "number", "desc": "上边距", "example": "10"},
                    "bottom": {"type": "number", "desc": "下边距", "example": "10"},
                    "inside": {"type": "number", "desc": "内侧边距", "example": "10"},
                    "outside": {"type": "number", "desc": "外侧边距", "example": "10"}
                }
            }
        }
    },
    "layers": {
        "desc": "图层设置(键名为图层标识,每个图层包含name/visible/locked/color)",
        "children": {
            "<layerKey>": {
                "desc": "图层定义,键名如art/design/clean/pageNumber/text/sfx",
                "children": {
                    "name": {"type": "string", "desc": "图层显示名称", "example": "Art"},
                    "visible": {"type": "bool", "desc": "图层是否可见", "example": "true"},
                    "locked": {"type": "bool", "desc": "图层是否锁定", "example": "false"},
                    "color": {"type": "string", "desc": "图层颜色标识", "example": "none"}
                }
            }
        }
    },
    "styleImport": {
        "desc": "样式导入设置(从InDesign模板导入段落/字符/对象样式和复合字体)",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用样式导入", "example": "true"},
            "styletemplatePath": {"type": "string", "desc": "样式模板文件(.indd)的绝对路径", "example": "M:/path/to/style_template.indd"},
            "importParagraphStyles": {"type": "bool", "desc": "是否导入段落样式", "example": "true"},
            "importCharacterStyles": {"type": "bool", "desc": "是否导入字符样式", "example": "true"},
            "importObjectStyles": {"type": "bool", "desc": "是否导入对象样式", "example": "true"},
            "importCompositeFonts": {"type": "bool", "desc": "是否导入复合字体", "example": "true"},
            "description": {"type": "string", "desc": "功能说明", "example": "填入样式模板文件indd路径"}
        }
    },
    "imageImport": {
        "desc": "图片导入设置(自动将图片放置到对应页面)",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用图片导入", "example": "true"},
            "artFolderPath": {"type": "string", "desc": "图片文件夹的绝对路径", "example": "M:/path/to/art_folder/"},
            "fileExtensions": {"type": "list", "desc": "支持的图片文件扩展名列表", "example": "[\".tif\",\".tiff\",\".psd\",\".jpg\",\".jpeg\",\".png\",\".pdf\"]"},
            "placeByFileName": {"type": "bool", "desc": "是否按文件名匹配放置(否则按页码顺序)", "example": "false"},
            "scaleFactor": {"type": "number", "desc": "图片缩放百分比", "example": "100"},
            "anchorPoint": {"type": "string", "desc": "锚点位置: center/topLeft/topRight/bottomLeft/bottomRight", "example": "center"},
            "createMissingFrames": {"type": "bool", "desc": "是否为缺失的图片框自动创建框架", "example": "false"},
            "fitOptions": {"type": "string", "desc": "适应选项: fillProportionally/fitProportionally/fitFrameToContent/fitContentToFrame/centerContent", "example": "fillProportionally"},
            "description": {"type": "string", "desc": "功能说明", "example": "填入图片文件路径"}
        }
    },
    "textImport": {
        "desc": "文本导入设置(从LabelPlus文本文件导入翻译文本)",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用文本导入", "example": "true"},
            "lpTxtPath": {"type": "string", "desc": "LabelPlus文本文件(.txt)的绝对路径", "example": "M:/path/to/translation_LPoutput.txt"},
            "description": {"type": "string", "desc": "功能说明", "example": "填入LabelPlus文本文件txt路径"},
            "singleLineMode": {"type": "bool", "desc": "单行模式:每条文本单独放置", "example": "true"},
            "multiLineMode": {"type": "bool", "desc": "多行模式:连续文本合并放置", "example": "false"},
            "pageOffset": {"type": "number", "desc": "页码偏移量(正负整数)", "example": "0"},
            "matchByNumber": {"type": "bool", "desc": "是否按页码编号匹配", "example": "false"},
            "fromStartToEnd": {"type": "bool", "desc": "是否从头到尾顺序放置", "example": "true"},
            "textFrameSize": {
                "desc": "文本框默认尺寸",
                "children": {
                    "width": {"type": "number", "desc": "文本框默认宽度(mm)", "example": "10"},
                    "height": {"type": "number", "desc": "文本框默认高度(mm)", "example": "25"}
                }
            },
            "replacements": {"type": "dict", "desc": "文本替换映射(键=原文,值=替换文),如 {\"！\":\"!\",\"？\":\"?\"}", "example": "{\"！\":\"!\",\"？\":\"?\"}"},
            "styleRules": {"type": "list", "desc": "样式匹配规则列表,每条包含match(匹配标记)/style(样式名)/enabled(是否启用)", "example": "[{\"match\":\"默认匹配\",\"style\":\"文本框居中-垂直-自动缩框\",\"enabled\":true}]"}
        }
    },
    "fontMapping": {
        "desc": "字体映射设置(将文本字体映射到InDesign字体)",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用字体映射", "example": "true"},
            "mapconfig": {"type": "string", "desc": "字体映射配置文件(.ini)的路径,留空则使用项目默认", "example": ""},
            "baseFontSize": {"type": "number", "desc": "基准字号(pt),用于字号缩放计算", "example": "9.2"},
            "description": {"type": "string", "desc": "功能说明", "example": "填入字体匹配配置ini的路径"}
        }
    },
    "segmentation": {
        "desc": "断句设置(使用jieba对中文文本进行断句处理)",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用断句功能", "example": "true"},
            "range": {"type": "string", "desc": "断句范围: entireDocument(全文档) / selection(选区)", "example": "entireDocument"},
            "pythonScriptPath": {"type": "string", "desc": "jieba断句脚本路径,留空则使用项目默认", "example": ""},
            "waitForUser": {"type": "bool", "desc": "断句完成后是否等待用户确认", "example": "true"},
            "description": {"type": "string", "desc": "功能说明", "example": "填入jieba断句脚本的位置"}
        }
    },
    "ghostFrameAlignment": {
        "desc": "辅助方框对齐设置(使用JSON定位数据自动对齐文本框)",
        "children": {
            "enabled": {"type": "bool", "desc": "是否启用辅助方框对齐", "example": "true"},
            "jsonPath": {"type": "string", "desc": "定位JSON文件路径,留空则自动查找", "example": ""},
            "autoFindJson": {"type": "bool", "desc": "是否自动在图片文件夹中查找JSON", "example": "true"},
            "jsonFileNamePattern": {"type": "string", "desc": "JSON文件名模式,[folderName]会被替换为文件夹名", "example": "imgtrans_[folderName].json"},
            "createGhostFrames": {"type": "bool", "desc": "是否创建辅助定位框", "example": "true"},
            "alignTextFrames": {"type": "bool", "desc": "是否对齐文本框到辅助框", "example": "true"},
            "alignRange": {"type": "string", "desc": "对齐范围: entireDocument / selection", "example": "entireDocument"},
            "hideGhostLayerAfterAlign": {"type": "bool", "desc": "对齐完成后是否隐藏辅助图层", "example": "true"},
            "ghostLayerName": {"type": "string", "desc": "辅助图层名称", "example": "[輔助定位] 完整矩形框"},
            "ghostColorName": {"type": "string", "desc": "辅助框颜色名称", "example": "GhostGreen"},
            "ghostStrokeWeight": {"type": "number", "desc": "辅助框描边粗细(pt)", "example": "1.5"},
            "description": {"type": "string", "desc": "功能说明", "example": "辅助方框对齐功能"}
        }
    },
    "output": {
        "desc": "输出设置(文档保存与日志)",
        "children": {
            "saveDocument": {"type": "bool", "desc": "是否保存文档", "example": "true"},
            "savePath": {"type": "string", "desc": "保存路径,留空则保存到原位置", "example": ""},
            "saveAsIdml": {"type": "bool", "desc": "是否同时保存为IDML格式", "example": "false"},
            "closeAfterSave": {"type": "bool", "desc": "保存后是否关闭文档", "example": "false"},
            "logPath": {"type": "string", "desc": "日志文件路径,留空则使用默认", "example": ""},
            "logEnabled": {"type": "bool", "desc": "是否启用日志记录", "example": "true"}
        }
    },
    "delays": {
        "desc": "延迟设置(毫秒),各步骤之间的等待时间,确保InDesign操作完成",
        "children": {
            "afterDocumentCreate": {"type": "number", "desc": "文档创建后延迟(ms)", "example": "500"},
            "afterLayerCreate": {"type": "number", "desc": "图层创建后延迟(ms)", "example": "300"},
            "afterStyleImport": {"type": "number", "desc": "样式导入后延迟(ms)", "example": "500"},
            "afterImagePlace": {"type": "number", "desc": "图片放置后延迟(ms)", "example": "200"},
            "afterTextImport": {"type": "number", "desc": "文本导入后延迟(ms)", "example": "500"},
            "afterStyleMatching": {"type": "number", "desc": "样式匹配后延迟(ms)", "example": "500"},
            "afterSegmentation": {"type": "number", "desc": "断句后延迟(ms)", "example": "1000"}
        }
    },
    "errorHandling": {
        "desc": "错误处理设置",
        "children": {
            "stopOnError": {"type": "bool", "desc": "遇到错误时是否停止执行", "example": "true"},
            "logErrors": {"type": "bool", "desc": "是否记录错误到日志", "example": "true"},
            "showAlerts": {"type": "bool", "desc": "是否显示错误警告弹窗", "example": "false"}
        }
    }
}


def get_nested_value(config_dict, key_path):
    d = config_dict
    for k in key_path.split('.'):
        if isinstance(d, dict) and k in d:
            d = d[k]
        else:
            return None
    return d


def set_nested_value(config_dict, key_path, value):
    d = config_dict
    keys = key_path.split('.')
    for k in keys[:-1]:
        if k not in d or not isinstance(d[k], dict):
            d[k] = {}
        d = d[k]
    d[keys[-1]] = value


def parse_value(value_str, key_path=None, config=None):
    if value_str.lower() == 'true':
        return True
    if value_str.lower() == 'false':
        return False
    try:
        if '.' in value_str:
            return float(value_str)
        return int(value_str)
    except ValueError:
        pass
    try:
        parsed = json.loads(value_str)
        return parsed
    except (json.JSONDecodeError, ValueError):
        pass
    return value_str


def load_config_file(config_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_config_file(config_path, config):
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


def find_default_config():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for name in ["manga_layout_config - template.json", "manga_layout_config.json"]:
        path = os.path.join(script_dir, name)
        if os.path.exists(path):
            return path
    return None


def cli_get(args):
    config_path = args.config or find_default_config()
    if not config_path or not os.path.exists(config_path):
        print(f"错误: 配置文件未找到", file=sys.stderr)
        sys.exit(1)
    config = load_config_file(config_path)
    value = get_nested_value(config, args.key)
    if value is None:
        print(f"错误: 键 '{args.key}' 不存在", file=sys.stderr)
        sys.exit(1)
    if isinstance(value, (dict, list)):
        print(json.dumps(value, ensure_ascii=False, indent=2))
    else:
        print(value)


def cli_set(args):
    config_path = args.config or find_default_config()
    if not config_path or not os.path.exists(config_path):
        print(f"错误: 配置文件未找到", file=sys.stderr)
        sys.exit(1)
    config = load_config_file(config_path)
    value = parse_value(args.value, args.key, config)
    set_nested_value(config, args.key, value)
    save_config_file(config_path, config)
    print(f"已设置 {args.key} = {json.dumps(value, ensure_ascii=False)}")


def cli_list(args):
    config_path = args.config or find_default_config()
    if not config_path or not os.path.exists(config_path):
        print(f"错误: 配置文件未找到", file=sys.stderr)
        sys.exit(1)
    config = load_config_file(config_path)
    if args.flat:
        flat_list = []
        def flatten(d, prefix=""):
            for k, v in d.items():
                full_key = f"{prefix}.{k}" if prefix else k
                if isinstance(v, dict):
                    flatten(v, full_key)
                else:
                    flat_list.append((full_key, v))
        flatten(config)
        max_key_len = max(len(k) for k, _ in flat_list) if flat_list else 0
        for key, val in flat_list:
            if isinstance(val, (dict, list)):
                val_str = json.dumps(val, ensure_ascii=False)
            else:
                val_str = str(val)
            print(f"  {key:<{max_key_len}}  =  {val_str}")
    else:
        print(json.dumps(config, ensure_ascii=False, indent=2))


def cli_save(args):
    config_path = args.config or find_default_config()
    if not config_path or not os.path.exists(config_path):
        print(f"错误: 配置文件未找到", file=sys.stderr)
        sys.exit(1)
    config = load_config_file(config_path)
    save_path = args.output or config_path
    save_config_file(save_path, config)
    print(f"配置已保存到: {save_path}")


def cli_schema(args):
    def print_schema(schema, prefix="", indent=0):
        pad = "  " * indent
        for key, info in schema.items():
            full_key = f"{prefix}.{key}" if prefix else key
            if "children" in info:
                print(f"{pad}{full_key}  —  {info.get('desc', '')}")
                print_schema(info["children"], full_key, indent + 1)
            else:
                type_str = info.get("type", "?")
                desc = info.get("desc", "")
                example = info.get("example", "")
                example_str = f" (示例: {example})" if example else ""
                print(f"{pad}{full_key}  [{type_str}]  {desc}{example_str}")
    print_schema(CONFIG_SCHEMA)


def cli_run(args):
    config_path = args.config or find_default_config()
    if not config_path:
        print("错误: 配置文件未找到", file=sys.stderr)
        sys.exit(1)
    vbs_path = os.path.join(os.path.dirname(config_path), "run_manga_layout.vbs")
    if not os.path.exists(vbs_path):
        print(f"错误: 脚本文件未找到: {vbs_path}", file=sys.stderr)
        sys.exit(1)
    subprocess.Popen(["cscript", vbs_path],
                     cwd=os.path.dirname(vbs_path),
                     creationflags=subprocess.CREATE_NEW_CONSOLE)
    print(f"已启动脚本: {vbs_path}")


def build_parser():
    parser = argparse.ArgumentParser(
        prog="config_editor",
        description="Manga Layout Config Editor — 图形界面与命令行工具\n\n"
                    "无子命令时启动GUI编辑器。使用子命令进行CLI操作。\n"
                    "所有key使用点号分隔的路径格式，如: documentSettings.pageWidth",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  %(prog)s                                          # 启动GUI\n"
            "  %(prog)s --config my_config.json                  # 指定配置文件启动GUI\n"
            "  %(prog)s get documentSettings.pageWidth           # 获取页面宽度\n"
            "  %(prog)s get textImport.replacements              # 获取替换映射(JSON)\n"
            "  %(prog)s set templateDocument.indtPath \"C:/tpl.indt\"  # 设置模板路径\n"
            "  %(prog)s set documentSettings.pageWidth 128       # 设置页面宽度\n"
            "  %(prog)s set imageImport.enabled true             # 启用图片导入\n"
            "  %(prog)s set imageImport.fileExtensions '[\".tif\",\".psd\"]'  # 设置列表值\n"
            "  %(prog)s list                                      # 以JSON格式列出全部配置\n"
            "  %(prog)s list --flat                               # 以扁平键值对列出全部配置\n"
            "  %(prog)s save                                      # 保存配置\n"
            "  %(prog)s save --output new_config.json            # 另存为\n"
            "  %(prog)s schema                                    # 查看配置schema说明\n"
            "  %(prog)s run                                       # 执行manga_layout脚本\n"
        )
    )
    parser.add_argument(
        "--config", "-c",
        metavar="PATH",
        help="指定配置文件路径(默认自动查找 manga_layout_config - template.json 或 manga_layout_config.json)"
    )

    subparsers = parser.add_subparsers(dest="command", help="子命令 (无子命令则启动GUI)")

    get_parser = subparsers.add_parser(
        "get",
        help="获取指定配置键的值",
        description="获取配置中指定键的值。键使用点号分隔路径，如 documentSettings.pageWidth。\n"
                    "返回值: 字符串/数字直接输出，字典/列表以JSON格式输出。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  %(prog)s documentSettings.pageWidth     # 输出: 128\n"
            "  %(prog)s templateDocument.enabled        # 输出: true\n"
            "  %(prog)s textImport.replacements         # 输出: {\"！\": \"!\", ...}\n"
            "  %(prog)s layers                          # 输出: 图层配置(JSON)\n"
        )
    )
    get_parser.add_argument("key", help="配置键路径(点号分隔)，如 documentSettings.pageWidth")

    set_parser = subparsers.add_parser(
        "set",
        help="设置指定配置键的值并保存",
        description="设置配置中指定键的值并立即保存到文件。\n"
                    "值类型自动推断: true/false→布尔, 数字→int/float, JSON字符串→对象/数组, 其他→字符串。\n"
                    "设置后自动保存配置文件。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例:\n"
            "  %(prog)s documentSettings.pageWidth 128           # 设置数字\n"
            "  %(prog)s templateDocument.enabled true             # 设置布尔值\n"
            "  %(prog)s templateDocument.indtPath \"C:/tpl.indt\"   # 设置字符串\n"
            "  %(prog)s imageImport.fileExtensions '[\".tif\"]'    # 设置列表(JSON)\n"
            "  %(prog)s textImport.replacements '{\"！\":\"!\"}'     # 设置字典(JSON)\n"
        )
    )
    set_parser.add_argument("key", help="配置键路径(点号分隔)，如 documentSettings.pageWidth")
    set_parser.add_argument("value", help="要设置的值(自动推断类型: 布尔/数字/JSON/字符串)")

    list_parser = subparsers.add_parser(
        "list",
        help="列出所有配置项",
        description="列出配置文件中的所有配置项。默认以JSON格式输出，使用--flat以扁平键值对输出。"
    )
    list_parser.add_argument(
        "--flat", "-f",
        action="store_true",
        help="以扁平键值对格式输出(每行一个 key = value)"
    )

    save_parser = subparsers.add_parser(
        "save",
        help="保存配置文件",
        description="保存配置到文件。默认保存到原路径，使用--output指定新路径。"
    )
    save_parser.add_argument(
        "--output", "-o",
        metavar="PATH",
        help="另存为指定路径"
    )

    subparsers.add_parser(
        "schema",
        help="输出配置schema说明(所有键的类型和描述)",
        description="输出完整的配置schema，包含所有键名、类型、描述和示例值。\n"
                    "适合AI或脚本了解配置结构后进行自动化操作。"
    )

    subparsers.add_parser(
        "run",
        help="执行 manga_layout.vbs 脚本",
        description="执行配置文件同目录下的 run_manga_layout.vbs 脚本。"
    )

    return parser


class ConfigEditor:
    def __init__(self, root, config_path):
        self.root = root
        self.root.title("Manga Layout Config Editor")
        self.config_path = config_path
        self.save_path = config_path
        self.config = {}
        self.template_config = {}
        self.widgets = {}
        self.replacement_vars = []
        self.style_rule_vars = []
        self.current_row = 0
        self.current_col = 0

        if config_path.endswith("manga_layout_config - template.json"):
            script_dir = os.path.dirname(config_path)
            self.save_path = os.path.join(script_dir, "manga_layout_config.json")

        self.load_template_config()

        self.main_frame = ttk.Frame(root, padding="3")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self.button_frame = ttk.Frame(root, padding="3")
        self.button_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(self.button_frame, text="打开配置", command=self.open_config, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.button_frame, text="执行脚本", command=self.run_manga_layout, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.button_frame, text="保存", command=self.save_config, width=8).pack(side=tk.RIGHT, padx=2)
        ttk.Button(self.button_frame, text="另存为...", command=self.save_as, width=8).pack(side=tk.RIGHT, padx=2)
        ttk.Button(self.button_frame, text="重新加载", command=self.reload_config, width=8).pack(side=tk.RIGHT, padx=2)

        self.load_config()
        self.create_widgets()

    def _create_tab(self, title):
        outer = ttk.Frame(self.notebook)
        self.notebook.add(outer, text=title)
        canvas = tk.Canvas(outer, highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas)
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill=tk.BOTH, expand=True)
        scrollbar.pack(side="right", fill=tk.Y)

        def _on_enter(event):
            canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        def _on_leave(event):
            canvas.unbind_all("<MouseWheel>")
        canvas.bind("<Enter>", _on_enter)
        canvas.bind("<Leave>", _on_leave)

        return scrollable

    def get_template_path(self):
        script_dir = os.path.dirname(os.path.abspath(__file__))
        template_path = os.path.join(script_dir, "manga_layout_config - template.json")
        if os.path.exists(template_path):
            return template_path
        return None

    def load_template_config(self):
        template_path = self.get_template_path()
        if template_path:
            try:
                with open(template_path, 'r', encoding='utf-8') as f:
                    self.template_config = json.load(f)
            except Exception as e:
                print(f"加载模板失败: {e}")
                self.template_config = {}
        else:
            self.template_config = {}

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

    def deep_merge(self, base, override):
        result = copy.deepcopy(base)
        for key, value in override.items():
            if key in result:
                if isinstance(result[key], dict) and isinstance(value, dict):
                    result[key] = self.deep_merge(result[key], value)
                elif isinstance(result[key], list) and isinstance(value, list):
                    result[key] = copy.deepcopy(value)
                else:
                    result[key] = copy.deepcopy(value)
            else:
                result[key] = copy.deepcopy(value)
        return result

    def collect_values(self):
        user_config = copy.deepcopy(self.template_config) if self.template_config else {}
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
            self.set_nested_value_in_dict(user_config, keys, value)
        replacements = {}
        for row, var_orig, var_new in self.replacement_vars:
            orig = var_orig.get()
            new = var_new.get()
            if orig:
                replacements[orig] = new
        if "textImport" not in user_config:
            user_config["textImport"] = {}
        user_config["textImport"]["replacements"] = replacements
        style_rules = []
        for row, var_enabled, var_match, var_style in self.style_rule_vars:
            style_rules.append({
                "match": var_match.get(),
                "style": var_style.get(),
                "enabled": var_enabled.get() == "true"
            })
        user_config["textImport"]["styleRules"] = style_rules
        return user_config

    def set_nested_value_in_dict(self, config_dict, keys, value):
        d = config_dict
        for key in keys[:-1]:
            if key not in d:
                d[key] = {}
            d = d[key]
        d[keys[-1]] = value

    def save_config(self, path=None):
        self.load_template_config()
        user_config = self.collect_values()
        if self.template_config:
            final_config = self.deep_merge(self.template_config, user_config)
        else:
            final_config = user_config
        save_path = path or self.save_path
        try:
            with open(save_path, 'w', encoding='utf-8') as f:
                json.dump(final_config, f, ensure_ascii=False, indent=4)
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
        self.load_template_config()
        self.load_config()
        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)
        self.widgets = {}
        self.current_row = 0
        self.current_col = 0
        self.create_widgets()

    def open_config(self):
        path = filedialog.askopenfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            title="选择配置文件"
        )
        if path:
            self.config_path = path
            if path.endswith("manga_layout_config - template.json"):
                script_dir = os.path.dirname(path)
                self.save_path = os.path.join(script_dir, "manga_layout_config.json")
            else:
                self.save_path = path
            self.reload_config()

    def run_manga_layout(self):
        vbs_path = os.path.join(os.path.dirname(self.config_path), "run_manga_layout.vbs")
        if not os.path.exists(vbs_path):
            messagebox.showerror("错误", f"脚本文件未找到: {vbs_path}")
            return
        try:
            subprocess.Popen(["cscript", vbs_path],
                             cwd=os.path.dirname(vbs_path),
                             creationflags=subprocess.CREATE_NEW_CONSOLE)
        except Exception as e:
            messagebox.showerror("错误", f"启动脚本失败: {e}")

    def add_group(self, parent, title, key, fields, sub_groups=None):
        frame = ttk.LabelFrame(parent, text=title, padding="3")
        frame.pack(fill=tk.X, padx=2, pady=2)
        col_frame = ttk.Frame(frame)
        col_frame.pack(fill=tk.X)
        col0 = ttk.Frame(col_frame)
        col0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        col1 = ttk.Frame(col_frame)
        col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mid = (len(fields) + 1) // 2
        for i, field in enumerate(fields):
            parent_col = col0 if i < mid else col1
            self.create_field(parent_col, field, key)
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
        if isinstance(field_info, tuple):
            key, label = field_info
        else:
            key = field_info
            label = field_info
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=1)
        value = self.config
        if parent_key:
            for k in parent_key.split('.'):
                value = value.get(k, {})
        value = value.get(key, "") if isinstance(value, dict) else ""
        ttk.Label(row, text=label, width=18, anchor='w', font=('', 8)).pack(side=tk.LEFT)
        widget_key = f"{parent_key}.{key}" if parent_key else key
        if isinstance(value, bool):
            var = tk.StringVar(value=str(value).lower())
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
        tab_basic = self._create_tab("基本")
        self.add_group(tab_basic, "基本信息", "",
                       [("version", "版本"), ("description", "描述")])
        self.add_group(tab_basic, "模板文档", "templateDocument",
                       [("enabled", "启用"), ("indtPath", "模板路径"), ("description", "说明")])
        self.add_group(tab_basic, "文档设置", "documentSettings",
                       [("pageWidth", "页面宽度"), ("pageHeight", "页面高度"), ("unit", "单位"),
                        ("pageBinding", "装订方向"), ("startPage", "起始页"), ("totalPages", "总页数"),
                        ("bleedTop", "出血上"), ("bleedBottom", "出血下"),
                        ("bleedInside", "出血内"), ("bleedOutside", "出血外")],
                       sub_groups=[("margin", "documentSettings.margin",
                                    [("top", "上"), ("bottom", "下"), ("inside", "内"), ("outside", "外")])])
        self.add_group(tab_basic, "错误处理", "errorHandling",
                       [("stopOnError", "出错停止"), ("logErrors", "记录错误"), ("showAlerts", "显示警告")])

        tab_layers = self._create_tab("图层")
        self.create_layers_group(tab_layers)

        tab_style = self._create_tab("样式")
        self.add_group(tab_style, "样式导入", "styleImport",
                       [("enabled", "启用"), ("styletemplatePath", "样式模板路径"),
                        ("importParagraphStyles", "段落样式"), ("importCharacterStyles", "字符样式"),
                        ("importObjectStyles", "对象样式"), ("importCompositeFonts", "复合字体"),
                        ("description", "说明")])
        self.add_group(tab_style, "字体映射", "fontMapping",
                       [("enabled", "启用"), ("mapconfig", "配置路径"),
                        ("baseFontSize", "基准字号"), ("description", "说明")])

        tab_image = self._create_tab("图片")
        self.add_group(tab_image, "图片导入", "imageImport",
                       [("enabled", "启用"), ("artFolderPath", "图片文件夹路径"),
                        ("fileExtensions", "文件扩展名"), ("placeByFileName", "按文件名放置"),
                        ("scaleFactor", "缩放比例"), ("anchorPoint", "锚点"),
                        ("createMissingFrames", "创建缺失框"), ("fitOptions", "适应选项"),
                        ("description", "说明")])

        tab_text = self._create_tab("文本")
        self.create_text_import(tab_text)

        tab_align = self._create_tab("对齐")
        self.add_group(tab_align, "辅助方框对齐", "ghostFrameAlignment",
                       [("enabled", "启用"), ("jsonPath", "JSON路径"), ("autoFindJson", "自动查找JSON"),
                        ("jsonFileNamePattern", "JSON文件名模式"), ("createGhostFrames", "创建辅助框"),
                        ("alignTextFrames", "对齐文本框"), ("alignRange", "对齐范围"),
                        ("hideGhostLayerAfterAlign", "对齐后隐藏"), ("ghostLayerName", "辅助图层名"),
                        ("ghostColorName", "辅助颜色名"), ("ghostStrokeWeight", "描边粗细"),
                        ("description", "说明")])
        self.add_group(tab_align, "断句设置", "segmentation",
                       [("enabled", "启用"), ("range", "范围"), ("pythonScriptPath", "脚本路径"),
                        ("waitForUser", "等待用户"), ("description", "说明")])

        tab_output = self._create_tab("输出")
        self.add_group(tab_output, "输出设置", "output",
                       [("saveDocument", "保存文档"), ("savePath", "保存路径"),
                        ("saveAsIdml", "存为IDML"), ("closeAfterSave", "保存后关闭"),
                        ("logPath", "日志路径"), ("logEnabled", "启用日志")])
        self.add_group(tab_output, "延迟设置 (ms)", "delays",
                       [("afterDocumentCreate", "文档创建后"), ("afterLayerCreate", "图层创建后"),
                        ("afterStyleImport", "样式导入后"), ("afterImagePlace", "图片放置后"),
                        ("afterTextImport", "文本导入后"), ("afterStyleMatching", "样式匹配后"),
                        ("afterSegmentation", "断句后")])

    def create_layers_group(self, parent):
        frame = ttk.LabelFrame(parent, text="图层设置", padding="3")
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

    def create_text_import(self, parent):
        frame = ttk.LabelFrame(parent, text="文本导入", padding="3")
        frame.pack(fill=tk.X, padx=2, pady=2)
        col_frame = ttk.Frame(frame)
        col_frame.pack(fill=tk.X)
        col0 = ttk.Frame(col_frame)
        col0.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        col1 = ttk.Frame(col_frame)
        col1.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        fields = [("enabled", "启用"), ("lpTxtPath", "LP文本路径"), ("description", "说明"),
                  ("singleLineMode", "单行模式"), ("multiLineMode", "多行模式"),
                  ("pageOffset", "页偏移"), ("matchByNumber", "按页码匹配"), ("fromStartToEnd", "从头到尾")]
        mid = (len(fields) + 1) // 2
        for i, f in enumerate(fields):
            self.create_field(col0 if i < mid else col1, f, "textImport")

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

        repl_frame = ttk.LabelFrame(frame, text="replacements (替换)", padding="2")
        repl_frame.pack(fill=tk.X, pady=2)
        replacements = self.config.get("textImport", {}).get("replacements", {})
        self.replacements_frame = repl_frame
        self.replacement_vars = []
        for orig, new in replacements.items():
            self.add_replacement_row(repl_frame, orig, new)
        ttk.Button(repl_frame, text="+ 添加", command=self.add_new_replacement).pack(anchor='w')

        rules_frame = ttk.LabelFrame(frame, text="styleRules (样式规则)", padding="2")
        rules_frame.pack(fill=tk.X, pady=2)
        style_rules = self.config.get("textImport", {}).get("styleRules", [])
        self.style_rules_frame = rules_frame
        self.style_rule_vars = []
        for i, rule in enumerate(style_rules):
            self.add_style_rule_row(rules_frame, rule)
        ttk.Button(rules_frame, text="+ 添加规则", command=self.add_new_style_rule).pack(anchor='w')

    def add_replacement_row(self, parent, orig="", new=""):
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
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=1)
        var_enabled = tk.StringVar(value=str(rule.get("enabled", True)).lower())
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


def run_gui(config_path):
    root = tk.Tk()
    root.geometry("650x750")
    root.minsize(500, 400)
    if not os.path.exists(config_path):
        messagebox.showerror("错误", f"未找到配置文件:\n{config_path}")
        root.destroy()
        return
    style = ttk.Style()
    style.configure('TLabel', padding=1)
    style.configure('TEntry', padding=1)
    style.configure('TLabelframe', padding=2)
    style.configure('TLabelframe.Label', font=('', 8, 'bold'))
    style.configure('TNotebook', padding=2)
    ConfigEditor(root, config_path)
    root.mainloop()


def main():
    parser = build_parser()
    args = parser.parse_args()

    if args.command == "get":
        cli_get(args)
    elif args.command == "set":
        cli_set(args)
    elif args.command == "list":
        cli_list(args)
    elif args.command == "save":
        cli_save(args)
    elif args.command == "schema":
        cli_schema(args)
    elif args.command == "run":
        cli_run(args)
    else:
        config_path = args.config
        if not config_path:
            config_path = find_default_config()
        if not config_path:
            print("错误: 未找到配置文件，请使用 --config 指定路径", file=sys.stderr)
            sys.exit(1)
        run_gui(config_path)


if __name__ == "__main__":
    main()
