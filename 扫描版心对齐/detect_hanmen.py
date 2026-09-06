# -*- coding: utf-8 -*-
"""
detect_hanmen.py — 扫描漫画页版心（hanmen）自动检测

功能：
  逐页检测扫描图中漫画版心的四边位置（最外侧长直线簇），
  用跨页共识（归一化中位数）兜底缺线页，输出 hanmen.json
  和 hanmen_debug/ 调试叠加图（红=检测OK / 黄=共识兜底 / 绿=需手动）。

  【页眉/页脚检测（可选）】加 --hf 启用，在版心检测之上增量做三件事：
    1. 页眉（卷标题 running header）：左窄带行分段取种子 → 短模板
       TM_CCOEFF_NORMED 全卷匹配 → 3 轮迭代重建模板。
    2. 页脚（页码）：双极性掩膜（白底黑字 / 黑底白字），默认走「相对页底」
       范式（2D 众数投票）；若投票峰票比过低或组内离散过大，自动切换到
       「相对分镜横线」范式（横线对齐种子 → 0-9 字形自学习 → 全卷回扫）。
    3. 分镜框线：在页眉带 / 页脚带内做 LINE / EDGE / 排线抑制三通道检测。
  ── 与版心框的关系 ────────────────────────────────────────────
    启用后，把「页眉/页脚带内检出的垂直边线」并入 X 轴贴边证据，
    共识框因此紧贴这些垂直边线。两条硬约束始终成立：
      · 框尺寸全书恒定（= 共识框尺寸），逐页只做整体平移，绝不缩放；
      · 共识框的 x1 / x2 仍是全书共用的一对常数（可选按竖线中位数微调一次）。

用法：
  python detect_hanmen.py <图片目录> [--hf] [--no-debug]
    --hf / --header-footer   启用页眉/页脚/分镜框线检测（默认关闭）
    --no-debug               不输出 hanmen_debug/ 叠加图

  双击 .py（或直接运行不传参）→ 弹出 Tkinter 简易界面：
    模式 A：传统版心检测（无页眉/页脚）
    模式 B：启用页眉/页脚检测（--hf），把共识框紧贴分镜框的垂直边线
    浏览…：选择扫描图片所在文件夹
    生成调试叠加图（默认勾选）→ 输出 hanmen_debug/
    点击确定后开始运行，进度显示在「运行日志」面板

支持格式：.jpg .jpeg .tif .tiff .png（含黑白 1-bit 二值 TIF）
依赖：Pillow、numpy、opencv-python

算法与阈值的实测依据见《扫描版心对齐/页眉页脚检测/页眉页脚检测方案.md》，
尤其 §7「已证伪的方案」——改动参数前请先读该节。
"""

import os
import re
import sys
import json
import itertools
import subprocess


def _ensure_dependencies():
    """检测第三方库，缺失则自动 pip 安装后继续运行。"""
    deps = [('numpy', 'numpy'), ('cv2', 'opencv-python'), ('PIL', 'Pillow')]
    missing = []
    for mod, pkg in deps:
        try:
            __import__(mod)
        except ImportError:
            missing.append(pkg)
    if not missing:
        return
    print('缺少依赖库：' + ', '.join(missing))
    print('正在自动安装（pip install ' + ' '.join(missing) + '）...')
    try:
        subprocess.check_call(
            [sys.executable, '-m', 'pip', 'install', '--disable-pip-version-check'] + missing)
    except Exception as e:
        print('自动安装失败：%s' % e)
        print('请手动执行："%s" -m pip install %s' % (sys.executable, ' '.join(missing)))
        try:
            input('\n按回车键退出...')
        except Exception:
            pass
        sys.exit(1)
    print('依赖安装完成。')


_ensure_dependencies()

import numpy as np
from PIL import Image
Image.MAX_IMAGE_PIXELS = None   # 关闭“解压炸弹”限制，支持 14000×20000 等超级大图
import cv2

# ---------------- 参数 ----------------

WORK_LONG = 1500          # 分析时缩放到的长边像素
SCAN_MIN_LEN = 250         # HoughLinesP / 角度过滤的最小长度（WORK_LONG 缩放图上的像素数）
BIN_THRESH = 170          # 灰度二值化阈值（< 该值视为墨迹）
HK_RATIO = 0.24           # 水平直线形态学核长 / 页宽
VK_RATIO = 0.12           # 垂直直线形态学核长 / 页高
CLUSTER_GAP = 6           # 投影聚类时允许的线宽间隙（小图像素）
STRENGTH_RATIO = 0.40     # 线簇强度阈值 = 最强簇 × 该比例
STRENGTH_MIN = 40         # 线簇强度绝对下限（小图像素）
# 版心物理尺寸恒定（同扫描仪同 DPI）：输出框尺寸一律取共识框、只做整体平移
# （贴边/居中）；「检测框尺寸」仅用于判定当页检测可信度（ok / 兜底）。
SIZE_TOL_W = 0.12         # 检测框宽度允许偏差（相对共识框宽）
SIZE_TOL_H = 0.10         # 检测框高度允许偏差（相对共识框高）
# 平移判定阈值全部为相对共识框对应尺寸的比例，自动适配不同分辨率/尺寸的原图：
ZONE_PCT = 0.10           # 在共识边 ±该比例带内寻找当页实测线
SHIFT_CAP_PCT_X = 0.09    # 水平平移上限（相对共识框宽）
SHIFT_CAP_PCT_Y = 0.05    # 垂直平移上限（相对共识框高）
AGREE_PCT = 0.035         # 多候选互相一致的最大散布（相对共识框对应尺寸）
INNER_SPAN_MIN = 0.55     # 居中模式：实测两线间距至少为共识框的该比例才算有效证据
INK_RATIO_MIN = 0.008     # 墨迹占比低于该值视为空白页
EXTS = ('.jpg', '.jpeg', '.tif', '.tiff', '.png')

# ============================================================
# 以下为「页眉 / 页脚 / 分镜框线检测」（可选功能，--hf 启用）的参数。
# 全部阈值出自《页眉页脚检测方案.md》实测，改动前请先读该文档的 §7「已证伪的方案」。
# 归一化参考系统一为图像宽 IW=1200（不按版心宽归一：版心宽在部分卷会乱跳）。
# ============================================================
HF_IW = 1200            # 归一化参考宽
HF_TOPR = 0.28          # 页眉工作区（顶部比例）
HF_BOTR = 0.26          # 页脚工作区（底部比例）
HF_PNL_TOPR = 0.20      # 分镜框线检测带：顶部
HF_PNL_BOTR = 0.26      # 分镜框线检测带：底部
HF_INK_T = 128          # 页眉墨迹阈值（单极性即可，页眉均为白底黑字）
HF_POS_T = 110          # 页脚/框线 正极性（白底黑字）
HF_NEG_T = 150          # 页脚/框线 负极性（黑底白字，反白）—— 不可省
HF_SNAP = True          # 是否用页眉/页脚附近的竖边线紧贴共识框
HF_REF_WIN = 0.12       # 共识边修正：在共识边 ±该比例（相对共识宽）内找竖线
HF_REF_CAP = 0.50       # 共识边修正量上限（相对图宽 W），防止灾难性漂移
HF_REF_CAP_Y = 0.15     # 共识 y 修正量上限（相对图高 H），防止页眉/页脚修正后框过小
HF_Y_MARGIN = 0.006     # 页眉下方 / 页脚上方留白（相对图高 H）
HF_V_BASE_W = 0.80      # 合并候选时「整页长竖线」相对「检测带竖线」的权重
HF_V_MERGE_PX = 0.004   # 竖线合并容差（相对图宽）

# --- 共识宽校准（成对竖线间距众数）---
# 必须成对投票，x1 / x2 各自独立找簇不保证配对，宽度会飘（方案 §4.7.1）。
HF_PW_BIN = 40.0        # 间距投票格宽（绝对像素）
HF_PW_EDGE = (0.03, 0.96)   # 有效竖线 x/W 区间（下界含窄边距卷，上界排除裁切线）
HF_PW_RANGE = (0.55, 0.90)  # 合理版心宽 / 图宽
HF_PW_MIN_N = 12        # 峰格最少支持页数
HF_PW_TOL = 0.12        # 与原共识宽的最大相对偏差（仅在证据弱时作为安全阀）
# 证据足够强时允许覆盖原共识宽（缺边多的卷基础共识本身不可信，见 §4.7.5）。
HF_PW_STRONG_N = 40     # 强证据阈值：峰格支持页数 ≥ 该值则允许大幅改宽
HF_PW_STRONG_RANGE = (0.62, 0.95)   # 强证据下的合理宽 / 图宽

# --- 共识高校准（同页「页眉基准线 → 页脚基准线」净跨距）---
# 规则与推导见方案 §4.7.4：共识框上下边必须落在页眉底边与页脚顶边之内。
HF_PH_BIN = 40.0        # 跨距投票格宽（绝对像素）
HF_PH_RANGE = (0.55, 0.95)  # 合理版心高 / 图高
HF_PH_MIN_N = 6         # 峰格支持页数达此值才走众数；不足则用全部同页样本中位

# --- 角点定位（共识框左上角 / 下角贴「框线交点」）见方案 §4.7.2 ---
HF_CORNER_TOL = 0.008   # 角点聚类容差（x 相对图宽 W，y 相对图高 H）
HF_CORNER_MIN_N = 4     # 角点稳定簇最少页数
HF_CORNER_RATIO = 0.08  # 稳定簇页数比例下限
HF_CORNER_CAP = 0.35    # 角点定位相对原位置的修正上限（相对图宽 / 图高）
HF_CORNER_SPAN = 0.35   # 页眉下方 / 页脚上方的候选搜索跨度（相对图高）
HF_PC_YW = 0.30         # 逐页选角：y 方向的距离权重
HF_PC_XGATE = 0.25      # 逐页选角：交点 x 与页眉侧锚点的最大偏离（相对图宽）

# --- 基准线「紧邻」约束（见方案 §4.7.6）---
HF_PC_NEAR = 0.045      # 基准线与页眉底边 / 页脚顶边的最大距离（相对图高 H）
HF_PC_BELOW = 0.030     # 允许基准线落在页脚顶边【下方】的容差（相对图高 H）
HF_PC_OFF_MIN_N = 8     # 统计「线y − 页脚顶边」中位偏移所需的最少有线页数
HF_PC_BOTH_TOL = 0.030  # 两端估计差在此内取均值，超出则信页脚（相对图高 H）
HF_PC_FXGATE = 0.030    # 页脚侧交点门控（相对图宽 W，远紧于页眉侧）

# --- 页眉 ---
HF_BX0, HF_BX1 = 0.05, 0.45   # 种子搜索左窄带（避开右侧黑网点块）
HF_HR = (0.012, 0.032)        # 文字带高 / sH
HF_WR = (0.10, 0.45)          # 文字带宽 / IW
HF_DENS = (0.03, 0.45)        # 带内墨迹密度
HF_NSEG_MIN = 5               # 列方向段数（≈字符数）下限
HF_YTOP_MIN = 0.008           # 上方留白下限 / sH
HF_SEG_HMIN = 0.006           # 低于该高度的段视为细线/噪点，跳过
HF_CORE_DXL = 0.04            # 核心种子 x 容差 / IW
HF_CORE_DHR = 0.005           # 核心种子高度容差
HF_TWS = 160                  # 短模板宽（只覆盖页眉左端，主动截掉右侧内容）
HF_TH_K = 1.15                # 模板高 = 中位字高 × 该系数
HF_CC = 0.70                  # TM_CCOEFF_NORMED 命中阈值
HF_XGATE = 0.04               # x 门控 / IW（不约束 y：页眉纵向位置会跟首格浮动）
HF_ROUNDS = 3                 # 模板迭代轮数
HF_MIN_SEED = 3               # 种子数下限，不足则判定本卷无页眉

# --- 页脚 公共 ---
HF_CORNW = 0.34               # 角区宽 / IW
HF_AREA_MIN = 6               # 连通块面积下限
HF_BLK_H = (0.0045, 0.021)    # 连通块字高 / sH
HF_BLK_W = 0.09               # 连通块单字宽上限 / IW
HF_FILL = 0.10                # 连通块填充率下限
HF_ROW_DYB = 0.005            # 聚行：底基线容差 / sH
HF_ROW_DH = 0.007             # 聚行：字高容差 / sH
HF_GAP_K = 0.8                # 子串相邻块间隙 ≤ 该系数 × 中位字高
HF_STR_W = 0.11               # 数字串总宽上限 / IW
HF_CUT_W = (0.40, 1.9)        # 谷点切分：每段宽 / 平均字宽
HF_CUT_T = 0.35               # 谷点阈值 = min + 该系数 × (max − min)
HF_CUT_MAX = 24               # 每个候选保留的切分组合上限（控制内存）
HF_CUT_PAGE_MAX = 48          # 每页保留的切分组合总数上限
HF_CUT_TOPK = 26              # 页脚字形切分前先按谷点深度留 TopK 列，避免组合爆炸

# --- 页脚 范式 A（相对页底）---
HF_A_H = (0.0070, 0.0170)     # 先验字高 / sH
HF_A_BOT = (0.030, 0.170)     # 先验距页底 / sH
HF_A_EDGE = (0.020, 0.200)    # 先验距外侧边 / IW
HF_BIN_B, HF_BIN_E = 0.010, 0.012   # 2D 投票格边长
HF_A_DH, HF_A_DB, HF_A_DE = 0.005, 0.022, 0.030   # 组内收敛容差
HF_A_MINRATIO = 0.45          # 峰票/页数 下限，低于此值判定范式 A 失效
HF_A_MAXSPREAD = 0.05         # 组内 botd 的 p90−p10 上限，超过则失效

# --- 页脚 范式 B（相对分镜横线）---
HF_B_H = (0.0060, 0.0125)     # 字高窗 / sH
HF_B_BOT = (0.010, 0.235)     # 距页底窗 / sH
HF_B_DY = (0.0010, 0.0110)    # 页码顶边 − 线底，归一化 / sH
HF_B_DX = 0.014               # 外缘与线端偏移上限 / IW
HF_GH, HF_GW = 20, 14         # 字形网格
HF_T_CC = 0.60                # 判定阈值（保 FP=0；0.45 换召回）
HF_LONG_K = 0.18              # 横线长核 / IW
HF_SHORT_K = 0.055            # 横线短核 / IW
HF_B_MIN_SEED = 6             # 字形学习所需的最少种子页

# --- 分镜框线 ---
HF_T_THN = (0.0010, 0.0060)   # 线厚 / sH
HF_T_DEN = 0.35               # 线两侧墨迹密度上限
HF_T_COV = 0.70               # 覆盖率下限
HF_T_LNH = 0.05               # 横线长度比下限（w/IW）
HF_T_LNV = 0.25               # 竖线长度比下限（h/bH）
HF_K_H = (0.18, 0.055)        # 横线核长 / IW（长核 + 短核，取并集）
HF_K_V = (0.60, 0.25)         # 竖线核长 / bH
HF_W_MAX = 0.030              # LINE 通道线宽上限 / sH
HF_T_STEP = 0.70              # EDGE 通道阶跃幅度下限
HF_E_GAP = 0.008              # EDGE 采样留空 / sH
HF_E_BAND = 0.010             # EDGE 采样带宽 / sH
HF_E_LNH = 0.10               # EDGE 横线核长 / IW
HF_E_LNV = 0.30               # EDGE 竖线核长 / bH
HF_E_WMAX = 0.020             # EDGE 线宽上限 / sH
HF_T_SHORT = 0.20             # 长度比 ≥ 该值不要求正交支持
HF_R_SUP = 0.016              # 正交支持搜索半径 / sH
HF_EDGE_PX = 0.012            # 横线端点抵达画幅边缘的判定带 / IW
HF_W_STRIPE = 0.040           # 排线抑制：法向半宽 / sH
HF_OV_STRIPE = 0.50           # 排线抑制：投影重叠比下限
HF_N_STRIPE = 4               # 排线抑制：法向邻域同向线数阈值（3 会误杀真框线）
HF_T_LONG = 0.45              # 排线抑制：长线豁免长度比

PAGE_NUM_RE = re.compile(r'(\d{3,4})(?=_?\d?\d?[a-zA-Z]?\.[A-Za-z]{3,4})')


def log(msg):
    print(msg, flush=True)


# ---------------- 读图 ----------------

def load_image(path):
    """用 Pillow 读图（兼容中文路径与 1-bit TIF），返回 (gray uint8, dpi, w, h)。"""
    im = Image.open(path)
    dpi = 1200
    try:
        info_dpi = im.info.get('dpi')
        if info_dpi and info_dpi[0]:
            dpi = int(round(float(info_dpi[0])))
    except Exception:
        pass
    gray = im.convert('L')
    arr = np.asarray(gray, dtype=np.uint8)
    h, w = arr.shape
    return arr, dpi, w, h


def is_binary(arr):
    """判断是否已经是黑白二值图（1-bit TIF 转 L 后只有 0/255）。"""
    uniq = np.unique(arr)
    if len(uniq) <= 2:
        return True
    return np.mean((arr > 10) & (arr < 245)) < 0.002


# ---------------- 直线检测 ----------------

def line_clusters(proj, thresh, gap=CLUSTER_GAP):
    """对一维投影做聚类，返回 [(中心位置, 峰值强度), ...]。"""
    idx = np.where(proj >= thresh)[0]
    if len(idx) == 0:
        return []
    clusters = []
    start = idx[0]
    prev = idx[0]
    for i in idx[1:]:
        if i - prev > gap:
            clusters.append((start, prev))
            start = i
        prev = i
    clusters.append((start, prev))
    out = []
    for a, b in clusters:
        seg = proj[a:b + 1]
        center = int(a + np.argmax(seg))
        out.append((center, float(seg.max())))
    return out


def estimate_skew(horiz_mask):
    """用水平直线像素估计整页倾斜角（度），失败返回 0。"""
    try:
        lines = cv2.HoughLinesP(horiz_mask, 1, np.pi / 180,
                                threshold=200, minLineLength=SCAN_MIN_LEN, maxLineGap=10)
        if lines is None:
            return 0.0
        angles = []
        for ln in lines[:, 0]:
            x1, y1, x2, y2 = ln
            dx, dy = x2 - x1, y2 - y1
            length = (dx * dx + dy * dy) ** 0.5
            if length < SCAN_MIN_LEN:
                continue
            ang = np.degrees(np.arctan2(dy, dx))
            if abs(ang) <= 2.0:
                angles.append(ang)
        if len(angles) < 5:
            return 0.0
        return float(np.median(angles))
    except Exception:
        return 0.0


def detect_edges(gray):
    """返回 (edges dict|None, angle, ink_ratio, debug info)。edges 为小图坐标。"""
    h, w = gray.shape
    binary = gray if is_binary(gray) else None
    if binary is None:
        _, binary = cv2.threshold(gray, BIN_THRESH, 255, cv2.THRESH_BINARY_INV)
    else:
        binary = cv2.compare(binary, 128, cv2.CMP_LT)  # 0(黑)->255

    ink_ratio = float(np.count_nonzero(binary)) / (w * h)

    # 长水平 / 垂直线
    hk = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, int(w * HK_RATIO)), 1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, int(h * VK_RATIO))))
    horiz = cv2.morphologyEx(binary, cv2.MORPH_OPEN, hk)
    vert = cv2.morphologyEx(binary, cv2.MORPH_OPEN, vk)

    angle = estimate_skew(horiz)

    proj_h = horiz.sum(axis=1).astype(np.float32) / 255.0
    proj_v = vert.sum(axis=0).astype(np.float32) / 255.0

    hc = line_clusters(proj_h, max(STRENGTH_MIN, proj_h.max() * STRENGTH_RATIO))
    vc = line_clusters(proj_v, max(STRENGTH_MIN, proj_v.max() * STRENGTH_RATIO))

    def pick(clusters, lo, hi, outermost):
        cand = [c for c in clusters if lo <= c[0] <= hi]
        if not cand:
            return None
        return min(cand, key=lambda c: c[0])[0] if outermost == 'min' \
            else max(cand, key=lambda c: c[0])[0]

    top = pick(hc, h * 0.04, h * 0.45, 'min')
    bot = pick(hc, h * 0.55, h * 0.92, 'max')
    left = pick(vc, w * 0.03, w * 0.45, 'min')
    right = pick(vc, w * 0.55, w * 0.90, 'max')

    edges = None
    if None not in (top, bot, left, right):
        # 几何合理性
        if (right - left) > w * 0.35 and (bot - top) > h * 0.45 \
                and (right - left) < (bot - top) * 0.95:
            edges = {'x1': left, 'y1': top, 'x2': right, 'y2': bot}

    return edges, angle, ink_ratio, (hc, vc)


def _strongest_edge_line(cands, center, zone, agree):
    """在共识边 center ± zone 带内取最强线簇位置。
    带内若有两条强度接近（≥80%）且相互分散（>agree）的强线，说明版心线
    与装饰线打架、有歧义，返回 None 不采用。"""
    near = sorted([(p, s) for p, s in cands if abs(p - center) <= zone],
                  key=lambda c: -c[1])
    if not near:
        return None
    p, s = near[0]
    if len(near) > 1 and near[1][1] >= 0.8 * s and abs(near[1][0] - p) > agree:
        return None
    return p


def _axis_shift(lines, c_lo, c_hi, size, zone, cap, agree, tol,
                det_lo=None, det_hi=None):
    """单轴平移量。框尺寸恒为 size（共识框尺寸），本函数只决定平移多少。

    lines 为当页全部线簇（仅在共识边 ±zone 带内采信）；det_lo/det_hi 为四边
    检测器已确认的边线位置（已经长线形态学验证，不受 zone 限制——误检内框
    的边线通常落在带外，需要靠它们做「居中」）。

    返回 (shift, mode)：
      snap   两边实测线一致（散布 ≤agree）→ 整框平移到中位位置（框居中于两线）
      center 两边实测线明显偏近（间距 < size×(1-tol)，误检内框）→ 不缩框，
             把共识框中点移到两线中点（间距 ≥ INNER_SPAN_MIN×size 才采纳）
      edge   只找到一条实测线 → 平移共识框贴住该边
      none   无可靠证据 → 不平移（保留共识位置）
    任何模式平移量超过 cap（出血页大标题/装饰线误导）都退回 none。
    """
    lo = _strongest_edge_line(lines, c_lo, zone, agree)
    if lo is None and det_lo is not None:
        lo = det_lo
    hi = _strongest_edge_line(lines, c_hi, zone, agree)
    if hi is None and det_hi is not None:
        hi = det_hi
    offs = [v - c for v, c in ((lo, c_lo), (hi, c_hi)) if v is not None]

    if len(offs) == 2:
        med = float(np.median(offs))
        if max(abs(o - med) for o in offs) <= agree:
            return (med, 'snap') if abs(med) <= cap else (0.0, 'none')
        # 两线不一致：若实测间距明显小于共识框 → 是误检的内框，
        # 不允许缩框，改用两线中点居中对齐共识框。
        span = hi - lo
        mid = (lo + hi) / 2.0 - (c_lo + c_hi) / 2.0
        if span < size * (1.0 - tol) and span >= size * INNER_SPAN_MIN \
                and abs(mid) <= cap:
            return mid, 'center'
        # 间距正常或偏大（夹了外侧装饰线）：仍取中位，但受平移上限保护
        return (med, 'snap') if abs(med) <= cap else (0.0, 'none')
    if len(offs) == 1:
        d = float(offs[0])
        return (d, 'edge') if abs(d) <= cap else (0.0, 'none')
    return 0.0, 'none'


def consensus_box_shift(vc_abs, hc_abs, cons, det=None, vc_hf=None):
    """版心框位置：尺寸恒定取共识框，只根据当页实测线整体平移（贴边/居中）。
    det 为四边检测器确认的 [x1,y1,x2,y2]（可空），其边线不受采信带限制。
    vc_hf 为「页眉/页脚带内检出的竖线」候选表（可空），启用页眉页脚检测时
    用它参与 X 轴贴边——这是「共识框紧贴垂直边线」的落点；为空则退回整页长竖线。
    返回 (dx, dy, mode_x, mode_y)，mode ∈ snap/center/edge/none。"""
    cx1, cy1, cx2, cy2 = cons
    cw, ch = cx2 - cx1, cy2 - cy1
    vc_x = vc_hf if vc_hf else vc_abs
    dx, mx = _axis_shift(vc_x, cx1, cx2, cw,
                         ZONE_PCT * cw, SHIFT_CAP_PCT_X * cw,
                         AGREE_PCT * cw, SIZE_TOL_W,
                         det[0] if det else None, det[2] if det else None)
    dy, my = _axis_shift(hc_abs, cy1, cy2, ch,
                         ZONE_PCT * ch, SHIFT_CAP_PCT_Y * ch,
                         AGREE_PCT * ch, SIZE_TOL_H,
                         det[1] if det else None, det[3] if det else None)
    return dx, dy, mx, my


# ---------------- 页眉 / 页脚 / 分镜框线检测 ----------------
# 归一化参考系：s = HF_IW / W，sH = H * s。所有 HF_* 阈值都在这个尺度上。

def _hf_row_segs(rows):
    """布尔序列 → [(起, 止), ...] 连续段。"""
    segs, s0 = [], None
    for i, v in enumerate(rows):
        if v and s0 is None:
            s0 = i
        elif not v and s0 is not None:
            segs.append((s0, i))
            s0 = None
    if s0 is not None:
        segs.append((s0, len(rows)))
    return segs


def _hf_band(gray, y0, y1, W, s):
    """把原图 [y0,y1) 条带缩放到 HF_IW 宽，返回 (缩放后灰度, 带高)。"""
    crop = gray[y0:y1, :]
    hh = max(1, int(round(crop.shape[0] * s)))
    return cv2.resize(crop, (HF_IW, hh), interpolation=cv2.INTER_AREA), hh


# ============ 1. 页眉（卷标题 running header）============

def hf_header_seed(sm, sH):
    """左窄带行分段取「最上方实体文字带」。不合格就地判无种子，不再向下搜索
    （否则会误取第一格分镜里的对白）。"""
    BX0, BX1 = int(HF_BX0 * HF_IW), int(HF_BX1 * HF_IW)
    band = (sm[:, BX0:BX1] > 0).astype(np.uint8)
    BW = BX1 - BX0
    rows = band.sum(axis=1) > max(2, int(0.01 * BW))
    # 容 1 行空隙，防笔画断裂把一段拆成两段
    rows = cv2.morphologyEx(
        rows.astype(np.uint8).reshape(-1, 1) * 255, cv2.MORPH_CLOSE,
        cv2.getStructuringElement(cv2.MORPH_RECT, (1, 3))).ravel() > 0
    for a, b in _hf_row_segs(rows):
        hr = (b - a) / sH
        if hr < HF_SEG_HMIN:
            continue
        sub = band[a:b, :]
        cols = np.where(sub.sum(axis=0) > 0)[0]
        if cols.size == 0:
            continue
        xl, xr = int(cols[0]) + BX0, int(cols[-1]) + BX0
        wr = (xr - xl) / float(HF_IW)
        dens = float(sub.mean())
        nseg = len(_hf_row_segs(sub.sum(axis=0) > 0))
        if (HF_HR[0] <= hr <= HF_HR[1] and HF_WR[0] <= wr <= HF_WR[1]
                and HF_DENS[0] <= dens <= HF_DENS[1]
                and nseg >= HF_NSEG_MIN and a >= HF_YTOP_MIN * sH):
            return dict(ytop=int(a), ybot=int(b), xl=xl, xr=xr,
                        h=int(b - a), hr=hr, wr=wr, dens=dens, nseg=nseg)
        return None
    return None


def hf_header_volume(records):
    """卷级页眉检测：核心种子 → 短模板 → 全卷 TM_CCOEFF_NORMED 匹配 → 迭代重建。
    records 需含 'hf_sm'（顶部墨迹带 uint8）、'hf_seed'、'page'、'name'。
    命中结果就地写入 r['hf_header']。"""
    for r in records:
        r['hf_header'] = None
    seeds = [(r, r['hf_seed']) for r in records if r.get('hf_seed')]
    info = dict(style='none', n_seed=len(seeds))
    if len(seeds) < HF_MIN_SEED:
        info['reason'] = 'seed_insufficient'
        return info

    xls = np.array([s['xl'] for _, s in seeds], float)
    hrs = np.array([s['hr'] for _, s in seeds], float)
    mx, mh = float(np.median(xls)), float(np.median(hrs))
    core = [(r, s) for r, s in seeds
            if abs(s['xl'] - mx) <= HF_IW * HF_CORE_DXL
            and abs(s['hr'] - mh) <= HF_CORE_DHR]
    if len(core) < HF_MIN_SEED:
        core = seeds
    TH = max(4, int(np.median([s['h'] for _, s in core]) * HF_TH_K))
    TWS = HF_TWS
    items = [(r, s['ybot'] - TH + int(TH * 0.08), s['xl']) for r, s in core]

    def build(its):
        acc = np.zeros((TH, TWS), np.float32)
        c = 0
        for r, ay, ax in its:
            sm = r['hf_sm']
            if ay < 0 or ax < 0 or ay + TH > sm.shape[0] or ax + TWS > sm.shape[1]:
                continue
            acc += (sm[ay:ay + TH, ax:ax + TWS] > 0).astype(np.float32)
            c += 1
        return (acc / c if c else None), c

    XA = max(0, int(mx - HF_IW * 0.08))
    XB = min(HF_IW, int(mx + TWS + HF_IW * 0.08))
    hits, tmpl, cnt = [], None, 0
    for rnd in range(HF_ROUNDS):
        tmpl, cnt = build(items)
        if tmpl is None:
            info['reason'] = 'template_failed'
            return info
        t8 = np.clip(tmpl * 255, 0, 255).astype(np.uint8)
        res = []
        for r in records:
            sm = r['hf_sm'][:, XA:XB]
            if sm.shape[0] <= TH or sm.shape[1] <= TWS:
                res.append((r, -1.0, 0, 0))
                continue
            _, mv, _, ml = cv2.minMaxLoc(
                cv2.matchTemplate(sm, t8, cv2.TM_CCOEFF_NORMED))
            res.append((r, float(mv), int(ml[1]), int(ml[0]) + XA))
        hits = [(r, cc, ay, ax) for r, cc, ay, ax in res
                if cc >= HF_CC and abs(ax - mx) <= HF_IW * HF_XGATE]
        if len(hits) < HF_MIN_SEED or rnd == HF_ROUNDS - 1:
            break
        items = [(r, ay, ax) for r, cc, ay, ax in hits]

    for r, cc, ay, ax in hits:
        r['hf_header'] = dict(box=[ax, ay, ax + TWS, ay + TH], cc=round(cc, 3))
    odd = sum(1 for r, cc, ay, ax in hits if r['page'] % 2)
    info.update(style='running_header' if len(hits) >= HF_MIN_SEED else 'none',
                n_core=len(core), rounds=HF_ROUNDS,
                template=dict(w=TWS, h=TH, x_med=round(mx, 1),
                              hr_med=round(mh, 5)),
                search_x=[XA, XB], n_superpose=cnt,
                hits=len(hits), odd=odd, even=len(hits) - odd)
    return info


# ============ 2. 页脚（页码）============

def _hf_glyph(patch, pol):
    """字形归一化：neg 极性反转成白底黑字 → 缩放到 GH×GW → z-score。"""
    if patch is None or patch.size == 0:
        return None
    p = patch.astype(np.float32)
    if pol == 'neg':
        p = 255.0 - p
    p = cv2.resize(p, (HF_GW, HF_GH), interpolation=cv2.INTER_AREA)
    sd = float(p.std())
    return (p - float(p.mean())) / (sd if sd > 1e-6 else 1.0)


def _hf_cut_sets(prof, wbb, nd):
    """按列墨迹投影找谷点，枚举把 bbox 切成恰好 nd 段的组合。"""
    if nd <= 1:
        return [[]]
    avg = wbb / float(nd)
    lo, hi = int(HF_CUT_W[0] * avg), int(HF_CUT_W[1] * avg)
    mn, mx = float(prof.min()), float(prof.max())
    thr = mn + HF_CUT_T * (mx - mn)
    cand = [i for i in range(2, wbb - 1) if prof[i] <= thr]
    if not cand:
        cand = list(range(2, wbb - 1))
    if len(cand) > HF_CUT_TOPK:
        cand = sorted(sorted(cand, key=lambda i: prof[i])[:HF_CUT_TOPK])
    out = []
    for combo in itertools.combinations(cand, nd - 1):
        bnds = [0] + list(combo) + [wbb]
        if all(lo <= b - a <= hi for a, b in zip(bnds[:-1], bnds[1:])):
            out.append(list(combo))
            if len(out) >= HF_CUT_MAX:
                break
    return out


def _hf_row_blocks(mask, sH, side):
    """角区连通块 → 按底基线聚行。用底基线而非中心线：页码含斜体，底部对齐最稳。"""
    cw = int(HF_CORNW * HF_IW)
    x0, x1 = (0, cw) if side == 'L' else (HF_IW - cw, HF_IW)
    sub = mask[:, x0:x1].astype(np.uint8)
    nl, lab, st, ct = cv2.connectedComponentsWithStats(sub, 8)
    bl = []
    for i in range(1, nl):
        x, y, ww, hh, area = st[i]
        if area < HF_AREA_MIN or ww < 2 or ww > HF_BLK_W * HF_IW:
            continue
        if not (HF_BLK_H[0] * sH <= hh <= HF_BLK_H[1] * sH):
            continue
        if area / float(ww * hh) < HF_FILL:
            continue
        bl.append(dict(x=int(x) + x0, y=int(y), w=int(ww), h=int(hh)))
    rows = []
    for b in sorted(bl, key=lambda q: q['y'] + q['h']):
        for r in rows:
            if abs((b['y'] + b['h']) - r['yb']) <= HF_ROW_DYB * sH \
                    and abs(b['h'] - r['h0']) <= HF_ROW_DH * sH:
                r['bs'].append(b)
                r['yb'] = float(np.mean([q['y'] + q['h'] for q in r['bs']]))
                break
        else:
            rows.append(dict(bs=[b], yb=float(b['y'] + b['h']), h0=b['h']))
    return rows


def hf_footer_cands(gm, sH, bH, side, nd):
    """双极性掩膜 → 连通块 → 聚行 → 行内连续子串枚举 → 候选（含字形切片）。

    必须做「行内子串枚举」而不是把整行当一个候选：同一基线上一个噪点就能把整行
    撑宽到超过串宽上限，导致整页候选为 0。
    """
    out = []
    ncut = 0
    for pol in ('pos', 'neg'):
        mask = (gm < HF_POS_T) if pol == 'pos' else (gm > HF_NEG_T)
        for r in _hf_row_blocks(mask, sH, side):
            bs = sorted(r['bs'], key=lambda q: q['x'])
            m = len(bs)
            hm = float(np.median([b['h'] for b in bs]))
            seen = set()
            for i in range(m):
                for j in range(i + 1, min(m, i + nd + 1) + 1):
                    seg = bs[i:j]
                    gap_ok = True
                    for a, b in zip(seg[:-1], seg[1:]):
                        if b['x'] - (a['x'] + a['w']) > HF_GAP_K * hm:
                            gap_ok = False
                            break
                    if not gap_ok:
                        continue
                    x0 = min(b['x'] for b in seg)
                    x1 = max(b['x'] + b['w'] for b in seg)
                    y0 = min(b['y'] for b in seg)
                    y1 = max(b['y'] + b['h'] for b in seg)
                    if (x1 - x0) > HF_STR_W * HF_IW:
                        continue
                    key = (x0, x1, y0, y1, len(seg), pol)
                    if key in seen:
                        continue
                    seen.add(key)
                    c = dict(pol=pol, n=len(seg), x0=int(x0), x1=int(x1),
                             y0=int(y0), y1=int(y1), h=int(y1 - y0))
                    c['hn'] = c['h'] / sH
                    c['botd'] = (bH - c['y1']) / sH
                    c['edged'] = ((c['x0'] if side == 'L' else HF_IW - c['x1'])
                                  / float(HF_IW))
                    c['gl'] = []
                    if c['n'] == nd:
                        for b in seg:
                            q = _hf_glyph(gm[b['y']:b['y'] + b['h'],
                                             b['x']:b['x'] + b['w']], pol)
                            if q is None:
                                c['gl'] = []
                                break
                            c['gl'].append(q)
                    c['cuts'] = []
                    if c['n'] != nd and ncut < HF_CUT_PAGE_MAX:
                        pw = gm[c['y0']:c['y1'], c['x0']:c['x1']]
                        prof = mask[c['y0']:c['y1'],
                                    c['x0']:c['x1']].sum(axis=0).astype(np.float32)
                        for combo in _hf_cut_sets(prof, c['x1'] - c['x0'], nd):
                            bnds = [0] + combo + [c['x1'] - c['x0']]
                            gs = []
                            for a, b in zip(bnds[:-1], bnds[1:]):
                                q = _hf_glyph(pw[:, a:b], pol)
                                if q is None:
                                    gs = []
                                    break
                                gs.append(q)
                            if len(gs) == nd:
                                c['cuts'].append(gs)
                                ncut += 1
                                if ncut >= HF_CUT_PAGE_MAX:
                                    break
                    if not c['gl'] and not c['cuts']:
                        continue
                    out.append(c)
    return out


def hf_footer_hlines(gm):
    """横线双通道并集：长核（整行框线）+ 短核（短框线）。缺一不可——
    只用长核漏短框线，只用短核漏部分整行线，两者命中集不重叠。"""
    out = []
    ink = (gm < HF_POS_T).astype(np.uint8)
    k = cv2.getStructuringElement(cv2.MORPH_RECT, (int(HF_LONG_K * HF_IW), 1))
    h = cv2.morphologyEx(ink, cv2.MORPH_OPEN, k)
    ys = np.where(h.sum(axis=1) > HF_LONG_K * HF_IW)[0]
    if ys.size:
        grp = [[ys[0]]]
        for y in ys[1:]:
            if y - grp[-1][-1] <= 4:
                grp[-1].append(y)
            else:
                grp.append([y])
        for g in grp:
            cols = np.where(h[g[0]:g[-1] + 1, :].sum(axis=0) > 0)[0]
            if cols.size:
                out.append((int(g[-1]), int(cols[0]), int(cols[-1])))
    for mk in ((gm < HF_POS_T), (gm > HF_NEG_T)):
        k = cv2.getStructuringElement(cv2.MORPH_RECT,
                                      (int(HF_SHORT_K * HF_IW), 1))
        hs = cv2.morphologyEx(mk.astype(np.uint8), cv2.MORPH_OPEN, k)
        hs = cv2.dilate(hs, cv2.getStructuringElement(cv2.MORPH_RECT, (1, 3)))
        nl, lab, st, ct = cv2.connectedComponentsWithStats(hs, 8)
        for i in range(1, nl):
            x, y, ww, hh, area = st[i]
            if ww < 0.05 * HF_IW or hh > HF_W_MAX * gm.shape[0]:
                continue
            # 必须按连通段各自取端点；跨整行取会把端点拉到画面边缘使 dx 判据失效
            out.append((int(y + hh - 1), int(x), int(x + ww - 1)))
    return out


def _hf_vote(items):
    """(botd, edged) 2D 直方图 + 3×3 邻域聚合取峰，返回 (峰内中位 [hn,botd,edged], 票数)。
    纯中位数会被每页 5-9 个噪声候选拉偏，众数投票只看密度峰，噪声无法成峰。"""
    if not items:
        return None
    box = {}
    for pg, c in items:
        k = (int(round(c['botd'] / HF_BIN_B)), int(round(c['edged'] / HF_BIN_E)))
        box.setdefault(k, []).append((pg, c))
    best, bn = None, -1
    for k in box:
        tot = set()
        for dy in (-1, 0, 1):
            for dx in (-1, 0, 1):
                for pg, c in box.get((k[0] + dy, k[1] + dx), []):
                    tot.add(pg)
        if len(tot) > bn:
            bn, best = len(tot), k
    grp = []
    for dy in (-1, 0, 1):
        for dx in (-1, 0, 1):
            grp += box.get((best[0] + dy, best[1] + dx), [])
    a = np.array([[c['hn'], c['botd'], c['edged']] for _, c in grp])
    return np.median(a, axis=0), bn


def hf_footer_paradigm_a(recs):
    """范式 A（相对页底）：绝对先验窗口 + 2D 众数投票 + side×parity 组内收敛。
    返回 (hits dict, 信息, 是否可用)。"""
    hits, info = {}, {}
    spreads, ratios = [], []
    for par in (1, 0):
        grp = [r for r in recs if r['page'] % 2 == par]
        if not grp:
            continue
        items = [(r['name'], c) for r in grp for c in r['cands']
                 if HF_A_H[0] <= c['hn'] <= HF_A_H[1]
                 and HF_A_BOT[0] <= c['botd'] <= HF_A_BOT[1]
                 and HF_A_EDGE[0] <= c['edged'] <= HF_A_EDGE[1]]
        v = _hf_vote(items)
        if not v:
            info[par] = dict(n=len(grp), ncand=len(items), vote=0, hits=0)
            continue
        med, nvote = v
        hp = {}
        for r in grp:
            best, bd = None, 1e9
            for c in r['cands']:
                dd = (abs(c['hn'] - med[0]) / HF_A_DH
                      + abs(c['botd'] - med[1]) / HF_A_DB
                      + abs(c['edged'] - med[2]) / HF_A_DE)
                if dd < bd:
                    bd, best = dd, c
            if best and (abs(best['hn'] - med[0]) <= HF_A_DH
                         and abs(best['botd'] - med[1]) <= HF_A_DB
                         and abs(best['edged'] - med[2]) <= HF_A_DE):
                best['fit'] = round(max(0.0, 1.0 - bd / 3.0), 3)
                hp[r['name']] = best
        info[par] = dict(n=len(grp), ncand=len(items), vote=nvote,
                         peak=[round(float(x), 5) for x in med], hits=len(hp))
        ratios.append(nvote / float(max(1, len(grp))))
        if hp:
            b = np.array([c['botd'] for c in hp.values()])
            spreads.append(float(np.percentile(b, 90) - np.percentile(b, 10)))
        hits.update(hp)
    spread = max(spreads) if spreads else 0.0
    ratio = min(ratios) if ratios else 0.0
    usable = bool(spreads) and spread <= HF_A_MAXSPREAD and ratio >= HF_A_MINRATIO
    return hits, info, usable, spread, ratio


def _hf_train(seed):
    """以 str(页号) 为标签累加 0-9 字形模板（逐块 z-score 后取平均再归一）。

    seed 形如 {文件名: (页号, [字形, ...])}；标签必须取页号而非 dict 的 key，
    否则位数不匹配、模板全空（见方案 §7）。
    """
    acc = dict((d, []) for d in '0123456789')
    for _name, (n, gl) in seed.items():
        ds = str(n)
        if len(ds) != len(gl):
            continue
        for d, g in zip(ds, gl):
            acc[d].append(g)
    tm = {}
    for d in acc:
        if acc[d]:
            m = np.mean(acc[d], axis=0)
            sd = float(m.std())
            tm[d] = (m - float(m.mean())) / (sd if sd > 1e-6 else 1.0)
    return tm, dict((d, len(acc[d])) for d in sorted(acc))


def _hf_score(c, n, tm):
    """逐位字形相关取「最小位」——任一位不匹配即整体否决，这是 FP=0 的来源。"""
    ds = str(n)
    best = -1.0
    for gl in ([c['gl']] if len(c['gl']) == len(ds) else []) + c['cuts']:
        ss = []
        for d, g in zip(ds, gl):
            if d not in tm:
                ss = None
                break
            ss.append(float((g * tm[d]).mean()))
        if ss:
            v = min(ss)
            if v > best:
                best = v
    return best


def hf_footer_paradigm_b(recs):
    """范式 B（相对分镜横线）：横线对齐种子 → 字形自学习 → 全卷回扫。单轮，不自举
    （自举会降低召回，见方案 §7）。"""
    seed = {}
    for r in recs:
        n = r['page']
        nd = len(str(n))
        side = r['side']
        anc = []
        for c in r['cands']:
            if not (HF_B_H[0] <= c['hn'] <= HF_B_H[1]
                    and HF_B_BOT[0] <= c['botd'] <= HF_B_BOT[1]):
                continue
            best = None
            for (ly, lxs, lxe) in r['hlines']:
                dy = (c['y0'] - ly) / r['sH']
                if not (HF_B_DY[0] <= dy <= HF_B_DY[1]):
                    continue
                dx = ((c['x0'] - lxs) if side == 'L'
                      else (lxe - c['x1'])) / float(HF_IW)
                if abs(dx) > HF_B_DX:
                    continue
                sc = abs(dx) / HF_B_DX + dy / HF_B_DY[1]
                if best is None or sc < best[0]:
                    best = (sc, dy, dx, (ly, lxs, lxe))
            if best is None:
                continue
            c['dy'], c['dx'], c['anc'] = best[1], best[2], best[3]
            if c['n'] == nd and c['gl']:
                anc.append((best[0], c))
        if anc:
            anc.sort(key=lambda t: t[0])
            # 值必须带页号：_hf_train 用它作字形标签（key 是文件名，不可当标签）
            seed[r['name']] = (n, anc[0][1]['gl'])
    info = dict(n_seed=len(seed))
    if len(seed) < HF_B_MIN_SEED:
        info['reason'] = 'seed_insufficient'
        return {}, info, {}
    tm, acc = _hf_train(seed)
    info['glyph_samples'] = acc
    hits = {}
    for r in recs:
        n, nd = r['page'], len(str(r['page']))
        bv, bc = -1.0, None
        for c in r['cands']:
            v = _hf_score(c, n, tm)
            if v > bv:
                bv, bc = v, c
        if bc is not None and bv >= HF_T_CC:
            hits[r['name']] = bc
    info['hits'] = len(hits)
    return hits, info, tm


def hf_footer_volume(records):
    """卷级页脚检测：先试范式 A，失效则自动切换到范式 B。
    就地写入 r['hf_footer']；返回卷级信息。"""
    for r in records:
        r['hf_footer'] = None
    recs = [dict(name=r['name'], page=r['page'], side=r['hf_side'],
                 sH=r['hf_sH'], cands=r['hf_cands'], hlines=r['hf_hlines'])
            for r in records]
    hitsA, infoA, usableA, spread, ratio = hf_footer_paradigm_a(recs)
    out = dict(paradigm='bottom_relative',
               a_info=infoA, a_spread=round(spread, 4), a_ratio=round(ratio, 3),
               a_usable=bool(usableA), a_hits=len(hitsA))
    if usableA:
        for r in records:
            c = hitsA.get(r['name'])
            if c is not None:
                r['hf_footer'] = dict(
                    box=[c['x0'], c['y0'], c['x1'], c['y1']],
                    side=r['hf_side'], pol=c['pol'], digits=c['n'],
                    cc=float(c.get('fit', 0.0)), anchor_line=None)
        out['hits'] = len(hitsA)
        return out

    hitsB, infoB, tm = hf_footer_paradigm_b(recs)
    out['paradigm'] = 'panel_line_relative'
    out['b_info'] = infoB
    if not hitsB:
        out['hits'] = 0
        out['reason'] = infoB.get('reason', 'no_hit')
        return out
    for r in records:
        c = hitsB.get(r['name'])
        if c is None:
            continue
        r['hf_footer'] = dict(
            box=[c['x0'], c['y0'], c['x1'], c['y1']],
            side=r['hf_side'], pol=c['pol'], digits=c['n'],
            cc=round(_hf_score(c, r['page'], tm), 3),
            anchor_line=(dict(y=c['anc'][0], x0=c['anc'][1], x1=c['anc'][2])
                         if c.get('anc') else None))
    out['hits'] = len(hitsB)
    out['glyph_templates'] = dict(
        (d, dict(samples=infoB.get('glyph_samples', {}).get(d, 0)))
        for d in '0123456789')
    return out


# ============ 3. 分镜框线（LINE / EDGE / STRIPE 三通道）============

def _hf_axis_pos(mask, orient, x, y, ww, hh):
    """中轴必须取墨迹峰值行/列，不能取几何中心——真线偏离几何中心 1px 就会
    让覆盖率掉到阈值以下被误杀。"""
    sub = mask[y:y + hh, x:x + ww]
    if sub.size == 0:
        return int(y + hh // 2) if orient == 'h' else int(x + ww // 2)
    if orient == 'h':
        return int(y + int(np.argmax(sub.mean(axis=1))))
    return int(x + int(np.argmax(sub.mean(axis=0))))


def _hf_runs_at(mask, orient, pos, coords, sH):
    lim = int(HF_W_MAX * sH)
    band = max(3, int(0.004 * sH))
    gap = 2
    ths, dA, dB = [], [], []
    for c in coords:
        line = mask[:, c] if orient == 'h' else mask[c, :]
        if pos >= line.size or not line[pos]:
            continue
        a = pos
        while a - 1 >= 0 and line[a - 1] and pos - a < lim:
            a -= 1
        b = pos
        while b + 1 < line.size and line[b + 1] and b - pos < lim:
            b += 1
        ths.append(b - a + 1)
        s0, e0 = max(0, a - gap - band), max(0, a - gap)
        s1, e1 = min(line.size, b + gap + 1), min(line.size, b + gap + 1 + band)
        if e0 > s0:
            dA.append(float(line[s0:e0].mean()))
        if e1 > s1:
            dB.append(float(line[s1:e1].mean()))
    if not ths:
        return None
    return (float(np.median(ths)),
            float(np.median(dA)) if dA else 1.0,
            float(np.median(dB)) if dB else 1.0,
            len(ths) / float(len(coords)))


def _hf_step_at(mask, orient, pos, coords, sH):
    gap = max(6, int(HF_E_GAP * sH))
    band = max(6, int(HF_E_BAND * sH))
    dA, dB = [], []
    for c in coords:
        line = mask[:, c] if orient == 'h' else mask[c, :]
        s0, e0 = pos - gap - band, pos - gap
        s1, e1 = pos + gap + 1, pos + gap + 1 + band
        if s0 < 0 or e1 > line.size:
            continue
        dA.append(float(line[s0:e0].mean()))
        dB.append(float(line[s1:e1].mean()))
    if not dA:
        return None
    a, b = float(np.median(dA)), float(np.median(dB))
    return a, b, abs(a - b), len(dA) / float(len(coords))


def _hf_line_segs(gm, sH, orient, pol):
    mask = ((gm < HF_POS_T) if pol == 'pos' else (gm > HF_NEG_T)).astype(np.uint8)
    bH = mask.shape[0]
    out = []
    krs = HF_K_H if orient == 'h' else HF_K_V
    for kr in krs:
        if orient == 'h':
            klen = max(3, int(kr * HF_IW))
            k = cv2.getStructuringElement(cv2.MORPH_RECT, (klen, 1))
            dil = (1, 3)
        else:
            klen = max(3, int(kr * bH))
            k = cv2.getStructuringElement(cv2.MORPH_RECT, (1, klen))
            dil = (3, 1)
        op = cv2.morphologyEx(mask, cv2.MORPH_OPEN, k)
        op = cv2.dilate(op, cv2.getStructuringElement(cv2.MORPH_RECT, dil))
        nl, lab, st, ct = cv2.connectedComponentsWithStats(op, 8)
        for i in range(1, nl):
            x, y, ww, hh, area = st[i]
            if orient == 'h':
                if ww < klen or hh > HF_W_MAX * sH:
                    continue
                lnr = ww / float(HF_IW)
                if lnr < HF_T_LNH:
                    continue
                pos = _hf_axis_pos(mask, 'h', x, y, ww, hh)
                cs = np.linspace(x + 1, x + ww - 2, 21).astype(int)
                r = _hf_runs_at(mask, 'h', pos, cs, sH)
            else:
                if hh < klen or ww > HF_W_MAX * sH:
                    continue
                lnr = hh / float(bH)
                if lnr < HF_T_LNV:
                    continue
                pos = _hf_axis_pos(mask, 'v', x, y, ww, hh)
                cs = np.linspace(y + 1, y + hh - 2, 21).astype(int)
                r = _hf_runs_at(mask, 'v', pos, cs, sH)
            if r is None:
                continue
            th, da, db, cov = r
            if not (HF_T_THN[0] <= th / sH <= HF_T_THN[1]):
                continue
            if max(da, db) > HF_T_DEN or cov < HF_T_COV:
                continue
            if orient == 'h':
                y, hh = pos - 1, 3
            else:
                x, ww = pos - 1, 3
            out.append(dict(kind='line', o=orient, pol=pol, x=int(x), y=int(y),
                            w=int(ww), h=int(hh), lnr=lnr, thn=th / sH, cov=cov))
    return out


def _hf_edge_segs(gm, sH, orient):
    """EDGE 通道（黑块/网点块边界）：不能挂在 LINE 通道上——一侧墨迹饱和，
    runs_at 的 run 会一直延伸进黑块，线厚直接爆表。改用阶跃幅度。"""
    ink = (gm < HF_POS_T).astype(np.uint8)
    bH = ink.shape[0]
    out = []
    if orient == 'h':
        d = (ink[:-1, :] != ink[1:, :]).astype(np.uint8)
        klen = max(3, int(HF_E_LNH * HF_IW))
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (klen, 1))
        dil = (1, 3)
    else:
        d = (ink[:, :-1] != ink[:, 1:]).astype(np.uint8)   # 纵向边界
        klen = max(3, int(HF_E_LNV * bH))
        k = cv2.getStructuringElement(cv2.MORPH_RECT, (1, klen))
        dil = (3, 1)
    op = cv2.morphologyEx(d, cv2.MORPH_OPEN, k)
    op = cv2.dilate(op, cv2.getStructuringElement(cv2.MORPH_RECT, dil))
    nl, lab, st, ct = cv2.connectedComponentsWithStats(op, 8)
    for i in range(1, nl):
        x, y, ww, hh, area = st[i]
        if orient == 'h':
            if ww < klen or hh > HF_E_WMAX * sH:
                continue
            lnr = ww / float(HF_IW)
            pos = _hf_axis_pos(d, 'h', x, y, ww, hh)
            cs = np.linspace(x + 1, x + ww - 2, 21).astype(int)
            r = _hf_step_at(ink, 'h', pos, cs, sH)
        else:
            if hh < klen or ww > HF_E_WMAX * sH:
                continue
            lnr = hh / float(bH)
            pos = _hf_axis_pos(d, 'v', x, y, ww, hh)
            cs = np.linspace(y + 1, y + hh - 2, 21).astype(int)
            r = _hf_step_at(ink, 'v', pos, cs, sH)
        if r is None:
            continue
        da, db, stp, cov = r
        if stp < HF_T_STEP or cov < HF_T_COV:
            continue
        if orient == 'h':
            y, hh = pos - 1, 3
        else:
            x, ww = pos - 1, 3
        out.append(dict(kind='edge', o=orient, pol='pos', x=int(x), y=int(y),
                        w=int(ww), h=int(hh), lnr=lnr, thn=0.0, cov=cov))
    return out


def _hf_ends(q):
    if q['o'] == 'h':
        yy = q['y'] + q['h'] / 2.0
        return (q['x'], yy), (q['x'] + q['w'], yy)
    xx = q['x'] + q['w'] / 2.0
    return (xx, q['y']), (xx, q['y'] + q['h'])


def _hf_dedup(segs):
    """同方向、中轴距 ≤8px、投影重叠 >50% 视为重复；优先保留 LINE 型。"""
    segs = sorted(segs, key=lambda q: (0 if q['kind'] == 'line' else 1,
                                       -(q['w'] if q['o'] == 'h' else q['h'])))
    keep = []
    for q in segs:
        dup = False
        for r in keep:
            if q['o'] != r['o']:
                continue
            if q['o'] == 'h':
                if abs((q['y'] + q['h'] / 2.0) - (r['y'] + r['h'] / 2.0)) <= 8 and \
                   min(q['x'] + q['w'], r['x'] + r['w']) - max(q['x'], r['x']) > 0.5 * q['w']:
                    dup = True
                    break
            else:
                if abs((q['x'] + q['w'] / 2.0) - (r['x'] + r['w'] / 2.0)) <= 8 and \
                   min(q['y'] + q['h'], r['y'] + r['h']) - max(q['y'], r['y']) > 0.5 * q['h']:
                    dup = True
                    break
        if not dup:
            keep.append(q)
    return keep


def _hf_orth_support(q, segs, sH, bH):
    R = max(4, int(HF_R_SUP * sH))
    n = 0
    for pt in _hf_ends(q):
        hit = False
        if q['o'] == 'h':
            if pt[0] <= HF_EDGE_PX * HF_IW or pt[0] >= HF_IW - HF_EDGE_PX * HF_IW:
                hit = True
        else:
            if pt[1] <= 2 or pt[1] >= bH - 3:
                hit = True
        if not hit:
            for r in segs:
                if r['o'] == q['o'] or r is q:
                    continue
                (ax, ay), (bx, by) = _hf_ends(r)
                if q['o'] == 'h':
                    if (ax - R <= pt[0] <= bx + R) and abs(ay - pt[1]) <= R:
                        hit = True
                        break
                else:
                    if (ay - R <= pt[1] <= by + R) and abs(ax - pt[0]) <= R:
                        hit = True
                        break
        if hit:
            n += 1
    return n


def _hf_drop_stripes(segs, sH):
    """排线纹理的局部特征与真框线逐项一致，唯一区别是「成组平行重复」。
    用法向邻域计数（不可用有序链式扫描：排线簇会分裂成多簇交错打断链条）。"""
    G = max(4, int(HF_W_STRIPE * sH))
    bad = set()

    def axis_of(q):
        return (q['y'] + q['h'] / 2.0) if q['o'] == 'h' else (q['x'] + q['w'] / 2.0)

    def span_of(q):
        return ((q['x'], q['x'] + q['w']) if q['o'] == 'h'
                else (q['y'], q['y'] + q['h']))

    for o in ('h', 'v'):
        grp = [q for q in segs if q['o'] == o]
        for a in grp:
            if a['lnr'] >= HF_T_LONG:      # 长线豁免：保护真框线主边
                continue
            n = 0
            for b in grp:
                if abs(axis_of(b) - axis_of(a)) > G:
                    continue
                (s0, e0), (s1, e1) = span_of(a), span_of(b)
                if min(e0, e1) - max(s0, s1) > HF_OV_STRIPE * min(e0 - s0, e1 - s1):
                    n += 1
            if n >= HF_N_STRIPE:
                bad.add(id(a))
    return [q for q in segs if id(q) not in bad], len(bad)


def hf_panel_lines(gm, sH, bH):
    """单条带的框线检测：LINE（双极性）+ EDGE（阶跃）→ 去重 → 正交支持
    → 排线抑制。返回线段列表。"""
    ls, es = [], []
    for pol in ('pos', 'neg'):
        for o in ('h', 'v'):
            ls += _hf_line_segs(gm, sH, o, pol)
    for o in ('h', 'v'):
        es += _hf_edge_segs(gm, sH, o)
    ok = _hf_dedup(ls + es)
    for q in ok:
        q['sup'] = _hf_orth_support(q, ok, sH, bH)
    fin = [q for q in ok if q['lnr'] >= HF_T_SHORT or q['sup'] >= 1]
    fin, nstr = _hf_drop_stripes(fin, sH)
    return fin


def hf_page_bands(gray, W, H, s, sH):
    """在页眉带 / 页脚带各跑一次框线检测。返回 [(带名, 线段列表), ...]。
    带名即归属标注（near_header / near_footer），不再另设判据。"""
    out = []
    for nm, y0, y1 in (('near_header', 0, int(H * HF_PNL_TOPR)),
                       ('near_footer', int(H * (1 - HF_PNL_BOTR)), H)):
        gm, bH = _hf_band(gray, y0, y1, W, s)
        out.append((nm, y0, hf_panel_lines(gm, sH, bH)))
    return out


def hf_band_vlines(bands, s):
    """从两个带的线段里取竖线，换算回绝对像素 → [(x_abs, 强度), ...]。"""
    out = []
    for nm, y0, segs in bands:
        for q in segs:
            if q['o'] != 'v':
                continue
            out.append(((q['x'] + q['w'] / 2.0) / s, float(q['lnr'])))
    return out


def hf_merge_vlines(pts, W):
    """竖线候选表聚类合并：位置接近的合并，两带都命中的加权。"""
    if not pts:
        return []
    tol = max(4.0, HF_V_MERGE_PX * W)
    out = []
    for x, st in sorted(pts):
        if out and x - out[-1][0] <= tol:
            n = out[-1][2]
            out[-1][0] = (out[-1][0] * n + x) / (n + 1.0)
            out[-1][1] += st
            out[-1][2] = n + 1
        else:
            out.append([x, st, 1])
    return [(c[0], c[1] * (1.3 if c[2] > 1 else 1.0)) for c in out]


def merge_vc(vc_abs, vc_hf, W):
    """把「页眉/页脚带竖线」并入「整页长竖线」候选表。
    两个来源的量纲不同，各自按自身峰值归一后再合并；位置接近的视为同一条，强度相加。"""
    if not vc_hf:
        return vc_abs
    if not vc_abs:
        return list(vc_hf)
    mh = max(s for _, s in vc_hf)
    ma = max(s for _, s in vc_abs)
    k = (mh / ma) if ma > 0 else 1.0
    pts = sorted([(p, s * k * HF_V_BASE_W) for p, s in vc_abs] + list(vc_hf))
    tol = max(4.0, HF_V_MERGE_PX * W)
    out = []
    for p, st in pts:
        if out and p - out[-1][0] <= tol:
            out[-1][1] += st
            out[-1][0] = (out[-1][0] + p) / 2.0
        else:
            out.append([p, st])
    return [(p, st) for p, st in out]


def _cluster_sorted(xs, tol):
    """对一维坐标做简单一维聚类，返回 [cluster_values]。"""
    if len(xs) == 0:
        return []
    s = np.sort(xs)
    clusters = []
    cur = [s[0]]
    for x in s[1:]:
        if x - cur[-1] <= tol:
            cur.append(x)
        else:
            clusters.append(cur)
            cur = [x]
    clusters.append(cur)
    return clusters


def hf_calib_consensus_width(cons, records):
    """用「成对竖线间距的众数」校准共识框宽度（并顺带给出左右边的绝对位置）。

    独立修正左右边可能取到不同对框线、宽度因此失真；成对间距众数才是真实
    版心宽。做法与阈值依据见方案 §4.7.1。

    返回 (新共识, info)。info['w'] 为采信的宽；未采信时 info['used']=False。
    """
    W = max(1, int(np.median([r['w'] for r in records if r.get('w')])))
    cw0 = cons[2] - cons[0]
    box = {}
    pairs = []
    for r in records:
        nm = r.get('name')
        rW = r.get('w') or W
        vs = r.get('hf_vlines') or []
        xs = sorted(x for x, st in vs
                    if HF_PW_EDGE[0] * rW <= x <= HF_PW_EDGE[1] * rW)
        for i in range(len(xs)):
            for j in range(i + 1, len(xs)):
                d = xs[j] - xs[i]
                if HF_PW_RANGE[0] * rW <= d <= HF_PW_RANGE[1] * rW:
                    box.setdefault(int(round(d / HF_PW_BIN)), set()).add(nm)
                    pairs.append((xs[i], xs[j], d))
    info = {'used': False, 'n_pair': len(pairs), 'w0': round(cw0, 1)}
    if not box:
        info['reason'] = 'no_pair'
        return cons, info
    best, bn = None, -1
    for k in box:
        tot = set()
        for dk in (-1, 0, 1):
            tot |= box.get(k + dk, set())
        if len(tot) > bn:
            bn, best = len(tot), k
    info['peak_w'] = round(best * HF_PW_BIN, 1)
    info['peak_pages'] = bn
    if bn < HF_PW_MIN_N:
        info['reason'] = 'peak_too_weak'
        return cons, info
    # 峰内（±1.5 格）所有对 → 宽 / 左 / 右 的中位数
    sel = [(a, b, d) for a, b, d in pairs
           if abs(d - best * HF_PW_BIN) <= 1.5 * HF_PW_BIN]
    if not sel:
        info['reason'] = 'peak_empty'
        return cons, info
    wmed = float(np.median([d for _, _, d in sel]))
    lmed = float(np.median([a for a, _, _ in sel]))
    rmed = float(np.median([b for _, b, _ in sel]))
    info.update(w=round(wmed, 1), x1=round(lmed, 1), x2=round(rmed, 1),
                n_sel=len(sel))
    strong = bn >= HF_PW_STRONG_N
    info['strong'] = bool(strong)
    if strong:
        if not (HF_PW_STRONG_RANGE[0] * W <= wmed <= HF_PW_STRONG_RANGE[1] * W):
            info['reason'] = 'strong_out_of_range'
            return cons, info
    elif abs(wmed - cw0) / max(1.0, cw0) > HF_PW_TOL:
        info['reason'] = 'delta_too_large'
        return cons, info
    # 采信：宽度改为 wmed，左右边取峰内中位（这一对本身就是配对的，天然贴齐）
    out = [lmed, cons[1], lmed + wmed, cons[3]]
    info['used'] = True
    info['dw'] = round(wmed - cw0, 1)
    return out, info


def hf_refine_consensus_x(cons, records, lock_width=False):
    """用页眉/页脚带的竖线修正共识框左右边，使共识框紧贴这些垂直边线。

    策略：整卷竖线聚类后，在「原共识边 ±HF_REF_WIN」窗口内找最近的稳定簇；
    窗口内无稳定簇时，退回原共识边（避免被页边/内部隔线拉偏）。只改 x1 / x2
    两个常数（全书共用一套），框尺寸仍然恒定，不做逐页缩放；修正量受
    HF_REF_CAP 限制，防止灾难性漂移。

    lock_width=True 时（宽度已由 hf_calib_consensus_width 校准）只做整体平移，
    不再各自拉边——否则校准好的宽度会被重新破坏。
    """
    cx1, cx2 = cons[0], cons[2]
    cw = cx2 - cx1
    W = max(1, int(np.median([r['w'] for r in records if r.get('w')])))
    cap = HF_REF_CAP * W   # 按图宽而非共识宽（共识宽本身可能就偏窄）
    win = HF_REF_WIN * cw  # 只采信原共识边附近的竖线
    tol = max(4.0, HF_V_MERGE_PX * W)
    n = 0
    all_x = []
    for r in records:
        vs = r.get('hf_vlines') or []
        if vs:
            n += 1
            all_x += [x for x, s in vs]
    if n < 5 or not all_x:
        return cons, [0.0, 0.0], (0, 0, n)
    arr = np.asarray(all_x, float)
    need = max(10, int(0.15 * n))

    def nearest_stable(xs, target):
        """在 xs 中找 target 附近 ±win 内最近的稳定簇；无则返回 None。"""
        if len(xs) == 0:
            return None, 0
        clusters = _cluster_sorted(xs, tol)
        sig = [c for c in clusters if len(c) >= need]
        if not sig:
            return None, 0
        best = None
        best_d = float('inf')
        for c in sig:
            med = float(np.median(c))
            if abs(med - target) <= win:
                d = abs(med - target)
                if d < best_d:
                    best_d, best = d, c
        if best is None:
            return None, 0
        return float(np.median(best)), len(best)

    # 左边界只看左半页（< 0.45W），右边界只看右半页（> 0.55W），避免跨半页误采
    L, nL = nearest_stable(arr[arr <= W * 0.45], cx1)
    R, nR = nearest_stable(arr[arr >= W * 0.55], cx2)

    out = list(cons)
    moved = [0.0, 0.0]
    if lock_width:
        # 宽度已由成对间距校准过 → 只允许【整体平移】，不得再各自拉边，
        # 否则刚校准好的宽度会被两边独立修正重新破坏。
        ds = [v - t for v, t in ((L, cx1), (R, cx2)) if v is not None]
        if ds:
            d = max(-cap, min(cap, float(np.mean(ds))))
            out[0], out[2] = cx1 + d, cx2 + d
            moved = [d, d]
        return out, moved, (nL, nR, n)
    if L is not None:
        d = L - cx1
        d = max(-cap, min(cap, d))
        out[0], moved[0] = cx1 + d, d
    if R is not None:
        d = R - cx2
        d = max(-cap, min(cap, d))
        out[2], moved[1] = cx2 + d, d
    return out, moved, (nL, nR, n)


def hf_refine_consensus_y(cons, records):
    """用页眉底边 / 页脚顶边修正共识框上下边，使共识框位于页眉下方、页脚上方。

    全书仍共用一套 y1 / y2；修正量受 HF_REF_CAP_Y 限制，并保证框高不低于
    原共识高的 55%，防止过度裁剪。页眉/页脚检测失效时，退回检测带内最强的
    水平框线作为边界证据。
    """
    cy1, cy2 = cons[1], cons[3]
    ch = cy2 - cy1
    W = max(1, int(np.median([r['w'] for r in records if r.get('w')])))
    H = max(1, int(np.median([r['h'] for r in records if r.get('h')])))
    cap = HF_REF_CAP_Y * H
    margin = HF_Y_MARGIN * H
    min_h = max(1, int(round(ch * 0.55)))

    def _header_bottom(r):
        hb = r.get('hf_header')
        if not hb or not hb.get('box'):
            return None
        s = r.get('hf_s', HF_IW / float(W))
        return hb['box'][3] / s

    def _footer_top(r):
        fb = r.get('hf_footer')
        if not fb or not fb.get('box'):
            return None
        s = r.get('hf_s', HF_IW / float(W))
        return H * (1.0 - HF_BOTR) + fb['box'][1] / s

    header_bottoms = [v for v in (_header_bottom(r) for r in records) if v is not None]
    footer_tops = [v for v in (_footer_top(r) for r in records) if v is not None]

    # fallback: 用页眉/页脚带的水平框线，但只取「紧贴原共识边、向内容侧移动」的线。
    # 对 y1 找页眉带中位于原 y1 下方且最近的一条；对 y2 找页脚带中位于原 y2 上方且最近的一条。
    # 这样避免把 y2 大幅上提到内部 Panel 分割线，只裁剪真正的页眉/页脚白边。
    if not header_bottoms:
        vals = []
        for r in records:
            bands = r.get('hf_bands') or []
            s = r.get('hf_s', HF_IW / float(W))
            for nm, y0, segs in bands:
                if nm != 'near_header':
                    continue
                best = None
                for q in segs:
                    if q['o'] == 'h' and q['lnr'] >= HF_T_SHORT:
                        y_abs = y0 + (q['y'] + q['h']) / s
                        # 必须位于原 y1 下方，且在 cap 范围内
                        if cy1 < y_abs <= cy1 + cap and (best is None or y_abs < best):
                            best = y_abs
                if best is not None:
                    vals.append(best)
        if vals:
            header_bottoms = vals
    # 注：y2 不启用水平框线 fallback。页脚带内的横线多为内部 Panel 分割线，
    # 用作内容下边界会过度裁剪内容（如 082/061 的底部插画/文字）。
    # 只有真正检测到页码（hf_footer）时才移动 y2；检测失败时保持原共识 y2。

    out = list(cons)
    moved = [0.0, 0.0]
    n_hb, n_ft = len(header_bottoms), len(footer_tops)

    # 页眉：用 p90 底边，保证大多数页的页眉都被排除；下方加留白
    if n_hb:
        new_y1 = float(np.percentile(header_bottoms, 90)) + margin
        d = new_y1 - cy1
        # 只向下移动（d>0），不向上
        if d > 0:
            d = min(d, cap)
            if cy2 - (cy1 + d) >= min_h:
                out[1], moved[0] = cy1 + d, d

    # 页脚：用 p10 顶边，保证大多数页的页脚都被排除；上方加留白
    if n_ft:
        new_y2 = float(np.percentile(footer_tops, 10)) - margin
        d = new_y2 - cy2
        # 只向上移动（d<0），不向下
        if d < 0:
            d = max(d, -cap)
            if (cy2 + d) - out[1] >= min_h:
                out[3], moved[1] = cy2 + d, d

    return out, moved, (n_hb, n_ft)


def hf_band_corners(segs, y0, s, y_lo=None, y_hi=None):
    """在一条带内求「水平框线 × 垂直框线」的交点（角点），换算回绝对像素。

    segs 是归一化带图（宽 IW）上的线段，y 坐标相对带起点 y0。
    角点取法：水平线取【下沿】（内容在线下方，框顶边贴线的下沿），
    垂直线取【中心】（框侧边压在线上）。两者必须真正相交才算角点。

    返回 [(x_abs, y_abs, strength), ...]。
    """
    hs = [q for q in segs if q['o'] == 'h']
    vs = [q for q in segs if q['o'] == 'v']
    if not hs or not vs:
        return []
    out = []
    for hq in hs:
        hx0, hx1 = hq['x'], hq['x'] + hq['w']       # 水平线 x 跨度（归一化）
        hyb = hq['y'] + hq['h']                     # 水平线下沿（归一化，带内）
        for vq in vs:
            vy0, vy1 = vq['y'], vq['y'] + vq['h']   # 垂线 y 跨度（归一化，带内）
            vxc = vq['x'] + vq['w'] / 2.0           # 垂线 x 中心（归一化）
            # 相交：垂线 x 落在水平线跨度内，且水平线下沿落在垂线跨度内
            if not (hx0 - 1 <= vxc <= hx1 + 1):
                continue
            if not (vy0 - 1 <= hyb <= vy1 + 1):
                continue
            x_abs = vxc / s
            y_abs = y0 + hyb / s
            if y_lo is not None and y_abs < y_lo:
                continue
            if y_hi is not None and y_abs > y_hi:
                continue
            out.append((x_abs, y_abs, float(hq['lnr'] * vq['lnr'])))
    return out


def hf_band_hlines(segs, y0, s, y_lo=None, y_hi=None):
    """带内水平框线 → [(x0_abs, x1_abs, y_abs, lnr), ...]，y 取【下沿】。

    用于「找不到横竖真交点」时的降级对齐，见方案 §4.7.6。
    """
    out = []
    for q in segs:
        if q['o'] != 'h':
            continue
        y_abs = y0 + (q['y'] + q['h']) / s
        if y_lo is not None and y_abs < y_lo:
            continue
        if y_hi is not None and y_abs > y_hi:
            continue
        out.append((q['x'] / s, (q['x'] + q['w']) / s, y_abs, float(q['lnr'])))
    return out


def hf_page_baselines(r, W, H):
    """【逐页】提取上下对齐基准线 —— 全书对齐规则的唯一证据源。

    对齐规则、搜索窗约束与降级策略见方案 §4.7.6。

    返回 dict:
      top = (y, x_or_None, kind)  上基准线；x 为 None 表示该轴无可信证据
      bot = (y, x_or_None, kind)  下基准线
    kind ∈ 'corner' / 'hline' / 'ftop' / 'ftop_adj'。
    任一项可能为 None（该端无页眉/页脚）。
    """
    s = r.get('hf_s')
    bands = r.get('hf_bands') or []
    if not s or not bands:
        return {'top': None, 'bot': None}
    rH = r.get('h') or H
    rW = r.get('w') or W
    hb = r.get('hf_header')
    fb = r.get('hf_footer')
    side = r.get('hf_side') or ('L' if (r.get('page') or 1) % 2 else 'R')
    xgate = HF_PC_XGATE * rW
    out = {'top': None, 'bot': None}

    # ---- 上基准线：页眉下方 ----
    if hb and hb.get('box'):
        hy_b = hb['box'][3] / s                 # 页眉框底边（绝对像素）
        # 页眉朝页面外侧那条边：奇页(L)取左边，偶页(R)取右边
        hx_o = (hb['box'][0] if side == 'L' else hb['box'][2]) / s
        pts, hls = [], []
        for nm, y0, segs in bands:
            if nm != 'near_header':
                continue
            pts += hf_band_corners(segs, y0, s, y_lo=hy_b,
                                   y_hi=hy_b + HF_CORNER_SPAN * rH)
            hls += hf_band_hlines(segs, y0, s, y_lo=hy_b,
                                  y_hi=hy_b + HF_CORNER_SPAN * rH)
        best, bd = None, float('inf')
        for (cx, cy, st) in pts:
            if cy < hy_b or abs(cx - hx_o) > xgate:
                continue
            d = abs(cx - hx_o) + abs(cy - hy_b) * HF_PC_YW
            if d < bd:
                bd, best = d, (cy, cx, 'corner')
        if best is not None:
            out['top'] = best
        else:
            # 降级：只用横线定 y；x 返回 None = 该轴不动（§4.7.6）
            cand = [t for t in hls if t[2] >= hy_b]
            if cand:
                cand.sort(key=lambda t: (abs(t[2] - hy_b), -t[3]))
                out['top'] = (cand[0][2], None, 'hline')

    # ---- 下基准线：页脚上方 ----
    if fb and fb.get('box'):
        fy_t = rH * (1.0 - HF_BOTR) + fb['box'][1] / s   # 页脚框顶边（绝对）
        # 页脚朝页面外侧那条边：奇页(L)取左，偶页(R)取右
        fx_o = (fb['box'][0] if side == 'L' else fb['box'][2]) / s
        # 搜索窗必须紧邻页脚（§4.7.6）
        y_lo = fy_t - HF_PC_NEAR * rH
        y_hi = fy_t + HF_PC_BELOW * rH
        pts, hls = [], []
        for nm, y0, segs in bands:
            if nm != 'near_footer':
                continue
            pts += hf_band_corners(segs, y0, s, y_lo=y_lo, y_hi=y_hi)
            hls += hf_band_hlines(segs, y0, s, y_lo=y_lo, y_hi=y_hi)
        ys = [(abs(t[2] - fy_t), -t[3], t[2]) for t in hls if y_lo <= t[2] <= y_hi]
        for (cx, cy, st) in pts:
            if y_lo <= cy <= y_hi:
                ys.append((abs(cy - fy_t), -st, cy))
        if ys:
            ys.sort(key=lambda t: (t[0], t[1]))
            by, kind = ys[0][2], 'hline'
        else:
            # 窗内无横线 → 用页脚框顶边兜底，卷级再统一补偏移（§4.7.6）
            by, kind = fy_t, 'ftop'
        # x：fx_o 本身即合法答案，交点只作精修，门控用 HF_PC_FXGATE
        fxgate = HF_PC_FXGATE * rW
        bx, bdx = None, float('inf')
        for (cx, cy, st) in pts:
            if abs(cy - by) > max(4.0, HF_CORNER_TOL * rH):
                continue
            dx = abs(cx - fx_o)
            if dx <= fxgate and dx < bdx:
                bdx, bx = dx, cx
        if bx is not None:
            out['bot'] = (by, bx, 'corner' if kind != 'ftop' else 'ftop')
        else:
            out['bot'] = (by, fx_o, kind)
    return out


def hf_fix_ftop_baselines(records, H):
    """卷级补偿：把 kind='ftop' 的下基准线从「页脚顶边」推回「横线位置」。

    用有线页统计 off = 线y − 页脚顶边 的中位数，补到无线页上，使两批页落在
    同一条物理线上。原理见方案 §4.7.6。

    就地改写 r['hf_base']['bot']，返回 (n_fix, off_med)。
    """
    offs = []
    for r in records:
        b = (r.get('hf_base') or {}).get('bot')
        fb = r.get('hf_footer')
        s = r.get('hf_s')
        if not b or not fb or not fb.get('box') or not s:
            continue
        if b[2] == 'ftop':
            continue
        rH = r.get('h') or H
        fy_t = rH * (1.0 - HF_BOTR) + fb['box'][1] / s
        offs.append(b[0] - fy_t)
    off = float(np.median(offs)) if len(offs) >= HF_PC_OFF_MIN_N else 0.0
    n = 0
    for r in records:
        base = r.get('hf_base') or {}
        b = base.get('bot')
        if not b or b[2] != 'ftop':
            continue
        base['bot'] = (b[0] + off, b[1], 'ftop_adj' if off else 'ftop')
        n += 1
    return n, off


def hf_calib_consensus_height(cons, records):
    """校准共识框高度：以「同页同时有页眉和页脚」的页实测净跨距为准。

    规则见方案 §4.7.4：共识框上下边必须落在页眉底边与页脚顶边【之内】，
    不得与页眉页脚重叠。同页跨距是唯一权威依据；同页样本不足时才降级到
    全书两端中位相减，并同样受净跨距上限约束。

    返回 (新共识, info)。未采信时 info['used'] = False，共识原样返回。
    """
    W = max(1, int(np.median([r['w'] for r in records if r.get('w')])))
    H = max(1, int(np.median([r['h'] for r in records if r.get('h')])))
    ch0 = cons[3] - cons[1]
    box = {}
    spans = []
    for r in records:
        bl = r.get('hf_base') or hf_page_baselines(r, W, H)
        t, b = bl.get('top'), bl.get('bot')
        if not t or not b:
            continue                      # 只统计「页眉页脚都有」的页
        rH = r.get('h') or H
        d = b[0] - t[0]
        if not (HF_PH_RANGE[0] * rH <= d <= HF_PH_RANGE[1] * rH):
            continue
        box.setdefault(int(round(d / HF_PH_BIN)), set()).add(r.get('name'))
        spans.append((t[0], b[0], d))
    info = {'used': False, 'n_both': len(spans), 'h0': round(ch0, 1)}
    # 净跨距上限：同页实测跨距的中位。共识高不得超过它，否则必然与页眉
    # 或页脚重叠（方案 §4.7.4）。同页样本为 0 时无上限可用。
    cap = float(np.median([d for _, _, d in spans])) if spans else None
    if cap is not None:
        info['span_med'] = round(cap, 1)

    def _fallback():
        """降级：全书上基准中位 − 全书下基准中位。

        两端样本来自不同批页面，属于混批统计，可能超出真实净跨距，
        因此结果须再受同页净跨距上限裁剪（见方案 §4.7.4）。
        """
        tys = [r['hf_base']['top'][0] for r in records
               if (r.get('hf_base') or {}).get('top')]
        bys = [r['hf_base']['bot'][0] for r in records
               if (r.get('hf_base') or {}).get('bot')]
        info['n_top'], info['n_bot'] = len(tys), len(bys)
        if len(tys) < HF_PH_MIN_N or len(bys) < HF_PH_MIN_N:
            info.setdefault('reason', 'sides_too_few')
            return cons, info
        tm = float(np.median(tys))
        bm = float(np.median(bys))
        hm = bm - tm
        if cap is not None and hm > cap:
            hm = cap
            info['capped'] = True
            info['cap'] = round(cap, 1)
            bm = tm + hm
        info.update(h=round(hm, 1), y1=round(tm, 1), y2=round(bm, 1),
                    mode='volume_median')
        if not (HF_PH_RANGE[0] * H <= hm <= HF_PH_RANGE[1] * H):
            info['reason'] = 'fallback_out_of_range'
            return cons, info
        info['used'] = True
        info['dh'] = round(hm - ch0, 1)
        return [cons[0], tm, cons[2], tm + hm], info

    if not box:
        info['reason'] = 'no_page_with_both'
        return _fallback()
    best, bn = None, -1
    for k in box:
        tot = set()
        for dk in (-1, 0, 1):
            tot |= box.get(k + dk, set())
        if len(tot) > bn:
            bn, best = len(tot), k
    info['peak_h'] = round(best * HF_PH_BIN, 1)
    info['peak_pages'] = bn
    if bn < HF_PH_MIN_N:
        # 峰不够强时不再整段丢弃：同页跨距仍是唯一直接证据，改用全部同页
        # 样本的中位（受 HF_PH_RANGE 过滤保护），而不是退回混批统计。
        info['weak_peak'] = True
        sel = list(spans)
    else:
        sel = [(a, b, d) for a, b, d in spans
               if abs(d - best * HF_PH_BIN) <= 1.5 * HF_PH_BIN] or list(spans)
    hmed = float(np.median([d for _, _, d in sel]))
    tmed = float(np.median([a for a, _, _ in sel]))
    bmed = float(np.median([b for _, b, _ in sel]))
    info.update(h=round(hmed, 1), y1=round(tmed, 1), y2=round(bmed, 1),
                n_sel=len(sel))
    if not (HF_PH_RANGE[0] * H <= hmed <= HF_PH_RANGE[1] * H):
        info['reason'] = 'span_out_of_range'
        return _fallback()
    info['mode'] = 'same_page_span'
    out = [cons[0], tmed, cons[2], tmed + hmed]
    info['used'] = True
    info['dh'] = round(hmed - ch0, 1)
    return out, info


def hf_refine_consensus_corner(cons, records):
    """用「页眉下方 / 页脚上方的框线交点（角）」定位共识框 —— 只改位置，不改尺寸。

    卷级统一定位，作为逐页角点定位的基线。规则见方案 §4.7.2。

    返回 (新共识, info)。info['mode'] ∈
      'header_corner' / 'footer_corner_L' / 'footer_corner_R' / 'none'。
    """
    W = max(1, int(np.median([r['w'] for r in records if r.get('w')])))
    H = max(1, int(np.median([r['h'] for r in records if r.get('h')])))
    cw = cons[2] - cons[0]
    ch = cons[3] - cons[1]
    tolx = max(4.0, HF_CORNER_TOL * W)
    toly = max(4.0, HF_CORNER_TOL * H)
    capx = HF_CORNER_CAP * W
    capy = HF_CORNER_CAP * H
    n_pages = max(1, len(records))
    need = max(HF_CORNER_MIN_N, int(HF_CORNER_RATIO * n_pages))

    def _header_bottom(r):
        hb = r.get('hf_header')
        if not hb or not hb.get('box'):
            return None
        s = r.get('hf_s', HF_IW / float(W))
        return hb['box'][3] / s

    def _footer_top(r):
        fb = r.get('hf_footer')
        if not fb or not fb.get('box'):
            return None
        s = r.get('hf_s', HF_IW / float(W))
        return H * (1.0 - HF_BOTR) + fb['box'][1] / s

    # 全书统一基准：页眉底边 / 页脚顶边的中位数。
    # 关键——不能按页各自为政（页眉常只命中奇页，未命中页若放开下界会把
    # 候选拖到页面最顶端，统计被严重污染，框反而往上跑）。所有页共用同一条
    # 下界 / 上界，才谈得上「全书一套框」的对齐。
    hbs = [v for v in (_header_bottom(r) for r in records) if v is not None]
    fts = [v for v in (_footer_top(r) for r in records) if v is not None]
    hb_med = float(np.median(hbs)) if hbs else None
    ft_med = float(np.median(fts)) if fts else None

    def _pick_stable(vals, target, tol, mode='nearest'):
        """在一维候选里取稳定簇中位数。
        mode: nearest = 离 target 最近；low = 最小（最靠上）；high = 最大（最靠下）。
        """
        if len(vals) < need:
            return None, 0
        clusters = _cluster_sorted(vals, tol)
        sig = [c for c in clusters if len(c) >= need]
        if not sig:
            return None, 0
        meds = [float(np.median(c)) for c in sig]
        if mode == 'low':
            i = int(np.argmin(meds))
        elif mode == 'high':
            i = int(np.argmax(meds))
        else:
            i = int(np.argmin([abs(m - target) for m in meds]))
        return meds[i], len(sig[i])

    # ---- 收集候选角点 ----
    tl_pts, bt_pts = [], []     # 左上角候选 / 下角候选
    for r in records:
        bands = r.get('hf_bands') or []
        s = r.get('hf_s', HF_IW / float(W))
        rH = r.get('h') or H
        for nm, y0, segs in bands:
            if nm == 'near_header' and hb_med is not None:
                # 页眉下方：下界 = 全书页眉底边中位数，向上再放 SPAN 高的搜索窗
                tl_pts += hf_band_corners(segs, y0, s,
                                          y_lo=hb_med,
                                          y_hi=hb_med + HF_CORNER_SPAN * rH)
            elif nm == 'near_footer' and ft_med is not None:
                # 页脚上方：上界 = 全书页脚顶边中位数，向下再放 SPAN 高的搜索窗
                bt_pts += hf_band_corners(segs, y0, s,
                                          y_lo=ft_med - HF_CORNER_SPAN * rH,
                                          y_hi=ft_med)

    info = {'mode': 'none', 'tl': None, 'bt': None, 'side': None,
            'need': need, 'n_tl': len(tl_pts), 'n_bt': len(bt_pts),
            'hb_med': round(hb_med, 1) if hb_med is not None else None,
            'ft_med': round(ft_med, 1) if ft_med is not None else None}

    # ---- ① 页眉下方角点 → 左上角 ----
    tl = None
    if tl_pts:
        # y 取页眉下方最近的一簇（框顶边紧贴页眉下方第一条框线）
        my, ny = _pick_stable([p[1] for p in tl_pts], cons[1], toly, 'low')
        # x 取离原共识左边界最近的稳定簇（框左边压在左边界垂线上）
        mx, nx = _pick_stable([p[0] for p in tl_pts], cons[0], tolx, 'nearest')
        if mx is not None and my is not None:
            tl = (mx, my)
            info['tl'] = [round(mx, 1), round(my, 1), nx, ny]

    # ---- ② 页脚上方角点 → 左下角 / 右下角 ----
    bt = None
    if bt_pts:
        # y 取页脚上方最近的一簇（框下边紧贴页码上方那条框线）
        my, ny = _pick_stable([p[1] for p in bt_pts], cons[3], toly, 'high')
        xs = [p[0] for p in bt_pts]
        # 左右判定：候选 x 的整体中位离原框哪条竖边更近
        med_x = float(np.median(xs))
        side = 'L' if abs(med_x - cons[0]) <= abs(med_x - cons[2]) else 'R'
        mx, nx = _pick_stable(xs, cons[0] if side == 'L' else cons[2],
                              tolx, 'nearest')
        if my is not None:
            bx = mx if mx is not None else med_x
            bt = (bx, my, side)
            info['bt'] = [round(bx, 1), round(my, 1), nx, ny]
            info['side'] = side

    # ---- ③ 定位：尺寸严格保持 (cw, ch)，只平移 (x1, y1) ----
    # 注意：不做「画面居中」fallback——两个角都找不到时保持原位置不动。
    out = list(cons)
    if tl is not None:
        dx = max(-capx, min(capx, tl[0] - cons[0]))
        dy = max(-capy, min(capy, tl[1] - cons[1]))
        out = [cons[0] + dx, cons[1] + dy,
               cons[0] + dx + cw, cons[1] + dy + ch]
        info['mode'] = 'header_corner'
        info['shift'] = [round(dx, 1), round(dy, 1)]
    elif bt is not None:
        bx, by, side = bt
        # 下角对齐：左下角 → x1 = bx；右下角 → x2 = bx ⇒ x1 = bx - cw
        nx1 = bx if side == 'L' else bx - cw
        ny1 = by - ch
        dx = max(-capx, min(capx, nx1 - cons[0]))
        dy = max(-capy, min(capy, ny1 - cons[1]))
        out = [cons[0] + dx, cons[1] + dy,
               cons[0] + dx + cw, cons[1] + dy + ch]
        info['mode'] = 'footer_corner_' + side
        info['shift'] = [round(dx, 1), round(dy, 1)]
    return out, info


def hf_page_corner_target(r, cw, ch, W, H):
    """【逐页】定位：返回本页框的左上角 (x1, y1)；None = 本页不参与角点定位。

    尺寸 (cw, ch) 由全书统一给定且严格不变，本函数只定位置。
    对齐优先级：有页眉 / 页脚时必须以其基准线为准，都没有才回退贴分镜框边；
    x 的左右归属由页号奇偶（hf_side）决定。完整规则见方案 §4.7.2 / §4.7.6。

    返回 (x1, y1) 里任一轴可为 None，含义是【该轴保持调用方的贴边结果】。
    """
    rW = r.get('w') or W
    side = r.get('hf_side') or ('L' if (r.get('page') or 1) % 2 else 'R')
    bl = r.get('hf_base')
    if bl is None:
        bl = hf_page_baselines(r, W, H)
    top, bot = bl.get('top'), bl.get('bot')
    if not top and not bot:
        return None, None

    # ---- y1：两端都有 → 取两估计的均值；分歧过大时以【页脚】为准 ----
    rH = r.get('h') or H
    if top and bot:
        t_est, b_est = top[0], bot[0] - ch
        if abs(t_est - b_est) <= HF_PC_BOTH_TOL * rH:
            ny1, ykind = 0.5 * (t_est + b_est), 'both'
        else:
            ny1, ykind = b_est, 'footer'
    elif top:
        ny1, ykind = top[0], 'header'
    else:
        ny1, ykind = bot[0] - ch, 'footer'

    # ---- x1：取外侧竖线证据，页脚优先（页脚在页面角上，更贴版心外缘）----
    xs = None
    xkind = ''
    if bot and bot[1] is not None:
        xs, xkind = bot[1], 'F'
    elif top and top[1] is not None:
        xs, xkind = top[1], 'H'

    nx1 = None
    if xs is not None:
        # 外侧 = 离书脊一侧。奇页(L) → 该竖线就是框左边；
        # 偶页(R) → 该竖线是框右边 ⇒ x1 = x − 框宽。
        nx1 = xs if side == 'L' else xs - cw

    kind = ykind + ('_' + side + xkind if nx1 is not None else '_hl')
    return (nx1, ny1), kind


# ---------------- 主流程 ----------------

def extract_page_num(name):
    m = PAGE_NUM_RE.search(name)
    if m:
        return int(m.group(1))
    m = re.search(r'(\d+)', os.path.splitext(name)[0])
    return int(m.group(1)) if m else 0


def clean_path(p):
    """清洗命令行/对话框路径：去首尾引号、空白、尾部反斜杠（避免 \" 转义引号问题）。"""
    p = p.strip().strip('"').strip("'").strip()
    if len(p) > 3 and p[-1] in '\\/':
        p = p.rstrip('\\/')
    return p


def choose_dir_dialog(initial=None):
    """弹出目录选择对话框，返回目录路径字符串；取消或失败返回 None。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        try:
            root.attributes('-topmost', True)
        except Exception:
            pass
        d = filedialog.askdirectory(title='选择扫描图片所在的文件夹',
                                    initialdir=initial or None)
        root.destroy()
        return d if d else None
    except Exception as e:
        log('无法弹出目录选择对话框（tkinter 不可用？）: %s' % e)
        return None


def pause_if_gui():
    """双击运行时窗口不会自动关闭，等待用户回车。"""
    try:
        input('\n已完成，按回车键退出...')
    except Exception:
        pass


def gui_main_dialog():
    """Tkinter 简易界面：模式 A/B + 路径选择 + 调试图选项 + 确定。
    用户点确定后，后台线程启动检测，前台滚动日志；
    完成后弹「完成」对话框，窗口转为「关闭」按钮供用户查看日志再退。

    返回值：
      None       —— 用户取消
      int 退出码 —— 检测完成（0 成功 / 非 0 出错）
    """
    try:
        import tkinter as tk
        from tkinter import filedialog, scrolledtext, messagebox
        import threading
    except Exception as e:
        log('无法启动 GUI（tkinter 不可用？）：%s' % e)
        return None

    version = '1.2'
    root = tk.Tk()
    root.title('扫描版心检测 %s' % version)
    root.geometry('820x620')
    try:
        root.attributes('-topmost', True)
    except Exception:
        pass

    # === 检测模式 ===
    frm_mode = tk.LabelFrame(root, text='检测模式', padx=10, pady=5)
    frm_mode.pack(fill='x', padx=10, pady=(10, 5))
    mode_var = tk.StringVar(value='A')
    tk.Radiobutton(
        frm_mode,
        text='模式 A：边框版心检测',
        variable=mode_var, value='A').pack(anchor='w')
    tk.Radiobutton(
        frm_mode,
        text='模式 B：启用页眉/页脚检测版心',
        variable=mode_var, value='B').pack(anchor='w')

    # === 路径 ===
    frm_path = tk.LabelFrame(root, text='扫描图片所在文件夹', padx=10, pady=5)
    frm_path.pack(fill='x', padx=10, pady=5)
    path_var = tk.StringVar()
    tk.Entry(frm_path, textvariable=path_var).pack(
        side='left', fill='x', expand=True, padx=(0, 5))

    def browse():
        d = filedialog.askdirectory(
            title='选择扫描图片所在的文件夹',
            initialdir=path_var.get() or None,
            parent=root)
        if d:
            path_var.set(d)

    tk.Button(frm_path, text='浏览…', command=browse).pack(side='left')

    # === 调试选项 ===
    debug_var = tk.BooleanVar(value=True)
    tk.Checkbutton(
        root,
        text='生成调试叠加图（hanmen_debug/）',
        variable=debug_var).pack(anchor='w', padx=15, pady=(0, 5))

    # === 日志面板 ===
    frm_log = tk.LabelFrame(root, text='运行日志', padx=5, pady=5)
    frm_log.pack(fill='both', expand=True, padx=10, pady=5)
    # 黑体（SimHei / 「黑体」）优先；其次微软雅黑；最后跟随系统默认。
    # 用 tkFont.families() 列表比对（f.actual('family') 在别名场景下
    # 会返回中文名 "黑体" 而非 "SimHei"，不可靠）。
    import tkinter.font as tkFont
    try:
        default_font = tkFont.nametofont('TkDefaultFont')
    except Exception:
        default_font = None
    available = set(tkFont.families())
    chosen_font = default_font
    for fam in ('SimHei', '黑体', 'Microsoft YaHei', '微软雅黑'):
        if fam in available:
            chosen_font = tkFont.Font(family=fam, size=10)
            break
    log_text = scrolledtext.ScrolledText(
        frm_log, height=15, state='disabled', wrap='word',
        font=chosen_font)
    log_text.pack(fill='both', expand=True)

    # === 按钮 ===
    frm_btn = tk.Frame(root)
    frm_btn.pack(fill='x', padx=10, pady=(0, 10))
    btn_cancel = tk.Button(frm_btn, text='取消', width=10)
    btn_cancel.pack(side='right', padx=5)
    btn_ok = tk.Button(frm_btn, text='确定', width=10)
    btn_ok.pack(side='right')

    # runs = 已完成的运行次数（>0 表示至少跑过一次，退出时按成功返回）
    state = {'running': False, 'done': False, 'exit_code': 0,
             'closed': False, 'runs': 0}

    def append_log(msg):
        if state['closed']:
            return
        log_text.after(0, _append_now, msg)

    def _append_now(msg):
        log_text.config(state='normal')
        log_text.insert('end', msg + '\n')
        log_text.see('end')
        log_text.config(state='disabled')

    def run_worker(img_dir, hf_on, make_debug):
        """后台线程：跑检测并把 log() 转发到 GUI。"""
        state['running'] = True
        # 替换 stdout，把 print 抓到 GUI
        import io
        gui_buf = io.StringIO()
        real_stdout = sys.stdout

        class _StdoutProxy:
            def __init__(self, buf, real):
                self._buf = buf
                self._real = real

            def write(self, s):
                # 把 s 按 \n 切成完整行 → 发到 GUI；最后一段（无换行结尾）
                # 留到缓冲区，由 flush() 时再发。每行独立 try，单行失败不影响其他。
                if not s:
                    return 0
                lines = s.split('\n')
                rest = lines[-1]
                for line in lines[:-1]:
                    try:
                        append_log(line)
                    except Exception:
                        pass
                if rest:
                    self._buf.write(rest)
                return len(s)

            def flush(self):
                # 把缓冲区里的残留行（无换行结尾）补一个换行发给 GUI
                rest = self._buf.getvalue()
                self._buf.seek(0)
                self._buf.truncate(0)
                if rest:
                    try:
                        append_log(rest)
                    except Exception:
                        pass
                try:
                    self._real.flush()
                except Exception:
                    pass

        sys.stdout = _StdoutProxy(gui_buf, real_stdout)
        try:
            _run_detection(img_dir, hf_on, make_debug, gui_mode=True,
                           tk_mode=True)
            state['exit_code'] = 0
        except SystemExit as e:
            state['exit_code'] = int(e.code) if isinstance(e.code, int) else 1
        except Exception as e:
            append_log('!!! 出错: %s' % e)
            import traceback
            append_log(traceback.format_exc())
            state['exit_code'] = 1
        finally:
            sys.stdout = real_stdout
            state['running'] = False
            state['done'] = True
            state['runs'] += 1
            if not state['closed']:
                root.after(0, _on_done)

    def _on_done():
        # 恢复成「可再次运行」状态：确定 = 再次检测，取消 = 关闭窗口。
        # 旧版把确定改成「关闭」并禁用取消，导致跑完一次后无法二次生成。
        state['done'] = False
        btn_ok.config(text='确定', state='normal')
        btn_cancel.config(text='关闭', state='normal')
        _append_now('—— 本次检测结束，可修改参数后再次点「确定」运行 ——')
        messagebox.showinfo(
            '完成',
            '检测完成（退出码 %d）。'% state['exit_code'],
            parent=root)

    def on_ok():
        if state['running']:
            messagebox.showwarning('提示', '正在运行，请稍候。', parent=root)
            return
        d = path_var.get().strip()
        if not d or not os.path.isdir(d):
            messagebox.showerror(
                '路径无效',
                '请选择一个有效的文件夹（包含扫描图片）。',
                parent=root)
            return
        # 第二次及以后运行：先加一条分隔线，日志不清空便于对比两次结果
        if state['runs'] > 0:
            _append_now('\n' + '=' * 60)
            _append_now('第 %d 次运行' % (state['runs'] + 1))
        btn_ok.config(state='disabled')
        btn_cancel.config(state='disabled')
        threading.Thread(
            target=run_worker,
            args=(d, mode_var.get() == 'B', debug_var.get()),
            daemon=True).start()

    def on_close():
        if state['running']:
            if not messagebox.askyesno(
                    '检测进行中',
                    '正在运行检测，确定要取消并关闭窗口吗？',
                    parent=root):
                return
        state['closed'] = True
        root.destroy()

    btn_ok.config(command=on_ok)
    btn_cancel.config(command=on_close)
    root.protocol('WM_DELETE_WINDOW', on_close)

    root.mainloop()
    # 判据用 runs（已完成的运行次数）而非 done —— done 在 _on_done 里被复位成
    # False 以便二次运行，若沿用它会把「已成功跑过」误判成「用户取消」。
    if state['runs'] == 0:
        return None  # 用户一次都没跑就关掉 = 取消
    return state['exit_code']


def main():
    """入口：解析命令行参数 / 启动 GUI，调度 _run_detection。"""
    args = [a for a in sys.argv[1:]]
    pos = [a for a in args if not a.startswith('--')]
    gui_mode = len(pos) < 1     # 双击 .py 或无参启动 → 弹 GUI
    make_debug = '--no-debug' not in args
    hf_on = ('--hf' in args) or ('--header-footer' in args)

    if gui_mode:
        # 优先 GUI；tkinter 缺失时回退到目录对话框
        rc = gui_main_dialog()
        if rc is None:
            try:
                import tkinter  # noqa: F401
            except ImportError:
                log('GUI 不可用（缺 tkinter），回退到目录对话框模式。')
                chosen = choose_dir_dialog()
                if not chosen:
                    log('未选择有效目录，已取消。')
                    sys.exit(1)
                img_dir = os.path.abspath(clean_path(chosen))
                _run_detection(img_dir, hf_on, make_debug, gui_mode=True)
                sys.exit(0)
            log('已取消。')
            sys.exit(0)
        sys.exit(rc)
        return

    # CLI 模式：必须给路径
    if not pos:
        log('命令行模式下请提供图片目录。')
        sys.exit(1)
    cand = clean_path(pos[0])
    if os.path.isdir(cand):
        img_dir = os.path.abspath(cand)
    elif os.path.isfile(cand):
        img_dir = os.path.dirname(os.path.abspath(cand))
    else:
        log('命令行给定的路径无效: ' + cand)
        sys.exit(1)

    _run_detection(img_dir, hf_on, make_debug, gui_mode=False)


def _run_detection(img_dir, hf_on, make_debug, gui_mode, tk_mode=False):
    """执行版心检测 + (可选) 页眉/页脚检测，落盘 hanmen.json。
    拆分自 main() 以便 GUI / CLI 复用同一套逻辑。

    tk_mode=True 表示由 Tkinter 后台线程调用，此时不得调 pause_if_gui()
    （input() 会在无控制台的窗口进程里永久阻塞）。
    """
    files = sorted(
        f for f in os.listdir(img_dir)
        if f.lower().endswith(EXTS) and os.path.isfile(os.path.join(img_dir, f))
    )
    if not files:
        log('目录中没有找到支持的图片文件: ' + img_dir)
        if gui_mode and not tk_mode:
            pause_if_gui()
        sys.exit(1)

    debug_dir = os.path.join(img_dir, 'hanmen_debug')
    if make_debug:
        os.makedirs(debug_dir, exist_ok=True)

    log('共发现 %d 张图片，开始检测...%s'
        % (len(files), '（已启用页眉/页脚检测）' if hf_on else ''))

    records = []   # 每项: dict(..., edges_abs=[x1,y1,x2,y2]全分辨率像素 或 None, ...)
    for idx, name in enumerate(files):
        path = os.path.join(img_dir, name)
        try:
            gray_full, dpi, W, H = load_image(path)
        except Exception as e:
            log('[%d/%d] 读取失败 %s: %s' % (idx + 1, len(files), name, e))
            records.append({'name': name, 'page': extract_page_num(name),
                            'w': 0, 'h': 0, 'dpi': 1200, 'edges_abs': None,
                            'angle': 0.0, 'ink': 0.0, 'clusters': None,
                            'small': None, 'status_raw': 'manual',
                            'reason': 'read_error'})
            continue

        scale = WORK_LONG / float(max(H, W))
        small = cv2.resize(gray_full, None, fx=scale, fy=scale,
                           interpolation=cv2.INTER_AREA)
        edges, angle, ink, clusters = detect_edges(small)

        # 空白页 / 封面（横图）直接 manual
        status = None
        reason = ''
        if W >= H:
            status, reason = 'manual', 'landscape_or_cover'
        elif ink < INK_RATIO_MIN:
            status, reason = 'manual', 'blank_page'
        elif edges is None:
            status, reason = None, 'edges_missing'   # 待共识阶段决定
        else:
            status, reason = 'detected', ''

        edges_abs = None
        if edges is not None:
            # 小图坐标 → 全分辨率像素
            edges_abs = [int(round(edges['x1'] / scale)),
                         int(round(edges['y1'] / scale)),
                         int(round(edges['x2'] / scale)),
                         int(round(edges['y2'] / scale))]
        hc0, vc0 = clusters
        hc_abs = [(y / scale, s) for y, s in hc0]
        vc_abs = [(x / scale, s) for x, s in vc0]

        rec = {'name': name, 'page': extract_page_num(name),
               'w': W, 'h': H, 'dpi': dpi, 'edges_abs': edges_abs,
               'angle': angle, 'ink': ink, 'clusters': clusters,
               'hc_abs': hc_abs, 'vc_abs': vc_abs,
               'small': small, 'status_raw': status, 'reason': reason,
               'hf_ok': False}

        # ---- 页眉 / 页脚 / 分镜框线（可选）----
        if hf_on and status != 'manual' and W and H and W < H:
            try:
                s = HF_IW / float(W)
                sH = H * s
                rec['hf_sH'] = sH
                rec['hf_s'] = s
                rec['hf_side'] = 'L' if rec['page'] % 2 else 'R'
                nd = max(1, len(str(rec['page'])))
                # 页眉：顶部墨迹带（保留到卷级模板匹配结束再释放）
                sm, _ = _hf_band((gray_full < HF_INK_T).astype(np.uint8) * 255,
                                 0, int(H * HF_TOPR), W, s)
                rec['hf_sm'] = sm
                rec['hf_seed'] = hf_header_seed(sm, sH)
                # 页脚：底部灰度带 → 候选（含字形切片）+ 横线
                gm_bot, bH = _hf_band(gray_full, int(H * (1 - HF_BOTR)), H, W, s)
                rec['hf_cands'] = hf_footer_cands(gm_bot, sH, bH,
                                                  rec['hf_side'], nd)
                rec['hf_hlines'] = hf_footer_hlines(gm_bot)
                # 分镜框线：页眉带 / 页脚带（竖线用于共识框贴边）
                bands = hf_page_bands(gray_full, W, H, s, sH)
                rec['hf_vlines'] = hf_merge_vlines(hf_band_vlines(bands, s), W)
                rec['hf_bands'] = bands
                rec['hf_ok'] = True
            except Exception as e:
                log('  页眉页脚检测失败 %s: %s' % (name, e))
                rec.setdefault('hf_cands', [])
                rec.setdefault('hf_hlines', [])
                rec.setdefault('hf_vlines', [])
                rec['hf_ok'] = False

        records.append(rec)
        log('[%d/%d] %s  ink=%.3f angle=%.2f  %s'
            % (idx + 1, len(files), name, ink, angle,
               status if status else 'pending'))

    # ---- 跨页共识（绝对像素；同一台扫描仪同一批扫描，物理版心像素位置一致）----
    valid = [r for r in records
             if r['edges_abs'] is not None and r['status_raw'] != 'manual']
    if not valid:
        log('警告：没有任何页面检测到版心四边，全部标记为 manual。')
        consensus = None
    else:
        arr = np.array([r['edges_abs'] for r in valid], dtype=np.float64)
        consensus = [float(np.median(arr[:, i])) for i in range(4)]
        std = [float(np.std(arr[:, i])) for i in range(4)]
        log('共识版心(绝对像素): x1=%.0f y1=%.0f x2=%.0f y2=%.0f  (基于 %d 页, 标准差 x1=%.0f y1=%.0f x2=%.0f y2=%.0f)'
            % (consensus[0], consensus[1], consensus[2], consensus[3],
               len(valid), std[0], std[1], std[2], std[3]))

    # ---- 卷级页眉 / 页脚统计 + 共识框竖线贴边（可选）----
    hf_info = None
    if hf_on:
        hf_recs = [r for r in records
                   if r.get('hf_ok') and r['status_raw'] != 'manual']
        if consensus is None or not hf_recs:
            log('页眉页脚检测：可用页面不足，跳过。')
            hf_info = {'enabled': False, 'reason': 'no_valid_page'}
        else:
            # ⓪ 宽度校准：成对竖线间距众数定宽，必须在 x 修正之前（§4.7.1）
            consensus, pw = hf_calib_consensus_width(consensus, hf_recs)
            if pw.get('used'):
                log('共识宽校准: %.0f → %.0f (%+.0f px)  峰宽=%.0f 支持 %d 页 / 对数 %d'
                    % (pw['w0'], pw['w'], pw['dw'], pw['peak_w'],
                       pw['peak_pages'], pw['n_sel']))
                log('  配对左右边: x1=%.0f x2=%.0f' % (pw['x1'], pw['x2']))
            else:
                log('共识宽校准: 未采信（%s，峰宽=%s 支持 %s 页，原宽 %.0f）'
                    % (pw.get('reason'), pw.get('peak_w'),
                       pw.get('peak_pages'), pw['w0']))

            # ① 共识框左右边紧贴页眉/页脚带的垂直边线（全书仍共用同一个框）
            base_x = [consensus[0], consensus[2]]
            consensus, moved_x, cnt_x = hf_refine_consensus_x(
                consensus, hf_recs, lock_width=bool(pw.get('used')))
            log('共识竖线修正: x1 %+.0f px (样本 %d)  x2 %+.0f px (样本 %d)  有效页 %d'
                % (moved_x[0], cnt_x[0], moved_x[1], cnt_x[1], cnt_x[2]))
            log('修正后共识: x1=%.0f x2=%.0f  宽=%.0f (原 %.0f)'
                % (consensus[0], consensus[2],
                   consensus[2] - consensus[0], base_x[1] - base_x[0]))

            # ② 页眉 / 页脚检测（必须先做，y 修正需要它们的命中结果）
            log('页眉检测中...')
            hdr = hf_header_volume(hf_recs)
            log('  种子 %d 页，命中 %s 页（奇 %s / 偶 %s）'
                % (hdr.get('n_seed', 0), hdr.get('hits', 0),
                   hdr.get('odd', 0), hdr.get('even', 0)))

            log('页脚检测中...')
            ftr = hf_footer_volume(hf_recs)
            log('  范式 %s，命中 %d 页（A 可用=%s 离散=%.3f 峰票比=%.2f）'
                % (ftr.get('paradigm'), ftr.get('hits', 0),
                   ftr.get('a_usable'), ftr.get('a_spread', 0.0),
                   ftr.get('a_ratio', 0.0)))

            # ③ 逐页基准线（页眉下方 / 页脚上方）—— 高度校准与逐页对齐共用同一份证据
            _bW = max(1, int(np.median([r['w'] for r in hf_recs if r.get('w')])))
            _bH = max(1, int(np.median([r['h'] for r in hf_recs if r.get('h')])))
            n_both = 0
            for r in hf_recs:
                r['hf_base'] = hf_page_baselines(r, _bW, _bH)
                if r['hf_base'].get('top') and r['hf_base'].get('bot'):
                    n_both += 1
            log('逐页基准线: 有上基准 %d 页，有下基准 %d 页，两者都有 %d 页'
                % (sum(1 for r in hf_recs if r['hf_base'].get('top')),
                   sum(1 for r in hf_recs if r['hf_base'].get('bot')), n_both))
            n_fix, off_med = hf_fix_ftop_baselines(hf_recs, _bH)
            if n_fix:
                log('  页脚顶边兜底 %d 页，用有线页中位偏移 %+.0f px 补齐'
                    % (n_fix, off_med))

            # ④ 高度校准：以同页净跨距为准，必须在 y 修正之前（§4.7.4）
            base_h0 = consensus[3] - consensus[1]
            consensus, ph = hf_calib_consensus_height(consensus, hf_recs)
            if ph.get('used'):
                log('共识高校准: %.0f → %.0f (%+.0f px)  依据=%s 峰高=%s 支持 %s 页'
                    % (ph['h0'], ph['h'], ph['dh'],
                       ph.get('mode', 'same_page_span'),
                       ph.get('peak_h'), ph.get('n_sel', ph.get('n_top'))))
                log('  配对上下边: y1=%.0f y2=%.0f  (双端页 %d，上基准 %s 页，下基准 %s 页)'
                    % (ph['y1'], ph['y2'], ph.get('n_both', 0),
                       ph.get('n_top', '-'), ph.get('n_bot', '-')))
            else:
                log('共识高校准: 未采信（%s，峰高=%s 支持 %s 页，双端页 %d，原高 %.0f）'
                    % (ph.get('reason'), ph.get('peak_h'), ph.get('peak_pages'),
                       ph.get('n_both', 0), base_h0))

            # ⑤ 共识框上下边位于页眉下方 / 页脚上方；高度已采信时跳过（§4.7.4）
            base_y = [consensus[1], consensus[3]]
            if ph.get('used'):
                moved_y, cnt_y = [0.0, 0.0], (0, 0)
                log('共识横线修正: 跳过（高度已由同页净跨距锁定）')
            else:
                consensus, moved_y, cnt_y = hf_refine_consensus_y(consensus, hf_recs)
                log('共识横线修正: y1 %+.0f px (页眉样本 %d)  y2 %+.0f px (页脚样本 %d)'
                    % (moved_y[0], cnt_y[0], moved_y[1], cnt_y[1]))
            log('修正后共识: y1=%.0f y2=%.0f  高=%.0f (原 %.0f)'
                % (consensus[1], consensus[3],
                   consensus[3] - consensus[1], base_y[1] - base_y[0]))


            # ④ 角点定位：共识框左上角贴「页眉下方的框线交点」，
            #    否则下角贴「页脚上方的框线交点」。只平移，尺寸严格不变。
            base_c = [consensus[0], consensus[1]]
            consensus, corner = hf_refine_consensus_corner(consensus, hf_recs)
            if corner.get('mode') == 'header_corner':
                log('共识角点定位: 左上角贴页眉下方交点 → (%.0f, %.0f)  样本 x%d/y%d  平移 %+.0f,%+.0f px'
                    % (corner['tl'][0], corner['tl'][1], corner['tl'][2],
                       corner['tl'][3], corner['shift'][0], corner['shift'][1]))
            elif corner.get('mode', '').startswith('footer_corner'):
                log('共识角点定位: 下角(%s)贴页脚上方交点 → (%.0f, %.0f)  样本 x%d/y%d  平移 %+.0f,%+.0f px'
                    % (corner['side'], corner['bt'][0], corner['bt'][1],
                       corner['bt'][2], corner['bt'][3],
                       corner['shift'][0], corner['shift'][1]))
            else:
                log('共识角点定位: 未找到页眉/页脚交点（候选 上%d 下%d，需 %d 页），保持原位'
                    % (corner.get('n_tl', 0), corner.get('n_bt', 0),
                       corner.get('need', 0)))
            log('  基准: 页眉底边中位=%s  页脚顶边中位=%s'
                % (corner.get('hb_med'), corner.get('ft_med')))
            log('角点定位后: x1=%.0f y1=%.0f  尺寸 %.0f×%.0f (保持不变)'
                % (consensus[0], consensus[1],
                   consensus[2] - consensus[0], consensus[3] - consensus[1]))

            hf_info = {'enabled': True,
                       'width_calib': pw,
                       'height_calib': ph,
                       'consensus_x_shift': [round(moved_x[0], 1), round(moved_x[1], 1)],
                       'consensus_x_samples': list(cnt_x),
                       'consensus_y_shift': [round(moved_y[0], 1), round(moved_y[1], 1)],
                       'consensus_y_samples': list(cnt_y),
                       'corner': corner,
                       'header': hdr, 'footer': ftr}
            # 页眉模板匹配已完成，释放顶部墨迹带
            for r in hf_recs:
                r.pop('hf_sm', None)

    # ---- 逐页定状态 + 输出 ----
    pages_out = {}
    counts = {'ok': 0, 'fallback_consensus': 0, 'manual': 0}
    manual_list = []
    corner_stat = {}        # 逐页角点对齐命中统计 {模式: 页数}

    cw = consensus[2] - consensus[0]
    ch = consensus[3] - consensus[1]

    fallback_list = []
    for r in records:
        name = r['name']
        W, H = r['w'], r['h']
        final_abs = None
        status = 'manual'
        confidence = 0.0
        note = ''
        vc_hf = None
        corner_mode, corner_pt = None, None

        if r['status_raw'] == 'manual' or consensus is None:
            status = 'manual'
            confidence = 0.0
        else:
            # 所有页版心框尺寸恒定 = 共识框尺寸，只做整体平移（贴边/居中），
            # 不因为当页检测框偏小就缩框——版心物理大小每页相同。
            # 启用页眉页脚检测时，X 轴的贴边证据改用「页眉/页脚带内检出的竖线」，
            # 使共识框紧贴那些垂直边线；该页无竖线时自动退回整页长竖线。
            vc_hf = None
            if hf_on:
                vc_hf = merge_vc(r.get('vc_abs', []),
                                 r.get('hf_vlines') or [], W)
            dx, dy, mx, my = consensus_box_shift(
                r.get('vc_abs', []), r.get('hc_abs', []), consensus,
                r.get('edges_abs'), vc_hf)
            final_abs = [int(round(consensus[0] + dx)),
                         int(round(consensus[1] + dy)),
                         int(round(consensus[2] + dx)),
                         int(round(consensus[3] + dy))]

            # ---- 逐页角点对齐（B 模式核心）----
            # 全书统一只用来定尺寸 (cw, ch)；位置【每一页独立】对齐：
            # 有页眉 → 左上角贴页眉框右下角正下方的交点；只有页脚 → 下角贴
            # 页脚框上方的交点；都没有 → 保持上面算出的框线贴边 / 居中结果。
            corner_mode, corner_pt = None, None
            if hf_on and r.get('hf_ok'):
                tgt, cmode = hf_page_corner_target(r, cw, ch, W, H)
                if tgt is not None:
                    bx1 = consensus[0] + dx
                    by1 = consensus[1] + dy
                    # sanity cap：只允许相对「常规贴边结果」再平移有限量，
                    # 防止选到页面另一侧的伪交点造成灾难性漂移
                    capx = HF_CORNER_CAP * (W or 1)
                    capy = HF_CORNER_CAP * (H or 1)
                    # tgt 的任一轴可能是 None，表示该轴保持常规贴边结果不动
                    # （横线降级只给 y，硬套横线端点当 x 会偏移数百 px）
                    nx1, ny1 = bx1, by1
                    if tgt[0] is not None:
                        nx1 = bx1 + max(-capx, min(capx, tgt[0] - bx1))
                        mx = 'corner'
                    if tgt[1] is not None:
                        ny1 = by1 + max(-capy, min(capy, tgt[1] - by1))
                        my = 'corner'
                    final_abs = [int(round(nx1)), int(round(ny1)),
                                 int(round(nx1 + cw)), int(round(ny1 + ch))]
                    corner_mode, corner_pt = cmode, (nx1, ny1)
            # 边界 clamp：HF 修正后共识可能离页面边缘很近，per-page 贴边可能
            # 把框推出页面。尺寸恒定优先，x1/y1 至少 0，x2/y2 最多 W/H。
            cw_eff = final_abs[2] - final_abs[0]
            ch_eff = final_abs[3] - final_abs[1]
            if final_abs[0] < 0:
                final_abs[0] = 0
                final_abs[2] = cw_eff
            if final_abs[2] > W:
                final_abs[2] = W
                final_abs[0] = W - cw_eff
            if final_abs[1] < 0:
                final_abs[1] = 0
                final_abs[3] = ch_eff
            if final_abs[3] > H:
                final_abs[3] = H
                final_abs[1] = H - ch_eff
            pos_note = 'x=%s,y=%s,shift=%.0f,%.0f' % (
                mx, my, final_abs[0] - consensus[0], final_abs[1] - consensus[1])
            mode_rank = {'corner': 3, 'snap': 2, 'center': 2, 'edge': 1, 'none': 0}
            n_good = mode_rank.get(mx, 0) + mode_rank.get(my, 0)

            if r['edges_abs'] is not None:
                e = r['edges_abs']
                ew, eh = e[2] - e[0], e[3] - e[1]
                dw = abs(ew - cw) / cw
                dh = abs(eh - ch) / ch
                if dw <= SIZE_TOL_W and dh <= SIZE_TOL_H:
                    # 四边检测成功且尺寸与共识一致
                    status = 'ok'
                    confidence = round(1.0 - 0.5 * max(dw / SIZE_TOL_W, dh / SIZE_TOL_H), 3)
                else:
                    # 检测到框但尺寸明显不符（误检内框/特殊版式）→ 兜底，
                    # 框仍取共识尺寸，位置按贴边/居中平移
                    status = 'fallback_consensus'
                    confidence = round(0.35 + 0.08 * n_good, 2)
                    note = 'size_mismatch(dw=%.2f,dh=%.2f,%s)' % (dw, dh, pos_note)
            else:
                # 缺边但内容正常：共识尺寸框 + 贴边/居中平移
                status = 'fallback_consensus'
                confidence = round(0.30 + 0.10 * n_good, 2)
                note = 'edges_missing(%s)' % pos_note

        counts[status] += 1
        if corner_mode:
            corner_stat[corner_mode] = corner_stat.get(corner_mode, 0) + 1
        if status == 'manual':
            manual_list.append('%s(%s)' % (name, r.get('reason', '')))
        elif status == 'fallback_consensus':
            fallback_list.append('%s %s' % (name, note))

        entry = {
            'page': r['page'],
            'width': W,
            'height': H,
            'status': status,
            'confidence': confidence,
            'angle_deg': round(r.get('angle', 0.0), 3),
        }
        if status == 'fallback_consensus':
            entry['note'] = note
        if final_abs is not None and W and H:
            x1, y1, x2, y2 = final_abs
            entry['hanmen'] = [int(x1), int(y1), int(x2), int(y2)]
            entry['hanmen_norm'] = [round(x1 / W, 5), round(y1 / H, 5),
                                    round(x2 / W, 5), round(y2 / H, 5)]

        # ---- 页眉 / 页脚 / 分镜框线 ----
        if hf_on:
            hfe = {}
            if r.get('hf_header'):
                hfe['header'] = r['hf_header']
            if r.get('hf_footer'):
                hfe['footer'] = r['hf_footer']
            if r.get('hf_bands') is not None:
                pl = {}
                for nm, by0, segs in r['hf_bands']:
                    pl[nm] = []
                    for q in segs:
                        ax0 = int(round((q['x'] + q['w'] / 2.0) / r['hf_s']))
                        ay0 = int(round(by0 + q['y'] / r['hf_s']))
                        ay1 = int(round(by0 + (q['y'] + q['h']) / r['hf_s']))
                        ax1 = int(round((q['x'] + q['w']) / r['hf_s']))
                        if q['o'] == 'h':
                            p0, p1 = [int(round(q['x'] / r['hf_s'])), ay0], \
                                     [int(round((q['x'] + q['w']) / r['hf_s'])), ay0]
                        else:
                            p0, p1 = [ax0, ay0], [ax0, ay1]
                        pl[nm].append(dict(o=q['o'], kind=q['kind'], pol=q['pol'],
                                           p0=p0, p1=p1,
                                           lnr=round(q['lnr'], 4),
                                           thn=round(q['thn'], 5),
                                           cov=round(q['cov'], 2), sup=q['sup']))
                hfe['panel_lines'] = pl
            if r.get('hf_vlines'):
                hfe['vlines_abs'] = [int(round(x)) for x, st in r['hf_vlines']]
            if final_abs is not None:
                hfe['snap_x'] = [[int(final_abs[0]),
                                  'xsnap_hf' if vc_hf else 'xsnap'],
                                 [int(final_abs[2]),
                                  'xsnap_hf' if vc_hf else 'xsnap']]
                hfe['snap_mode'] = [mx, my]
            if corner_mode:
                # 逐页角点对齐的模式与实际落点，便于按页排查对齐依据
                hfe['corner_mode'] = corner_mode
                if corner_pt:
                    hfe['corner_pt'] = [int(round(corner_pt[0])),
                                        int(round(corner_pt[1]))]
            entry['hf'] = hfe
        pages_out[name] = entry

        # ---- debug 叠加图 ----
        if make_debug and r.get('small') is not None:
            small = r['small']
            vis = cv2.cvtColor(small, cv2.COLOR_GRAY2BGR)
            hc, vc = r['clusters'] or ([], [])
            for y, s in hc:
                cv2.line(vis, (0, y), (small.shape[1], y), (255, 80, 0), 1)
            for x, s in vc:
                cv2.line(vis, (x, 0), (x, small.shape[0]), (0, 140, 255), 1)
            color = {'ok': (0, 0, 255),
                     'fallback_consensus': (0, 200, 255),
                     'manual': (0, 220, 0)}[status]
            if final_abs is not None:
                sh, sw = small.shape
                sx = sw / float(W)
                sy = sh / float(H)
                p1 = (int(final_abs[0] * sx), int(final_abs[1] * sy))
                p2 = (int(final_abs[2] * sx), int(final_abs[3] * sy))
                cv2.rectangle(vis, p1, p2, color, 3)
            # 页眉 / 页脚 / 分镜框线叠加（黄=分镜框线，品红=页码，蓝=页眉）
            if hf_on and W and H and r.get('hf_s'):
                sh, sw = small.shape
                kx = sw / float(HF_IW)                       # 归一化 x → 小图
                ky = (sh / float(H)) * (float(W) / HF_IW)    # 归一化 y → 小图
                for nm, by0, segs in (r.get('hf_bands') or []):
                    oy = by0 * (sh / float(H))
                    for q in segs:
                        c = (0, 190, 190) if q['o'] == 'h' else (0, 215, 255)
                        cv2.line(vis,
                                 (int(q['x'] * kx), int(oy + q['y'] * ky)),
                                 (int((q['x'] + q['w']) * kx),
                                  int(oy + (q['y'] + q['h']) * ky)),
                                 c, 2)
                hb = r.get('hf_header')
                if hb:
                    b = hb['box']
                    cv2.rectangle(vis, (int(b[0] * kx), int(b[1] * ky)),
                                  (int(b[2] * kx), int(b[3] * ky)), (255, 0, 0), 2)
                fb = r.get('hf_footer')
                if fb:
                    b = fb['box']
                    oy = int(H * (1 - HF_BOTR)) * (sh / float(H))
                    cv2.rectangle(vis, (int(b[0] * kx), int(oy + b[1] * ky)),
                                  (int(b[2] * kx), int(oy + b[3] * ky)),
                                  (200, 0, 200), 2)
                # 角点：青色小十字=本页候选交点，亮青大十字=全书选中的定位角
                # （用于肉眼核查共识框角是否真的压在框线交点上）
                cinfo = (hf_info or {}).get('corner') if hf_on else None
                if cinfo:
                    hbm = cinfo.get('hb_med')
                    ftm = cinfo.get('ft_med')
                    sx2 = sw / float(W)
                    sy2 = sh / float(H)
                    for nm, by0, segs in (r.get('hf_bands') or []):
                        if nm == 'near_header' and hbm is not None:
                            pts = hf_band_corners(segs, by0, r['hf_s'],
                                                  y_lo=hbm,
                                                  y_hi=hbm + HF_CORNER_SPAN * H)
                        elif nm == 'near_footer' and ftm is not None:
                            pts = hf_band_corners(segs, by0, r['hf_s'],
                                                  y_lo=ftm - HF_CORNER_SPAN * H,
                                                  y_hi=ftm)
                        else:
                            pts = []
                        for (cx, cy, st) in pts:
                            cv2.drawMarker(vis, (int(cx * sx2), int(cy * sy2)),
                                           (255, 255, 0),
                                           cv2.MARKER_CROSS, 11, 2)
                    # 大十字：本页【实际选中】的定位角。
                    # 逐页对齐命中时用本页的 corner_pt（必然落在上面的小十字之一上），
                    # 否则退回全书统一的角，供肉眼对比。
                    sel_pt = corner_pt
                    if sel_pt is None:
                        sel = cinfo.get('tl') or cinfo.get('bt')
                        sel_pt = (sel[0], sel[1]) if sel else None
                    if sel_pt:
                        cv2.drawMarker(vis, (int(sel_pt[0] * sx2),
                                             int(sel_pt[1] * sy2)),
                                       (0, 255, 255),
                                       cv2.MARKER_CROSS, 27, 3)

            cv2.putText(vis, status, (20, 60), cv2.FONT_HERSHEY_SIMPLEX,
                        1.6, color, 3)
            outp = os.path.join(debug_dir, os.path.splitext(name)[0] + '_debug.png')
            ok2, buf = cv2.imencode('.png', vis)
            if ok2:
                buf.tofile(outp)

    # 归一化共识（用中位页宽/页高换算，仅供参考）
    consensus_norm = None
    if consensus:
        mw = float(np.median([r['w'] for r in records if r['w']]))
        mh = float(np.median([r['h'] for r in records if r['h']]))
        consensus_norm = [round(consensus[0] / mw, 5), round(consensus[1] / mh, 5),
                          round(consensus[2] / mw, 5), round(consensus[3] / mh, 5)]

    out_json = {
        'version': 2,
        'image_dir': img_dir.replace('\\', '/'),
        'dpi': int(np.median([r['dpi'] for r in records if r['w']])) if any(r['w'] for r in records) else 1200,
        'consensus_abs': [int(round(v)) for v in consensus] if consensus else None,
        'consensus_norm': consensus_norm,
        'pages': pages_out,
    }
    if hf_info:
        out_json['header_footer'] = hf_info
    json_path = os.path.join(img_dir, 'hanmen.json')

    def _jsonable(o):
        if isinstance(o, dict):
            return dict((str(k), _jsonable(v)) for k, v in o.items())
        if isinstance(o, (list, tuple)):
            return [_jsonable(v) for v in o]
        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return float(o)
        if isinstance(o, (np.bool_,)):
            return bool(o)
        return o

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(_jsonable(out_json), f, ensure_ascii=False, indent=1)

    log('')
    log('检测完成：共 %d 页 | ok=%d | 共识兜底=%d | 需手动=%d'
        % (len(records), counts['ok'], counts['fallback_consensus'], counts['manual']))
    if corner_stat:
        tot = sum(corner_stat.values())
        log('逐页角点对齐：%d/%d 页命中（%s）'
            % (tot, len(records),
               '  '.join('%s=%d' % (k, v) for k, v in sorted(corner_stat.items()))))
    else:
        log('逐页角点对齐：0 页命中（全部走框线贴边 / 居中）')
    if hf_info and hf_info.get('enabled'):
        log('页眉：命中 %d 页（奇 %d / 偶 %d）  模板 %s'
            % (hf_info['header'].get('hits', 0), hf_info['header'].get('odd', 0),
               hf_info['header'].get('even', 0),
               hf_info['header'].get('template')))
        ftr = hf_info['footer']
        log('页脚：范式 %s  命中 %d 页' % (ftr.get('paradigm'), ftr.get('hits', 0)))
    if fallback_list:
        log('共识兜底页面（建议抽查调试图）：')
        for m in fallback_list:
            log('  ' + m)
    if manual_list:
        log('需手动处理页面：')
        for m in manual_list:
            log('  ' + m)
    log('JSON 已写入: ' + json_path)
    if make_debug:
        log('调试叠加图目录: ' + debug_dir)

    if gui_mode and not tk_mode:
        pause_if_gui()


if __name__ == '__main__':
    try:
        main()
    except SystemExit:
        raise
    except Exception:
        import traceback
        traceback.print_exc()
        if len(sys.argv) < 2:
            try:
                input('\n运行出错，请把以上错误信息截图反馈，按回车键退出...')
            except Exception:
                pass
        sys.exit(1)
