# -*- coding: utf-8 -*-
"""
detect_hanmen.py — 扫描漫画页版心（hanmen）自动检测

功能：
  逐页检测扫描图中漫画版心的四边位置（最外侧长直线簇），
  用跨页共识（归一化中位数）兜底缺线页，输出 hanmen.json
  和 hanmen_debug/ 调试叠加图（红=检测OK / 黄=共识兜底 / 绿=需手动）。

用法：
  python detect_hanmen.py <图片目录> [--no-debug]

支持格式：.jpg .jpeg .tif .tiff .png（含黑白 1-bit 二值 TIF）
依赖：Pillow、numpy、opencv-python
"""

import os
import re
import sys
import json
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
                                threshold=200, minLineLength=250, maxLineGap=10)
        if lines is None:
            return 0.0
        angles = []
        for ln in lines[:, 0]:
            x1, y1, x2, y2 = ln
            dx, dy = x2 - x1, y2 - y1
            length = (dx * dx + dy * dy) ** 0.5
            if length < 250:
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


def consensus_box_shift(vc_abs, hc_abs, cons, det=None):
    """版心框位置：尺寸恒定取共识框，只根据当页实测线整体平移（贴边/居中）。
    det 为四边检测器确认的 [x1,y1,x2,y2]（可空），其边线不受采信带限制。
    返回 (dx, dy, mode_x, mode_y)，mode ∈ snap/center/edge/none。
    所有阈值均为共识框尺寸的比例，自动适配不同分辨率原图。"""
    cx1, cy1, cx2, cy2 = cons
    cw, ch = cx2 - cx1, cy2 - cy1
    dx, mx = _axis_shift(vc_abs, cx1, cx2, cw,
                         ZONE_PCT * cw, SHIFT_CAP_PCT_X * cw,
                         AGREE_PCT * cw, SIZE_TOL_W,
                         det[0] if det else None, det[2] if det else None)
    dy, my = _axis_shift(hc_abs, cy1, cy2, ch,
                         ZONE_PCT * ch, SHIFT_CAP_PCT_Y * ch,
                         AGREE_PCT * ch, SIZE_TOL_H,
                         det[1] if det else None, det[3] if det else None)
    return dx, dy, mx, my


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


def main():
    make_debug = '--no-debug' not in sys.argv
    gui_mode = len(sys.argv) < 2  # 双击 .py 启动时没有命令行参数

    img_dir = None
    if len(sys.argv) >= 2:
        cand = clean_path(sys.argv[1])
        if os.path.isdir(cand):
            img_dir = os.path.abspath(cand)
        elif os.path.isfile(cand):
            img_dir = os.path.dirname(os.path.abspath(cand))
        else:
            log('命令行给定的路径无效: ' + cand)

    if img_dir is None:
        log('请在弹出的对话框中选择扫描图片所在的文件夹...')
        chosen = choose_dir_dialog()
        if chosen:
            img_dir = os.path.abspath(clean_path(chosen))

    if not img_dir or not os.path.isdir(img_dir):
        log('未选择有效目录，已取消。')
        if gui_mode:
            pause_if_gui()
        sys.exit(1)

    files = sorted(
        f for f in os.listdir(img_dir)
        if f.lower().endswith(EXTS) and os.path.isfile(os.path.join(img_dir, f))
    )
    if not files:
        log('目录中没有找到支持的图片文件: ' + img_dir)
        if gui_mode:
            pause_if_gui()
        sys.exit(1)

    debug_dir = os.path.join(img_dir, 'hanmen_debug')
    if make_debug:
        os.makedirs(debug_dir, exist_ok=True)

    log('共发现 %d 张图片，开始检测...' % len(files))

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

        records.append({'name': name, 'page': extract_page_num(name),
                        'w': W, 'h': H, 'dpi': dpi, 'edges_abs': edges_abs,
                        'angle': angle, 'ink': ink, 'clusters': clusters,
                        'hc_abs': hc_abs, 'vc_abs': vc_abs,
                        'small': small, 'status_raw': status, 'reason': reason})
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

    # ---- 逐页定状态 + 输出 ----
    pages_out = {}
    counts = {'ok': 0, 'fallback_consensus': 0, 'manual': 0}
    manual_list = []

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

        if r['status_raw'] == 'manual' or consensus is None:
            status = 'manual'
            confidence = 0.0
        else:
            # 所有页版心框尺寸恒定 = 共识框尺寸，只做整体平移（贴边/居中），
            # 不因为当页检测框偏小就缩框——版心物理大小每页相同。
            dx, dy, mx, my = consensus_box_shift(
                r.get('vc_abs', []), r.get('hc_abs', []), consensus,
                r.get('edges_abs'))
            final_abs = [int(round(consensus[0] + dx)),
                         int(round(consensus[1] + dy)),
                         int(round(consensus[2] + dx)),
                         int(round(consensus[3] + dy))]
            pos_note = 'x=%s,y=%s,shift=%.0f,%.0f' % (mx, my, dx, dy)
            mode_rank = {'snap': 2, 'center': 2, 'edge': 1, 'none': 0}
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
    json_path = os.path.join(img_dir, 'hanmen.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(out_json, f, ensure_ascii=False, indent=1)

    log('')
    log('检测完成：共 %d 页 | ok=%d | 共识兜底=%d | 需手动=%d'
        % (len(records), counts['ok'], counts['fallback_consensus'], counts['manual']))
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

    if gui_mode:
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
