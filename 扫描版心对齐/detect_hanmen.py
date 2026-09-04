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
# 版心物理尺寸恒定（同扫描仪同 DPI），但页面在扫描台上的摆放位置有平移，
# 因此按「检测框尺寸」判定可信度，位置以当页检测为准。
SIZE_TOL_W = 0.12         # 检测框宽度允许偏差（相对共识框宽）
SIZE_TOL_H = 0.10         # 检测框高度允许偏差（相对共识框高）
DELTA_ZONE = 500          # 兜底页平移估计：在共识边 ±该范围（px）内寻找当页线簇
SHIFT_AGREE = 150         # 平移候选互相一致的最大散布（px）
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


def estimate_shift(vc_abs, hc_abs, cons):
    """兜底页平移估计：在共识版心四边 ±DELTA_ZONE 带内寻找当页线簇，
    用线簇与对应共识边的偏差估计当页扫描平移量 (dx, dy)。
    为防止出血页/大标题等内部装饰线误导，仅当 ≥2 个候选互相一致
    （距中位数 ≤SHIFT_AGREE px）且位移量合理时才采纳，否则返回 0。"""
    cx1, cy1, cx2, cy2 = cons

    def decide(cands, max_shift):
        if len(cands) < 2:
            return 0.0, 0
        med = float(np.median(cands))
        if abs(med) > max_shift:
            return 0.0, len(cands)
        if max(abs(c - med) for c in cands) > SHIFT_AGREE:
            return 0.0, len(cands)
        return med, len(cands)

    dxs = []
    for x, s in vc_abs:
        if abs(x - cx1) <= DELTA_ZONE:
            dxs.append(x - cx1)
        elif abs(x - cx2) <= DELTA_ZONE:
            dxs.append(x - cx2)
    dys = []
    for y, s in hc_abs:
        if abs(y - cy1) <= DELTA_ZONE:
            dys.append(y - cy1)
        elif abs(y - cy2) <= DELTA_ZONE:
            dys.append(y - cy2)
    dx, nx = decide(dxs, 400)
    dy, ny = decide(dys, 250)
    return dx, dy, nx, ny


def fallback_snap_box(vc_abs, hc_abs, cons):
    """兜底版心框：优先把共识框各边直接「吸附」到当页实测线。

    原理：缺线页通常仍能检测到 2~3 根版心边线（只是没凑齐 4 根或几何
    校验没过）。对每条共识边，在其 ±DELTA_ZONE 带内找当页最强的线簇
    作为该边的实测位置——能实测的边直接钉在实线上；找不到的边保留共识
    位置。同一轴上已实测边相对共识边的位移中位数视为当页扫描平移量
    （同批扫描版心物理尺寸不变，只有整页摆放平移），补到缺失边上。

    防护：
    - 带内两条强线强度接近且位置分散（≥SHIFT_AGREE）→ 该边有歧义，不吸附；
    - 轴平移量超上限（x 400 / y 250 px）→ 该轴吸附结果不可信（出血页
      大标题/装饰线），整轴退回共识位置；
    - 轴两边都实测但框尺寸偏差超 SIZE_TOL → 夹了内部框线，整轴退回；
    - 一条边都没吸附时退回 estimate_shift 的整体平移估计。

    返回 (box_abs[x1,y1,x2,y2], n_snapped, sides_str, dx, dy)。
    sides_str 为已吸附边的标签组合（L=左 T=上 R=右 B=下）。
    """
    cx1, cy1, cx2, cy2 = cons
    cw = cx2 - cx1
    ch = cy2 - cy1

    def best_line(cands, center):
        near = [(p, s) for p, s in cands if abs(p - center) <= DELTA_ZONE]
        if not near:
            return None
        near.sort(key=lambda c: -c[1])
        p, s = near[0]
        # 次强线强度接近(≥80%)且位置明显分散 → 两条候选强线打架，不可靠
        if len(near) > 1 and near[1][1] >= 0.8 * s and abs(near[1][0] - p) > SHIFT_AGREE:
            return None
        return p

    lx = best_line(vc_abs, cx1)
    rx = best_line(vc_abs, cx2)
    ty = best_line(hc_abs, cy1)
    by = best_line(hc_abs, cy2)

    dxs = [v - c for v, c in ((lx, cx1), (rx, cx2)) if v is not None]
    dys = [v - c for v, c in ((ty, cy1), (by, cy2)) if v is not None]
    dx = float(np.median(dxs)) if dxs else None
    dy = float(np.median(dys)) if dys else None

    # 轴位移超上限 → 该轴吸附不可信，整轴退回
    if dx is not None and abs(dx) > 400:
        lx = rx = dx = None
    if dy is not None and abs(dy) > 250:
        ty = by = dy = None

    # 两边都实测但框尺寸偏差过大 → 夹了内部框线，整轴退回
    if lx is not None and rx is not None and abs(rx - lx - cw) / cw > SIZE_TOL_W:
        lx = rx = dx = None
    if ty is not None and by is not None and abs(by - ty - ch) / ch > SIZE_TOL_H:
        ty = by = dy = None

    sides = ''.join(lab for lab, v in (('L', lx), ('T', ty), ('R', rx), ('B', by))
                    if v is not None)
    n_snap = len(sides)

    # 一条边都没吸附：退回旧的整体平移估计（≥2 候选一致才平移）
    if n_snap == 0:
        dx0, dy0, _, _ = estimate_shift(vc_abs, hc_abs, cons)
        return [int(round(cx1 + dx0)), int(round(cy1 + dy0)),
                int(round(cx2 + dx0)), int(round(cy2 + dy0))], 0, '', dx0, dy0

    fx = dx if dx is not None else 0.0
    fy = dy if dy is not None else 0.0
    x1 = lx if lx is not None else cx1 + fx
    x2 = rx if rx is not None else cx2 + fx
    y1 = ty if ty is not None else cy1 + fy
    y2 = by if by is not None else cy2 + fy
    return [int(round(x1)), int(round(y1)), int(round(x2)), int(round(y2))], \
        n_snap, sides, fx, fy


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
        elif r['edges_abs'] is not None:
            e = r['edges_abs']
            ew, eh = e[2] - e[0], e[3] - e[1]
            dw = abs(ew - cw) / cw
            dh = abs(eh - ch) / ch
            if dw <= SIZE_TOL_W and dh <= SIZE_TOL_H:
                # 尺寸与共识一致 → 位置以当页检测为准（容忍扫描摆放平移）
                status = 'ok'
                final_abs = e
                confidence = round(1.0 - 0.5 * max(dw / SIZE_TOL_W, dh / SIZE_TOL_H), 3)
            else:
                # 尺寸不符（误检/特殊版式）：共识框 + 实测边吸附
                final_abs, n_snap, sides, dx, dy = fallback_snap_box(
                    r.get('vc_abs', []), r.get('hc_abs', []), consensus)
                status = 'fallback_consensus'
                confidence = round(min(0.4 + 0.08 * n_snap, 0.75), 2)
                note = 'size_mismatch(dw=%.2f,dh=%.2f,snap=%d:%s,shift=%.0f,%.0f)' \
                       % (dw, dh, n_snap, sides or '-', dx, dy)
        else:
            # 缺边但内容正常：共识框 + 实测边吸附（检测到几根边就钉几根）
            final_abs, n_snap, sides, dx, dy = fallback_snap_box(
                r.get('vc_abs', []), r.get('hc_abs', []), consensus)
            status = 'fallback_consensus'
            confidence = round(min(0.3 + 0.12 * n_snap, 0.75), 2)
            note = 'edges_missing(snap=%d:%s,shift=%.0f,%.0f)' \
                   % (n_snap, sides or '-', dx, dy)

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
