/*
 * 置入并对齐版心.jsx
 *
 * 配合 detect_hanmen.py 生成的 hanmen.json 使用：
 *   1. 读取每页扫描图中检测到的版心框（像素坐标）；
 *   2. 把图片置入与页面等大的图框；
 *   3. 缩放并平移图片，使扫描版心与 InDesign 页面边距参考线完全重合；
 *   4. 所需缩放率超出「基准缩放率 ± 容差」的页面视为特殊页，保持不动，
 *      结束时汇总问题页清单，提示人工检查。
 *
 * 日志写入 JSON 同目录下的 hanmen_place_log.txt。
 */

#include "../Library/KTUlib.jsx"

// ---------------- 全局状态 ----------------
var version = "v1.0";
var SCRIPT_NAME = '置入并对齐版心'+ version;
var doc = app.activeDocument;
var bookSize = doc.pages.count();
var isLtR = doc.documentPreferences.pageBinding == PageBindingOptions.LEFT_TO_RIGHT;

var artLayer = doc.layers.itemByName('Art').isValid ?
    doc.layers.itemByName('Art') :
    doc.layers.add({ name: 'Art' });

// 从文件名提取页码（与 Place Art.js 相同规则）
function extractPageNum(path) {
    var regex = /\d{3,4}(?=\_?\d?[a-zA-Z]?\.[A-Za-z]{3,4})/;
    var m = regex.exec(path);
    return m ? parseInt(m[0], 10) : 0;
}

// ---------------- JSON 读取 ----------------
function loadHanmenJson(fileObj) {
    fileObj.encoding = 'UTF-8';
    if (!fileObj.open('r')) {
        alert('无法打开文件：' + fileObj.fsName);
        return null;
    }
    var txt = fileObj.read();
    fileObj.close();
    var data;
    try {
        data = JSON.parse(txt);
    } catch (e) {
        // 老版本 ExtendScript 没有原生 JSON
        data = eval('(' + txt + ')');
    }
    if (!data || !data.pages) {
        alert('JSON 内容无效：找不到 pages 字段');
        return null;
    }
    return data;
}

// ---------------- 对话框 ----------------
function showDialog(jsonPath) {
    var w = new Window('dialog', SCRIPT_NAME);
    w.alignChildren = ['fill', 'top'];

    w.add('statictext', undefined,
        '读取 detect_hanmen.py 生成的 hanmen.json，将扫描图置入并对齐到页面边距线。');

    // JSON 选择
    var jsonGrp = w.add('group');
    jsonGrp.alignChildren = ['left', 'center'];
    jsonGrp.add('statictext', undefined, 'hanmen.json：');
    var jsonPathText = jsonGrp.add('edittext', undefined, jsonPath);
    jsonPathText.characters = 42;
    var browseBtn = jsonGrp.add('button', undefined, '浏览...');
    var chosenJson = new File(jsonPath);
    browseBtn.onClick = function () {
        var filter = File.fs == 'Windows' ? 'JSON 文件:*.json' : function (f) {
            return f instanceof File && /\.json$/i.test(f.name);
        };
        var f = File.openDialog('选择 hanmen.json', filter);
        if (f) {
            chosenJson = f;
            jsonPathText.text = f.fsName;
        }
    };

    // 缩放参数
    var scalePanel = w.add('panel', undefined, '缩放率判定（页与页之间允许的缩放波动）');
    scalePanel.alignChildren = ['left', 'top'];
    var g1 = scalePanel.add('group');
    g1.add('statictext', undefined, '基准缩放率(%)：');
    var baseInput = g1.add('edittext', undefined, '100');
    baseInput.characters = 6;
    g1.add('statictext', undefined, '  容差下限(%)：');
    var lowInput = g1.add('edittext', undefined, '98');
    lowInput.characters = 5;
    g1.add('statictext', undefined, '  上限(%)：');
    var highInput = g1.add('edittext', undefined, '102');
    highInput.characters = 5;
    var uniformChk = scalePanel.add('checkbox', undefined,
        '长宽等比例缩放（水平/垂直取同一缩放率=两者平均值；取消勾选则水平、垂直独立缩放）');
    uniformChk.value = true;

    // 页面范围 / 反向
    var optPanel = w.add('panel', undefined, '页面选项');
    optPanel.alignChildren = ['left', 'top'];
    var mangaChk = optPanel.add('checkbox', undefined,
        '漫画反向放置（仅当文档为「从左到右」装订时生效；日漫右翻页文档无需勾选也正确）');
    mangaChk.value = true;
    var bleedChk = optPanel.add('checkbox', undefined,
        '主页无图框而新建图框时，按文档出血设置向四周外扩（跨页文档订口侧不扩）');
    bleedChk.value = true;

    var rangeGrp = optPanel.add('group');
    rangeGrp.alignChildren = ['left', 'center'];
    var allRadio = rangeGrp.add('radiobutton', undefined, '所有页面');
    var rangeRadio = rangeGrp.add('radiobutton', undefined, '页面范围：');
    var rangeInput = rangeGrp.add('edittext', undefined, '');
    rangeInput.characters = 16;
    allRadio.value = true;

    // 超范围处理
    var outPanel = w.add('panel', undefined, '特殊页（所需缩放率超出容差带）处理方式');
    outPanel.alignChildren = ['left', 'top'];
    var skipRadio = outPanel.add('radiobutton', undefined, '跳过，不置入（推荐，稍后人工处理）');
    var placeRawRadio = outPanel.add('radiobutton', undefined, '仅置入，不对齐（保持 100% 原尺寸）');
    skipRadio.value = true;

    var btns = w.add('group');
    btns.alignment = ['right', 'bottom'];
    btns.add('button', undefined, '确定', { name: 'ok' });
    btns.add('button', undefined, '取消', { name: 'cancel' });

    if (w.show() != 1) return null;

    var base = parseFloat(baseInput.text);
    var low = parseFloat(lowInput.text);
    var high = parseFloat(highInput.text);
    if (isNaN(base) || isNaN(low) || isNaN(high) || low > high) {
        alert('缩放率参数无效，请输入数字且下限 ≤ 上限。');
        return null;
    }

    var pageNums = null; // null = 全部
    if (rangeRadio.value) {
        pageNums = parsePageRange(rangeInput.text);
        if (!pageNums || pageNums.length === 0) {
            alert('页面范围无效，示例："12, 32-35"');
            return null;
        }
    }

    return {
        jsonFile: chosenJson,
        basePct: base,
        lowPct: low,
        highPct: high,
        uniformScale: uniformChk.value,
        extendBleed: bleedChk.value,
        mangaReverse: mangaChk.value,
        pageNums: pageNums,
        skipOutOfRange: skipRadio.value
    };
}

// "12, 32-35" -> [12,32,33,34,35]
function parsePageRange(str) {
    var res = [];
    var parts = String(str).split(',');
    for (var i = 0; i < parts.length; i++) {
        var seg = parts[i].replace(/\s/g, '');
        if (seg === '') continue;
        var nums = seg.split('-');
        var a = parseInt(nums[0], 10);
        var b = nums.length > 1 ? parseInt(nums[1], 10) : a;
        if (isNaN(a) || isNaN(b) || b < a) return null;
        for (var n = a; n <= b; n++) res.push(n);
    }
    return res;
}

// ---------------- 页面几何 ----------------

// 返回边距框（版心线）在粘贴板坐标系中的 [上, 左, 下, 右]，单位 pt
function getMarginRect(page) {
    var pb = page.bounds;                 // [top, left, bottom, right]
    var mp = page.marginPreferences;
    return [
        pb[0] + mp.top,
        pb[1] + mp.left,
        pb[2] - mp.bottom,
        pb[3] - mp.right
    ];
}

// 新建图框用对象样式：无填充、无描边（InDesign 默认图框可能带黑描边）
var OBJECT_STYLE_NAME = '版心对齐_无填充无描边';

// 中文版 InDesign 的 "None" 色板名为「无」，需多语言名查找
function getNoneSwatch() {
    var names = ['None', '无', '[无]', '[None]', 'Aucun', 'Ohne'];
    for (var i = 0; i < names.length; i++) {
        var s = doc.swatches.itemByName(names[i]);
        if (s.isValid) return s;
    }
    return null;
}

// 把对象（图框或对象样式）的填充/描边清掉。描边宽设 0 不依赖色板名，必定生效。
function clearFillStroke(obj) {
    var sw = getNoneSwatch();
    try { if (sw) obj.strokeColor = sw; } catch (e) {}
    try { if (sw) obj.fillColor = sw; } catch (e2) {}
    try { obj.strokeWeight = 0; } catch (e3) {}
}

function getNoFrameObjectStyle() {
    var st = doc.objectStyles.itemByName(OBJECT_STYLE_NAME);
    if (!st.isValid) {
        st = doc.objectStyles.add({ name: OBJECT_STYLE_NAME });
    }
    // 每次运行都重新写入无填充无描边定义：
    // 旧版本脚本可能已在文档中留下带黑描边的同名样式，不能只在新建时设置。
    clearFillStroke(st);
    return st;
}

// 图框边界：与页面重合；useBleed 时按文档出血设置外扩。
// 跨页（facingPages）文档：订口侧不扩，仅外三边 + 外侧扩；单页文档：四周均扩。
function getFrameBounds(page, useBleed) {
    var pb = page.bounds;  // [上, 左, 下, 右]
    if (!useBleed) return [pb[0], pb[1], pb[2], pb[3]];
    var dp = doc.documentPreferences;
    var y1 = pb[0] - dp.documentBleedTopOffset;
    var y2 = pb[2] + dp.documentBleedBottomOffset;
    var x1, x2;
    if (dp.facingPages) {
        if (page.side == PageSideOptions.RIGHT_HAND) {
            // 右页：左缘是订口（不扩），右缘是外侧
            x1 = pb[1];
            x2 = pb[3] + dp.documentBleedOutsideOrRightOffset;
        } else {
            // 左页：左缘是外侧，右缘是订口（不扩）
            x1 = pb[1] - dp.documentBleedOutsideOrRightOffset;
            x2 = pb[3];
        }
    } else {
        x1 = pb[1] - dp.documentBleedInsideOrLeftOffset;
        x2 = pb[3] + dp.documentBleedOutsideOrRightOffset;
    }
    return [y1, x1, y2, x2];
}

function findRectangle(arr, anyLayer) {
    for (var i = arr.length - 1; i >= 0; i--) {
        var it = arr[i];
        if (it instanceof Rectangle && !it.itemLayer.locked &&
            (anyLayer || it.itemLayer == artLayer)) {
            return it;
        }
    }
    return null;
}

// 找或建图框：
//   1) 页面 Art 层已有带图框 → 返回 hasImage（跳过，避免重复置入）
//   2) 主页上有图框 → 覆盖（override）主页项目后使用
//   3) 页面 Art 层有空图框 → 直接使用
//   4) 都没有 → 新建（页面+出血外扩，应用无填充无描边对象样式）
function getOrCreateFrame(page, useBleed) {
    // 1) 页面上任何未锁定图层的图框已含图像 → 跳过（避免重复置入；
    //    含上一轮已覆盖主页后置入的情况）
    var rects = page.rectangles;
    for (var k = rects.length - 1; k >= 0; k--) {
        if (!rects[k].itemLayer.locked && rects[k].graphics.length > 0) {
            return { frame: rects[k], hasImage: true };
        }
    }

    // 2) 主页图框：覆盖主页项目
    if (page.appliedMaster !== null && page.appliedMaster.isValid) {
        var masterFrame = findRectangle(page.masterPageItems, true);
        if (masterFrame) {
            var ov = masterFrame.override(page);
            if (ov && ov.isValid) {
                return { frame: ov, hasImage: ov.graphics.length > 0 };
            }
        }
    }

    // 3) Art 层空图框
    var empty = findRectangle(page.rectangles, false);
    if (empty) return { frame: empty, hasImage: false };

    // 4) 新建。
    // 注意：page.rectangles.add() 构造参数里的坐标按「页面相对坐标」解释，
    // 右页会整体偏移一个页宽；必须先建空框，再用 geometricBounds 属性赋
    // 跨页标尺原点下的绝对坐标（与 KTUlib 的做法一致）。
    var frame = page.rectangles.add({ itemLayer: artLayer });
    frame.geometricBounds = getFrameBounds(page, useBleed);
    // 应用「无填充无描边」对象样式并清除全部本地优先选项。
    // 参考 Adobe 社区脚本 ClearStyleOverrides.jsx：
    // applyObjectStyle(style, true) 应用样式时清除覆盖，再显式调一次
    // clearObjectStyleOverrides()。此后不再对图框直接设置填充/描边属性，
    // 否则又会形成新的本地覆盖。
    try {
        frame.applyObjectStyle(getNoFrameObjectStyle(), true);
        frame.clearObjectStyleOverrides();
    } catch (e) {}
    return { frame: frame, hasImage: false };
}

// ---------------- 主流程 ----------------
function run() {
    var usersUnits = app.scriptPreferences.measurementUnit;
    app.scriptPreferences.measurementUnit = MeasurementUnits.POINTS;
    // 切到跨页标尺原点：page.bounds / geometricBounds 统一使用跨页绝对坐标，
    // 新建图框赋坐标时右页才不会偏移（与 KTUlib 的做法一致），结束后还原。
    var oldOrigin = doc.viewPreferences.rulerOrigin;
    doc.viewPreferences.rulerOrigin = RulerOrigin.SPREAD_ORIGIN;

    var logLines = [];
    function log(msg) {
        logLines.push(msg);
    }

    try {
        // 默认 JSON 路径：脚本目录或图片目录旁
        var defaultJson = new File($.fileName).parent.fsName + '/hanmen.json';
        var opts = showDialog(defaultJson);
        if (!opts) {
            return;
        }

        var data = loadHanmenJson(opts.jsonFile);
        if (!data) {
            return;
        }

        var dpi = data.dpi || 1200;
        var imgDir = data.image_dir || opts.jsonFile.parent.fsName;
        imgDir = imgDir.replace(/\\/g, '/');

        // 文件名页码 -> 页面条目
        var byFileNum = {};
        for (var key in data.pages) {
            var pn = extractPageNum(key);
            if (pn && !byFileNum[pn]) byFileNum[pn] = { name: key, info: data.pages[key] };
        }

        log('==== ' + SCRIPT_NAME + ' ' + new Date() + ' ====');
        log('文档: ' + doc.name + '  页数: ' + bookSize +
            '  装订: ' + (isLtR ? '从左到右' : '从右到左') +
            '  漫画反向: ' + (opts.mangaReverse ? '是' : '否'));
        log('JSON: ' + opts.jsonFile.fsName);
        log('图片目录: ' + imgDir + '  DPI: ' + dpi);
        log('基准缩放率: ' + opts.basePct + '%  容差带: ' +
            opts.lowPct + '% ~ ' + opts.highPct + '%  ' +
            '超范围页: ' + (opts.skipOutOfRange ? '跳过不置入' : '仅置入不对齐'));
        log('缩放方式: ' + (opts.uniformScale ? '长宽等比例(平均)' : '长宽独立') +
            '  新建图框出血外扩: ' + (opts.extendBleed ? '是' : '否') +
            '  跨页文档: ' + (doc.documentPreferences.facingPages ? '是' : '否'));
        log('');

        // 进度条
        var progress = new Window('palette', SCRIPT_NAME + '...');
        progress.minimumSize = { width: 400, height: 70 };
        var pbar = progress.add('progressbar', undefined, 0, 100);
        var ptext = progress.add('statictext', undefined, '准备中...');
        progress.show();

        var counts = { aligned: 0, fallback: 0, outRange: 0, manual: 0, missing: 0, exists: 0 };
        var problemPages = [];   // 需人工检查的文档页
        var fallbackPages = [];  // 共识兜底对齐的文档页（需逐页抽查）

        var docPages = opts.pageNums || [];
        if (!opts.pageNums) {
            for (var p = 1; p <= bookSize; p++) docPages.push(p);
        }

        for (var idx = 0; idx < docPages.length; idx++) {
            var docPageNum = docPages[idx];
            var page = doc.pages.itemByName(String(docPageNum));
            if (!page.isValid) {
                log('[文档页 ' + docPageNum + '] 跳过：文档中不存在该页');
                counts.missing++;
                problemPages.push('文档页' + docPageNum + '(页面不存在)');
                continue;
            }

            // 文档页 -> 扫描文件页码
            var filePageNum = (opts.mangaReverse && isLtR) ?
                (bookSize - docPageNum + 1) : docPageNum;

            pbar.value = Math.round(idx / docPages.length * 100);
            ptext.text = '文档页 ' + docPageNum + ' / 扫描页 ' + filePageNum;
            progress.update();

            var entry = byFileNum[filePageNum];
            if (!entry) {
                log('[文档页 ' + docPageNum + ' <- 扫描页 ' + filePageNum + '] 跳过：JSON 中没有对应图片');
                counts.missing++;
                problemPages.push('p.' + docPageNum + '(无JSON数据)');
                continue;
            }
            var info = entry.info;
            var imgFile = new File(imgDir + '/' + entry.name);
            if (!imgFile.exists) {
                log('[文档页 ' + docPageNum + '] 跳过：找不到图片文件 ' + entry.name);
                counts.missing++;
                problemPages.push('p.' + docPageNum + '(文件缺失:' + entry.name + ')');
                continue;
            }
            if (info.status == 'manual' || !info.hanmen) {
                log('[文档页 ' + docPageNum + ' <- ' + entry.name + '] 跳过：检测状态为 manual（封面/空白/横图），需人工处理');
                counts.manual++;
                problemPages.push('p.' + docPageNum + '(特殊页:' + entry.name + ')');
                continue;
            }

            var got = getOrCreateFrame(page, opts.extendBleed);
            if (got.hasImage) {
                log('[文档页 ' + docPageNum + ' <- ' + entry.name + '] 跳过：图框中已有图像');
                counts.exists++;
                problemPages.push('p.' + docPageNum + '(已有图像)');
                continue;
            }
            var frame = got.frame;
            // 不重置图框边界：主页框保持主页尺寸，新建框已按页面+出血生成
            frame.place(imgFile);
            var graphic = frame.graphics[0];

            // ---- 计算所需缩放率 ----
            var margin = getMarginRect(page);
            var mW = margin[3] - margin[1];   // 边距框宽 pt
            var mH = margin[2] - margin[0];   // 边距框高 pt
            var hx1 = info.hanmen[0], hy1 = info.hanmen[1];
            var hx2 = info.hanmen[2], hy2 = info.hanmen[3];

            // 以置入后图像的实际显示尺寸反算 像素→pt 比率，
            // 不依赖 dpi 元数据（避免 1-bit TIF 缺 dpi 时 InDesign 按 72dpi 置入）
            var gb0 = graphic.geometricBounds;
            var ptPerPxX = (gb0[3] - gb0[1]) / info.width;
            var ptPerPxY = (gb0[2] - gb0[0]) / info.height;

            var hanmenWpt = (hx2 - hx1) * ptPerPxX;
            var hanmenHpt = (hy2 - hy1) * ptPerPxY;
            var sx = mW / hanmenWpt;          // 水平所需缩放（相对当前显示尺寸）
            var sy = mH / hanmenHpt;          // 垂直所需缩放
            var sxPct = sx * graphic.absoluteHorizontalScale;  // 绝对缩放百分比
            var syPct = sy * graphic.absoluteVerticalScale;

            // 等比例：取两者平均值作为统一缩放率（长宽同一数值）
            var sUse, sPct;
            if (opts.uniformScale) {
                sUse = (sx + sy) / 2;
                sPct = (sxPct + syPct) / 2;
            } else {
                sUse = null;
                sPct = null;
            }
            var useXPct = opts.uniformScale ? sPct : sxPct;
            var useYPct = opts.uniformScale ? sPct : syPct;

            // 容差带以基准缩放率为中心（base=100、98~102 时，即所需缩放须在 98%~102%）
            var bandLo = opts.basePct * opts.lowPct / 100;
            var bandHi = opts.basePct * opts.highPct / 100;
            var inBand = useXPct >= bandLo && useXPct <= bandHi &&
                         useYPct >= bandLo && useYPct <= bandHi;

            var scaleLog = opts.uniformScale ?
                sPct.toFixed(2) + '%' :
                sxPct.toFixed(2) + '%/' + syPct.toFixed(2) + '%';
            var skewNote = (Math.abs(info.angle_deg || 0) >= 0.2) ? ' 倾斜' + info.angle_deg + '度需人工' : '';

            if (!inBand) {
                if (opts.skipOutOfRange) {
                    // 删除刚置入的图，保持页面原样
                    graphic.remove();
                    log('[文档页 ' + docPageNum + ' <- ' + entry.name +
                        '] 超范围未置入：所需缩放 ' + scaleLog +
                        '，允许带 ' + bandLo.toFixed(1) + '%~' + bandHi.toFixed(1) + '%');
                    counts.outRange++;
                    problemPages.push('p.' + docPageNum + '(缩放' + scaleLog + ')');
                    continue;
                } else {
                    // 保留原尺寸置入，不对齐
                    log('[文档页 ' + docPageNum + ' <- ' + entry.name +
                        '] 超范围仅置入未对齐：所需缩放 ' + scaleLog);
                    counts.outRange++;
                    problemPages.push('p.' + docPageNum + '(未对齐,缩放' + scaleLog + ')');
                    continue;
                }
            }

            // ---- 缩放（等比例时水平/垂直使用同一缩放率）----
            graphic.absoluteHorizontalScale = useXPct;
            graphic.absoluteVerticalScale = useYPct;
            var moveSx = opts.uniformScale ? sUse : sx;
            var moveSy = opts.uniformScale ? sUse : sy;

            // ---- 平移：使版心左上角落到边距框左上角 ----
            var gb = graphic.geometricBounds;  // [top, left, bottom, right]
            var imgLeft = gb[1];
            var imgTop = gb[0];
            var dx = margin[1] - (imgLeft + hx1 * ptPerPxX * moveSx);
            var dy = margin[0] - (imgTop + hy1 * ptPerPxY * moveSy);
            graphic.move(undefined, [dx, dy]);

            var tag = info.status == 'fallback_consensus' ? '兜底' : '对齐';
            log('[文档页 ' + docPageNum + ' <- ' + entry.name + '] ' + tag +
                ' 缩放 ' + scaleLog +
                '  平移 (' + dx.toFixed(1) + ', ' + dy.toFixed(1) + ')pt' +
                (info.note ? '  [' + info.note + ']' : '') + skewNote);

            if (info.status == 'fallback_consensus') {
                counts.fallback++;
                fallbackPages.push(docPageNum);
            } else {
                counts.aligned++;
            }
        }

        progress.close();

        log('');
        log('完成统计：精确对齐 ' + counts.aligned + ' 页，共识兜底对齐 ' + counts.fallback +
            ' 页（建议抽查），超范围特殊页 ' + counts.outRange + ' 页，' +
            'manual 跳过 ' + counts.manual + ' 页，数据/文件缺失 ' + counts.missing +
            ' 页，已有图像跳过 ' + counts.exists + ' 页。');
        if (fallbackPages.length > 0) {
            log('');
            log('==== 共识兜底对齐页面（共 ' + fallbackPages.length +
                ' 页，请逐页对照 hanmen_debug 黄框抽查）====');
            log(fallbackPages.join(', '));
        }

        // 写日志
        var logFile = new File(opts.jsonFile.parent.fsName + '/hanmen_place_log.txt');
        logFile.encoding = 'UTF-8';
        if (logFile.open('w')) {
            logFile.write(logLines.join('\r\n'));
            logFile.close();
        }

        // ---- 汇总弹窗 ----
        var msg = '处理完成：\n' +
            '  精确对齐：' + counts.aligned + ' 页\n' +
            '  共识兜底对齐：' + counts.fallback + ' 页（建议抽查）\n' +
            '  超范围特殊页：' + counts.outRange + ' 页\n' +
            '  特殊页跳过(manual)：' + counts.manual + ' 页\n' +
            '  数据/文件缺失：' + counts.missing + ' 页\n' +
            '  已有图像跳过：' + counts.exists + ' 页\n\n';
        if (fallbackPages.length > 0) {
            msg += '需抽查的兜底页（共 ' + fallbackPages.length + ' 页）：\n' +
                fallbackPages.join(', ') + '\n\n';
        }
        if (problemPages.length > 0) {
            var show = problemPages.slice(0, 40).join('、');
            if (problemPages.length > 40) show += ' 等共 ' + problemPages.length + ' 页';
            msg += '其他需人工检查的页面：\n' + show + '\n\n';
        }
        msg += '详细日志：' + logFile.fsName;
        alert(msg, SCRIPT_NAME);

    } finally {
        app.scriptPreferences.measurementUnit = usersUnits;
        doc.viewPreferences.rulerOrigin = oldOrigin;
    }
}

KTUDoScriptAsUndoable(run, SCRIPT_NAME);
