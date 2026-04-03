// 拆分为单字框 Split to Single Char Frames
// 将选中的文本框按字符拆分为多个独立文本框，每框一字。
// 做法：duplicate() 原框后只修改文字内容，位置/样式/图层/所在页面全部原样保留。
// 原文本框不删除，仅隐藏（visible = false）。
// 每个复制框在 X 轴方向依次小幅偏移，避免完全重叠。
//
// 用法：选中一个或多个文本框后运行脚本。

(function () {

    // ── 参数 ────────────────────────────────────────────
    var X_OFFSET_STEP = 4; // 每个新框在 X 轴方向的累计偏移量（单位同文档，通常为 pt）

    // ── 主逻辑 ──────────────────────────────────────────
    function splitToSingleCharFrames(tf) {
        if (!(tf instanceof TextFrame)) {
            return;
        }

        var rawText = tf.contents;

        var chars = [];
        for (var i = 0; i < rawText.length; i++) {
            var c = rawText.charAt(i);
            if (c !== "\n" && c !== "\r" && c !== "\t" && c !== "") {
                chars.push(c);
            }
        }

        if (chars.length === 0) {
            return;
        }

        tf.visible = false;

        for (var i = 0; i < chars.length; i++) {
            var newTF = tf.duplicate();
            newTF.contents = chars[i];
            if (X_OFFSET_STEP !== 0) {
                newTF.move(undefined, [i * X_OFFSET_STEP, 0]);
            }
            newTF.visible = true;
        }
    }

    function main() {
        var sel = app.selection;
        if (!sel || sel.length === 0) {
            alert("请先选中一个或多个文本框再运行此脚本。");
            return;
        }

        var targets = [];
        for (var i = 0; i < sel.length; i++) {
            var target = sel[i];
            if (!(target instanceof TextFrame)) {
                if (target.parentTextFrames && target.parentTextFrames.length > 0) {
                    target = target.parentTextFrames[0];
                } else if (target.parent && target.parent instanceof TextFrame) {
                    target = target.parent;
                } else {
                    continue;
                }
            }
            targets.push(target);
        }

        if (targets.length === 0) {
            alert("请先选中一个或多个文本框再运行此脚本。");
            return;
        }

        app.doScript(
            function () {
                for (var i = 0; i < targets.length; i++) {
                    splitToSingleCharFrames(targets[i]);
                }
            },
            ScriptLanguage.JAVASCRIPT,
            [],
            UndoModes.ENTIRE_SCRIPT,
            "拆分为单字框"
        );
    }

    main();

})();
