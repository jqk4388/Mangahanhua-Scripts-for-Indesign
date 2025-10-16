#targetengine "session"
// 调试脚本 - 比较两个字符的属性
try {
    // 检查是否有选中的文本
    if (app.selection.length > 0 && app.selection[0].constructor.name === "InsertionPoint") {
        var selection = app.selection[0];
        var story = selection.parentStory;
        var index = selection.index;

        // 确保有足够的字符可以比较
        if (story.contents.length >= index + 2) {
            var char1 = story.characters[index];
            var char2 = story.characters[index + 1];

            // 递归比较属性
            function compareProperties(obj1, obj2, prefix, depth) {
                if (depth > 3) return; // 限制递归深度为3层
                for (var prop in obj1) {
                    if (obj1.hasOwnProperty(prop)) {
                        try {
                            var value1 = obj1[prop];
                            var value2 = obj2[prop];
                            var fullPropName = prefix ? prefix + "." + prop : prop;

                            if (typeof value1 === "object" && typeof value2 === "object" && value1 && value2) {
                                compareProperties(value1, value2, fullPropName, depth + 1);
                            } else if (value1 != value2) {
                                $.writeln(fullPropName + ":");
                                $.writeln("  字符1: " + value1);
                                $.writeln("  字符2: " + value2);
                            }
                        } catch (e) {
                            // 忽略无法访问的属性
                        }
                    }
                }
            }

            $.writeln("=== 字符比较结果 ===");
            $.writeln("字符1: '" + char1.contents + "'");
            $.writeln("字符2: '" + char2.contents + "'");
            $.writeln("-------------------");

            // 开始比较属性
            compareProperties(char1, char2, "", 1);
        } else {
            alert("请确保选择位置后有至少两个字符可供比较。");
        }
    } else {
        alert("请先选择文本位置。");
    }
} catch (e) {
    alert("发生错误：" + e);
}