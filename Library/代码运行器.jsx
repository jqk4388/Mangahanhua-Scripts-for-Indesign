#targetengine "session"
var version = "v1.1";
var win;

// 创建窗口界面
function createWindow() {
    // 注意：不要在此函数外直接声明顶层窗口变量，函数仅返回新建的 palette 实例
    var _win = new Window("palette", "代码运行器 "+version); 

    // 创建代码输入区域
    _win.add("statictext", undefined, "代码粘贴区域：");
    var codeInput = _win.add("edittext", [0, 0, 400, 200], "", { multiline: true });

    // 显示运行时间
    var timeGroup = _win.add("group");
    timeGroup.add("statictext", undefined, "运行时间：");
    var timeDisplay = timeGroup.add("statictext", undefined, "0s");

    // 按钮组
    var buttonGroup = _win.add("group");
    buttonGroup.orientation = "row";

    // 运行按钮
    var runButton = buttonGroup.add("button", undefined, "运行");
    runButton.onClick = function () {
        try {
            var startTime = new Date();
            eval(codeInput.text);
            var endTime = new Date();
            timeDisplay.text = ((endTime - startTime) / 1000) + "s";
        } catch (e) {
            alert("运行错误: " + e.message);
        }
    };

    // 保存按钮（使用 dialog 进行文件名输入，保持模态保存对话以保证用户输入）
    var saveButton = buttonGroup.add("button", undefined, "保存");
    saveButton.onClick = function () {
        try {
            var saveDialog = new Window("dialog", "保存新脚本");
            saveDialog.add("statictext", undefined, "请输入新脚本名：");
            var scriptNameInput = saveDialog.add("edittext", [0, 0, 300, 25], "");
            scriptNameInput.active = true;
            var saveGroup = saveDialog.add("group");
            saveGroup.orientation = "row";
            var saveConfirm = saveGroup.add("button", undefined, "确定");
            var saveCancel = saveGroup.add("button", undefined, "取消");

            saveConfirm.onClick = function () {
                if (scriptNameInput.text === "") {
                    alert("脚本名称不能为空！");
                } else {
                    saveDialog.close(1);
                }
            };
            saveCancel.onClick = function () {
                saveDialog.close(0);
            };

            if (saveDialog.show() === 1) {
                var fileName = scriptNameInput.text + ".jsx";
                var file = File(Folder.userData + "/" + fileName);
                if (file.open("w", "TEXT", "UTF-8")) {
                    file.encoding = "UTF-8"
                    file.write(codeInput.text);
                    file.close();
                    alert("脚本已保存到: " + file.fsName);
                } else {
                    alert("无法保存脚本");
                }
            }
        } catch (e) {
            alert("保存失败: " + e.message);
        }
    };

    // 我的脚本按钮
    var myScriptsButton = buttonGroup.add("button", undefined, "我的脚本");
    myScriptsButton.onClick = function () {
        try {
            Folder(Folder.userData).execute();
        } catch (e) {
            alert("无法打开脚本文件夹: " + e.message);
        }
    };

    // 清除按钮
    var clearButton = buttonGroup.add("button", undefined, "清除");
    clearButton.onClick = function () {
        codeInput.text = "";
        timeDisplay.text = "0s";
    };

    // 关闭按钮
    var closeButton = buttonGroup.add("button", undefined, "关闭");
    closeButton.onClick = function () {
        _win.close();
    };

    return _win;
}

// 调色板类型的窗口不能有主函数
win = createWindow();
win.show();
