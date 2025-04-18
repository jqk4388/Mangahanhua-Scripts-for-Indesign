//DESCRIPTION:插入 Unicode 字符

/*
关于脚本：
此脚本将在 InDesign 中的当前插入点插入一个或多个 Unicode 字符序列。
此脚本的配置方式有些特殊，需要通过重命名 UnicodeInjector.jsx 文件的副本来完成。

其理念是制作此脚本的一个或多个副本，
根据所需目标重命名这些副本，并将它们放置到脚本文件夹中。

例如，如果将此脚本的副本重命名为：
   U+0061 U+0062 插入 Unicode 字符.jsx

将其复制到 InDesign 脚本文件夹中，每次将文本光标放置在某些文本中，然后双击脚本
（或者如果分配了快捷键，则按下快捷键），它将插入 Unicode 字符 U+0061 和 U+0062（即 'A' 和 'B'）。

使用方法：
制作 UnicodeInjector.jsx 文件的副本。
查找您希望插入的 Unicode 字符（例如，假设我们想插入一个细空格，U+200A）。
将副本重命名为包含所需的 Unicode 以及一些描述性文本 - 例如 U+200A 插入细空格.jsx。
将 'U+200A 插入细空格.jsx' 脚本移动到您的 InDesign 脚本文件夹中。
打开 InDesign - 脚本已准备好使用。通过双击任何文本框中的文本光标定位，然后双击脚本名称，在脚本面板中，Unicode 字符将被插入。

您可以通过“键盘快捷键...”菜单为脚本分配快捷键。

更多信息
---------
此脚本将忽略文件名中的任何描述性文本。
因此，在上述示例中，文本 '插入 Unicode 字符' 将被忽略 - 仅解释 U+0061 和 U+0062。

脚本理解以下符号：
 U+nnnn（U、+，然后是 1 到 4 个十六进制数字）
   示例：U+0041, u+20eF, u+61
 0xnnnn（0、x，然后是 1 到 4 个十六进制数字）
   示例：0x0041, 0x20Ef, 0X61
 0dnnnnn（0、d，然后是 1 到 5 个十进制数字）
   示例：0d65, 0d8431, 0D97


// ----------------
关于 Rorohiko：
Rorohiko 专注于提高印刷、出版和网络工作流程的效率。
此脚本是我们为客户创建的定制解决方案的免费示例。
如果您的工作流程因无聊或重复的任务而受阻，请咨询 sales@rorohiko.com。
我们为客户编写的脚本通常在几周或几个月内就能收回成本。

(C) 2011-2019 Rorohiko Ltd.
版权所有。
作者：Kris Coppieters
kris@rorohiko.com
*/

// 这是从 ExtendScript Toolkit 运行脚本时使用的“虚拟”脚本名称
const kDebugSampleScriptName = "U+0061 U+0062 插入 Unicode 字符.jsx";

const kUnicodeRegEx = /^(.*?u\+([0-9a-f]{1,4}))(.*)$/i;
const kHexaRegEx = /^(.*?0x([0-9a-f]{1,4}))(.*)$/i;
const kDeciRegEx = /^(.*?0d([1-9][0-9]*))(.*)$/i;

const kErr_NoError = 0;
const kErr_NoSelection = 1;
const kErr_MissingCodes = 2;
const kErr_NoDocument = 3;

var error = kErr_NoError;

do
{
    if (app.documents.length == 0) 
    {
        error = kErr_NoDocument;        
        break;
    }
    
    if (app.selection == null)
    {
        error = kErr_NoSelection;
        break;
    }

    if (! (app.selection instanceof Array))
    {
        error = kErr_NoSelection;
        break;
    }

    if (app.selection.length != 1)
    {
        error = kErr_NoSelection;
        break;
    }

    if (! (app.selection[0] instanceof InsertionPoint))
    {
        error = kErr_NoSelection;
        break;
    }

    var fileName; 
    var err;
    try
    {
        if (app.activeScript instanceof File)
        {
            fileName = app.activeScript.name;
        }
        else 
        {
            fileName = kDebugSampleScriptName;
        }
    }
    catch (err)
    {
        fileName = kDebugSampleScriptName;
    }
    
    var fileNameWasPrefixedWithCode;
    var charCode;
    var numberBase;
    var numericalCharCode;
    var matchingRegEx;
    
    error = kErr_MissingCodes;

    // 开始一个循环，从文件名中逐个剥离代码前缀，直到处理完所有代码
    do
    {
        fileNameWasPrefixedWithCode = false;
    
        if (fileName.match(kUnicodeRegEx) != null)
        {
            matchingRegEx = kUnicodeRegEx;
            numberBase = 16;
        }
        else if (fileName.match(kHexaRegEx) != null)
        {
            matchingRegEx = kHexaRegEx;
            numberBase = 16;
        }
        else if (fileName.match(kDeciRegEx) != null)
        {
            matchingRegEx = kDeciRegEx;
            numberBase = 10;
        }
        else
        {
            matchingRegEx = null;
        }

        if (matchingRegEx != null)
        {
            charCode = fileName.replace(matchingRegEx,"$2");

            if (charCode != "")
            {
                numericalCharCode = parseInt(charCode, numberBase);
                if (! isNaN(numericalCharCode) && numericalCharCode != 0)
                {
                    error = kErr_NoError;
                    var word = app.selection[0];
                    word.contents = String.fromCharCode(numericalCharCode);
                    // 设置上标格式，注意上标应用对象是选中部分而不是单独字符
                    word.position = Position.SUPERSCRIPT;
                    

                    fileNameWasPrefixedWithCode = true;
                    // 剥离刚处理的代码，然后循环回去尝试获取下一个代码（如果有）
                    fileName = fileName.replace(matchingRegEx,"$3");
                }
            }
        }
    }
    while (fileNameWasPrefixedWithCode);
}
while (false);

switch (error)
{
case kErr_NoSelection:
    alert("将文本光标放置在文本框中，但不要选择文本。");
    break;
case kErr_MissingCodes:
    alert("重命名脚本以进行配置。要查看配置说明，请右键单击脚本并选择“编辑脚本”。");
    break;
case kErr_NoDocument:
    alert("至少打开一个文档以运行此脚本。");
    break;

}