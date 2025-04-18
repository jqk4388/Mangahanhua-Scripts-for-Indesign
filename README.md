# 漫画排版 InDesign 脚本集

本套 InDesign 自动化脚本旨在减少日漫英译排版中的重复劳动和误差。
***
脚本安装路径：`C:\Users\xxx\AppData\Roaming\Adobe\InDesign\Version 20.0-J\zh_CN\Scripts\Scripts Panel`

在脚本窗口中右键用户文件夹打开资源管理器把本项目放在`Scripts Panel`文件夹下
***
## 重要说明

**这些脚本需作为套件配合使用。** 大多数脚本依赖 `Library/KTUlib.jsx` 文件中的函数，需确保该文件与脚本处于同级目录。

**完整包下载方式**：点击文件列表上方的绿色 **Code** 按钮，选择 "Download Zip"。解压到 InDesign 脚本文件夹根目录即可使用。
***
# 新增汉化专用脚本
* 1.开局新建图层图框.jsx
* 2.导入三个样式.jsx
* 3.LabelPlusTXT导入ID脚本
* 4.放置图像
* 5.导出脚本：1400分辨率（印刷tif图），268分辨率png（B6开本的web图源），前三页彩页分辨率600的RGB
## 文本修改
* 自动断句脚本：
    * 结巴断句支持中文断句，有外挂词典逻辑复杂速度慢，大概1秒1个文本框。可以指定断句的范围，有UI界面
    * 简单断句适合选中文本框断句，速度快。
        * 安装方法 `pip install jieba`
        * 安装完在cmd控制台输入 `python -m jieba` 显示建立词典缓存说明安装成功。
    * 移到前一行：光标放在行首，将第一个字符移到前一行，用于快速修改断句。
* 文章转框架做蒙版.js 选中一个文本框脚本将其复制到剪贴板，同一个位置新建框架，根据提示粘贴到内部。之后可以用钢笔修改锚点制作简单的文字蒙版。
* 括号：自动包裹括号脚本新增简中适用的括号形式。插入是在文字左右插入，包裹作用于整个文本框。
* 自动加拼音：用圆括号（）包裹的文本作为拼音，以行为单位自动在括号前面的文字上添加拼音，自动调整行距144%。
* 自动着重号：用【】（）等括号包裹的文本作为加重音文字，对其内容加着重号，自动调整行距144%。
* 黑白字转换：对文字颜色进行翻转，效果包括填充&描边&外发光&编组对象的填充&描边。
* 收集文档中的文本带坐标导出ID2LPtxt.jsx 全部嵌字完用，在桌面上导出一个LPtxt，包括字体字号位置坐标。
* 注释脚本
    * 选中转水平注释：文本框内选中一段话，转为水平注释。（需要提前导入水平注释的样式）
    * 插入注释标志符号※或全角星号＊（脚本文件名的Unicode编码决定插入符号）
## 页面调整
* 两个图框置入交换：适用于左右两页置入反的情况

## 待实现功能
- 移动全部文本框至文字图层
- 制作LPtxt导入脚本的UI界面，并按页码对应文稿
- 制作一个通用的批量脚本，有UI界面，可以选择批量处理的对象以及运用哪个脚本。


# 原版英化脚本功能说明
## 文档初始化

**添加参考线与页码**：在 A-母版页添加水平/垂直参考线，创建带页码的 B-母版页。

**初始化文档**：新建文档后立即运行，功能包括：
1. 创建分层：参考线、文字、特效、页码、修图、设计、画稿（自下而上）
2. 创建 A/B 母版页
3. 在 A-母版设置参考线，B-母版继承 A-母版
4. 在 A-母版放置裁切尺寸的画框
5. 在 B-母版添加自动页码
6. 提示选择画稿素材文件夹（支持嵌套子文件夹递归搜索）
7. 解析文件名确定页面顺序和总页数
8. 添加页面至最近的 16 页倍数
9. 自动将画稿置入对应页面的画框

## 画稿处理

**创建母版画框**：在 A-母版页创建两个出血边距（外边缘 1/8 英寸）的画框，用于统一管理全文档画框属性。

**全文档画稿移至底层**：移动所有画框至最底层。

**当前页画稿移至底层**：仅移动当前页画框至底层。

**自动置入画稿**：实验性功能，根据文件名匹配自动置入画稿（稳定性较差）。

**母版画框比例调整**：批量调整通过母版画框创建的画框缩放比例（需注意以下限制）：
- 基于当前 100% 比例进行调整
- 需配合母版画框使用
- 个别页面可能需要手动微调
- 手动修改过的画框将脱离母版关联

## 参考线调整

**添加参考线**：在 A-母版页四周添加 1/2 英寸参考线。

**水平边距微调**：每次运行使跨页上下参考线靠近/远离 1pt。

**垂直边距微调**：每次运行使跨页左右参考线靠近/远离 1pt。

## 页码与装订

**切换装订方向**：左开本/右开本切换。

**反向书页顺序**：整体反转文档页序，注意首尾页内容可能错位。

**左/右页页码开关**：为当前跨页的左/右页应用/取消 B-母版页码样式。

## 文字特效

**字符渐大**：使文字逐字增大（音量增强效果）

**字符渐小**：使文字逐字缩小（音量减弱效果）

**选中包裹**：用各种类型括号包裹选中文字

**中部膨胀**：文字中间大两端小

**中部收缩**：文字中间小两端大

**随机错位**：字符随机上下偏移

**下行阶梯**：字符逐字降低基线

**上行阶梯**：字符逐字升高基线

**随机旋转**：字符交替随机旋转

**垂直堆叠**：字符竖向排列（每字一行）

![文字特效示例](文本特效/Text%20Design%20Examples.png)

## 其他工具

**新建右开本文档**：创建预设右开本的新文档（适用于旧版 InDesign）

**创建译注缩略图**：将选中对象创建为带样式的缩略图组（注：可能产生溢流文本错误）

**查找空文本框**：定位未填写文字的文本框

**显示尺寸信息**：显示选中对象的几何边界（调试用）

**锁定/解锁全部对象**：全局锁定/解锁

**选择左/右页内容**：选中当前跨页左/右页所有对象

**匹配画框尺寸**：调整选中画框至页面尺寸+出血边距（自动识别装订方向）

**压缩文本行**：当前文本行水平缩放减少 5%（最小建议 80%）

**取消换行**：删除选中文本框中的所有换行

**代码运行器**：简易的代码运行器，调试代码用

**调试光标后两字符**：可以简单判断两个字符发生了什么变化，方便找出变量名
