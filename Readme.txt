0.简介
    程序使用最简单的VB6开发工具直接拖动控件，来完成Python的TKinter的
    GUI布局，可以在VB界面上设置控件的一些属性，最终自动生成必要的代码
   （包括回调函数框架），代码生成后仅需要在对应的回调函数中增加相应的
    逻辑代码即可。
    这个工具不支持全部TKiner控件，仅选择一些常用的，和VB控件有直接对应
    关系或比较类似的（参见下面的控件说明）。

1.适用对象
    仅适用于学习了TKinter并不想太麻烦写GUI代码，也不想用其他工具和框架
    比如wxPython,PyQt4的同学。
    也仅适用于界面不太复杂的小程序开发，界面复杂的还是适用wxPython等
    框架吧。
    因为TKinter为Python标准库，使用TKinter完成的Python程序可以称为
    “绿色软件”，不需要目标机器上安装wxPython,PyQt4等框架，只要有Python
    的机器就能运行。
    如果软件逻辑不复杂，一个*.py搞定，不像其他框架，需要几个文件。
    此工具为业余时间简单倒腾而成，BUG肯定是不少的，不过幸好有源代码，
    如果有熟悉VB的同学可以自己修改。

2.使用方法简介
  2.1 首先注册此插件，可以使用自带的安装程序，或自己手动完成。
  2.2 打开VB6，新建一个标准EXE工程，在窗体上设计自己的GUI布局，这个
    工作估计没有VB基础的同学都可以完成，同时可以设置相应的控件属性。
  2.3 如果使用自带安装程序安装了插件，现在VB的工具条上应该有一个新
    图标（一片橙红色羽毛），如果没有，到菜单"外接程序"|"外接程序管理器"
    里面启动Visual Tkinter，Visual Tkinter图标和菜单应该会出现。
  2.4 启动Visual Tkinter后，先按“刷新窗体列表”按钮，列出当前工程的所有
    窗体和控件列表。
  2.5 逐个确认各控件的输出属性，在要输出的选项前打钩，如果必要，
    可以在属性列表中双击修改属性的值。（一般情况不需要再修改控件属性）。
    VisualTkinter尽量的将VB控件属性翻译成Tkinter控件属性，比如字体、颜色
    初始值、外观、状态等，甚至包括按钮类和菜单的快捷键设置等待。
    当然了，如果部分属性没有对应关系的，需要在VisualTkinter界面上设置。
  2.6 按“生成代码”按钮则在代码预览窗口生成代码，可以双击代码预览窗口
    放大阅读，也可以直接修改代码。
  2.7 确认完成后可以将代码拷贝到剪贴板或保持到文件。
    布局可以使用百分比定位（相对定位）或绝对坐标定位（按像素定位），
    百分比定位为有一个好处，主界面大小变化后，控件也可以相对变化大小。
    如果不希望控件大小变化，可以选择绝对坐标定位。
    注：如果修改了以前设计的界面，可以选择仅输出main函数或界面生成类。
    不影响外部已经实现的逻辑代码。
  2.8 如果程序有多个GUI界面，可以在VB工程中添加窗体，就可以选择输出
    哪个窗体。
  2.9 针对结构化代码，如果要在Python代码中引用和修改其他控件的值，
    可以使用全局字典gComps，这个字典保存了所有的GUI元素和一些对应的
    控件变量，可以直接使用形如gComps["Text1Var"].set("new Text")的代码
    来访问对应控件。
    如果输出的是面向对象代码，则可以在界面派生类Application中直接访问
    对应的控件。
  2.10 一般的GUI框架都会将UI部分和逻辑代码部分分别放在不同的文件中，在
    逻辑代码文件中导入UI文件，实现修改UI不影响逻辑代码。因为对于实现
    简单的程序来说，我偏爱单文件，所以我将UI类和逻辑代码类都放在同一个
    文件中，在修改界面后，你可以直接覆盖对应的Application_ui类即可实现
    界面的变更，不过如果增加了新的事件回调函数，需要在子类Application
    中增加才行。

3.目前支持的控件列表
  3.1 Label
    标签条在VB和Python中基本一样。
  3.2 TextBox
    Python文本框有两种：Entry和Text，如果VB的TextBox的MultiLine=False，则
    生成Entry，否则生成Text。
  3.3 Frame
    对应Python的LabelFrame控件，一样可以作为其他控件的容器，不过注意的一点
    是如果使用到了Frame控件，则建议使用坐标相对定位布局控件（代码生成选项），
    如果要使用绝对坐标，则VB设计窗体的ScaleMode要设置为3(vbPixels)，否则
    LabelFrame中的控件布局错误。这可能是VB的一个BUG，因为Frame控件能做容器，
    但不能设置容器内的坐标单位，默认固定为Twips。
    （相比之下，PictureBox能做容器，并且能设置ScaleMode）
  3.4 CommandButton
    对应Python的Button，没有太多区别。
    为了代码简洁，窗体的退出按钮可以设置Cancel属性为True，然后程序自动生成
    对应Tkinter的quit回调，这样就不需要再实现一个回调函数。
    在VB里面字母前增加一个"&"符号可以直接绑定一个快捷键Alt+对应字母，
    VisualTkinter也支持此设置，自动生成对应的事件绑定代码。
  3.5 CheckBox
    多选按钮对应Python的Checkbutton。
  3.6 OptionButton
    单选按钮对应Python的Radiobutton。
  3.7 ComboBox
    组合框在Tkinter中没有对应的控件，比较类似的只有OptionMenu，类似ComboBox
    的Style=2 (Dropdown List)时的表现，一个下拉列表，只能在列表中选择一个值，
    不能直接输入。所以建议在VB的ComboBox中写下所有的下拉列表值。
    如果启用了TTK主题扩展库支持，则直接对应到TTK的Combobox，外形和行为基本
    一致。
  3.8 ListBox
    列表框对应Python的Listbox，行为也类似，可以在设计阶段设置初始列表。
  3.9 HScrollBar, VScrollBar
    滚动条在Python中为Scrollbar，通过设置orient来控制水平还是垂直。
  3.10 Slider
    类似对应Python中的Scale。
  3.11 PictureBox
    简单对应到Python中的Canvas。
  3.12 Menu
    这下就可以使用VB的菜单编辑器来设计Python的菜单了。
    在VB中的菜单标题为"-"是分隔条。
    也可以在正常的菜单标题中增加(&+字母)的方式添加快捷
    方式。
  ===================================================
  以下的控件需要在VB的'控件工具箱'中按右键添加'部件'，选择
  'Microsoft Windows Common Controls 6.0'
  ====================================================
  3.13 ProgressBar
    对应到Python的Progressbar，需要启用TTK主题扩展（默认）
  3.14 TreeView
    对应到Python的Treeview，树形显示控件，可以选择是否显示标题行,
    需要启用TTK主题扩展（默认）
  3.15 TabStrip
    选项卡控件，对应到Python的Notebook，因为VB的IDE限制，无法在设计阶段就
    放置到每个选项卡内的控件，所以此工具简单生成两个默认选项卡，每个选项卡
    里面放了一个Label控件，你可以照此添加其他控件。
    或使用一个新的VB窗体布置好所需要控件，生成代码后手工拷贝到Python的对应
    选项卡的Frame控件代码下面。
    需要启用TTK主题扩展（默认）
  -----------------------------------------------------
  3.16 CommonDialog
    这个控件也算支持，如果VB窗体中有这个控件，则在Python代码中导入
    filedialog、simpledialog、colorchooser这三个模块，这三个模块提供简单的
    文件选择、输入框、颜色选择对话框功能。
    需要在控件工具箱增加"Microsoft Common Dialog Control 6.0"

  PS:打开TTK支持后，控件的有些属性设置无效，尽管可以设置，但是外观不变

4. 版本历史
  v1.2
    1.增加多语种支持，语言文件为VisualTkinter.dll目录下的language.lng，
      版本发布时支持简体中文、繁体中文、英文。
      如果没有语言文件，显示软件内置的简体中文。
    2.增加状态栏控件支持，因为TK和TTK都不支持Statusbar，就自己使用Label
      简单模拟了一个，支持多窗格，控件类定义直接添加到Python源码。
    3.支持主窗口的属性设置，比如图标等。
  v1.1
    1.增加TTK主题扩展库支持，代码不变，界面更漂亮，更Native
    2.增加进度条Progressbar,树形控件Treeview,选项卡控件Notebook
      这几个控件都需要TTK支持。
    3.更新安装程序，可以完整卸载此ADDIN了。
  v1.0 第一个版本
    支持控件列表：Label, Entry, LabelFrame, Button, Checkbutton, Radiobutton,
    OptionMenu, Combobox, Listbox, Scrollbar, Scale, Canvas, Menu
