Readme of english version refers to [Readme_EN.md](https://github.com/cdhigh/tkinter-designer/blob/master/Readme_EN.md)
--------------------------------

# 简介
这是一个VB6的Addin（外接程序），用于使用VB6开发工具直接拖放控件，直接可视化完成Python的Tkinter的GUI布局和设计，可以在VB界面上设置控件的大部分属性，最终自动生成必要的代码（包括回调函数框架），代码生成后仅需要在对应的回调函数中增加相应的逻辑功能代码即可。
![Screenshot](https://raw.githubusercontent.com/cdhigh/vb6tkinter/master/Setup/Screenshots/Vb6Tkinter_ScrPrnt.png)

这个工具支持绝大部分Tkinter控件，可应付一般GUI的需求。
（列表参见下面的控件说明）。



# 适用用户  
* 适用于学习了Tkinter并不想太麻烦手写GUI生成和排版代码，也不想用其他第三方工具和框架比如wxPython/PyQt的同学。
* 适用于界面不太复杂的小程序开发，界面复杂的还是适用wxPython/PyQt等框架吧。
* 因为Tkinter为Python标准库，使用Tkinter完成的Python程序可以称为"绿色软件"，不需要目标机器上安装wxPython/PyQt等框架，只要有Python的机器就能运行。即使使用pyinstaller/cx_freeze等打包成exe，最终文件也不会很大。
* 如果软件逻辑不是很复杂，通常一个 py 文件搞定，不像其他一些辅助框架，需要几个文件。    
  > （如果不希望py直接解释运行时弹出黑漆漆的命令行窗口，后缀名请改为pyw）     



# 安装   
有两个方式，可以任选一个：   
1. 整合安装方式：   
在 <https://github.com/cdhigh/Vb6Tkinter/releases> 下载整合了精简版VB6的安装包，已预先打好各种补丁，双击执行安装即可，不需要进行下面第二节的操作。   

2. 独立安装方式：  
2.1. 在 <https://github.com/cdhigh/Vb6Tkinter/releases> 下载插件压缩包，解压到某个目录。   
2.2. 先自行在其他网站下载并安装VB6，可以是完整版或[精简版](https://pan.baidu.com/s/1xwMNvUsaeWOexjdo-J3VAA?pwd=7db6)    
 2.2.1. 如果是 Windows 7 及以上版本, 用 Vb6_SP6_Fix_for_Win7.zip 中的 Vb6.exe 覆盖原来的Vb6.exe   
 2.2.2. 如果是 Windows 10 及以上版本，请应用 VB6_AppCompat 补丁    
2.3. 注册此插件，可以使用自带的注册工具"Setup for Vb6Tkinter.exe"，或自己手动完成。   
 2.3.1. 运行 `regsvr32 /s diretory\Vb6Tkinter.dll`   
 2.3.2. 在C:\Windows\VBADDIN.INI的段[Add-Ins32]增加一行：`Vb6Tkinter.Connect=3`     

3. 一旦你安装和注册完成后，以后的升级可以仅下载新版本的 Vb6Tkinter.dll/Vb6Tkinter.lng 覆盖旧的文件即可。




# 使用方法简介   
1. 打开VB6，新建一个标准EXE工程，在窗体上设计自己的GUI布局，这个工作估计没有VB基础的同学都可以完成，同时可以设置相应的控件属性。  
   如果不希望窗体大小可变，建议在VB6中修改窗体的BorderStyle属性为"Fixed Single"或"Fixed Dialog"。

2. 如果使用自带安装程序安装了插件，现在VB的工具条上应该有一个新图标（一片橙红色羽毛）， 
   如果没有，到菜单"外接程序"|"外接程序管理器"里面启动 Vb6Tkinter， Vb6Tkinter 图标和菜单应该会出现。
   
3. 启动Vb6Tkinter后，先按“刷新窗体列表”按钮，列出当前工程的所有窗体和控件列表。
   
4. 逐个确认各控件的输出属性，在要输出的选项前打钩，如果必要，可以在属性列表中双击修改属性的值。（一般情况不需要再修改控件属性）。
    Vb6Tkinter尽量的将VB控件属性翻译成Tkinter控件属性，比如字体、颜色、初始值、外观、状态等，甚至包括按钮类和菜单的快捷键设置等等。当然了，如果部分属性没有对应关系的，需要在Vb6Tkinter界面上设置。
   
5. 按“生成代码”按钮则在代码预览窗口生成代码，可以双击代码预览窗口放大阅读，也可以直接修改代码。

6. 确认完成后可以将代码拷贝到剪贴板或保存到文件。
    布局可以使用百分比定位（相对定位）或绝对坐标定位（按像素定位），百分比定位为有一个好处，主界面大小变化后，控件也可以相对变化大小。如果不希望主界面大小变化后控件跟随变化，可以选择绝对坐标定位。
    注：如果修改了以前设计的界面，可以选择仅输出main函数或界面生成类。不影响外部已经实现的逻辑代码。
7. 如果程序有多个GUI界面，可以在VB工程中添加窗体，就可以选择产生哪个窗体的对应代码。

8. 针对结构化代码，如果要在Python代码中引用和修改其他控件的值，可以使用全局字典gComps，这个字典保存了所有的GUI元素和一些对应的控件变量，可以直接使用形如gComps["Text1Var"].set("new Text")的代码来访问对应控件。
    如果输出的是面向对象代码，则可以在界面派生类Application中使用self.widgetName方式直接访问对应的控件。
    
9. 一般的GUI框架都会将UI部分和逻辑代码部分分别放在不同的文件中，在逻辑代码文件中导入UI文件，实现修改UI不影响逻辑代码。因为对于实现简单的程序来说，我偏爱单文件，所以我将UI类和逻辑代码类都放在同一个文件中，在修改界面后，你可以直接覆盖对应的Application_ui类即可实现界面的变更，不过如果增加了新的事件回调函数，需要在子类Application中增加才行。




# 目前支持的控件列表
1. **Form**
    VB6中的窗体(Form)对应到Tkinter的Frame，用于主窗体呈现。
    为了更好的效果，建议在VB6里面设置窗体的BorderStyle属性为"Fixed Single"或"Fixed Dialog"。
2. **Label**    
    标签条在VB和Python中基本一样。可以在文本中插入\n来换行，如果启用ttk，还可以设置wraplength属性。
3. **TextBox**    
    Python文本框有两种：Entry和Text，如果VB的TextBox的MultiLine=False，则
    生成Entry，否则生成Text。
4. **Frame**    
    对应Python的Frame或LabelFrame控件，做为其他控件的容器，或做为界面元素视觉分类。    
    有Caption时生成LabelFrame，否则生成Frame。
5. **CommandButton**    
    对应Python的Button，没有太多区别。
    为了代码简洁，窗体的退出按钮可以设置Cancel属性为True，然后程序自动生成
    对应Tkinter的destroy回调，这样就不需要再实现一个回调函数。
    在VB里面字母前增加一个"&"符号可以直接绑定一个快捷键 'Alt+对应字母'，
    Vb6Tkinter也支持此设置，自动生成对应的事件绑定代码。
    其他控件比如CheckBox等有"标题"属性的控件一样如此处理。
6. **CheckBox**    
    多选按钮对应Python的Checkbutton。
7. **OptionButton**    
    单选按钮对应Python的Radiobutton。
    tkinter中Radiobutton的分组方法和VB有些不一样（分组意味着组内的单选按钮自动
    互斥，用户选择一个则其他的自动取消）。
    在VB中，如果使用Frame将几个OptionButton圈起来，则这几个OptionButton自动成为一组。
    Vb6Tkinter也支持这样的操作，同一父控件的Radiobutton自动成为一组。
    如果需要手动处理，你要将需要分成一组的Radiobutton的variable属性设置为同一个变量，
    然后各个Radiobutton的value值要不一样（默认为控件名），具体是什么值你可以随便设置，
    反正不一样就行，最简单的就是1/2/3等，或者一个设置为man，另一个设置为woman，
    在对应的Radiobutton被选择后，variable变量自动设置为对应的value值，读取即可
    知道哪个Radiobutton被选中了，反之设置variable变量会导致对应的Radiobutton
    被选中。
8. **ComboBox**    
    组合框在Tkinter中没有对应的控件，比较类似的只有OptionMenu，类似ComboBox
    的Style=2 (Dropdown List)时的表现，一个下拉列表，只能在列表中选择一个值，
    不能直接输入。所以建议在VB的ComboBox中写下所有的下拉列表值。
    如果启用了TTK主题扩展库支持，则直接对应到TTK的Combobox，外形和行为基本
    一致。
9. **ListBox**    
    列表框对应Python的Listbox，行为也类似，可以在设计阶段设置初始列表。
    如果需要滚动，则在适当位置创建滚动条，如果滚动条紧靠着列表框的右边或下边，
    并且长度(水平滚动条)或高度(垂直滚动条)差不多，则滚动条和列表框自动绑定，
    如果没有自动绑定，则可以在Addin界面选择其xscrollcommand或yscrollcommand
    属性为对应滚动条的.set方法。
10. **HScrollBar, VScrollBar**    
    滚动条在Python中为Scrollbar，通过设置orient来控制水平还是垂直。
11. **PictureBox**    
    简单对应到Python中的Canvas，用做其他控件的容器或画图容器使用。
    如果需要滚动，则在适当位置创建滚动条，如果滚动条紧靠着图像框的右边或下边，
    并且长度(水平滚动条)或高度(垂直滚动条)差不多，则滚动条和图像框自动绑定，
    如果没有自动绑定，可以在Addin界面选择其xscrollcommand和yscrollcommand
    属性为对应滚动条的.set方法。
12. **Menu**    
    可以使用VB的菜单编辑器来设计Python的菜单。
    在VB中的菜单标题设置为"-"则创建分隔条。
    也可以在正常的菜单标题中增加(&+字母)的方式添加 'Alt快捷键'。
    除 'Alt快捷键' 外，在VB菜单编辑器中选择菜单对应的快捷键则会直接显示快捷键
    信息在菜单标题后面，并自动注册对应的bind命令。
13. **Line**    
    可以用于组织复杂界面，仅支持水平或垂直线。
    
	> *以下的控件需要在VB的'控件工具箱'中按右键添加'部件'，选择 "Microsoft Windows Common Controls 6.0"*

14. **Slider**    
    类似对应Python中的Scale。
15. **ProgressBar**    
    对应到Python的Progressbar，需要启用TTK主题扩展（默认）
16. **TreeView**    
    对应到Python的Treeview，树形显示控件，可以选择是否显示标题行,
    需要启用TTK主题扩展（默认）
    如果需要滚动，则在适当位置创建滚动条，如果滚动条紧靠着TreeView的右边或下边，
    并且长度(水平滚动条)或高度(垂直滚动条)差不多，则滚动条和TreeView自动绑定，
    如果没有自动绑定，可以在Addin界面选择其xscrollcommand和yscrollcommand
    属性为对应滚动条的.set方法。
17. **TabStrip**    
    选项卡控件，对应到Python的Notebook，需要启用TTK主题扩展（默认）。
    如果要布局各个页面内的控件，按以下步骤：    
    1. 每个选项页对应一个Frame或PictureBox，命名为：TabStrip的名字
    加 `__Tab` (双下划线)，再加一个序号，从1开始，比如TabStrip的名字为TabStrip1，
    则你可以在窗体上创建一个PictureBox，命名为 `TabStrip1__Tab1` (注意大小写)。    
    2. 然后在PictureBox/Frame内摆放你需要的其他控件，生成代码后此容器内自动添加
    到对应的选项页，Vb6Tkinter会在后台为您做这一切。
    标签页对应的PictureBox/Frame可以放置在窗体的可视范围外，也就是说设计好
    对应的选项页后，缩小IDE中的窗体为你需要的大小。    
    注意：    
    * 如果使用相对坐标，PictureBox或Frame容器的大小请和TabStrip内部大小一致或
    接近，否则选项页内的控件将会通过拉伸或收缩来适配可伸缩来适配可用空间，这样有些
    控件看起来会比较怪。如果使用绝对坐标，则PictureBox/Frame可以不用和TabStrip
    一样大，PictureBox/Frame内的控件将以TabStrip的左上角为原点放置，大小和长宽比例
    会和设计时一致。
    所以还是建议如果有TabStrip控件的话，使用绝对坐标。  
    * Frame和PictureBox均可作为容器，如果使用Frame作为容器，则其标题可以作为选项页
    标题，如果你没有设置选项页标题的话。（选项卡控件的标题设置优先）
18. **CommonDialog**    
    这个控件也算支持，如果VB窗体中有这个控件，则在Python代码中导入  
    filedialog、simpledialog、colorchooser这三个模块，这三个模块提供简单的  
    文件选择、输入框、颜色选择对话框功能。  
    需要在控件工具箱增加 "Microsoft Common Dialog Control 6.0"   




# 本应用对tkinter的扩展和"封装"  
1. tkinter没有Statusbar控件，所以我使用Label实现了一个简单的Statusbar，在VB窗体中添加Statusbar后会插入这部分实现代码。   
   （VB需要先添加“Microsoft Windows Common Controls 6.0”部件才有Statusbar）
2. tkinter控件没有Tooltip(鼠标悬停提示)属性，所以我实现了一个简单的Tooltip类给tkinter控件加上Tooltip功能，设置VB控件的
   ToolTipText属性即自动给对应的控件创建一个Tooltip类。ToolTipText支持自动换行，如果要手动控制换行，也可以使用 '\n'。
3. 隐藏反人类的TK控件设置和获取控件显示值的机制（需要通过控件变量textvariable），给Entry/Label/Button/Checkbutton/Radiobutton控件
   添加更符合直觉的 setText()/text() 方法，可以直接设置和获取其控件显示的文本值。    
   （Checkbutton/Radiobutton默认不添加，如需要，可以在Vb6Tkinter将textvariable打勾，因为很少需要运行时修改这两个控件的文本。）
4. 同样类似第三条，给Checkbutton/Radiobutton添加 setValue()/value() 方法，参数为1/0。    
   `self.Text1.setText('new text')`    
   `print(self.Text1.text())`    
   `print(self.Check1.value())`    
   `self.Option1.setValue(1)`    
   `print(self.Option1.value())`   




# 其他建议  
1. 我最喜欢VB IDE的窗体设计器的其中一点是它有一个菜单项 "**格式**"，用法是选择多个控件后， 通过此菜单项可以设置控件的大小对齐相对位置等，善用这个功能在界面设计中非常省心省力，而且界面会更美观。
   * 对齐： 左对齐 顶对齐等
   * 统一尺寸： 宽度相同，高度相同，两者都相同
   * 水平间距： 相同间距，递增，递减等
   * 垂直间距： 相同间距，递增，递减等
   * 在窗体中居中对齐：水平对齐，垂直对齐
2. 插件不支持使用控件数组，界面可以显示，但是后面的同名控件名会覆盖前面定义的，导致在代码中无法再和此控件互动。
3. VB窗体的ScaleMode建议保持默认值(vbTwips)，如果要设置为其他值，则Frame控件内就不要再放Frame控件了，否则其内部的控件布局错误。
4. 如果需要简体汉字界面，则需要Vb6Tkinter.lng文件在Vb6Tkinter.dll同一目录。
5. 界面设计往往需要经历多次修改，如果某些属性不使用Vb6Tkinter生成的默认值，则有可能下次修改界面时忘记同步修改，导致界面错误，此时可以修改对应控件的Tag属性为下面两个格式任何一个：
- `p@属性1@属性2@属性n`
- `p@属性1=值1@属性2=值2@属性n=值n`
每个属性的值可以忽略，忽略值属性则自动选中对应属性，不修改值。
6. 此插件支持更多的一些便捷特性，分散在版本历史中，这里就不一一列举了，有需要的可以去了解。




# ttk库额外说明
  ttk是tkinter的主题扩展库，也是标准库，看起来很漂亮，在不同操作系统下界面呈现为本地化风格，建议使用，    
  只是要注意以下几个ttk的BUG：    
1. TTK的Entry和Combobox控件背景色设置无效（可以设置，不报错，但是界面不变）。
2. LabelFrame和Notebook控件的字体单独设置无效，但是可以设置ttk的全局字体属性来改变，比如：self.style.configure('.', font=('宋体',12))。




# 版本历史
*  v1.9
    1. 支持生成GUI国际化代码（gettext）
*  v1.8.1
    1. 根据Caption属性自动转换Frame为LabelFrame或Frame
*  v1.8
    1. 删除结构化代码生成逻辑，仅生成面向对象代码
    2. ListBox添加 (MultiSelect -> selectmode) 属性对应转换
*  v1.7.2
    1. bugfix: 面向过程代码无法预览
*  v1.7.1  
    1. Combobox的Change事件映射到Tkinter的ComboboxSelected
    2. 多行Text的Change事件映射到Tkinter的Modified
    3. 修正base64编码的一个小bug，之前的版本将末尾填充的=去掉了，导致有些解码库报错
    4. 修正生成的状态栏代码缺了一部分的缺陷
*  v1.7
    1. 插件名从 TkinterDesigner 更名为 Vb6Tkinter
    2. 增加在线检查版本更新功能
    3. 如果仅需要Python3代码，修改部分输出的代码更符合Python3习惯
    4. 默认取消配置项"兼容Python2/3代码"，默认输出Python3代码
*  v1.6.8
    1. Notebook控件的鼠标点击事件对应到 <<NotebookTabChanged>> 事件
*  v1.6.7
    1. 如窗体设置为屏幕中心，将设置窗体位置代码移动到初始化函数开头，避免启动时窗体有一个移动过程。
    2. 删除生成的代码中的部分提示信息。
*  v1.6.6
    1. 如果目录下存在 icon.gif 文件，则自动用做窗体图标(编码为base64嵌入代码)。
*  v1.6.5
    1. 如果生成的代码超过65k，则不显示在文本框中，直接要求保存到磁盘文件。
*  v1.6.4
    1. Listbox在VB中有Click函数的话，自动生成 <<ListboxSelect>> 绑定。
*  v1.6.3
    1. 支持设计时在窗体上设置分组中其中一个Radiobutton的选中状态。
*  v1.6.2
    1. 支持ToolTipText属性，实现鼠标悬停弹出提示框功能，Vb6Tkinter没有Tooltip功能，自己实现了一个Tooltip类来支持。
*  v1.6.1
    1. 给Checkbutton/Radiobutton添加setValue()/value()函数，参数为整型1/0。
*  v1.6
    1. 简单扩展tkinter，隐藏控件的textvariable操作，给控件动态添加setText()/text()函数用于设置和获取对应的字符串。
*  v1.5.2
    1. RadioButton的value值默认设置为其控件名，这样直接xxxRadioVar.set(控件名) 就可以选择对应单选框。
*  v1.5.1
    1. bugfix:修正在特定条件下因控件创建顺序问题导致Python变量在定义前引用的错误。
    2. RadioButton的variable名字修改为ParentName + 'RadioVar'，同时value默认为TabIndex。
    3. 保存代码文件的格式从带BOM的utf-8修改为不带BOM的utf-8。
*  v1.5
    1. 添加一个功能：可以将一个磁盘文件编码为Base64字符串，可以用于将一些资源文件保存到python源文件中。
*  v1.4.13
    1. 改变判断python版本的代码，增强兼容性。
*  v1.4.12
    1. combox不需要设置relheight或height属性。
*  v1.4.11
    1. 修改_MouseMove 事件对应为tkinter的<Motion>事件。
*  v1.4.10
    1. 增加窗体的Tag标签处理。
*  v1.4.9
    1. 修正拉丁语系环境下控件小数点变成逗号的问题。
*  v1.4.6
    1. 可以在VB中设置窗体初始坐标。
    2. bugfix:修正LabelFrame的标题字体和颜色设置无效的问题。
*  v1.4.5
    1. Canvas控件增加scrollregion属性。
*  v1.4.4
    1. bugfix:在windows7 64bit下获取python安装目录失败的bug（用于预览界面）。
    2. bugfix:在窗体目录下同时放置一个ico和一个gif文件时窗体图标文件设置错误。
*  v1.4.3
    1. 增加一个方便一段时间后再次修改GUI的特性：使用控件的Tag属性来保存修改的值。
       方法是如果有一些属性不采用默认值，则在Tag属性中采用如下格式填写：
       p@属性1@属性2@属性n 或 p@属性1=值1@属性2=值2@属性n=值n
       每个属性的值是可以忽略的，忽略了值的属性则自动选中对应属性，不修改值。
    2. 增加窗体启动位置属性，可选择启动时在屏幕上居中。
*  v1.4.2
    1. bugfix:修正在配置列表框中选择下列列表时未更新就切换控件导致错误的问题。
    2. Text控件增加spacing1,spacing2,spacing3属性。
*  v1.4.1
    1. 支持自动绑定滚动条到对应控件，只需要在需要滚动的控件右边或下边紧靠着放置合适
       长度的滚动条，则滚动条自动绑定之对应控件，不需要再手工选择配置。
    2. Checkbutton的控件变量由StringVar类型改变为IntVar类型（Tkinter默认）。
    3. ComboBox的textvariable变量默认选择。
*  v1.4
    1. 支持VB绿色精简版。
*  v1.3.3
    1. 增加窗体的TOPMOST/ALPHA属性设置。
*  v1.3.2
    1. 如果VB窗体目录下有一个ico/gif文件，则自动将其作为窗体图标。
       （注意：如果目录有多个图标文件，则你要自己在下拉列表中选择一个。）
    2. 支持没有后缀名的主窗体图标（需要手动填写图标文件名）。
    3. 增加cursor属性，用于设置控件的鼠标指针。
    4. Form增加bindcommand和windowstate的处理。
    5. 按钮类控件的下划线回调函数使用tk内置的invoke()代替外部实现的xxx_Cmd()，
       使用invoke()为模拟用户点击，有更好的视觉反馈效果。
    6. bugfix: 修正Radiobutton分组时variable变量重复创建的BUG。
    7. bugfix: 修改Scale的digits等几个属性在ttk样式和创建函数中重复出现的问题。
*  v1.3.1
    1. 增加对VB代码的简单分析，代码中有对应控件的一些事件处理函数则自动生成
       tkinter对应的事件注册和回调框架，比如如果VB代码存在Text1_Change函数，则
       自动注册和生成控件Text1的Change事件处理回调函数。
    2. 增加对OptionMenu控件的属性设置，适当的参数调整可以让OptionMenu更美观。
*  v1.3
    1. 增加对Line控件的支持，可用于组织界面，内部实现为Separator控件，仅支持
       水平或垂直样式，如果在VB窗体上画了斜线，则使用其在水平方向或垂直方向的
       投影。需要启用TTK主题库。
    2. 增加一个重要特性：可以拖放设计Notebook(选项卡控件)的各选项页内控件。
       方法和步骤参加上面的TabStrip控件说明，简单来说就是使用PictureBox或
       Frame控件来作为各选项页的容器设计，命名类似：TabStrip1__Tab1等。
       这个特性让此ADDIN设计复杂界面成为可能，因为很多复杂的GUI用到选项卡
       控件来整理其他小控件，特别是各种配置页面。
    3. 控件的命令回调函数可以直接使用匿名函数lambda。
    4. 完善控件的字体处理，现在除ttk.LabelFrame和Notebook控件因ttk库的BUG外，
       其余控件均已实现字体的完美处理。
    5. 增加Treeview的滚动条绑定处理。
    6. 增加代码处理Frame控件的ScaleMode一直保持为vbTwips的BUG，现在可以允许
       窗体存在Frame的情况下设置窗体ScaleMode和使用绝对坐标定位。
    7. 增加系统颜色翻译成tkinter颜色的处理，现在控件颜色可以选择各种系统颜色，
       或在调色板内直接选择。
    6. bugfix：将全局菜单快捷键Delete写成DeletE的问题。
    7. bugfix：如果ADDIN启动时就没有启用TTK，并且在产生代码前没有修改TTK选项，
       则ADDIN还是使用Combobox代替OptionMenu，而tkinter没有Combobox控件。
    8. bugfix：解决自定义列表框中'第一次'显示下拉组合框时数据显示不全的问题。
    9. 还有一些小的修正。
*  v1.2.8
    1. 增加滚动条和列表框/多行文本框/图片框的绑定处理，方法是在窗体上对应
       控件的旁边放上滚动条，然后在ADDIN界面的控件属性xscrollcommand和
       yscrollcommand中选择对应滚动条的set方法即可。
*  v1.2.7
    1. 控件选项增加'字体名'属性处理（之前的版本仅处理大小粗体斜体属性）。
    2. 完善XP风格按钮外观，增加键盘操作。
*  v1.2.6
    1. 增加界面预览功能。
    2. 卸载程序一并删除注册表中保存的配置项，保证完全卸载。
    3. 换了一个清爽一点的XP风格按钮。
*  v1.2.5
    1. 增加一个选项：'Unicode字符串增加前缀u'，注意：增加前缀u会无法兼容
       python3.3之前的3.x版本(3.3支持u前缀了)，所以针对2.x的UNICODE字符串，
       可以引入：from __future__ import unicode_literals，实现2.x/3.x的
       UNICODE字符串兼容。
    2. Bugfix: 修正生成菜单代码失败的问题(v1.1引入的BUG)
    3. 界面美化：渐变按钮，窗体大小可改变。
*  v1.2.4
    Notebook控件根据TabStrip控件(VB)的选项卡设置自己的选项卡数量
*  v1.2.3
    1. 增加窗体消息拦截属性，可以拦截窗体消息，比如可以禁止窗体关闭按钮等。
    2. 删除安装程序中设置VB的IDE为XP样式的代码，以避免360误报有病毒，如果
       需要VB的IDE为XP样式，可以自己在网上找一个manifest文件改名为
       VB6.EXE.manifest，放到VB6目录下。
*  v1.2.2
    1. 对应一些属性值，如果只有有限的可选值，则可以在下列列表中选择。
*  v1.2
    1. 增加多语种支持，语言文件为Vb6Tkinter.dll目录下的Vb6Tkinter.lng，
       版本发布时支持简体中文、英文。
       如果没有语言文件，显示软件内置的简体中文。
    2. 增加状态栏控件支持，因为TK和TTK都不支持Statusbar，就自己使用Label
       简单模拟了一个，支持多窗格，控件类定义直接添加到Python源码。
    3. 支持主窗口的属性设置，比如图标等。
*  v1.1
    1. 增加TTK主题扩展库支持，代码不变，界面更漂亮，更Native
    2. 增加进度条Progressbar,树形控件Treeview,选项卡控件Notebook
       这几个控件都需要TTK支持。
    3. 更新安装程序，可以完整卸载此ADDIN了。
*  v1.0 第一个版本
    支持控件列表：Label, Entry, LabelFrame, Button, Checkbutton, Radiobutton,
    OptionMenu, Combobox, Listbox, Scrollbar, Scale, Canvas, Menu

  