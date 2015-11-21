# Introduce
The program is a addin of VB6(Visual Basic 6), It allows you design your GUI layout of Tkinter (standard GUI library of Python) in VB6 Integrated Development Environment.
The addin will generate python code of GUI framework for you. the only thing you need to do is that add your logical code in Callback of GUI framework.

![screenshot](https://raw.githubusercontent.com/cdhigh/tkinter-designer/master/Setup/Screenshots/TkinterDesigner_ScrPrnt_EN.JPG)

# Install and Usage
1. Install VB6.
    (if your os is windows 7 or later, please install VB6 SP6 version and 
    overwrite VB6.EXE using Vb6_SP6_Fix_for_Win7.exe)
    
2. Download Tkinter Designer from github.
    [https://github.com/cdhigh/tkinter-designer](https://github.com/cdhigh/tkinter-designer)
    
3. Copy directory 'Bin' to other directory you want, execute 'Setup for TkinterDesigner.exe' to register TkinterDesigner.dll. you may need language.lng too.

4. Open VB6, create a standard EXE project, design your layout firstly,
    and than click the icon 'Tkinter Designer' in toolbox of VB6.
    
5. Enjoy TkinterDesigner.

# List of VB Controls Supported
1. Label
    Has same appearance and behavior in Tkinter.
2. TextBox
    If the property MultiLine=False, the addin generate code for widget
    'Entry' of Tkinter, otherwise for widget 'Text'.
3. Frame
    Similar to widget LabelFrame of Tkinter. It can be container of other
    widget.
4. CommandButton
    It represent the widget Button of Tkinter.
    You can also add a '&' befor a corresponding letter to make a keyboard
    shortcut like 'Alt-letter', the addin generate code for shortcurts too.
5. CheckBox
    Similar to widget Checkbutton.
6. OptionButton
    Similar to widget Radiobutton.
7. ComboBox
    If the menu 'Use TTK library' checked, it translate to widget Combobox
    of TTK, otherwise, it translate to widget OptionMenu of Tkinter.
8. ListBox
    Similar to widget Listbox of tkinter. you can add list of text in IDE.
9. HScrollBar, VScrollBar
    The widget Scrollbar of tkinter.
10. Slider
    Similar to widget Scale of tkinter.
11. PictureBox
    It tranlate to Canvas.
12. Menu
    Has same appearance and behavior in Tkinter.
    you can create a separator by setting the caption of menuitem to '-'.
    Add a keyboard shortcut by using format of '&+letter'.
    
> Adding the component 'Microsoft Windows Common Controls 6.0' to toolbox of VB IDE to support the other controls.

13. ProgressBar
    Similar to widget Progressbar of TTK library.
14. TreeView
    Similar to widget Treeview of TTK library.
15. TabStrip
    Similar to widget Notebook of TTK libray.
    If you want to design the tabs of Notebook in VB IDE, please use a
    PictureBox or Frame as container for widgets, you have to name the 
    PictureBox or Frame following format:
    "Name of TabStrip" + "\__Tab" + sequence number(start in 1)
    For example, if name of TabStrip is 'TabStrip1'(by default), you can
    Create a PictureBox named 'TabStrip1\__Tab1' as container for Tab1 of 
    Notebook widget.

16. CommonDialog
    if the form has this control, then the addin will create code for 
    import modules filedialog, simpledialog, colorchooser.    

# History
* v1.5.2
    1. value of RadioButton changed to it's name.
* v1.5.1
    1. bugfix: Reference of some variable before defined. 
    2. Name of variable of RadioButton changed to 'ParentName' + 'RadioVar'.
    3. Save code file to disk using format utf-8 (no BOM).
* v1.5
    1. feature added: encode a file to Base64 string.
* v1.4.13
    1. Change codeline of judge of python's version for better compatibility.
* v1.4.12
    1. combox don't need property 'height' or 'relheight'.
* v1.4.11
    1. change the map of event 'xxx_MouseMove' to '<Motin> in tkinter.
* v1.4.10
    1. properties tag of form is working now.
* v1.4.9
    1. bugfix: decimal seperator from comma to point in latin language.
* v1.4.6
    1. Setup the coordinate of form in VB.
    2. set the font and color of caption of labelframe.
* v1.4.5
    1.Add a property 'scrollregion' to Canvas.
* v1.4.4
    1. Its a minor bugfix version.
* v1.4.3
    1. use properties tag of widgets to save same config. format is 
       p@property1@property2 or p@property1=value1@property2=value2
    2. Add a property 'startup position' to Form.
* v1.4.2
    1. bugfix:error when change the widget before update the configure of 
       widget.
    2. add properties spacing1, spacing2, spacing3 to widget 'Text'.
* v1.4.1
    1. Support scrollbar bind with widgets automatic.
    2. Control-Variable of Checkbutton changed to IntVar.
    3. textvariable of Combobox selected defaultly.
* v1.4
    1. Can used in VB6 green and compact version.
* v1.3.3
    1. Add properties 'Topmost' and 'Alpha' for Form.
* v1.3.2
    1. Support icon file without suffix.
    2. Add property 'cursor'.
* v1.3.1
    1. Bind callback functions via analysing VB code.
    2. Can configure widget OptionMenu now.
* v1.3
    1. Add a feature the allow design all tabs of Notebook widget in VB IDE.
       refer to description of TabStrip.
    2. Can add a lambda function in callback field.
    3. bugfix: add a wrong 'DeletE' event when you set the shortcut is 'Delete'.
    4. bugfix: substitute Combobox for OptionMenu widget when ttk is disabled.
    5. bugfix: some text disaperence when first time of display of dropdown.
* v1.2.8
    1. Add a process for banding of scrollbar and listbox/canvas/text widgets.
* v1.2.7
    1. Add family name of font to properties of widgets. (if need)
    2. improve the apparence of XP-Style button.
* v1.2.6
    1. Add feature 'preview'.
    2. Delete all settings saved in register when uninstall the addin.
    3. Substitute a XP-Style button control for shade-button.
* v1.2.5
    1. Add a option 'Add a prefix u to unicode string'.
    2. Bugfix: can't generate code for menu.
    3. the mainform is resizeable now, and substitute a shade-button for standard-button.
* v1.2.4
    Widget Notebook create tabs as same of TabStrip(VB)
* v1.2.3
    1. Add a property 'protocol' to Form.
    2. Delete the option 'Build file for XP style of VB6.EXE', you can do it
       youself by using google. the keyword is 'vb xp manifest' etc.
* v1.2.2
    Provide a combobox to choose some value in it.
* v1.2
    1.Supports feature multi-language.
    2.Add supports of Statusbar.
    3.Add properties of Form, such as icon etc.
* v1.1
    1.Add supports to TTK library.
    2.Add supports to widget Progressbar, Treeview, Notebook.
    3.Update setup program, can uninstall the addin now.
* v1.0 First version
    Supports : Label, Entry, LabelFrame, Button, Checkbutton, Radiobutton,
    OptionMenu, Combobox, Listbox, Scrollbar, Scale, Canvas, Menu

