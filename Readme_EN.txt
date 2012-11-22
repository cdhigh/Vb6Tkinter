1.Introduce
    The program is a addin of VB6(Visual Basic 6), It allow you design
    your GUI layout of Tkinter (standard GUI library of Python) in VB6 
    Integrated Development Environment.
    The addin will generate python code of GUI framework for you. the 
    only thing you need to do is add your logical code in Callback of 
    GUI framework.

2. List of VB Controls Supported
  2.1 Label
    Has same appearance and behavior in Tkinter.
  2.2 TextBox
    If the property MultiLine=False, the addin generate code for widget
    'Entry' of Tkinter, otherwise for widget 'Text'.
  2.3 Frame
    Similar to widget LabelFrame of Tkinter. It can be container of other
    widget.
  2.4 CommandButton
    It represent the widget Button of Tkinter.
    You can also add a '&' befor a corresponding letter to make a keyboard
    shortcut like 'Alt-letter', the addin generate code for shortcurts too.
  2.5 CheckBox
    Similar to widget Checkbutton.
  2.6 OptionButton
    Similar to widget Radiobutton.
  2.7 ComboBox
    If the menu 'Use TTK library' checked, it translate to widget Combobox
    of TTK, otherwise, it translate to widget OptionMenu of Tkinter.
  2.8 ListBox
    Similar to widget Listbox of tkinter. you can add list of text in IDE.
  2.9 HScrollBar, VScrollBar
    The widget Scrollbar of tkinter.
  2.10 Slider
    Similar to widget Scale of tkinter.
  2.11 PictureBox
    It tranlate to Canvas.
  2.12 Menu
    Has same appearance and behavior in Tkinter.
    you can create a separator by setting the caption of menuitem to '-'.
    Add a keyboard shortcut by using format of '&+letter'.
  ===================================================
  Adding the component 'Microsoft Windows Common Controls 6.0' to toolbox 
  of VB IDE to support the other controls.
  ====================================================
  2.13 ProgressBar
    Similar to widget Progressbar of TTK library.
  2.14 TreeView
    Similar to widget Treeview of TTK library.
  2.15 TabStrip
    Similar to widget Notebook of TTK libray.
  -----------------------------------------------------
  2.16 CommonDialog
    if the form has this control, then the addin will create code for 
    import modules filedialog, simpledialog, colorchooser.    

3. History
  v1.2.6
    1. Add feature 'preview'.
    2. Delete all settings saved in register when uninstall the addin.
    3. Substitute a XP-Style button control for shade-button.
  v1.2.5
    1. Add a option 'Add a prefix u to unicode string'.
    2. Bugfix: can't generate code for menu.
    3. the mainform is resizeable now, and substitute a shade-button for standard-button.
  v1.2.4
    Widget Notebook create tabs as same of TabStrip(VB)
  v1.2.3
    1. Add a property 'protocol' to Form.
    2. Delete the option 'Build file for XP style of VB6.EXE', you can do it
       youself by using google. the keyword is 'vb xp manifest' etc.
  v1.2.2
    Provide a combobox to choose some value in it.
  v1.2
    1.Supports feature multi-language.
    2.Add supports of Statusbar.
    3.Add properties of Form, such as icon etc.
  v1.1
    1.Add supports to TTK library.
    2.Add supports to widget Progressbar, Treeview, Notebook.
    3.Update setup program, can uninstall the addin now.
  v1.0 First version
    Supports : Label, Entry, LabelFrame, Button, Checkbutton, Radiobutton,
    OptionMenu, Combobox, Listbox, Scrollbar, Scale, Canvas, Menu

