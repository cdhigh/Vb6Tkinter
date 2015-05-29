VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBaseControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'实现所有控件类的公用代码，因为VB不支持真正的继承，所以其他控件类将内嵌这个类，而不是继承

Private m_dic As Dictionary                                                     '当前要生成代码的属性/值对
Private m_Type As String                                                        '直接对应到PYTHON的控件类型
Private m_Name As String                                                        '控件名
Private m_Parent As String
Private m_Value As String                                                       ' 控件值（如果与的话）
Private m_StyleName As String                                                   '样式基类
Private m_ScaleMode As Long

'输出PYTHON代码,
'sCmdFunc: 输出参数，事件处理回调代码;
'rel:是否使用相对坐标,
'oop:是否使用面向对象编程
'usettk:是否使用TTK主题扩展
Public Sub toString(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean)
    
    If oop Then
        toStringOOP sOut, sCmdFunc, rel, IIf(Len(m_StyleName), usettk, False)
    Else
        toStringStructure sOut, sCmdFunc, rel, IIf(Len(m_StyleName), usettk, False)
    End If
    
End Sub

'输出结构化代码
Public Sub toStringStructure(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, usettk As Boolean, Optional sOtherParams As String = "")
    
    Dim s() As String, i As Long, extra As String, sStyle As String
    Dim sTmp As String, dTmp As Double, sCmdName As String, sUnderlineCmd As String
    
    '如果需要变量绑定，则先创建对应的变量
    If Len(Dic("textvariable")) Then
        sOut.Append Space(4) & Dic("textvariable") & " = StringVar(value=" & U(Dic("text")) & ")"
    End If
    If Len(Dic("variable")) Then
        If m_Type = "Radiobutton" Then  '一组单选按钮的variable都是同一个，不要重复创建
            If Not sOut.ExistString(Space(4) & Dic("variable") & " = StringVar()") Then
                sOut.Append Space(4) & Dic("variable") & " = StringVar()"
            End If
        ElseIf m_Type = "Checkbutton" Then '单选按钮的variable使用IntVar比较好
            sOut.Append Space(4) & Dic("variable") & " = IntVar(value=" & m_Value & ")"
        Else
            sOut.Append Space(4) & Dic("variable") & " = StringVar(value=" & U(m_Value) & ")"
        End If
    End If
    If Len(Dic("listvariable")) Then
        sOut.Append Space(4) & Dic("listvariable") & " = StringVar(value=" & U(m_Value) & ")"
    End If
    If Not usettk And Len(Dic("font")) Then
        If Left$(Dic("font"), 1) = "(" Then
            sOut.Append Space(4) & m_Name & "Font = Font(font=" & Dic("font") & ")"
        Else '如果用户忘了加括号，这里加上
            sOut.Append Space(4) & m_Name & "Font = Font(font=(" & Dic("font") & "))"
        End If
    End If
    
    If Len(Dic("columns")) Then
        sOut.Append Space(4) & Dic("columns") & " = []  " & L("l_cmtTodoCols", "#TODO在这里添加标题列表，第一列固定为树形显示")
    End If
    If Len(Dic("displaycolumns")) And InStr(1, Dic("displaycolumns"), "#all") <= 0 Then
        sOut.Append Space(4) & Dic("displaycolumns") & "[]    " & L("l_cmtTodoDisCols", "#TODO在这里添加显示标题列表，第一列固定为树形显示")
    End If
    
    '组合框将displayrows转换为height属性
    If Len(Dic("displayrows")) Then
        sOtherParams = sOtherParams & IIf(Len(sOtherParams), ", ", "") & "height=" & Dic("displayrows")
    End If
    
    If usettk Then                                                              '创建STYLE对象
        sStyle = GetStyleParams(False, usettk)
        If Len(sStyle) Then
            sOut.Append Space(4) & "style.configure('T" & m_Name & "." & m_StyleName & "', " & sStyle & ")"
            If m_Type = "LabelFrame" Then ' LabelFrame的字体和前景色要设置到Label
                sOut.Append Space(4) & "style.configure('T" & m_Name & "." & m_StyleName & ".Label" & "', " & sStyle & ")"
            End If
            sOtherParams = sOtherParams & IIf(Len(sOtherParams), ", ", "") & "style='T" & m_Name & "." & m_StyleName & "'"
        End If
    End If
    
    extra = GetExtraParams(False, usettk)
    extra = extra & IIf(Len(extra) > 0 And Len(sOtherParams) > 0, ", ", "") & sOtherParams
    sOut.Append Space(4) & m_Name & " = " & m_Type & "(" & m_Parent & IIf(Len(extra), ", ", "") & extra & ")"
    sOut.Append Space(4) & m_Name & ".place(" & GetPositionParams(rel) & ")"
    sOut.Append Space(4) & "gComps['" & m_Name & "'] = " & m_Name
    
    '如果需要变量绑定，为了方便在main()函数外使用，同时将变量引用放入全局控件字典
    If Len(Dic("textvariable")) Then
        sOut.Append Space(4) & "gComps['" & Dic("textvariable") & "'] = " & Dic("textvariable")
    End If
    If Len(Dic("variable")) Then
        sOut.Append Space(4) & "gComps['" & Dic("variable") & "'] = " & Dic("variable")
    End If
    If Len(Dic("listvariable")) Then
        sOut.Append Space(4) & "gComps['" & Dic("listvariable") & "'] = " & Dic("listvariable")
    End If
    
    '创建事件处理函数框架（如果有），事件里面有点号说明是调用系统的函数，不需要生成函数体
    If Len(Dic("command")) > 0 And InStr(1, Dic("command"), ".") <= 0 And InStr(1, Dic("command"), "lambda ") <= 0 Then
        sCmdFunc.Append CreateFuncDef(Dic("command"), "event=None")
    End If
    If Len(Dic("postcommand")) Then
        sCmdFunc.Append CreateFuncDef(Dic("postcommand"), "event=None")
    End If
    
    '处理下划线快捷方式
    If Len(Dic("underline")) > 0 And Dic("underline") <> "-1" And IsNumeric(Dic("underline")) Then
        If m_Type = "Button" Or m_Type = "Checkbutton" Or m_Type = "Radiobutton" Then
            sUnderlineCmd = "lambda e: " & m_Name & ".focus_set() or " & m_Name & ".invoke()"
        ElseIf Len(Dic("command")) > 0 Then
            sUnderlineCmd = Dic("command")
        Else
            sUnderlineCmd = "lambda e: " & m_Name & ".focus_set()"
        End If
        
        If Len(sUnderlineCmd) Then
            sTmp = Dic("text")
            If Len(sTmp) = 0 Then sTmp = Dic("label")
            If Len(sTmp) And CLng(Dic("underline")) < Len(sTmp) Then
                sOut.Append Space(4) & WTOP & ".bind_all('<Alt-" & Mid(sTmp, CLng(Dic("underline")) + 1, 1) & ">', " & sUnderlineCmd & ")"
                If Mid(sTmp, CLng(Dic("underline")) + 1, 1) >= "a" And Mid(sTmp, CLng(Dic("underline")) + 1, 1) <= "z" Then
                    sOut.Append Space(4) & WTOP & ".bind_all('<Alt-" & UCase(Mid(sTmp, CLng(Dic("underline")) + 1, 1)) & ">', " & sUnderlineCmd & ")"
                ElseIf Mid(sTmp, CLng(Dic("underline")) + 1, 1) >= "A" And Mid(sTmp, CLng(Dic("underline")) + 1, 1) <= "Z" Then
                    sOut.Append Space(4) & WTOP & ".bind_all('<Alt-" & LCase(Mid(sTmp, CLng(Dic("underline")) + 1, 1)) & ">', " & sUnderlineCmd & ")"
                End If
            End If
        End If
    End If
    
    If Len(Dic("bindcommand")) Then                                             '有需要使用bind语句绑定的其他事件处理
        sTmp = Dic("bindcommand")
        sTmp = Replace(sTmp, "'", "")                                           '自动去掉括号和空格，如果有的话
        sTmp = Replace(sTmp, Chr(34), "")
        sTmp = Replace(sTmp, " ", "")
        
        s = Split(sTmp, ",")
        For i = 0 To UBound(s)
            s(i) = Trim(s(i))
            If s(i) = "<Change>" Then '专门处理自定义的这个事件，用控件变量监视器模拟
                If (m_Type = "Combobox" Or m_Type = "Entry" Or m_Type = "Label") And (Len(Dic("textvariable")) > 0) Then
                    sCmdName = m_Name & "_Change"
                    sOut.Append Space(4) & Dic("textvariable") & ".trace('w', " & sCmdName & ")"
                    sCmdFunc.Append CreateFuncDef(sCmdName, "*args")
                End If
            ElseIf Left(s(i), 1) = "<" And Right(s(i), 1) = ">" Then
                
                sCmdName = m_Name & "_" & Replace(Replace(Replace(s(i), "<", ""), ">", ""), "-", "_")
                
                sOut.Append Space(4) & m_Name & ".bind('" & s(i) & "', " & sCmdName & ")"
                
                sCmdFunc.Append CreateFuncDef(sCmdName, "event")
                
                'Python是大小写敏感的，对应快捷键也一样，如果设置的快捷键包含字母键，则将对应的大写/小写也一起绑定
                If Right(s(i), 3) >= "-a>" And Right(s(i), 3) <= "-z>" Then
                    
                    s(i) = Left(s(i), Len(s(i)) - 2) & UCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '对应字母变大写
                    sOut.Append Space(4) & m_Name & ".bind('" & s(i) & "', " & sCmdName & ")"
                    
                ElseIf Right(s(i), 3) >= "-A>" And Right(s(i), 3) <= "-Z>" Then
                    
                    s(i) = Left(s(i), Len(s(i)) - 2) & LCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '对应字母变小写
                    sOut.Append Space(4) & m_Name & ".bind('" & s(i) & "', " & sCmdName & ")"
                    
                End If
            End If
        Next
    End If
    
End Sub

'输出面向对象代码
Public Sub toStringOOP(ByRef sOut As cStrBuilder, ByRef sCmdFunc As cStrBuilder, rel As Boolean, usettk As Boolean, Optional sOtherParams As String = "")
    
    Dim s() As String, i As Long, extra As String
    Dim sTmp As String, sStyle As String, sCmdName As String, sUnderlineCmd As String
    
    '如果需要变量绑定，则先创建对应的变量
    If Len(Dic("textvariable")) Then
        sOut.Append Space(8) & "self." & Dic("textvariable") & " = StringVar(value=" & U(Dic("text")) & ")"
    End If
    If Len(Dic("variable")) Then
        If m_Type = "Radiobutton" Then  '一组单选按钮的variable都是同一个，不要重复创建
            If Not sOut.ExistString(Space(8) & "self." & Dic("variable") & " = StringVar()") Then
                sOut.Append Space(8) & "self." & Dic("variable") & " = StringVar()"
            End If
        ElseIf m_Type = "Checkbutton" Then '单选按钮的variable使用IntVar比较好
            sOut.Append Space(8) & "self." & Dic("variable") & " = IntVar(value=" & m_Value & ")"
        Else
            sOut.Append Space(8) & "self." & Dic("variable") & " = StringVar(value=" & U(m_Value) & ")"
        End If
    End If
    If Len(Dic("listvariable")) Then
        sOut.Append Space(8) & "self." & Dic("listvariable") & " = StringVar(value=" & U(m_Value) & ")"
    End If
    If Not usettk And Len(Dic("font")) Then
        If Left$(Dic("font"), 1) = "(" Then
            sOut.Append Space(8) & "self." & m_Name & "Font = Font(font=" & Dic("font") & ")"
        Else
            sOut.Append Space(8) & "self." & m_Name & "Font = Font(font=(" & Dic("font") & "))"
        End If
    End If
    
    If Len(Dic("columns")) Then
        sOut.Append Space(8) & "self." & Dic("columns") & " = []    " & L("l_cmtTodoCols", "#TODO在这里添加标题列表，第一列固定为树形显示")
    End If
    If Len(Dic("displaycolumns")) And InStr(1, Dic("displaycolumns"), "#all") <= 0 Then
        sOut.Append Space(8) & "self." & Dic("displaycolumns") & " = []    " & L("l_cmtTodoDisCols", "#TODO在这里添加显示标题列表，第一列固定为树形显示")
    End If
    
    '组合框将displayrows转换为height属性
    If Len(Dic("displayrows")) Then
        sOtherParams = sOtherParams & IIf(Len(sOtherParams), ", ", "") & "height=" & Dic("displayrows")
    End If
    
    If usettk Then                                                              '创建STYLE对象
        sStyle = GetStyleParams(True, usettk)
        If Len(sStyle) Then
            sOut.Append Space(8) & "self.style.configure('T" & m_Name & "." & m_StyleName & "', " & sStyle & ")"
            If m_Type = "LabelFrame" Then ' LabelFrame的字体和前景色要设置到Label
                sOut.Append Space(8) & "self.style.configure('T" & m_Name & "." & m_StyleName & ".Label" & "', " & sStyle & ")"
            End If
            sOtherParams = sOtherParams & IIf(Len(sOtherParams), ", ", "") & "style='T" & m_Name & "." & m_StyleName & "'"
        End If
    End If
    
    extra = GetExtraParams(True, usettk)
    extra = extra & IIf(Len(extra) > 0 And Len(sOtherParams) > 0, ", ", "") & sOtherParams
    sOut.Append Space(8) & "self." & m_Name & " = " & m_Type & "(self." & m_Parent & IIf(Len(extra), ", ", "") & extra & ")"
    sOut.Append Space(8) & "self." & m_Name & ".place(" & GetPositionParams(rel) & ")"
    
    '创建事件处理函数框架（如果有），事件里面有点号说明是调用系统的函数，不需要生成函数体
    If Len(Dic("command")) > 0 And InStr(1, Dic("command"), ".") <= 0 And InStr(1, Dic("command"), "lambda ") <= 0 Then
        sCmdFunc.Append CreateFuncDefOOP(Dic("command"), "event=None")
    End If
    If Len(Dic("postcommand")) Then
        sCmdFunc.Append CreateFuncDefOOP(Dic("postcommand"), "event=None")
    End If
    
    '处理下划线快捷方式
    If Len(Dic("underline")) > 0 And Dic("underline") <> "-1" And IsNumeric(Dic("underline")) Then
        If m_Type = "Button" Or m_Type = "Checkbutton" Or m_Type = "Radiobutton" Then
            sUnderlineCmd = "lambda e: self." & m_Name & ".focus_set() or self." & m_Name & ".invoke()"
        ElseIf Len(Dic("command")) > 0 Then
            sUnderlineCmd = "self." & Dic("command")
        Else
            sUnderlineCmd = "lambda e: self." & m_Name & ".focus_set()"
        End If
        
        If Len(sUnderlineCmd) Then
            sTmp = Dic("text")
            If Len(sTmp) = 0 Then sTmp = Dic("label")
            If Len(sTmp) And CLng(Dic("underline")) < Len(sTmp) Then
                sOut.Append Space(8) & "self." & WTOP & ".bind_all('<Alt-" & Mid(sTmp, CLng(Dic("underline")) + 1, 1) & ">', " & sUnderlineCmd & ")"
                If Mid(sTmp, CLng(Dic("underline")) + 1, 1) >= "a" And Mid(sTmp, CLng(Dic("underline")) + 1, 1) <= "z" Then
                    sOut.Append Space(8) & "self." & WTOP & ".bind_all('<Alt-" & UCase(Mid(sTmp, CLng(Dic("underline")) + 1, 1)) & ">', " & sUnderlineCmd & ")"
                ElseIf Mid(sTmp, CLng(Dic("underline")) + 1, 1) >= "A" And Mid(sTmp, CLng(Dic("underline")) + 1, 1) <= "Z" Then
                    sOut.Append Space(8) & "self." & WTOP & ".bind_all('<Alt-" & LCase(Mid(sTmp, CLng(Dic("underline")) + 1, 1)) & ">', " & sUnderlineCmd & ")"
                End If
            End If
        End If
    End If
    
    If Len(Dic("bindcommand")) Then                                             '有需要使用bind语句绑定的其他事件处理
        sTmp = Dic("bindcommand")
        sTmp = Replace(sTmp, "'", "")                                           '自动去掉括号和空格，如果有的话
        sTmp = Replace(sTmp, Chr(34), "")
        sTmp = Replace(sTmp, " ", "")
        
        s = Split(sTmp, ",")
        For i = 0 To UBound(s)
            s(i) = Trim(s(i))
            If s(i) = "<Change>" Then '专门处理自定义的这个事件，用控件变量监视器模拟
                If (m_Type = "Combobox" Or m_Type = "Entry" Or m_Type = "Label") And (Len(Dic("textvariable")) > 0) Then
                    sCmdName = m_Name & "_Change"
                    sOut.Append Space(8) & "self." & Dic("textvariable") & ".trace('w', self." & sCmdName & ")"
                    sCmdFunc.Append CreateFuncDefOOP(sCmdName, "*args")
                End If
            ElseIf Left(s(i), 1) = "<" And Right(s(i), 1) = ">" Then
                
                sCmdName = m_Name & "_" & Replace(Replace(Replace(s(i), "<", ""), ">", ""), "-", "_")
                
                sOut.Append Space(8) & "self." & m_Name & ".bind('" & s(i) & "', self." & sCmdName & ")"
                sCmdFunc.Append CreateFuncDefOOP(sCmdName, "event")
                
                'Python是大小写敏感的，对应快捷键也一样，如果设置的快捷键包含字母键，则将对应的大写/小写也一起绑定
                If Right(s(i), 3) >= "-a>" And Right(s(i), 3) <= "-z>" Then
                    
                    s(i) = Left(s(i), Len(s(i)) - 2) & UCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '变大写
                    sOut.Append Space(8) & "self." & m_Name & ".bind('" & s(i) & "', self." & sCmdName & ")"
                    
                ElseIf Right(s(i), 3) >= "-A>" And Right(s(i), 3) <= "-Z>" Then
                    
                    s(i) = Left(s(i), Len(s(i)) - 2) & LCase(Mid(s(i), Len(s(i)) - 1, 1)) & ">" '变小写
                    sOut.Append Space(8) & "self." & m_Name & ".bind('" & s(i) & "', self." & sCmdName & ")"
                    
                End If
            End If
        Next
    End If
    
End Sub

'根据rel(是否采用相对坐标)，生成对应的控件位置信息
Public Function GetPositionParams(rel As Boolean) As String
    If rel Then
        GetPositionParams = "relx=" & Commas2Points(Dic("relx")) & ", rely=" & Commas2Points(Dic("rely")) & _
            ", relwidth=" & Commas2Points(Dic("relwidth"))
        If m_Type <> "Combobox" Then 'Combobox不需要设置height属性
            GetPositionParams = GetPositionParams & ", relheight=" & Commas2Points(Dic("relheight"))
        End If
    Else
        GetPositionParams = "x=" & Commas2Points(Dic("x")) & ", y=" & Commas2Points(Dic("y")) & _
            ", width=" & Commas2Points(Dic("width"))
        If m_Type <> "Combobox" Then 'Combobox不需要设置height属性
            GetPositionParams = GetPositionParams & ", height=" & Commas2Points(Dic("height"))
        End If
    End If
End Function

'除了必选参数外，这个函数生成用户选择的其他参数列表
Public Function GetExtraParams(oop As Boolean, usettk As Boolean) As String
    
    Dim cfg As Variant, k As Variant, ks As Variant, sValue As String, s As New cStrBuilder
    
    Set ks = m_dic.Keys
    
    For Each k In ks
        If isExtra(k, usettk) And Len(Dic(k)) Then
            '需要使用引号括起来的属性，如果用户忘了，则在这里自动添加
            If k = "text" Or k = "label" Then
                If m_Type = "Entry" And Len(Dic("textvariable")) > 0 And usettk Then 'ttk模式下Entry使用textvariable显示字符，而不是使用text属性
                    '忽略text属性
                    sValue = ""
                Else
                    sValue = U(Dic(k))
                End If
            ElseIf InStr(1, ",fg,bg,anchor,justify,show,state,activestyle,labelanchor,mode,cursor,highlightbackground,highlightcolor,selectbackground,selectforeground,", _
                        "," & k & ",") Then
                sValue = Quote(Dic(k))
            Else
                sValue = Dic(k)
            End If
            
            If Len(sValue) Then
                s.Append k & "=" & sValue
            End If
        End If
    Next
    
    If Len(Dic("columns")) Then s.Append "columns=" & IIf(oop, "self.", "") & Dic("columns")
    If Len(Dic("displaycolumns")) Then
        If InStr(1, Dic("displaycolumns"), "#all") <= 0 Then
            s.Append "displaycolumns=" & IIf(oop, "self.", "") & Dic("displaycolumns")
        Else
            s.Append "displaycolumns='#all'"
        End If
    End If
    
    If Len(Dic("textvariable")) Then s.Append "textvariable=" & IIf(oop, "self.", "") & Dic("textvariable")
    If Len(Dic("variable")) Then s.Append "variable=" & IIf(oop, "self.", "") & Dic("variable")
    If Len(Dic("listvariable")) Then s.Append "listvariable=" & IIf(oop, "self.", "") & Dic("listvariable")
    If Len(Dic("values")) Then s.Append "values=" & IIf(oop, "self.", "") & Dic("values")
    
    If Len(Dic("command")) Then
        If InStr(1, Dic("command"), "lambda ") > 0 Then '匿名函数，不用加self.
            s.Append "command=" & Dic("command")
        Else
            s.Append "command=" & IIf(oop, "self.", "") & Dic("command")
        End If
    End If
    
    If Len(Dic("postcommand")) Then s.Append "postcommand=" & IIf(oop, "self.", "") & Dic("postcommand")
    If Len(Dic("xscrollcommand")) Then s.Append "xscrollcommand=" & IIf(oop, "self.", "") & Dic("xscrollcommand")
    If Len(Dic("yscrollcommand")) Then s.Append "yscrollcommand=" & IIf(oop, "self.", "") & Dic("yscrollcommand")
    If Len(Dic("font")) Then
        If usettk Then 'TTK模式的大多数控件的font一般都要写在样式里面，但Entry/Combobox的font要写在创建函数中才管用
            If m_Type = "Entry" Or m_Type = "Combobox" Then
                s.Append "font=" & IIf(Left$(Dic("font"), 1) = "(", Dic("font"), "(" & Dic("font") & ")") '自动加括号
            End If
        Else
            s.Append "font=" & IIf(oop, "self.", "") & m_Name & "Font"
        End If
    End If
    
    GetExtraParams = s.toString(", ")
    
End Function

'判断一个属性是否是额外参数，会根据TTK来判断更多的属性
Private Function isExtra(ByVal sK As String, usettk As Boolean) As Boolean
    Const EXTRA_STRING As String = ",x,y,relx,rely,width,height,relwidth,relheight,command,bindcommand,xscrollcommand," & _
        "yscrollcommand,postcommand,font,textvariable,variable,listvariable,values,displayrows,columns,displaycolumns,"
    Const EXTRA_STRING_TTK As String = EXTRA_STRING & ",fg,bg,bd,relief,activerelief,overrelief,anchor,jump,indicatoron,resolution,digits," & _
        "sliderlength,sliderrelief,showvalue,tickinterval,"
    
    If usettk Then
        isExtra = (InStr(1, EXTRA_STRING_TTK, "," & sK & ",") <= 0)
    Else
        isExtra = (InStr(1, EXTRA_STRING, "," & sK & ",") <= 0)
    End If
End Function

'如果使用了TTK扩展，使用这个函数获取TTK相关参数并创建合法的字符参数列表
Private Function GetStyleParams(oop As Boolean, usettk As Boolean) As String
    Dim s As New cStrBuilder
    
    If Len(Dic("relief")) Then s.Append "relief=" & Dic("relief")
    If Len(Dic("activerelief")) Then s.Append "activerelief=" & Dic("activerelief")
    If Len(Dic("overrelief")) Then s.Append "overrelief=" & Dic("overrelief")
    If Len(Dic("anchor")) Then s.Append "anchor=" & Quote(Dic("anchor"))
    If Len(Dic("fg")) Then s.Append "foreground=" & Quote(Dic("fg"))
    If Len(Dic("bg")) Then s.Append "background=" & Quote(Dic("bg"))
    If Len(Dic("bd")) Then s.Append "borderwidth=" & Dic("bd")
    If Len(Dic("jump")) Then s.Append "jump=" & Dic("jump")
    If Len(Dic("indicatoron")) Then s.Append "indicatoron=" & Dic("indicatoron")
    If Len(Dic("font")) Then
        'ttk.Entry的font要写在构建函数中才管用,ttk.LabelFrame的font要设置到.Label属性才管用
        If m_Type <> "Entry" And m_Type <> "Combobox" Then
            If Left$(Dic("font"), 1) = "(" Then
                s.Append "font=" & Dic("font")
            Else  '如果用户忘了加括号，这里加上
                s.Append "font=(" & Dic("font") & ")"
            End If
        End If
    End If
    If Len(Dic("showvalue")) Then s.Append "showvalue=" & Dic("showvalue")
    If Len(Dic("tickinterval")) Then s.Append "tickinterval=" & Dic("tickinterval")
    If Len(Dic("sliderrelief")) Then s.Append "sliderrelief=" & Dic("sliderrelief")
    If Len(Dic("sliderlength")) Then s.Append "sliderlength=" & Dic("sliderlength")
    If Len(Dic("digits")) Then s.Append "digits=" & Dic("digits")
    
    GetStyleParams = s.toString(", ")
    
End Function

'根据代码模块中的函数声明，自动填写bindcommand域
'dMethods：控件名为键，使用逗号分隔的控件事件处理函数名字符串
Public Function GetBindCommandStr(dMethods As Dictionary) As String
    Dim s As String, sOut As cStrBuilder
    If Not dMethods.Exists(m_Name) Then Exit Function
    
    Set sOut = New cStrBuilder
    s = dMethods.Item(m_Name)
    
    '这几个控件的Click事件由command属性设置比较好，不需要bind
    If m_Type = "Button" Or m_Type = "Checkbutton" Or m_Type = "Radiobutton" Then
        If InStr(1, s, "," & m_Name & "_MouseDown,") > 0 Then sOut.Append "<Button-1>"
    Else
        If InStr(1, s, "," & m_Name & "_Click,") > 0 Or InStr(1, s, "," & m_Name & "_MouseDown,") > 0 Then sOut.Append "<Button-1>"
    End If
    
    If InStr(1, s, "," & m_Name & "_DblClick,") > 0 Then sOut.Append "<Double-Button-1>"
    If InStr(1, s, "," & m_Name & "_Resize,") > 0 Then sOut.Append "<Configure>"
    If InStr(1, s, "," & m_Name & "_GotFocus,") > 0 Then sOut.Append "<FocusIn>"
    If InStr(1, s, "," & m_Name & "_LostFocus,") > 0 Then sOut.Append "<FocusOut>"
    If InStr(1, s, "," & m_Name & "_KeyPress,") > 0 Or InStr(1, s, "," & m_Name & "_KeyDown,") > 0 Then sOut.Append "<KeyPress>"
    If InStr(1, s, "," & m_Name & "_KeyUp,") > 0 Then sOut.Append "<KeyRelease>"
    If InStr(1, s, "," & m_Name & "_MouseUp,") > 0 Then sOut.Append "<ButtonRelease-1>"
    If InStr(1, s, "," & m_Name & "_Enter,") > 0 Or InStr(1, s, "," & m_Name & "_MouseMove,") > 0 Then sOut.Append "<Motion>"
    If InStr(1, s, "," & m_Name & "_Leave,") > 0 Then sOut.Append "<Leave>"
    
    If InStr(1, s, "," & m_Name & "_Change,") > 0 Then sOut.Append "<Change>"  '这个事件由内部处理，不是标准tk事件
    
    If m_Type = "Treeview" Then
        If InStr(1, s, "," & m_Name & "_NodeClick,") > 0 Then sOut.Append "<<TreeviewSelect>>"
        If InStr(1, s, "," & m_Name & "_Collapse,") > 0 Then sOut.Append "<<TreeviewClose>>"
        If InStr(1, s, "," & m_Name & "_Expand,") > 0 Then sOut.Append "<<TreeviewOpen>>"
    ElseIf m_Type = "Notebook" Then
        If InStr(1, s, "," & m_Name & "_BeforeClick,") > 0 Then sOut.Append "<<NotebookTabChanged>>"
    End If
    
    GetBindCommandStr = sOut.toString(",")
End Function

Public Function IsExistCommand(dMethods As Dictionary, sCmd As String) As Boolean
    If dMethods.Exists(m_Name) Then IsExistCommand = (InStr(1, dMethods.Item(m_Name), "," & m_Name & "_" & sCmd & ",") > 0)
End Function

'设置/获取字典的值
Public Property Get Dic(ByVal sKey As String) As String
Attribute Dic.VB_UserMemId = 0
    If m_dic.Exists(sKey) Then Dic = m_dic(sKey)
End Property

Public Property Let Dic(ByVal sKey As String, ByVal sValue As String)
    m_dic(sKey) = sValue
End Property

Public Sub Remove(ByVal sKey As String)
    m_dic.Remove (sKey)
End Sub

'将用户选择的配置更新到对象中，参数为使用"|"分割的很多对属性/值对
Public Sub SetConfig(sAttrs As String)
    Dim sa() As String, i As Long
    
    sa = Split(sAttrs, "|")
    Debug.Assert (UBound(sa) Mod 1 = 0)
    
    m_dic.RemoveAll
    
    For i = 0 To UBound(sa) - 1 Step 2
        m_dic(sa(i)) = sa(i + 1)
    Next
    
End Sub


'修改或增加单个配置项，属性/值由"|"分隔
Public Sub SetSingleConfig(sAttr As String)
    Dim sa() As String
    sa = Split(sAttr, "|")
    Debug.Assert (UBound(sa) = 1)
    m_dic(sa(0)) = sa(1)
End Sub

'设置属性值的可能值列表
'返回值：0-没有可选值，1-有一个严格限制的可选值列表，2-除提供的可选值列表外，还可以手动输入其他值
'输出：sa()可选值列表数组
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 1
    Select Case sAttr
        Case "anchor"
            sa = Split("'w','n','s','e','nw','ne','sw','se','center'", ",")
        Case "relief", "overrelief"
            sa = Split("FLAT,GROOVE,RAISED,RIDGE,SOLID,SUNKEN", ",")
        Case "takefocus"
            sa = Split("1,0", ",")
        Case "state"
            sa = Split("'normal','disabled'", ",")
        Case "justify"
            sa = Split("'left','right','center'", ",")
        Case "orient"
            sa = Split("'horizontal','vertical'", ",")
        Case "cursor"
            sa = Split("'arrow','bottom_left_corner','bottom_right_corner','center_ptr','circle','clock','cross'," & _
                "'crosshair','dot','double_arrow','exchange','fleur','hand1','hand2','icon','left_ptr','plus'," & _
                "'question_arrow','sb_h_double_arrow','sb_v_double_arrow','sizing','tcross','watch','xterm','X_cursor'", ",")
            GetAttrValueList = 2
        Case Else
            GetAttrValueList = 0
    End Select
            
End Function

'返回属性在线帮助
Public Function Tips(sAttr As String) As String
    
    Tips = sAttr & vbCrLf
    Select Case sAttr:
    Case "text", "label":
        Tips = Tips & L("l_TipText", "控件的显示文本。")
    Case "x", "y":
        Tips = Tips & L("l_TipXY", "控件的放置位置。")
    Case "width":
        Tips = Tips & L("l_TipWidth", "控件的宽度。")
    Case "height":
        Tips = Tips & L("l_TipHeight", "控件的高度。")
    Case "relx", "rely":
        Tips = Tips & L("l_TipRelXY", "控件相对于父窗口的位置，值范围为0-1。")
    Case "relwidth":
        Tips = Tips & L("l_TipRelWidth", "控件相对于父窗口的宽度，值范围为0-1。")
    Case "relheight":
        Tips = Tips & L("l_TipRelHeight", "控件相对于父窗口的高度，值范围为0-1。")
    Case "fg":
        Tips = Tips & L("l_TipFg", "前景色，格式为#RGB或red等单词，比如'#FFEA00'。")
    Case "bg":
        Tips = Tips & L("l_TipBg", "背景色，格式为#RGB或red等单词，比如'#FFEA00'。")
    Case "bd":
        Tips = Tips & L("l_TipBd", "边框宽度，单位为像素。")
    Case "anchor":
        Tips = Tips & L("l_TipAnchor", "控件内文字的对齐方式，可能值为：\n'w'|'n'|'s'|'e'|'nw'|'ne'|'sw'|'se'|'center'.")
    Case "relief":
        Tips = Tips & L("l_TipRelief", "外观效果，可选值有：FLAT, GROOVE, RAISED, RIDGE, SOLID, SUNKEN。")
    Case "overrelief":
        Tips = Tips & L("l_TipOverRelief", "鼠标悬停在控件上的外观效果，可选值有：FLAT, GROOVE, RAISED, RIDGE, SOLID, SUNKEN。")
    Case "takefocus":
        Tips = Tips & L("l_TipTakeFocus", "是否允许TAB选择控件。设置为1为允许，设置为0则TAB跳过此控件。")
    Case "state":
        Tips = Tips & L("l_TipState", "控件状态，可选值有：'normal', 'disabled'")
    Case "underline":
        Tips = Tips & L("l_TipUnderline", "在哪个字母上添加下划线，做为快捷方式的表示，第一个字母索引为0。")
    Case "justify":
        Tips = Tips & L("l_TipJustify", "控制多行字符的对齐方式，可能值为：" & vbCrLf & "left, right, center。")
    Case "padding":
        Tips = Tips & L("l_TipPadding", "控件内部额外间距，一般不需要设置。如果需要设置，请提供'左 上 右 下'四个值，或三个值'左 上 右'(下=上)，两个值'左 上'(右=左,下=上)，一个值'左'(上=下=右=左)。")
    Case "orient":
        Tips = Tips & L("l_TipOrient", "控件方向类型，全大写不用括号，小写则需要括号括起来，可选值有：'horizontal', 'vertical'。")
    Case "cursor":
        Tips = Tips & L("l_TipCursor", "鼠标指针。可以使用tk内置的指针，或在文件名前增加一个@符号使用外部指针，如'@custom.cur'。")
    Case "variable":
        Tips = Tips & L("l_TipVariable", "控件值和一个变量绑定，通过变量可以获取或设置控件的值。\n如果需要这个特性，这个属性设置为变量名。")
    Case "textvariable":
        Tips = Tips & L("l_TipTextVariable", "将显示文本和一个变量绑定，变量变化时控件的显示文本将改变，反正亦然。\n如果需要这个特性，这个属性设置为变量名。")
    Case "command":
        Tips = Tips & L("l_TipCommand", "触发事件处理函数。触发不带参数。也可以设置为top.destroy等tk方法或使用lambda函数。")
    Case "bindcommand":
        Tips = Tips & L("l_TipBindCommand", "使用bind()绑定的事件处理列表，绑定多个则使用逗号分隔(比如：<Control-C>,<F8>,<Alt-A>)。\n双击文本框查看详情。")
    Case "font":
        Tips = Tips & L("l_TipFont", "文字字体属性。是一个元组。\n元组前两个元素为：字体名,字体大小,之后的几个元素可选：'bold', 'italic', 'underline', 'overstrike'")
    Case "xscrollcommand":
        Tips = Tips & L("l_TipXScrlCmd", "水平滚动事件处理，如果需要水平滚动，设置这个值为水平滚动条的set方法。")
    Case "yscrollcommand":
        Tips = Tips & L("l_TipYScrlCmd", "上下滚动事件处理，如果需要上下滚动，设置这个值为垂直滚动条的set方法。")
    Case "xscrollincrement":
        Tips = Tips & L("l_TipXScrlIncre", "水平每次滚动增量，默认为0，可以随意滚动，可以设置为一个大于零的整数。")
    Case "yscrollincrement":
        Tips = Tips & L("l_TipYScrlIncre", "上下每次滚动增量，默认为0，可以随意滚动，可以设置为一个大于零的整数。")
    Case "scrollregion"
        Tips = Tips & L("l_TipScrlregion", "控件滚动范围（可以理解为控件实际上有多大），格式：(w,n,e,s)。\n如果要设置Canvas可以滚动，则必须设置此属性。")
    Case "confine"
        Tips = Tips & L("l_TipConfine", "是否限制在scrollregion内滚动，默认为true。")
    Case "highlightbackground"
        Tips = Tips & L("l_TipHlbg", "控件失去焦点时边框的颜色。")
    Case "highlightcolor"
        Tips = Tips & L("l_TipHlColor", "控件选中状态时边框的颜色。")
    Case "highlightthickness"
        Tips = Tips & L("l_TipHlThickness", "焦点状态边框宽度，默认为1。")
    Case "selectbackground"
        Tips = Tips & L("l_TipSltbg", "控件内选中项的背景色。")
    Case "selectborderwidth"
        Tips = Tips & L("l_TipSltbd", "控件内选中项的边框宽度。")
    Case "selectforeground"
        Tips = Tips & L("l_TipSltfg", "控件内选中项的前景色。")
    Case Else:
        Tips = Tips & L("l_TipUnknown", "未知属性。")
    End Select
    
End Function

Public Property Let ctlType(s As String)
    m_Type = s
End Property

'设置控件的父窗口，默认是top
Public Property Let Parent(s As String)
    m_Parent = s
End Property

'返回控件的父窗口，默认是top
Public Property Get Parent() As String
    Parent = m_Parent
End Property

'类实例所代表的控件类的名字
Public Property Get Name() As String
    Name = m_Name
End Property
Public Property Let Name(s As String)
    m_Name = s
End Property

Public Property Let Value(s As String)
    m_Value = s
End Property

Public Property Let StyleName(s As String)
    m_StyleName = s
End Property

Public Property Let ScaleMode(nV As Long)
    m_ScaleMode = nV
End Property
Public Property Get ScaleMode() As Long
    ScaleMode = m_ScaleMode
End Property

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    vSer.Serializer m_dic
End Function

Public Function Deserializer(vSer As clsSerialization)
    vSer.Deserializer m_dic
End Function

Public Function Keys() As Collection
    Set Keys = New Collection
    Dim k As Variant
    For Each k In m_dic.Keys
        Keys.Add k
    Next
End Function

Private Sub Class_Initialize()
    m_Type = ""
    m_Name = ""
    m_Value = ""
    m_StyleName = ""
    m_Parent = WTOP
    m_ScaleMode = vbTwips
    Set m_dic = New Dictionary
End Sub

'构建一个函数空骨架; FuncName: 函数名，Params: 参数，sbody: 函数体
Public Function CreateFuncDef(FuncName As String, Optional sparams As String = "", Optional sbody As String = "") As String
    If Len(FuncName) = 0 Then Exit Function
    CreateFuncDef = "def " & FuncName & "(" & sparams & "):" & vbCrLf
    If Len(sbody) Then
        CreateFuncDef = CreateFuncDef & sbody & vbCrLf
    Else
    #If DebugVer Then
        CreateFuncDef = CreateFuncDef & Space(4) & "print('" & FuncName & "')" & vbCrLf
    #Else
        CreateFuncDef = CreateFuncDef & Space(4) & "#TODO, Please finish the function here!" & vbCrLf
        CreateFuncDef = CreateFuncDef & Space(4) & "pass" & vbCrLf
    #End If
    End If
End Function

'构建一个函数空骨架（面向对象代码）; FuncName: 函数名，Params: 参数，sbody: 函数体
Public Function CreateFuncDefOOP(FuncName As String, Optional sparams As String = "", Optional sbody As String = "") As String
    If Len(FuncName) = 0 Then Exit Function
    CreateFuncDefOOP = "    def " & FuncName & "(self" & IIf(Len(sparams), ", ", "") & sparams & "):" & vbCrLf
    If Len(sbody) Then
        CreateFuncDefOOP = CreateFuncDefOOP & sbody & vbCrLf
    Else
    #If DebugVer Then
        CreateFuncDefOOP = CreateFuncDefOOP & Space(8) & "print('" & FuncName & "')" & vbCrLf
    #Else
        CreateFuncDefOOP = CreateFuncDefOOP & Space(8) & "#TODO, Please finish the function here!" & vbCrLf
        CreateFuncDefOOP = CreateFuncDefOOP & Space(8) & "pass" & vbCrLf
    #End If
    End If
End Function


'将各种单位转换为像素
Public Function toPixelX(nX As Long) As Long
    If m_ScaleMode = vbTwips Then
        toPixelX = Twip2PixelX(nX)
    ElseIf m_ScaleMode = vbPoints Then
        toPixelX = Point2PixelX(nX)
    Else
        toPixelX = nX
    End If
End Function

Public Function toPixelY(nY As Long) As Long
    If m_ScaleMode = vbTwips Then
        toPixelY = Twip2PixelY(nY)
    ElseIf m_ScaleMode = vbPoints Then
        toPixelY = Point2PixelY(nY)
    Else
        toPixelY = nY
    End If
End Function

'将VB的鼠标指针值翻译为tkinter的鼠标指针名
Public Function GetCursorName(nCursor As Long) As String
    Select Case nCursor
        Case vbArrow: GetCursorName = "'arrow'"
        Case vbCrosshair: GetCursorName = "'cross'"
        Case vbIbeam: GetCursorName = "'xterm'"
        Case vbSizePointer: GetCursorName = "'fleur'"
        Case vbSizeNESW: GetCursorName = "'bottom_left_corner'"
        Case vbSizeNS: GetCursorName = "'sb_v_double_arrow'"
        Case vbSizeNWSE: GetCursorName = "'bottom_right_corner'"
        Case vbSizeWE: GetCursorName = "'sb_h_double_arrow'"
        Case vbUpArrow: GetCursorName = "'center_ptr'"
        Case vbHourglass: GetCursorName = "'clock'"
        Case vbNoDrop: GetCursorName = "'X_cursor'"
        Case vbArrowHourglass: GetCursorName = "'watch'"
        Case vbArrowQuestion: GetCursorName = "'question_arrow'"
        Case vbSizeAll: GetCursorName = "'sizing'"
        Case Else: GetCursorName = ""
    End Select
End Function

'For latin, decimal 123,45 changed to 123.45
Function Commas2Points(ByVal InS As String) As String
    Commas2Points = Replace(InS, ",", ".")
End Function
