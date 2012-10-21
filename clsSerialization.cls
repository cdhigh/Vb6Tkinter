VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerialization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'对象持续化类
Private mInnerArray() As String
Private mInnerArrayItems As Long
Private mInnerArrayCur As Long

'设置要解码的字符串
Public Property Let SerialString(newStr As String)
    ResetStrArray
    mInnerArray = Split(newStr, vbCrLf)
    mInnerArrayItems = UBound(mInnerArray) + 1
    mInnerArrayCur = 0
End Property

Public Property Get SerialString() As String
    If mInnerArrayItems > 0 Then
        SerialString = Join(mInnerArray, vbCrLf)
    Else
        SerialString = ""
    End If
End Property


Public Sub Serializer(ParamArray vObjs() As Variant)
    Dim i As Long
    For i = LBound(vObjs) To UBound(vObjs)
        Dim tType As String
        tType = GetType(vObjs(i))
        SaveData tType                                                          'Save Type
        
        Select Case tType
        Case "Object"
            SaveData TypeName(vObjs(i))
            vObjs(i).Serializer Me
            
        Case "Variant"
            'Data = TypeName(vObjs(i))
            SaveData vObjs(i)
        Case "Nothing"
            'Do Nothing
        End Select
    Next i
End Sub

Public Sub Deserializer(ParamArray vObjs() As Variant)
    Dim i As Long, tType As String
    For i = LBound(vObjs) To UBound(vObjs)
        Dim Classification As String
        Classification = GetData()
        Select Case Classification
        Case "Object"
            tType = GetData()
            vObjs(i).Deserializer Me
            
        Case "Variant"
            vObjs(i) = GetData()
        Case "Nothing"
        End Select
    Next i
End Sub



Private Function GetType(vObj As Variant) As String
    If IsObject(vObj) = True Then
        GetType = "Object"
        Exit Function
    End If
    
    If TypeName(vObj) = "Nothing" Then
        GetType = "Nothing"
        Exit Function
    End If
    If GetType = "" Then
        GetType = "Variant"
        Exit Function
    End If
    
End Function

Private Sub Class_Initialize()
    ResetStrArray
End Sub

Private Sub Class_Terminate()
    ResetStrArray
End Sub

Private Function GetData() As Variant
    GetData = GetOneInnerString
End Function

Private Sub SaveData(vData As Variant)
    AppendInnerString CStr(vData)
End Sub


Private Function GetOneInnerString() As String
    If mInnerArrayCur < mInnerArrayItems Then
        GetOneInnerString = mInnerArray(mInnerArrayCur)
        mInnerArrayCur = mInnerArrayCur + 1
    End If
End Function

Private Sub AppendInnerString(ByVal newStr As String)
    ReDim Preserve mInnerArray(mInnerArrayItems) As String
    mInnerArray(mInnerArrayItems) = newStr
    mInnerArrayItems = mInnerArrayItems + 1
End Sub

Private Sub ResetStrArray()
    mInnerArrayItems = 0
    mInnerArrayCur = 0
    Erase mInnerArray
End Sub

Public Sub Reset()
    ResetStrArray
End Sub
