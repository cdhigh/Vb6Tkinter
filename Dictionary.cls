VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'对SCRIPTING.DICTIONARY的简单封装

Private m_dic As Object

Public Sub Add(ByVal sK As String, sV As Variant)
    m_dic.Add sK, sV
End Sub

Public Property Get Item(ByVal sK As String) As String
Attribute Item.VB_UserMemId = 0
    Item = m_dic.Item(sK)
End Property

Public Property Let Item(ByVal sK As String, sV As String)
    m_dic.Item(sK) = sV
End Property

Public Function Exists(ByVal sK As String) As Boolean
    Exists = m_dic.Exists(sK)
End Function


Public Function Keys() As Collection
    'Set Keys = m_Dic.Keys
    Set Keys = New Collection
    Dim k As Variant
    For Each k In m_dic
        Keys.Add k
    Next
End Function

Public Function Items() As Collection
    'Set Items = m_Dic.Items
    Set Items = New Collection
    Dim k As Variant
    For Each k In m_dic
        Items.Add m_dic.Item(k)
    Next
End Function

Public Sub Remove(sKey As String)
    If m_dic.Exists(sKey) Then m_dic.Remove sKey
End Sub

Public Sub RemoveAll()
    m_dic.RemoveAll
End Sub

Public Property Get Count() As Long
    Count = m_dic.Count
End Property

'对象序列化函数
Public Function Serializer(vSer As clsSerialization)
    
    Dim k As Variant
    
    vSer.Serializer m_dic.Count
    
    For Each k In m_dic
        vSer.Serializer k, m_dic(k)
    Next
    
End Function

Public Function Deserializer(vSer As clsSerialization)
    Dim nCount As Long, i As Long, k  As Variant, v As Variant
    
    m_dic.RemoveAll
    
    vSer.Deserializer nCount
    
    For i = 0 To nCount - 1
        vSer.Deserializer k, v
        m_dic.Add k, v
    Next
    
End Function


Private Sub Class_Initialize()
    Set m_dic = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_dic = Nothing
End Sub

