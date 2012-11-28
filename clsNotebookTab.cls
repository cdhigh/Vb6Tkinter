VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNotebookTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'ѡ��еĵ�ҳ���������úͱ�����Ӧ�����Ժ�ֵ�������ɶ�ӦPYTHON��TKinter����

Private m_Comps() As Object  '�����ѡ��ҳ�����еĿؼ�
Private m_CanbeOutByMainForm As Boolean
Private m_SubWidgets() As Object
Private m_numSubWidgets As Long
Private m_Name As String
Private m_Caption As String
Private m_Parent  As String

'���PYTHON����,
'sCmdFunc: ����������¼������ص�����;
'rel:�Ƿ�ʹ���������,
'oop:�Ƿ�ʹ�����������
'usettk:�Ƿ�ʹ��TTK������չ
Public Function toString(ByRef sCmdFunc As cStrBuilder, rel As Boolean, oop As Boolean, usettk As Boolean) As String
    
    Dim i As Long, sOut As New cStrBuilder
    
    '���Ȱ�������ϵ��������
    SortWidgets m_SubWidgets, m_numSubWidgets
    
    For i = 0 To m_numSubWidgets - 1
        sOut.Append m_SubWidgets(i).toString(sCmdFunc, rel, oop, usettk)
    Next
    
    toString = sOut.toString(vbCrLf)
    
End Function


'���������Ҫ���ϵ������������ʼ��������
Public Sub InitConfig(ByRef o As Object, parentWidth As Long, parentHeight As Long)
    
    m_Name = o.Properties("Name")
    
    On Error Resume Next
    m_Caption = o.Properties("Caption")
    On Error GoTo 0
    
    Erase m_SubWidgets
    m_numSubWidgets = 0
    
End Sub

'����һ���ӿؼ�
Public Sub AddSubWidget(oWidget As Object)
    ReDim Preserve m_SubWidgets(m_numSubWidgets) As Object
    Set m_SubWidgets(m_numSubWidgets) = oWidget
    m_numSubWidgets = m_numSubWidgets + 1
End Sub

'��������ֵ�Ŀ���ֵ�б�
'����ֵ��0-û�п�ѡֵ��1-��һ���ϸ����ƵĿ�ѡֵ�б���2-���ṩ�Ŀ�ѡֵ�б��⣬�������ֶ���������ֵ
'�����sa()��ѡֵ�б�����
Public Function GetAttrValueList(sAttr As String, ByRef sa() As String) As Long
    GetAttrValueList = 0
End Function

Public Function Tips(sAttr As String) As String
    Tips = ""
End Function

'���û�ѡ������ø��µ������У�����Ϊʹ��"|"�ָ�ĺܶ������/ֵ��
Public Sub SetConfig(sAttrs As String)
    '''''
End Sub

'�޸Ļ����ӵ������������/ֵ��"|"�ָ�
Public Sub SetSingleConfig(sAttr As String)
    ''''''
End Sub

'����һ�����ϣ�ÿ����Ŀ��Ԫ��"������|ֵ|�Ƿ�Ĭ��ѡ��"
'�����������������������Բ����б���
Public Function Allitems() As Collection
    Set Allitems = New Collection
End Function

Private Sub Class_Initialize()
    m_numSubWidgets = 0
    m_CanbeOutByMainForm = True
    m_Parent = WTOP
End Sub

Public Property Let Parent(s As String)
    m_Parent = s
End Property

Public Property Get Parent() As String
    Parent = m_Parent
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

'���ڸı���Ĭ�϶�Ӧ��widget���ͣ��޸�widget���ͺ�ע�������б��ĺϷ���
Public Function SetWidgetType(sType As String, sStyleName As String)
    ''''
End Function

'ȷ�������������ܷ������toString()���������룬Ĭ��ΪTrue������ΪFalse˵�����������������ô���
Public Property Get EnableOutByMainForm() As Boolean
    EnableOutByMainForm = m_CanbeOutByMainForm
End Property
Public Property Let EnableOutByMainForm(bEnable As Boolean)
    m_CanbeOutByMainForm = bEnable
End Property

'�������л�����
Public Function Serializer(vSer As clsSerialization)
    'vSer.Serializer m_Base
End Function

Public Function Deserializer(vSer As clsSerialization)
    'vSer.Deserializer m_Base
End Function

Public Property Get Description() As String
    Description = L("l_DescNotebookTab", "ѡ���ҳ���󣬶�Ӧ��Tkinter.ttk��Notebook�ؼ������ÿһҳ��û�п����õ����ԡ�")
End Property

Public Property Let ScaleMode(nV As Long)
    'm_Base.ScaleMode = nV
End Property

'����ģ��Ƚ�����ĺ�����ʵ�������ж����������������ϵ
'�ñ��������һ������Ƚϣ��߼����Ϊ'������-��һ������'
'����ֵ���壺
'<0����ʾ��������Ҫ����һ������ǰ�������
'=0����ʾ���߽�û��������ϵ������ǰ��˳����Ӱ��
'>0����һ������Ҫ��������롣
'������߼�����������صĳ���
Public Function Compare(ByRef Obj As Object) As Long
    If Parent = Obj.Name Then '���ؼ����������
        Compare = 1
    ElseIf Obj.Parent = Name Then
        Compare = -1
    ElseIf Parent = WTOP And Obj.Parent <> WTOP Then '����ؼ������
        Compare = -1
    ElseIf Parent <> WTOP And Obj.Parent = WTOP Then
        Compare = 1
    Else
        Compare = 0
    End If
End Function

