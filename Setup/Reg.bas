Attribute VB_Name = "Reg"

'使用示例：

'新建串值
'SetStringValue "HKEY_LOCAL_MACHINE", "String Value", "Hello Visual Basic programmer"

'SetStringValue "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Internat", "internat.exe"

'新建二进制值
'SetBinaryValue "HKEY_LOCAL_MACHINE", "Binary Value", Chr$(&H1) + Chr$(&H2) + Chr$(&H3) + Chr$(&H4)

'新建 DWORD 值
'SetDWORDValue "HKEY_LOCAL_MACHINE", "DWORD Value", "1"

'读取串值
'GetStringValue("HKEY_LOCAL_MACHINE", "String Value")

'读取二进制值
'GetBinaryValue("HKEY_LOCAL_MACHINE", "Binary Value")
'If rtn = Chr$(&H1) + Chr$(&H2) + Chr$(&H3) + Chr$(&H4) Then

'读取 DWORD 值
'GetDWORDValue("HKEY_LOCAL_MACHINE", "DWORD Value")

'删除键值
'DelValue("HKEY_LOCAL_MACHINE", "String Value")

'新建主键
'CreateKey "HKEY_LOCAL_MACHINE\Registry Editor"

'删除主键
'DeleteKey "HKEY_LOCAL_MACHINE\Registry Editor"         '删除当前键,如果有分支则会失败.
'DeleteKey "HKEY_LOCAL_MACHINE\Registry Editor", True   '删除包括其下所有分支,如果有的话.

Type FILETIME
    lLowDateTime As Long
    lHighDateTime As Long
End Type

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal ulOptions As Long, _
        ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, _
        ByVal lpData As String, lpcbData As Long) As Long
Private Declare Function RegQueryValueExA Lib "advapi32.dll" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal lpReserved As Long, lpType As Long, _
        ByRef lpData As Long, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByVal lpData As String, _
        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExA Lib "advapi32.dll" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByRef lpData As Long, _
        ByVal cbData As Long) As Long
Private Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String, _
        ByVal Reserved As Long, _
        ByVal dwType As Long, _
        ByRef lpData As Byte, _
        ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
     ByVal hKey As Long, _
     ByVal dwIndex As Long, _
     ByVal lpName As String, _
     ByRef lpcbName As Long, _
     ByVal lpReserved As Long, _
     ByVal lpClass As String, _
     ByRef lpcbClass As Long, _
     ByVal lpftLastWriteTime As Long) As Long

Private Const ERROR_SUCCESS = 0&
Private Const ERROR_BADDB = 1009&
Private Const ERROR_BADKEY = 1010&
Private Const ERROR_CANTOPEN = 1011&
Private Const ERROR_CANTREAD = 1012&
Private Const ERROR_CANTWRITE = 1013&
Private Const ERROR_OUTOFMEMORY = 14&
Private Const ERROR_INVALID_PARAMETER = 87&
Private Const ERROR_ACCESS_DENIED = 5&
Private Const ERROR_NO_MORE_ITEMS = 259&
Private Const ERROR_MORE_DATA = 234&

Private Const REG_NONE = 0&
Private Const REG_SZ = 1&
Private Const REG_EXPAND_SZ = 2&
Private Const REG_BINARY = 3&
Private Const REG_DWORD = 4&
Private Const REG_DWORD_LITTLE_ENDIAN = 4&
Private Const REG_DWORD_BIG_ENDIAN = 5&
Private Const REG_LINK = 6&
Private Const REG_MULTI_SZ = 7&
Private Const REG_RESOURCE_LIST = 8&
Private Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Private Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Private Const KEY_QUERY_VALUE = &H1&
Private Const KEY_SET_VALUE = &H2&
Private Const KEY_CREATE_SUB_KEY = &H4&
Private Const KEY_ENUMERATE_SUB_KEYS = &H8&
Private Const KEY_NOTIFY = &H10&
Private Const KEY_CREATE_LINK = &H20&
Private Const READ_CONTROL = &H20000
Private Const WRITE_DAC = &H40000
Private Const WRITE_OWNER = &H80000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const STANDARD_RIGHTS_READ = READ_CONTROL
Private Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Private Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Private Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Private Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Private Const KEY_EXECUTE = KEY_READ

Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const HKEY_PERFORMANCE_DATA = &H80000004
Private Const HKEY_CURRENT_CONFIG = &H80000005
Private Const HKEY_DYN_DATA = &H80000006

Dim hKey As Long, MainKeyHandle As Long
Dim Rtn As Long, lBuffer As Long, sBuffer As String
Dim lBufferSize As Long
Dim lDataSize As Long
Dim ByteArray() As Byte
Dim sBaseName As String

'This constant determins wether or not to display error messages to the
'user. I have set the default value to False as an error message can and
'does become irritating after a while. Turn this value to true if you want
'to debug your programming code when reading and writing to your system
'registry, as any errors will be displayed in a message box.

Const DisplayErrorMsg = False

Function GetMainKeyHandle(MainKeyName As String) As Long
    Select Case MainKeyName
        Case "HKEY_CLASSES_ROOT"
            GetMainKeyHandle = HKEY_CLASSES_ROOT
        Case "HKEY_CURRENT_USER"
            GetMainKeyHandle = HKEY_CURRENT_USER
        Case "HKEY_LOCAL_MACHINE"
            GetMainKeyHandle = HKEY_LOCAL_MACHINE
        Case "HKEY_USERS"
            GetMainKeyHandle = HKEY_USERS
        Case "HKEY_PERFORMANCE_DATA"
            GetMainKeyHandle = HKEY_PERFORMANCE_DATA
        Case "HKEY_CURRENT_CONFIG"
            GetMainKeyHandle = HKEY_CURRENT_CONFIG
        Case "HKEY_DYN_DATA"
            GetMainKeyHandle = HKEY_DYN_DATA
    End Select
End Function

Function ErrorMsg(lErrorCode As Long) As String

    'If an error does accurr, and the user wants error messages displayed, then
    'display one of the following error messages

    Select Case lErrorCode
    Case 1009, 1015
        GetErrorMsg = "The Registry Database is corrupt!"
    Case 2, 1010
        GetErrorMsg = "Bad Key Name"
    Case 1011
        GetErrorMsg = "Can't Open Key"
    Case 4, 1012
        GetErrorMsg = "Can't Read Key"
    Case 5
        GetErrorMsg = "Access to this key is denied"
    Case 1013
        GetErrorMsg = "Can't Write Key"
    Case 8, 14
        GetErrorMsg = "Out of memory"
    Case 87
        GetErrorMsg = "Invalid Parameter"
    Case 234
        GetErrorMsg = "There is more data than the buffer has been allocated to hold."
    Case Else
        GetErrorMsg = "Undefined Error Code: " & Str$(lErrorCode)
    End Select

End Function

Function ClearKK(Str As String)

Dim I As Long

Dim K As String

    For I = 1 To Len(Str)

        K = Mid(Str, I, 1)

        If K = Chr(0) Then

            Str = Mid(Str, 1, I)

            Exit For

        End If

    Next

End Function

Function GetStringValue(SubKey As String, Entry As String)

    Call ParseKey(SubKey, MainKeyHandle)

    If MainKeyHandle Then
        Rtn = RegCreateKey(MainKeyHandle, SubKey, hKey)    'open the key
        If Rtn = ERROR_SUCCESS Then    'if the key could be opened then
            sBuffer = Space(255)    'make a buffer
            lBufferSize = LenB(StrConv(sBuffer, vbFromUnicode))
            Rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize)    'get the value from the registry
            If Rtn = ERROR_SUCCESS Then    'if the value could be retreived then
                Rtn = RegCloseKey(hKey)    'close the key
                sBuffer = Trim(sBuffer)
                ClearKK sBuffer
                GetStringValue = Trim(Left(sBuffer, LenB(StrConv(sBuffer, vbFromUnicode)) - 1))    'return the value to the user
            Else    'otherwise, if the value couldnt be retreived
                GetStringValue = ""    'return Error to the user
                If DisplayErrorMsg = True Then    'if the user wants errors displayed then
                    MsgBox ErrorMsg(Rtn)    'tell the user what was wrong
                End If
            End If
        Else    'otherwise, if the key couldnt be opened
            GetStringValue = ""    'return Error to the user
            If DisplayErrorMsg = True Then    'if the user wants errors displayed then
                MsgBox ErrorMsg(Rtn)    'tell the user what was wrong
            End If
        End If
    End If

End Function

Private Sub ParseKey(KeyName As String, Keyhandle As Long)

    Rtn = InStr(KeyName, "\")    'return if "\" is contained in the Keyname

    If Left(KeyName, 5) <> "HKEY_" Or Right(KeyName, 1) = "\" Then    'if the is a "\" at the end of the Keyname then
        MsgBox "Incorrect Format:" + Chr(10) + Chr(10) + KeyName    'display error to the user
        Exit Sub    'exit the procedure
    ElseIf Rtn = 0 Then    'if the Keyname contains no "\"
        Keyhandle = GetMainKeyHandle(KeyName)
        KeyName = ""    'leave Keyname blank
    Else    'otherwise, Keyname contains "\"
        Keyhandle = GetMainKeyHandle(Left(KeyName, Rtn - 1))    'seperate the Keyname
        KeyName = Right(KeyName, LenB(StrConv(KeyName, vbFromUnicode)) - Rtn)
    End If

End Sub


