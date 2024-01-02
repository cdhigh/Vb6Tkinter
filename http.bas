Attribute VB_Name = "http"
Option Explicit
'连接网络获取信息，需要引用 "Microsoft WinHTTP Services, version 5.1"

'参数：sUrl，URL
'参数：postData，post内容(如果为POST)
'参数：method，方法，为"POST"或"GET"
'参数：cookies，post或get带上的cookie
'返回：post或get返回的内容
Public Function HttpGetResponse(sUrl As String, Optional ByVal postData As String = "", Optional ByVal method As String = "GET", Optional ByVal cookies As String = "") As String
    Dim request As winhttp.WinHttpRequest
    If Len(Trim(cookies)) = 0 Then
        cookies = "a:x," ' cookie为空则随便弄个cookie，不然容易报错
    End If
    
    On Error Resume Next
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    With request
        .Open UCase(method), sUrl, True 'True:同步接收数据
        .SetTimeouts 30000, 30000, 30000, 30000 '设置超时时间为30s
        .Option(WinHttpRequestOption_SslErrorIgnoreFlags) = &H3300 '非常重要(忽略错误)
        .SetRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        .SetRequestHeader "Accept", "text/html, application/xhtml+xml, */*"
        .SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .SetRequestHeader "Cookie", cookies
        .SetRequestHeader "Content-Length", Len(postData)
        .Send postData '' 开始发送
        
        .WaitForResponse '等待请求
        'MsgBox WinHttp.Status'请求状态
        HttpGetResponse = .ResponseText '得到返回文本(或者是其它)
    End With
    Set request = Nothing
End Function


