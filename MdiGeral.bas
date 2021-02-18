Attribute VB_Name = "MdiGeral"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
Public Cn As ADODB.Connection



Public Function ReadUrl(ByVal Url As String, Optional ByVal Metodo As String = "GET") As String
    
    Dim Xml As Object, RetornoXml As String
   
    On Error GoTo ReadUrl_Error
    DoEvents
    'Set Xml = CreateObject("Microsoft.XMLHTTP")
    Set Xml = CreateObject("MSXML2.XMLHTTP.3.0")
    
    Xml.Open Metodo, Url, False
    DoEvents
    Xml.Send
    DoEvents
    RetornoXml = Xml.responseText
    DoEvents
    ReadUrl = RetornoXml
    
    Set Xml = Nothing
    
    On Error GoTo 0
    Exit Function

ReadUrl_Error:
    'RetornoErro "ReadUrl"
    
End Function
