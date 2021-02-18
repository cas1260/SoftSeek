Attribute VB_Name = "Módulo1"
Option Explicit
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_OPEN_TYPE_PROXY = 3

Private Const scUserAgent = "VB Project"
Private Const INTERNET_FLAG_RELOAD = &H80000000

Private Declare Function InternetOpen Lib "wininet.dll" _
  Alias "InternetOpenA" (ByVal sAgent As String, _
  ByVal lAccessType As Long, ByVal sProxyName As String, _
  ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
  Alias "InternetOpenUrlA" (ByVal hOpen As Long, _
  ByVal sUrl As String, ByVal sHeaders As String, _
  ByVal lLength As Long, ByVal lFlags As Long, _
  ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" _
  (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) _
  As Integer

Private Declare Function InternetCloseHandle _
   Lib "wininet.dll" (ByVal hInet As Long) As Integer



'###########################################################################################

Public Function OpenURL(ByVal sUrl As String) As String
'****************************************************
'PURPOSE:       Returns Contents (including all HTML) from
'               a web page
'PARAMETER:     sURL (e.g., http://www.freevbcode.com)
'RETURN VALUE:  Contents of requested page, or
'               empty string if sURL is not available
'COMMENTS:  This is an alternative to using the Internet Transfer
'           Control 's OpenURL method.  That control has a bug
'           Whereby not all the contents of the page will be
'           returned in certain circumstances
'*****************************************************

    Dim hOpen               As Long
    Dim hOpenUrl            As Long
    Dim bDoLoop             As Boolean
    Dim bRet                As Boolean
    Dim sReadBuffer         As String * 2048
    Dim lNumberOfBytesRead  As Long
    Dim sBuffer             As String

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, _
    vbNullString, vbNullString, 0)

hOpenUrl = InternetOpenUrl(hOpen, sUrl, vbNullString, 0, _
   INTERNET_FLAG_RELOAD, 0)

    bDoLoop = True
    While bDoLoop
        DoEvents
        sReadBuffer = vbNullString
        bRet = InternetReadFile(hOpenUrl, sReadBuffer, _
           Len(sReadBuffer), lNumberOfBytesRead)
        sBuffer = sBuffer & Left$(sReadBuffer, _
             lNumberOfBytesRead)
        If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
        DoEvents
    Wend
      
    If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
    If hOpen <> 0 Then InternetCloseHandle (hOpen)
    OpenURL = sBuffer

End Function
