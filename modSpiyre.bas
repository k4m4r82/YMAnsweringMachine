Attribute VB_Name = "modSpiyre"
Option Explicit
'CODED BY JOHN CASEY; SPIYRE@MSN.COM
'DONT FORGET REFERENCE TO "MICROSOFT HTML OBJECT LIBRARY"

Private Declare Function ObjectFromLresult Lib "oleacc" (ByVal lResult As Long, riid As UUID, ByVal wParam As Long, ppvObject As Any) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Type UUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
   
Private Const SMTO_ABORTIFHUNG = &H2

Private Function IEDOMFromhWnd(ByVal hwnd As Long) As IHTMLDocument
    Dim IID_IHTMLDocument   As UUID
    Dim lRes                As Long 'if = 0 isn't inet window.
    Dim lMsg                As Long
    Dim hr                  As Long
    
    '---END-DECLARES---------
    lMsg = RegisterWindowMessage("WM_HTML_GETOBJECT") 'Register Wnd Message
    Call SendMessageTimeout(hwnd, lMsg, 0, 0, SMTO_ABORTIFHUNG, 1000, lRes) 'Get's Object
    
    '---CHECKS-FOR-WINDOW----
    hr = ObjectFromLresult(lRes, IID_IHTMLDocument, 0, IEDOMFromhWnd)
End Function

Public Function getIEText(ByVal hwnd As Long) As String
    Dim doc As IHTMLDocument2
    
    'On Error Resume Next
    
    If hwnd <> 0 Then
        Set doc = IEDOMFromhWnd(hwnd)
    Else
        getIEText = "-[TEXT CANNOT BE FOUND]-"
        Exit Function
    
    End If
    
    '---CHECKS-FOR-HWND------
    If doc.body.innerText = vbNullString Then
        getIEText = "ERROR! [WINDOW DOESN'T CONTAIN HTML]"
        Exit Function
    End If
    '---CHECKS-FOR-HTML-EMBEDDED
    
    getIEText = doc.body.innerText
End Function




