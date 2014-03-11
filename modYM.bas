Attribute VB_Name = "modYM"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD

Dim hwndIMClass         As Long
Dim hwndYIMInputWindow  As Long

Public Function getYMID(ByVal hwndYMMainClass As Long) As String
    Dim titleBar        As String
    Dim ymID            As String
    Dim arrTitleBar()   As String
    
    titleBar = String$(100, Chr$(0))
    GetWindowText hwndYMMainClass, titleBar, 100
    titleBar = Left$(titleBar, InStr(titleBar, Chr$(0)) - 1)
    
    If InStr(1, titleBar, " (") > 0 Then 'lawan chat terdaftar di Messenger List
        'ex YM8    : KoKom Armag3d0n (k4m4r82) - Instant Message
        '   YM9/10 : KoKom Armag3d0n (k4m4r82)
        
        arrTitleBar = Split(titleBar, " (")
        ymID = arrTitleBar(0)
        
    Else
        Select Case ymVersion
            Case "8" 'ex : KoKom Armag3d0n - Instant Message
                arrTitleBar = Split(titleBar, " - ")
                ymID = arrTitleBar(0)
                
            Case "9", "10" 'ex : KoKom Armag3d0n
                ymID = titleBar
        End Select
    End If
            
    getYMID = ymID
End Function

Public Function getYMMessage(ByVal hwndYMMainClass As Long) As String
    Dim hwndYHTMLContainer  As Long
    Dim hwndIEServer        As Long
    
    Dim ymID                As String
    Dim msg                 As String
    
    Dim arrMsg()            As String
    Dim arrValidMsg()       As String
    Dim validMsg            As String
    Dim i                   As Long
    
    Select Case ymVersion
        Case "8"
            'urutkan kelas yg harus dilalui untuk membaca pesan yang masuk
            'YSearchMenuWndClass ->  IMClass -> YHTMLContainer -> Internet Explorer_Server
            
            If hwndYMMainClass <> 0 Then
                ymID = getYMID(hwndYMMainClass)
        
                hwndIMClass = FindWindowEx(hwndYMMainClass, 0&, "IMClass", vbNullString)
                hwndYHTMLContainer = FindWindowEx(hwndIMClass, 0&, "YHTMLContainer", vbNullString)
                hwndIEServer = FindWindowEx(hwndYHTMLContainer, 0&, "Internet Explorer_Server", vbNullString)
        
                msg = getIEText(hwndIEServer)
                
                arrMsg = Split(msg, Chr(10))
                For i = LBound(arrMsg) To UBound(arrMsg)
                    If Len(arrMsg(i)) > 0 Then
                        If Left(arrMsg(i), Len(ymID) + 2) = ymID & ": " Then
                            arrValidMsg = Split(arrMsg(i), ": ")
                            validMsg = arrValidMsg(UBound(arrValidMsg))
                            Exit For
                        End If
                    End If
                Next i
        
                validMsg = Replace(validMsg, Chr(13), "")
                getYMMessage = validMsg
            End If
    
        Case "9", "10"
            'urutkan kelas yg harus dilalui untuk membaca pesan yang masuk
            'Y!M 9  : ATL:007C07F0 -> YHTMLContainer -> Internet Explorer_Server
            'Y!M 10 : CConvWndBase -> YHTMLContainer -> Internet Explorer_Server
            
            If hwndYMMainClass <> 0 Then
                ymID = getYMID(hwndYMMainClass)
        
                hwndYHTMLContainer = FindWindowEx(hwndYMMainClass, 0&, "YHTMLContainer", vbNullString)
                hwndIEServer = FindWindowEx(hwndYHTMLContainer, 0&, "Internet Explorer_Server", vbNullString)
        
                msg = getIEText(hwndIEServer)
                
                arrMsg = Split(msg, Chr(10))
                For i = LBound(arrMsg) To UBound(arrMsg)
                    If Len(arrMsg(i)) > 0 Then
                        If Left(arrMsg(i), Len(ymID) + 2) = ymID & " (" Then
                            arrValidMsg = Split(arrMsg(i), "): ")
                            validMsg = arrValidMsg(UBound(arrValidMsg))
                            Exit For
                        End If
                    End If
                Next i
        
                validMsg = Replace(validMsg, Chr(13), "")
                getYMMessage = validMsg
            End If
    
        Case Else
            'silahkan coba sendiri versi ym yg lain :)
    End Select
End Function

Public Sub ymChatSend(ByVal hwndYMMainClass As Long, ByVal msgToSend As String)
    Select Case ymVersion
        Case "8"
            'urutkan kelas yg harus dilalui untuk membalas pesan yang masuk
            'YSearchMenuWndClass ->  IMClass -> YIMInputWindow
            
            If hwndYMMainClass <> 0 Then
                hwndIMClass = FindWindowEx(hwndYMMainClass, 0&, "IMClass", vbNullString)
                hwndYIMInputWindow = FindWindowEx(hwndIMClass, 0&, "YIMInputWindow", vbNullString)
            End If
            
        Case "9", "10"
            'urutkan kelas yg harus dilalui untuk membalas pesan yang masuk
            'Y!M 9  : ATL:007C07F0 -> YIMInputWindow
            'Y!M 10 : CConvWndBase -> YIMInputWindow
            If hwndYMMainClass <> 0 Then hwndYIMInputWindow = FindWindowEx(hwndYMMainClass, 0&, "YIMInputWindow", vbNullString)
    End Select
    
    If hwndYIMInputWindow <> 0 Then
        Call SendMessageByString(hwndYIMInputWindow, WM_SETTEXT, 0&, msgToSend)
        Call SendMessage(hwndYIMInputWindow, WM_KEYDOWN, VK_RETURN, 0&) 'otomatis menekan tombol Send
    End If
End Sub

Public Sub closeYM(ByVal hwndYMMainClass As Long)
    PostMessage hwndYMMainClass, &H10, 0, 0
    DestroyWindow hwndYMMainClass
End Sub


