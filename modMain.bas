Attribute VB_Name = "modMain"
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Public ymVersion    As String
Public appVersion   As String

Private Function getYMVersion() As String
    Dim fso             As Scripting.FileSystemObject
    Dim YMExe           As String
    Dim arrYMVersion()  As String
    
    'output -> "C:\Program Files\Yahoo!\Messenger\YahooMessenger.exe" %1
    YMExe = getFromWindowsRegistry(HKEY_CLASSES_ROOT, "ymsgr\shell\open\command", "")
    YMExe = Replace(YMExe, " %1", "") 'hapus karakter spasi+%1
    YMExe = Replace(YMExe, Chr(34), "") 'hapus karakter petik "
    
    If Len(YMExe) > 0 Then
        Set fso = New Scripting.FileSystemObject
        arrYMVersion = Split(fso.GetFileVersion(YMExe), ".") ' ex : 10.0.0.1102, kita ambil mayor versionnya aja = 10
        Set fso = Nothing
        
    Else
        ReDim arrYMVersion(0)
    End If
    
    getYMVersion = arrYMVersion(0)
End Function

Public Sub Main()
    Dim ret As Boolean
        
    appVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    ymVersion = getYMVersion
    If Not (Len(ymVersion) > 0) Then
        MsgBox "Y!M belum terinstall, aplikasi tidak bisa dilanjutkan", vbExclamation, "Peringatan"
        End
    End If
        
    ret = konekToServer
    
    frmMain.Show
End Sub

