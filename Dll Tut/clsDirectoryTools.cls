VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirectoryTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Api to retreive the Windows Directory
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Api to retreive the Windows System Directory
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Function:  GetSysDir
'Input(s):  None
'Returns:
    '1. Success: Windows System Directory
    '2. Error:   Error Number and Description
Public Function GetSysDir() As String
    Dim SysDir As String
    Dim File As String
    Dim Res As Integer
    SysDir = Space$(20)
    Res = GetSystemDirectory(SysDir, 20)
    File = Left$(SysDir, InStr(1, SysDir, Chr$(0)) - 1)
    GetSysDir = Trim$(File) & "\"
End Function

'Function:  GetWinDir
'Input(s):  None
'Returns:
    '1. Success: Windows Directory
    '2. Error:   Error Number and Description
Public Function GetWinDir() As String
    Dim WinDir As String
    Dim File As String
    Dim Res As Integer
    WinDir = Space$(20)
    Res = GetWindowsDirectory(WinDir, 20)
    File = Left$(WinDir, InStr(1, WinDir, Chr$(0)) - 1)
    GetWinDir = Trim$(File) & "\"
End Function
