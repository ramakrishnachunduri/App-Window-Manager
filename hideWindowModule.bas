Attribute VB_Name = "hideWindowModule"
Option Explicit


Private Declare Function FindWindow% Lib "User" (ByVal lpClassName As Any, _
        ByVal lpWindowName As Any)
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
       Global Const conHwndTopmost = -1
       Global Const conSwpNoActivate = &H10
       Global Const conSwpShowWindow = &H40
       Global Const HWND_BOTTOM = 1
       Global Const SWP_NOSIZE = &H1
       Global Const SWP_DRAWFRAME = &H20
       Global Const SWP_HIDEWINDOW = &H80
Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
       
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long
       Public Const WM_CLOSE = &H10
       
Declare Function ShowWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
       Public Const SW_HIDE = 0
       Public Const SW_MAXIMIZE = 3
       Public Const SW_SHOW = 5
       Public Const SW_MINIMIZE = 6
       Public Const SW_RESTORE = 9
Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public OldText As String

Function RetHandle(Str As String) As Long
    Dim i As Integer
    Dim TempStr As String
    'MsgBox Str
    For i = 1 To Len(Str)
       TempStr = Right(Str, i)
       If Left(TempStr, 1) = " " Then
          ' found the space
          RetHandle = CLng(LTrim(TempStr))
          Exit Function
       End If
    Next i
    
End Function



 
