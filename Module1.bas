Attribute VB_Name = "Module1"
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'This module is required by CDrag_Drop class.
'For an explanation of the following routines see the
'accompanied ReadMe file.
'                               -Author     Muhammad Abubakar
'                                       <joehacker@yahoo.com>
'                                       http://go.to/abubakar

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
Public Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_DROPFILES = &H233
Public PrevWndFunc As Long

Public obj As CDrag_Drop

Public Function WndProc(ByVal Hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    'Dim obj As New CDrag_Drop
    
    Dim n As Long, iLoop As Long, FileInfo As Long
    Dim Buffer As String * 256, tmp As String
    Dim length As Long
    If msg = WM_DROPFILES Then
        obj.ClearFileNames
        FileInfo = wParam
        n = DragQueryFile(FileInfo, -1&, vbNullString, 0)
        For iLoop = 0 To n - 1
            length = DragQueryFile(FileInfo, iLoop, ByVal Buffer, 256)
            Buffer = Trim(Buffer)
            obj.AddInFileNames Buffer
        Next
        
        obj.NowRaiseEvent
        
        DragFinish FileInfo 'wParam
        WndProc = 0
    Else
        WndProc = CallWindowProc(PrevWndFunc, Hwnd, msg, wParam, lParam)
    End If
    
    
End Function

