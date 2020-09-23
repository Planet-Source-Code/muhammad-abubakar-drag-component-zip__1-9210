VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDrag_Drop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'       THE DRAG CLASS
'       ~~~~~~~~~~~~~~

'You can use this class in any of your projects, in anyform
'you like, modify it according to your needs, but give credit
'where credit is due.
'                               -Author:    Muhammad Abubakar
'                                       <joehacker@yahoo.com>
'                                       http://go.to/abubakar

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Public Event FilesDroped()
Private m_DragHwnd As Long
Private m_FileCount As Integer
Private FileNames() As String
Private Working As Boolean
Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal Hwnd As Long, ByVal fAccept As Long)

Friend Sub AddInFileNames(Buffer As String)
    ReDim Preserve FileNames(0 To m_FileCount)
    FileNames(m_FileCount) = Buffer
    m_FileCount = m_FileCount + 1
    'Debug.Print "file recieved : " & Buffer
    
End Sub

Friend Sub NowRaiseEvent()
    RaiseEvent FilesDroped
End Sub

Friend Sub ClearFileNames()
    ReDim FileNames(0)
    m_FileCount = 0

End Sub
Public Function StartDrag() As Long
    'This will start monitoring for the message of WM_DROPFILES
    'If already working then we wont subclass again
    If Working = False Then
        If m_DragHwnd > 0 Then
            DragAcceptFiles m_DragHwnd, True
            'Set obj = Me
            
            PrevWndFunc = SetWindowLong(m_DragHwnd, GWL_WNDPROC, AddressOf WndProc)
            StartDrag = 1 'Successfully started
            Working = True
            
        Else
            StartDrag = 0 'Unsuccessful, handle not given
            
        End If
    Else
        StartDrag = 2
    End If
    
End Function
Public Property Get DragHwnd() As Long
    m_DragHwnd = DragHwnd

End Property

Public Property Let DragHwnd(ByVal Hwnd As Long)
    
    If Not Working Then m_DragHwnd = Hwnd
    
End Property
Public Function StopDrag() As Long
    'Stop subclassing and monitoring of WM_DROPFILES message.
    
    If Working = True Then
        SetWindowLong m_DragHwnd, GWL_WNDPROC, PrevWndFunc
        DragAcceptFiles m_DragHwnd, False
        Working = False
        StopDrag = 1 'successfully stoped subclassing
        
    Else
        StopDrag = 0 'It was already not subclassed so no need to unsubclass
        
    End If
End Function
Public Function FileName(index As Integer) As String
    If index >= 0 And index <= m_FileCount Then
        FileName = FileNames(index)
    Else
        FileName = ""
    End If
    
End Function

Private Sub Class_Initialize()
    m_DragHwnd = 0
    m_FileCount = 0
    'obj is declared in BAS- <CDrag_Drop_Module> of type CDrag_Drop
    Set obj = Me
    
End Sub

Private Sub Class_Terminate()
    If Working = True Then StopDrag
    
End Sub

Public Property Get FileCount() As Integer
    FileCount = m_FileCount
    
End Property
