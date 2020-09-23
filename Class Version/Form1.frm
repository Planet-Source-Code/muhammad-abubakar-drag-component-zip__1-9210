VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "XFiles"
   ClientHeight    =   4860
   ClientLeft      =   1470
   ClientTop       =   2460
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   8625
   Begin RichTextLib.RichTextBox RT 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7223
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   4485
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "10:49 PM"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ListBox List1 
      Height          =   255
      ItemData        =   "Form1.frx":00E4
      Left            =   120
      List            =   "Form1.frx":00E6
      TabIndex        =   1
      Top             =   0
      Width           =   8415
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "All Files | *.*"
   End
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu New 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu SaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu Go 
         Caption         =   "http://&go.to/Abubakar"
         Shortcut        =   ^G
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'   A text editor sort of App using the Drag Class.
'                       -Author:        Muhammad Abubakar
'                                       <joehacker@yahoo.com>
'                                       http://go.to/abubakar

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Private size As New CResize
Private WithEvents FileClass As CDrag_Drop
Attribute FileClass.VB_VarHelpID = -1
Private f1 As Boolean, fChange As Boolean, op As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal Hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Exit_Click()

    Unload Me
    
End Sub

Private Sub FileClass_FilesDroped()
    On Error Resume Next
    
    List1.Clear
    List1.AddItem FileClass.FileName(0)
    
    FLoad FileClass.FileName(0)
    
End Sub

Private Sub Form_Load()
    Set FileClass = New CDrag_Drop
    With size
        .hParam = Height
        .wParam = Width
        .Map List1, RS_WidthOnly
        .Map RT, RS_Height_Width
    End With

    If Command <> "" Then
        Me.Show
        List1.AddItem Command
        FLoad Command
        'f1 = True
        fChange = False
        op = True
    End If
        
    FileClass.DragHwnd = List1.Hwnd
    FileClass.StartDrag
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim r As Integer
    If fChange = True Then
        r = Ask
        Select Case r
            Case vbYes:
                
                Save_Click
                
            Case vbCancel:
                Cancel = 1
                'Exit Sub
                
        End Select
    End If
End Sub
Private Function Ask() As Integer
    Ask = MsgBox("The contents of the file have been changed." & _
        vbCrLf & "Do you want to save the changes ?", vbYesNoCancel, "Save Changes")
End Function
Private Sub Form_Resize()
    size.rSize Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'The window is being subclassed so we must stop subclassing
    'before leaving or CRASH BOOM BANG!!!
    'On the other hand, even if you dont call StopDrag method, nothing
    'will happen as the subclassing is automatically turned safely
    'off as the class terminates.
    
    FileClass.StopDrag
    Set size = Nothing
    Set FileClass = Nothing
    
End Sub

Private Sub Go_Click()
    
    ShellExecute Me.Hwnd, "Open", "http://go.to/abubakar", "", App.Path, 1
    
End Sub

Private Sub New_Click()
    Dim r As Integer
    If fChange = True Then
        r = Ask
        Select Case r
            Case vbYes:
                Save_Click
                
            Case vbCancel:
                Exit Sub
        End Select
    End If
    
    op = False
    RT.Text = ""
    List1.Clear
    fChange = False
    
    
End Sub

Private Sub Open_Click()
    On Error GoTo AI
    Dim str As String
    Dim r As Integer
    If fChange = True Then
        r = Ask
        Select Case r
            Case vbYes:
                Save_Click
                
            Case vbCancel:
                Exit Sub
        End Select
    End If
    
    CD.ShowOpen
    str = CD.FileName
    If str <> "" Then
        f1 = True
        FLoad str
        List1.Clear
        List1.AddItem str
        
        fChange = False
        op = True
    End If
    Exit Sub
AI:
    MsgBox "There was an error opening the file!", vbCritical, "File Open Error"
    
End Sub

Private Sub RT_Change()
    If f1 = True Then
        f1 = False
    Else
        fChange = True
    End If
    
End Sub

Private Sub Save_Click()
    If op = True Then
        FSave List1.List(0)
        fChange = False
    Else
        CD.ShowSave
        If CD.FileName <> "" Then FSave CD.FileName: fChange = False
    End If
    
End Sub

Private Sub SaveAs_Click()
    CD.ShowSave
    If CD.FileName <> "" Then
        FSave CD.FileName
        fChange = False
    End If
    
    
End Sub
'if you want to save and load the files in the RTF format then change
'rtfText to rtfRTF below
Private Sub FSave(FName As String)
    RT.SaveFile FName, rtfText 'rtfRTF
End Sub
Private Sub FLoad(FName As String)
    RT.LoadFile FName, rtfText 'rtfRTF
    
End Sub
