VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Drag Component Test"
   ClientHeight    =   3420
   ClientLeft      =   1170
   ClientTop       =   1935
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   3420
   ScaleWidth      =   8820
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2790
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S&top Draging"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Start Draging"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Drag any number of files over the ListBox below"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Component Not Working"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'Test Project for the Drag Component.
'
'                               --- Author, Muhammad Abubakar
'                                       <joehacker@yahoo.com>
'                                       http://go.to/abubakar

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit

Private WithEvents iClass As CDrag_Drop
Attribute iClass.VB_VarHelpID = -1
Private size As New CResize

Private Sub Command1_Click()
    iClass.DragHwnd = List1.hWnd
    iClass.StartDrag
    Label1 = "Component Working..."
    
End Sub

Private Sub Command2_Click()
    iClass.StopDrag
    Label1 = "Component Not Working"
    
End Sub

Private Sub Form_Load()
    Set iClass = New CDrag_Drop
    With size
        .hParam = Me.Height
        .wParam = Me.Width
        .Map Command1, RS_LeftOnly
        .Map Command2, RS_Top_Left
        .Map Label1, RS_LeftOnly
        .Map List1, RS_Height_Width
    End With
    
End Sub

Private Sub Form_Resize()
    size.rSize Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    iClass.StopDrag
    
End Sub

Private Sub iClass_FilesDroped()
    Dim i As Integer
    List1.Clear
    
    With iClass
        For i = 0 To .FileCount - 1
            List1.AddItem .FileName(i)
        Next
    End With
    
End Sub
