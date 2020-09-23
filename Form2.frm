VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Simple Version"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form2"
   ScaleHeight     =   4230
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   5895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   2895
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   255
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1575
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "G O"
      Height          =   615
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
List1.Clear
List2.Clear
List2.AddItem Dir1.Path
Dim a As Integer
a = 0

Do Until a >= List2.ListCount
    DoEvents
    Dir1.Path = List2.List(a)

For b = 0 To File1.ListCount - 1
    If Right(LCase(File1.List(b)), 3) = "mp3" Then
    List1.AddItem Dir1.Path & "\" & File1.List(b)
    End If
Next


For i = 0 To Dir1.ListCount - 1
    List2.AddItem Dir1.List(i)
Next
a = a + 1
Loop
Dir1.Path = Drive1.Drive & "\"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Dir1.Path = Drive1.Drive & "\"
End Sub

Private Sub Form_Resize()
On Error Resume Next
List1.Width = Form2.Width - 100
List1.Height = Form2.Height - 2400
End Sub
