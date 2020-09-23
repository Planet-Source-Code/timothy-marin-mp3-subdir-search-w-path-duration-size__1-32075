VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form Form1 
   Caption         =   "Mp3 Search"
   ClientHeight    =   5265
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   5265
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Text            =   "Dave Matthews"
      Top             =   480
      Width           =   2295
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Search For Name"
      Height          =   255
      Left            =   3000
      TabIndex        =   6
      Top             =   480
      Width           =   1695
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3255
      Left            =   0
      TabIndex        =   5
      Top             =   2040
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5741
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Path"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "File Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Duration"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Search"
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2895
   End
   Begin VB.DirListBox Dir1 
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   1680
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Height          =   255
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   255
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   615
      Left            =   3000
      TabIndex        =   13
      Top             =   1320
      Width           =   4215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
   Begin VB.Label Label5 
      Caption         =   "Files."
      Height          =   255
      Left            =   6360
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   5400
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Directories... Found"
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "0"
      Height          =   255
      Left            =   3000
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Status..."
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   720
      Width           =   1335
   End
   Begin VB.Menu mnus 
      Caption         =   "View Simple Version"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Text1.Enabled = False Then
Text1.Enabled = True
Else
Text1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
lv1.ListItems.Clear
Label2 = "0"
Label4 = "0"
List2.Clear
List2.AddItem Dir1.Path
Dim a As Integer
a = 0
Do Until a >= List2.ListCount
    Label2.Caption = Label2.Caption + 1
    DoEvents
    Dir1.Path = List2.List(a)

For b = 0 To File1.ListCount - 1
    Dim min1 As Integer
    Dim sec1 As Integer
    
    If Right(LCase(File1.List(b)), 3) = "mp3" Then
    If Text1.Enabled = True Then
    If InStr(LCase(File1.List(b)), LCase(Text1.Text)) Then
    Label4 = Label4 + 1
    
    MediaPlayer1.FileName = Dir1.Path & "\" & File1.List(b)
    min1 = MediaPlayer1.Duration \ 60
    sec1 = MediaPlayer1.Duration - (min1 * 60)
    lv1.ListItems.Add , , File1.List(b)
    lv1.ListItems.Item(lv1.ListItems.Count).SubItems(1) = Dir1.Path
    lv1.ListItems.Item(lv1.ListItems.Count).SubItems(2) = FileLen(Dir1.Path & "\" & File1.List(b))
    lv1.ListItems.Item(lv1.ListItems.Count).SubItems(3) = min1 & ":" & sec1
    MediaPlayer1.FileName = ""
    End If
    Else
    Label4 = Label4 + 1
    
        MediaPlayer1.FileName = Dir1.Path & "\" & File1.List(b)
    min1 = MediaPlayer1.Duration \ 60
    sec1 = MediaPlayer1.Duration - (min1 * 60)
    
    lv1.ListItems.Add , , File1.List(b)
    lv1.ListItems.Item(lv1.ListItems.Count).SubItems(1) = Dir1.Path
    lv1.ListItems.Item(lv1.ListItems.Count).SubItems(2) = FileLen(Dir1.Path & "\" & File1.List(b))
    lv1.ListItems.Item(lv1.ListItems.Count).SubItems(3) = min1 & ":" & sec1
    MediaPlayer1.FileName = ""
    
    End If
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
On Error Resume Next
Dir1.Path = "c:\"
End Sub

Private Sub Form_Resize()
On Error Resume Next
lv1.Width = Form1.Width - 100
lv1.Height = Form1.Height - 2710
End Sub

Private Sub lv1_Click()
On Error Resume Next
MediaPlayer1.FileName = lv1.ListItems.Item(lv1.SelectedItem.Index).SubItems(1) & "\" & lv1.SelectedItem
End Sub

Private Sub mnus_Click()
Form2.Show
End Sub
