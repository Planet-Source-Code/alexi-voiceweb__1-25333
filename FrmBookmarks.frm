VERSION 4.00
Begin VB.Form FrmBookmarks 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   2295
   ClientTop       =   2445
   ClientWidth     =   7830
   Height          =   6060
   Icon            =   "FrmBookmarks.frx":0000
   Left            =   2235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7830
   Top             =   2100
   Width           =   7950
   Begin VB.PictureBox PicCtls 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   7695
      TabIndex        =   5
      Top             =   4920
      Width           =   7695
      Begin VB.Label LblUPDn 
         Alignment       =   2  'Center
         Caption         =   "DOWN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   1
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.PictureBox PicCtls 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   0
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   7575
      TabIndex        =   3
      Top             =   0
      Width           =   7575
      Begin VB.Label LblUPDn 
         Alignment       =   2  'Center
         Caption         =   "UP"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   7575
      End
   End
   Begin VB.PictureBox PicBookmarks 
      BorderStyle     =   0  'None
      Height          =   7455
      Left            =   120
      ScaleHeight     =   7455
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   720
      Width           =   7575
      Begin VB.Label LblBook 
         Caption         =   "Label1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   11055
      End
   End
   Begin VB.Timer TmrSpeak 
      Interval        =   300
      Left            =   3240
      Top             =   2760
   End
   Begin VB.FileListBox File1 
      Height          =   1425
      Left            =   1320
      Pattern         =   "*.url"
      TabIndex        =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "FrmBookmarks"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Dim DoneUpd As Integer
Dim Anounced As Integer
Dim tmr As Double
Dim MoveUp As Boolean

Private Sub Form_Load()
DoneUpd = -1
loadBookmarks
FrmMain.Enabled = False
Me.Top = (Screen.Height / 2 - (Me.Height / 2))
Me.Left = (Screen.Width / 2 - (Me.Width / 2))
End Sub

Sub loadBookmarks()
Dim X As Integer
Dim tmp As String
Dim tmp1 As String
Dim pr As String

On Error Resume Next
File1.Path = "c:\windows\favorites"
File1.Refresh
t = File1.ListCount
For X = 0 To t - 1
    tmp = File1.List(X)
    Open "c:\windows\favorites\" + tmp For Input As #1
    tmp1 = Input(FileLen("c:\windows\favorites\" + tmp), #1)
    Close 1
    ur = InStr(tmp1, "URL=")
    ub = InStr(ur, tmp1, Chr(10))
    pr = Trim(Str(X + 1))
    Do While (Len(pr) < 4)
        pr = pr + " "
    Loop
    'If (ur > 0) Then
        tmp1 = Mid(tmp1, ur + 4, ub - ur - 5)
        Load LblBook(X + 1)
        LblBook(X).Caption = pr + UCase(Left(tmp, 1)) + Mid(tmp, 2, Len(tmp) - 5)
        LblBook(X).Top = FrmBookmarks.LblBook(X - 1).Top + FrmBookmarks.LblBook(X - 1).Height
        LblBook(X).Visible = True
        LblBook(X).Tag = tmp1
    'End If
Next X
    PicBookmarks.Height = LblBook(X - 1).Top + LblBook(X - 1).Height + 10
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicCtls(0).BorderStyle = 0
PicCtls(1).BorderStyle = 0
DoneUpd = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
FrmMain.Enabled = True
End Sub

Private Sub LblBook_Click(Index As Integer)
FrmMain.Enabled = True
FrmMain.Visible = True
Me.Hide
DoEvents
DoneUpd = -1
FrmMain.Web.Navigate (LblBook(Index).Tag)
Unload Me
End Sub

Private Sub LblBook_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (DoneUpd = Index) Then Exit Sub
DoneUpd = Index
Anounced = -1
tmr = Timer
On Error Resume Next
PicCtls(0).BorderStyle = 0
PicCtls(1).BorderStyle = 0
For X = 0 To LblBook.Count - 1
    If (X <> Index And LblBook(X).BackColor <> &H8000000F) Then
        LblBook(X).BackColor = &H8000000F
        LblBook(X).ForeColor = &H80000012
    End If
Next X
'LblBook(Index + 1).BackColor = &H8000000F
'LblBook(Index + 1).ForeColor = &H80000012


LblBook(Index).BackColor = &H800000
LblBook(Index).ForeColor = &HFFFFFF
End Sub


Private Sub LblUPDn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If (DoneUpd = -2) Then Exit Sub
DoneUpd = -2
On Error Resume Next
For X = 0 To LblBook.Count - 1
    If (LblBook(X).BackColor <> &H8000000F) Then
        LblBook(X).BackColor = &H8000000F
        LblBook(X).ForeColor = &H80000012
    End If
Next X
PicCtls(Index).BorderStyle = 1
If (Index > 0) Then
    PicCtls(0).BorderStyle = 0
    MoveUp = False
Else
    MoveUp = True
    PicCtls(1).BorderStyle = 0
End If

End Sub


Private Sub TmrSpeak_Timer()
TmrSpeak.Enabled = False
If (Timer - tmr) > 1 And DoneUpd >= 0 And Anounced <> DoneUpd Then
    Genie.Speak Mid(LblBook(DoneUpd).Caption, 5)
    Anounced = DoneUpd
End If
If (DoneUpd = -2) Then
    If (MoveUp) Then
        If (PicBookmarks.Top) < PicCtls(0).Top + PicCtls(0).Height - 40 Then
            PicBookmarks.Top = PicBookmarks.Top + LblBook(0).Height
        End If
    Else
        If (PicBookmarks.Top + PicBookmarks.Height) > PicCtls(1).Top + 40 Then
            PicBookmarks.Top = PicBookmarks.Top - LblBook(0).Height
        End If
    End If
End If

TmrSpeak.Enabled = True
End Sub


