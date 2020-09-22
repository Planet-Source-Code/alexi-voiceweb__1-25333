VERSION 4.00
Begin VB.Form FrmMain 
   Caption         =   "VoiceWeb"
   ClientHeight    =   5940
   ClientLeft      =   2010
   ClientTop       =   1755
   ClientWidth     =   6690
   Height          =   6345
   Icon            =   "VoiceWeb.frx":0000
   KeyPreview      =   -1  'True
   Left            =   1950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   6690
   Top             =   1410
   Width           =   6810
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdCommand 
      Caption         =   "&Add"
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
      Index           =   8
      Left            =   6120
      TabIndex        =   10
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "B&ookmarks"
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
      Index           =   7
      Left            =   6840
      TabIndex        =   9
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "&Refresh"
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
      Index           =   6
      Left            =   7440
      TabIndex        =   8
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox TxtUrl 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   7
      Text            =   "HTTP://www.excite.com"
      Top             =   1560
      Width           =   6495
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "&Stop"
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
      Index           =   5
      Left            =   4560
      TabIndex        =   6
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "S&top R"
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
      Index           =   4
      Left            =   4560
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "Se&lection"
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
      Index           =   3
      Left            =   1920
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "R&ead"
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "&Forward>"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton CmdCommand 
      Caption         =   "<&Back"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1695
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   5760
      Top             =   4800
      _cx             =   
      _cy             =   
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   5530
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Private Sub CmdCommand_Click(Index As Integer)
On Error Resume Next

Select Case Index
Case 0
    Web.GoBack
Case 1
    Web.GoForward
Case 2
    Genie.Speak Web.Document.body.outerText
Case 3
    Genie.Speak Web.Document.selection.createRange.Text
Case 4
    Genie.Stop
Case 5
    Web.Stop
Case 6
    Web.Refresh
Case 7
    FrmBookmarks.Show
Case 8
    Dim webbr As New SHDocVwCtl.ShellUIHelper
    webbr.AddFavorite Web.Document.location, Web.Document.Title
End Select
End Sub

Private Sub Form_Activate()
TxtUrl.SetFocus
End Sub

Private Sub Form_Click()
'web.ShowBrowserBar 2
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If (KeyCode = 116) Then Web.Refresh

End Sub

Private Sub Form_Load()
On Error Resume Next
    Agent1.Characters.Load "Genie", GNPATH
    Set Genie = Agent1.Characters("Genie")
    Genie.LanguageID = &H409

Genie.Left = Screen.Width / Screen.TwipsPerPixelX - Genie.Width
Genie.Top = 20
Genie.Show
Genie.Balloon = False
Genie.Speak "\emp\Hello, let's surf the net!?"
Web.Navigate "HTTP://www.excite.com"


End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Web.Width = Me.Width - 205
    Web.Height = Me.Height - TxtUrl.Top - TxtUrl.Height - 400
    TxtUrl.Width = Me.Width - 105
   
End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub TxtUrl_GotFocus()
TxtUrl.SelStart = 0
TxtUrl.SelLength = Len(TxtUrl.Text)
End Sub



Private Sub TxtUrl_KeyPress(KeyAscii As Integer)
On Error Resume Next
If (KeyAscii = 13) Then
    KeyAscii = 0
    Web.Navigate TxtUrl.Text
End If
End Sub


Private Sub web_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
TxtUrl.Text = Web.Document.location
End Sub

