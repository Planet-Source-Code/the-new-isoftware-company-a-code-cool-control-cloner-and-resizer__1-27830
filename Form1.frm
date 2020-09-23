VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Cool Control Cloner and Resizer"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Text            =   "Wow!  It was cloned!"
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   7
      Left            =   6000
      MousePointer    =   9  'Size W E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   6
      Left            =   6000
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   5
      Left            =   6240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   4
      Left            =   6480
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   3
      Left            =   6480
      MousePointer    =   9  'Size W E
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   2
      Left            =   6480
      MousePointer    =   6  'Size NE SW
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   1
      Left            =   6240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.PictureBox handle 
      BackColor       =   &H00C00000&
      FillColor       =   &H00FFFFFF&
      Height          =   135
      Index           =   0
      Left            =   6000
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   75
      ScaleWidth      =   75
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "New Text && Resize"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Shape"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "The New iSoftware Company"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   4680
      Width           =   4935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Source Items"
      Height          =   195
      Left            =   2160
      TabIndex        =   11
      Top             =   120
      Width           =   930
   End
   Begin VB.Shape Shape 
      Height          =   615
      Index           =   0
      Left            =   3240
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CopyControl Shape, True, Shape(0).Top + Shape(0).Height + 100, Shape(0).Left, Shape(0).Width, Shape(0).Width
End Sub

Private Sub Command2_Click()
CopyControlWithResize Text, True, True, handle, 2000, 2000, 2000, 2000

End Sub

Private Sub handle_LostFocus(Index As Integer)
    X = 0
    Do Until X = 8
    handle(X).Visible = False
    X = X + 1
    Loop
End Sub

Private Sub handle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
ControlResize Text(Text.Count), handle, Index
End Sub
