VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1305
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "paul@paulmather.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1012
      MouseIcon       =   "frmAbout.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1800
      Width           =   1920
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.paulmather.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   847
      MouseIcon       =   "frmAbout.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2040
      Width           =   2250
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   525
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   3240
      Picture         =   "frmAbout.frx":091E
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblComment 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   675
      Left            =   165
      TabIndex        =   3
      Top             =   1080
      Width           =   3615
   End
   Begin VB.Image imgDate 
      Height          =   465
      Left            =   240
      Picture         =   "frmAbout.frx":0D60
      Stretch         =   -1  'True
      Top             =   120
      Width           =   465
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "v1.0.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   105
      TabIndex        =   1
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Calendar Control"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Me.Caption = "About " & App.FileDescription
    
    lblTitle.Caption = "Calendar Control" ' App.FileDescription
    lblVersion.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
    
    lblAuthor.Caption = "Written by Paul Mather"
    lblComment.Caption = "This is a replacement for the MonthView Control in mscomct2.ocx"

End Sub

Private Sub lblEmail_Click()
    Call ShellExecute(0&, vbNullString, "mailto:paul@paulmather.net", vbNullString, vbNullString, vbNormalFocus)
    Unload Me
End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblURL.ForeColor = vbBlack
    lblEmail.ForeColor = vbBlue
End Sub

Private Sub lblURL_Click()
    Call ShellExecute(0&, vbNullString, "http://www.paulmather.net", vbNullString, vbNullString, vbNormalFocus)
    Unload Me
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEmail.ForeColor = vbBlack
    lblURL.ForeColor = vbBlue
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEmail.ForeColor = vbBlack
    lblURL.ForeColor = vbBlack
End Sub

