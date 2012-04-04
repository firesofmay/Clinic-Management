VERSION 5.00
Begin VB.Form Startup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saphhire"
   ClientHeight    =   5940
   ClientLeft      =   6060
   ClientTop       =   2610
   ClientWidth     =   5880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Startup.frx":0000
   ScaleHeight     =   5940
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   3480
      Picture         =   "Startup.frx":4AD7
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   495
      Left            =   5280
      Picture         =   "Startup.frx":5824
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   3120
      PasswordChar    =   "#"
      TabIndex        =   1
      Text            =   "admin"
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Text            =   "iam"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USER NAME"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "Startup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    End
End Sub

'CHECKING IF PASSWORD IS CORRECT THEN LOAD LOGINUSER AND START CONNECTION ELSE DENY ACCESS
Private Sub cmdSub_Click()
    If txtInfo(0).Text = "iam" And txtInfo(1).Text = "admin" Then
        LoginUser.Show
        Call modLoading
        Me.Hide
    Else
        MsgBox "Incorrect Login/Password"
    End If
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
        txtInfo(Index).SelStart = 0
        txtInfo(Index).SelLength = Len(txtInfo(Index).Text)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    'KEYASCII : 65 TO 90 = ABC....Z
    'KEYASCII : 97 TO 127 = abc....z
    'KEYASCII : 32 = 'SPACE'
    'KEYASCII : 8 = 'BACKSPACE'
    'KEYASCII : 48-58 = 0-9
    
    'USERNAME AND PASSWORD
    If Index = 1 Or Index = 0 Then
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
    End If
End Sub

