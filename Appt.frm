VERSION 5.00
Begin VB.Form Appt 
   Caption         =   "Appointment"
   ClientHeight    =   4545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   4545
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbDr 
      Height          =   315
      ItemData        =   "Appt.frx":0000
      Left            =   3120
      List            =   "Appt.frx":0025
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdSub 
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   3600
      Width           =   1455
   End
   Begin VB.ComboBox cmbY 
      Height          =   315
      ItemData        =   "Appt.frx":00A1
      Left            =   5520
      List            =   "Appt.frx":00C9
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cmbM 
      Height          =   315
      ItemData        =   "Appt.frx":0115
      Left            =   4320
      List            =   "Appt.frx":013D
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.ComboBox cmbT 
      Height          =   315
      ItemData        =   "Appt.frx":017D
      Left            =   3120
      List            =   "Appt.frx":01A2
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.ComboBox cmbD 
      Height          =   315
      ItemData        =   "Appt.frx":021E
      Left            =   3120
      List            =   "Appt.frx":027F
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Doctor Name"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Time"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Date   (DD/MMM/YYYY)"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Appt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbD_Change()
    MsgBox cmbD.Index
End Sub

Private Sub cmdSub_Click()
    Me.Hide
End Sub
