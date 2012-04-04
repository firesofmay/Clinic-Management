VERSION 5.00
Begin VB.Form chkPatID 
   Caption         =   "checking of patient ID"
   ClientHeight    =   3405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   8235
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSub 
      Caption         =   "Submit"
      Height          =   495
      Left            =   6720
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtID 
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Patient ID :"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "chkPatID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSub_Click()
    If txtID.Text = 12 Then
        MsgBox "Enter correct Patient ID!!!"
    Else
        billing.Show
    
End Sub
