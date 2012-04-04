VERSION 5.00
Begin VB.Form Prescription 
   Caption         =   "Prescription"
   ClientHeight    =   9540
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   14520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Prescription.frx":0000
   ScaleHeight     =   9540
   ScaleWidth      =   14520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   13200
      Picture         =   "Prescription.frx":46DD0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8760
      Width           =   855
   End
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   11640
      Picture         =   "Prescription.frx":4751C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8760
      Width           =   1335
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   6
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   5
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Timer cmdSubEnable 
      Interval        =   500
      Left            =   480
      Top             =   2640
   End
   Begin VB.ListBox lstHistory 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      ItemData        =   "Prescription.frx":48269
      Left            =   11640
      List            =   "Prescription.frx":4826B
      TabIndex        =   18
      Top             =   960
      Width           =   2295
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   4
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   3
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   2
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   1
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   0
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtPresc 
      Height          =   4095
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   7455
   End
   Begin VB.ListBox lstConductTest 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      ItemData        =   "Prescription.frx":4826D
      Left            =   9240
      List            =   "Prescription.frx":4826F
      TabIndex        =   2
      Top             =   4440
      Width           =   2175
   End
   Begin VB.ListBox lstTotalTest 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      ItemData        =   "Prescription.frx":48271
      Left            =   11760
      List            =   "Prescription.frx":48273
      TabIndex        =   1
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   7
      Left            =   6240
      TabIndex        =   23
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   6
      Left            =   840
      TabIndex        =   22
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Medical History"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   19
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Test List"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Index           =   2
      Left            =   12135
      TabIndex        =   17
      Top             =   3720
      Width           =   1425
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Be Conducted"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1080
      Index           =   1
      Left            =   9375
      TabIndex        =   16
      Top             =   3480
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   4
      Left            =   6240
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Index           =   3
      Left            =   6240
      TabIndex        =   12
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Prescription ID"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Index           =   2
      Left            =   840
      TabIndex        =   10
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   1
      Left            =   840
      TabIndex        =   6
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prescription"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   780
      Index           =   0
      Left            =   3750
      TabIndex        =   5
      Top             =   3840
      Width           =   2235
   End
End
Attribute VB_Name = "Prescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
    Todayappt.Show
End Sub

Private Sub cmdSub_Click()
            
    'IN PRESC ADD PRESC ID, DOC ID, PAT ID, MED RECOMM IF GIVEN
    Presc.AddNew
    Presc(0) = txtInfo(4).Text
    Presc(1) = txtInfo(2).Text
    Presc(2) = txtInfo(3).Text
            
    If txtPresc.Text <> "" Then
        Presc(3) = UCase(txtPresc.Text)
    Else
        Presc(3) = Null
    End If
    
    Presc.Update
    
    'IN APPT ADD PRESC ID FOR THIS APPT
    Set Appt = New ADODB.Recordset
    Appt.Open "select * from appt where doc_id = " & Todayappt.doc_id & " and appt_date = '" & Todayappt.appt_date & "' and appt_time = '" & Todayappt.appt_time & "'", Ado, adOpenKeyset, adLockOptimistic
    Appt(3) = txtInfo(4).Text
    Appt.Update
    
    'IF TESTS TO BE CODUCTED THEN ADD TO TEST_RESULT - TEST ID, PRESC ID, PAT ID, RESULT = NULL
    If lstConductTest.ListCount >= 0 Then
        Set Test_Result = New ADODB.Recordset
        Test_Result.Open "select * from test_result", Ado, adOpenKeyset, adLockOptimistic
        
        For i = 0 To lstConductTest.ListCount - 1
                Test.MoveFirst
                
                While Not Test.EOF
                    If Test(1) = lstConductTest.List(i) Then
                        Test_Result.AddNew
                        Test_Result(0) = Test(0)
                        Test_Result(1) = Todayappt.pat_id
                        Test_Result(2) = txtInfo(4).Text
                        Test_Result(3) = Null
                        Test_Result.Update
                        GoTo over
                    Else
                        Test.MoveNext
                    End If
                Wend
over:
        Next i
    End If
    
    Unload Me
    LoginUser.Show


End Sub

Private Sub Form_Load()
    Set Test = New ADODB.Recordset
    Test.Open "select * from test", Ado, adOpenKeyset, adLockOptimistic
    
    Set Presc = New ADODB.Recordset
    Presc.Open "select * from presc", Ado, adOpenKeyset, adLockOptimistic
    
    Set History = New ADODB.Recordset
    History.Open "select disease from medical med, history his, patient pat where pat.pat_id = his.pat_id and med.med_id = his.med_id and pat.pat_id = " & Todayappt.pat_id, Ado, adOpenKeyset, adLockOptimistic
    
    j = 0
    While Not Test.EOF
            lstTotalTest.AddItem (Test(1)), j
            Test.MoveNext
            j = j + 1
    Wend
            
    j = 0
    While Not History.EOF
            lstHistory.AddItem (History(0)), j
            History.MoveNext
            j = j + 1
    Wend
            
            
    txtInfo(0).Text = Todayappt.doc_name
    txtInfo(1).Text = Todayappt.pat_name
    txtInfo(2).Text = Todayappt.doc_id
    txtInfo(3).Text = Todayappt.pat_id
    txtInfo(4).Text = Presc.RecordCount + 1
    txtInfo(5).Text = Todayappt.appt_time
    txtInfo(6).Text = Todayappt.appt_date
    
    
    
End Sub


'SIMILAR TO HISTORY AND MEDICAL LIST IN PATIENTINFO FORM
Private Sub lstConductTest_Click()
    lstTotalTest.AddItem (lstConductTest.List(lstConductTest.ListIndex))
    lstConductTest.RemoveItem (lstConductTest.ListIndex)
End Sub

'THIS ADDS A MEDICAL HISTORY TO THE PATIENT AND REMOVES IT FROM THE MEDICAL LIST
Private Sub lstTotalTest_Click()
    lstConductTest.AddItem (lstTotalTest.List(lstTotalTest.ListIndex))
    lstTotalTest.RemoveItem (lstTotalTest.ListIndex)
End Sub

Private Sub cmdsubenable_Timer()
    If txtPresc.Text <> "" Or lstConductTest.ListCount > 0 Then
        cmdSub.Enabled = True
    Else
        cmdSub.Enabled = False
    End If
End Sub


Private Sub txtPresc_GotFocus()
    txtPresc.SelStart = 0
    txtPresc.SelLength = Len(txtPresc.Text)
End Sub
