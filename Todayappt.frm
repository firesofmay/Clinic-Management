VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Todayappt 
   Caption         =   "Today's Appointment"
   ClientHeight    =   7215
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   10875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Todayappt.frx":0000
   ScaleHeight     =   7215
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   9720
      Picture         =   "Todayappt.frx":B4BC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   8280
      Picture         =   "Todayappt.frx":BC08
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Height          =   405
      Index           =   7
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   8640
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   480
      Width           =   2175
   End
   Begin VB.ComboBox cmbDr 
      Height          =   315
      ItemData        =   "Todayappt.frx":C955
      Left            =   600
      List            =   "Todayappt.frx":C957
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2760
      Width           =   2535
   End
   Begin VB.ListBox cmbT 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      ItemData        =   "Todayappt.frx":C959
      Left            =   8280
      List            =   "Todayappt.frx":C95B
      TabIndex        =   2
      Top             =   2400
      Width           =   2175
   End
   Begin VB.ListBox lstTemptime 
      Height          =   450
      ItemData        =   "Todayappt.frx":C95D
      Left            =   720
      List            =   "Todayappt.frx":C985
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2009
      Month           =   9
      Day             =   24
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   8640
      TabIndex        =   10
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5760
      TabIndex        =   14
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label First 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   6240
      TabIndex        =   13
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label First 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   6240
      TabIndex        =   12
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   1320
      TabIndex        =   11
      Top             =   960
      Width           =   1575
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Width           =   2055
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Index           =   0
      Left            =   9000
      TabIndex        =   6
      Top             =   1920
      Width           =   735
   End
End
Attribute VB_Name = "Todayappt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Patient, Person, Doctor, History, Appt, Test, Presc, Medical As ADODB.Recordset
Dim tempstr As String
Public doc_id, pat_id As Integer
Public pat_name, doc_name, appt_date, appt_time As String

Private Sub Calendar1_Click()
        
    txtInfo(7).Text = ""
    txtInfo(1).Text = ""
    txtInfo(2).Text = ""
    
    appt_date = Calendar1.Day
    Select Case (Calendar1.Month)
    Case 1
       appt_date = appt_date & "-Jan-"
    Case 2
       appt_date = appt_date & "-Feb-"
    Case 3
       appt_date = appt_date & "-Mar-"
    Case 4
       appt_date = appt_date & "-Apr-"
    Case 5
       appt_date = appt_date & "-May-"
    Case 6
       appt_date = appt_date & "-Jun-"
    Case 7
       appt_date = appt_date & "-Jul-"
    Case 8
       appt_date = appt_date & "-Aug-"
    Case 9
       appt_date = appt_date & "-Sep-"
    Case 10
       appt_date = appt_date & "-Oct-"
    Case 11
       appt_date = appt_date & "-Nov-"
    Case 12
       appt_date = appt_date & "-Dec-"
    End Select
    
    appt_date = appt_date & Calendar1.Year
    
    Call cmbDr_Click
    
End Sub

Private Sub cmbDr_Click()
    
    txtInfo(7).Text = ""
    txtInfo(1).Text = ""
    txtInfo(2).Text = ""
    
    If cmbDr.ListIndex >= 0 Then
        cmbT.Enabled = True
        
        'TO GET DOCTOR_ID OF THE DOCTOR NAME SELECTED FROM THE DROP DOWN LIST OF DR.
        Set Doctor = New ADODB.Recordset
        Doctor.Open "select d.doc_id, fname, lname from doctor d, person p where p.per_id = d.per_id", Ado, adOpenKeyset, adLockOptimistic
        
        For i = 0 To cmbDr.ListIndex - 1
            Doctor.MoveNext
        Next i
        
        doc_id = Doctor(0)
        doc_name = Doctor(1) & " " & Doctor(2)
        
    For i = 0 To cmbT.ListCount - 1
        cmbT.RemoveItem (0)
    Next i
    
    j = 0
        
    For i = 0 To lstTemptime.ListCount - 1
        tempstr = "select * from appt where presc_id is null and doc_id = " & doc_id & " and appt_date = '" & appt_date & "' and appt_time = '" & lstTemptime.List(i) & "'"
        
        Set Appt = New ADODB.Recordset
        Appt.Open tempstr, Ado, adOpenKeyset, adLockOptimistic
        If Appt.RecordCount = 1 Then
            cmbT.AddItem (lstTemptime.List(i)), j
            j = j + 1
        End If
        
    Next i
    
    End If
    
End Sub

Private Sub cmbT_Click()

'    On Error Resume Next

    If cmbT.ListCount >= 0 Then
        
        appt_time = cmbT.List(cmbT.ListIndex)
        
        Set Patient = New ADODB.Recordset
        Patient.Open "select pat_id, fname, lname from person per, patient pat where pat.per_id = per.per_id and pat.pat_id in (select pat_id from appt a, doctor d where a.appt_time = '" & appt_time & "' and a.appt_date = '" & appt_date & "' and d.doc_id = " & doc_id & ")", Ado, adOpenKeyset, adLockOptimistic
                        
        txtInfo(7).Text = Patient(0)
        txtInfo(1).Text = Patient(1)
        txtInfo(2).Text = Patient(2)
        pat_id = txtInfo(7).Text
        pat_name = txtInfo(1).Text & " " & txtInfo(2).Text

    
    Else
        txtInfo(7).Text = ""
        txtInfo(1).Text = ""
        txtInfo(2).Text = ""
        MsgBox "Error", vbCritical
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
    LoginUser.Show
End Sub

Private Sub cmdSub_Click()
    If txtInfo(7).Text <> "" Then
        Me.Hide
        Prescription.Show
    Else
        MsgBox "Enter Patient ID or Select Appointment Date/Time", vbCritical
    End If

End Sub

Private Sub Form_Load()
        
    Calendar1.Value = Date
        
    
'    'SHOWING PERSON ID AND PATIENT ID
'    txtinfo(7).Text = patient(0)
'    txtinfo(1).Text = patient(1)
'    txtinfo(2).Text = patient(2)
    
    
    'TO ADD DOCTORS NAME IN THE DOCTOR NAME DROP DOWN LIST
    Set Doctor = New ADODB.Recordset
    Doctor.Open "select fname, lname from doctor d, person p where p.per_id = d.per_id", Ado, adOpenKeyset, adLockOptimistic
    
    While Not Doctor.EOF
            'CONCATINATING FIRST AND LAST NAME OF THE DR.
            cmbDr.AddItem (Doctor(0) & " " & Doctor(1))
            Doctor.MoveNext
    Wend
           
    'TO SET THE appt_date VALUE TO DEFAULT TODAY'S DATE
    Call Calendar1_Click
    
             
End Sub


