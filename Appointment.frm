VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form Appointment 
   Caption         =   "Appointment"
   ClientHeight    =   8025
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11040
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "Appointment.frx":0000
   ScaleHeight     =   8025
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   8040
      Picture         =   "Appointment.frx":51F80
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   9600
      Picture         =   "Appointment.frx":52CCD
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   8280
      TabIndex        =   10
      Top             =   240
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   7
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   2175
   End
   Begin VB.ListBox lstTemptime 
      Height          =   450
      ItemData        =   "Appointment.frx":53419
      Left            =   600
      List            =   "Appointment.frx":53441
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox cmbDr 
      Height          =   315
      ItemData        =   "Appointment.frx":534C9
      Left            =   240
      List            =   "Appointment.frx":534CB
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2640
      Width           =   2295
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
      ItemData        =   "Appointment.frx":534CD
      Left            =   8280
      List            =   "Appointment.frx":534CF
      TabIndex        =   3
      Top             =   2400
      Width           =   2175
   End
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   3000
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
      Left            =   8280
      TabIndex        =   11
      Top             =   840
      Width           =   2175
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   9
      Left            =   600
      TabIndex        =   14
      Top             =   240
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   12
      Top             =   240
      Width           =   1935
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
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Appointment Time"
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
      Height          =   1095
      Index           =   0
      Left            =   8160
      TabIndex        =   6
      Top             =   1920
      Width           =   3855
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
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   2040
      Width           =   3495
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   13
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Appointment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Patient, Person, Doctor, History, Appt, Test, Presc, Medical As ADODB.Recordset
Dim appt_date As String
Dim doc_id As Integer
Dim temp As String

Private Sub Calendar1_Click()
        MsgBox "Appointment Before Today is not allowed"
        Calendar1.Value = appt_date
            
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
    
    If cmbDr.ListIndex >= 0 Then
        cmbT.Enabled = True
        
        'TO GET DOCTOR_ID OF THE DOCTOR NAME SELECTED FROM THE DROP DOWN LIST OF DR.
        Set Doctor = New ADODB.Recordset
        Doctor.Open "select d.doc_id from doctor d, person p where p.per_id = d.per_id", Ado, adOpenKeyset, adLockOptimistic
        
        For i = 0 To cmbDr.ListIndex - 1
            Doctor.MoveNext
        Next i
        
        doc_id = Doctor(0)
        
    For i = 0 To cmbT.ListCount - 1
        cmbT.RemoveItem (0)
    Next i
    
    j = 0
        
    For i = 0 To lstTemptime.ListCount - 1
        temp = "select * from appt where doc_id = " & doc_id & " and appt_date = '" & appt_date & "' and appt_time = '" & lstTemptime.List(i) & "'"
        
        Set Appt = New ADODB.Recordset
        Appt.Open temp, Ado, adOpenKeyset, adLockOptimistic
        
        If Appt.RecordCount = 0 Then
            cmbT.AddItem (lstTemptime.List(i)), j
            j = j + 1
        End If
        
    Next i
    
    End If
    
End Sub


Private Sub cmdCancel_Click()
    Unload Me
    CheckPatient.Show
End Sub

Private Sub cmdSub_Click()
    On Error GoTo disablecmdsub
         
    Set Appt = New ADODB.Recordset
    Appt.Open "select * from appt", Ado, adOpenKeyset, adLockOptimistic

    Appt.AddNew
    
                
    'INSERT THE DOCTOR ID IN THE APPT TABLE'S FIRST COLOUMN
    Appt(0) = doc_id
    
    'INSERT THE APPOINTMENT DATE IN 2ND COLOUMN
    Appt(1) = appt_date
    
    'INSERT THE TIME IN 3RD COLOUMN
    Appt(2) = cmbT.List(cmbT.ListIndex)
    
    'PRESC_ID IS NULL RIGHT NOW
    Appt(3) = Null
    
    'PATIENT_ID ACCEPT IT FROM ID_TABLE 2ND COLOUMN
    Appt(4) = txtInfo(7).Text
        
    Appt.Update
        
    Unload Me
    Unload CheckPatient
    LoginUser.Show

disablecmdsub:
    cmdSub.Enabled = False

End Sub


Private Sub Form_Load()
    'INITIALIZING VALUE OF APPT CALENDAR, APPT_DATE, DOB
    appt_date = Calendar1.Value
    
    Set Patient = New ADODB.Recordset
    Patient.Open "select fname,lname from person per,patient pat where per.per_id = pat.per_id and pat.pat_id = " & CheckPatient.patid, Ado, adOpenKeyset, adLockOptimistic
    
    'SHOWING PERSON ID AND PATIENT ID
    txtInfo(7).Text = CheckPatient.patid
    txtInfo(1).Text = Patient(0)
    txtInfo(2).Text = Patient(1)
    
    CheckPatient.Hide
    
        
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
    Exit Sub
                    
End Sub

