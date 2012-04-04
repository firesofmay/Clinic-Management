VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form PatientInfo 
   Caption         =   "Patient Info"
   ClientHeight    =   10260
   ClientLeft      =   1515
   ClientTop       =   870
   ClientWidth     =   16470
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "PatientInfo.frx":0000
   ScaleHeight     =   10260
   ScaleWidth      =   16470
   Begin VB.OptionButton Sex 
      BackColor       =   &H00000000&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   4920
      TabIndex        =   10
      Top             =   3720
      Width           =   975
   End
   Begin VB.ListBox lsthistory 
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
      ItemData        =   "PatientInfo.frx":46DD0
      Left            =   10560
      List            =   "PatientInfo.frx":46DD2
      TabIndex        =   16
      Top             =   4920
      Width           =   2175
   End
   Begin VB.ComboBox cmbDr 
      Height          =   315
      ItemData        =   "PatientInfo.frx":46DD4
      Left            =   8160
      List            =   "PatientInfo.frx":46DD6
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1560
      Width           =   2775
   End
   Begin VB.ListBox lstmedical 
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
      ItemData        =   "PatientInfo.frx":46DD8
      Left            =   13320
      List            =   "PatientInfo.frx":46DDA
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
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
      ItemData        =   "PatientInfo.frx":46DDC
      Left            =   7920
      List            =   "PatientInfo.frx":46DDE
      TabIndex        =   15
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   9
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   11
      Top             =   5760
      Width           =   2175
   End
   Begin VB.OptionButton Sex 
      BackColor       =   &H00000000&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   3840
      TabIndex        =   9
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   6
      Left            =   3720
      MaxLength       =   10
      TabIndex        =   8
      Top             =   4680
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   5
      Left            =   3720
      MaxLength       =   35
      TabIndex        =   7
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   8
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   6
      Top             =   5160
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   5
      Top             =   3240
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2640
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   7
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   0
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ListBox lstTemptime 
      Height          =   255
      ItemData        =   "PatientInfo.frx":46DE0
      Left            =   4605
      List            =   "PatientInfo.frx":46E08
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   -22800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   14160
      Picture         =   "PatientInfo.frx":46E90
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9000
      Width           =   855
   End
   Begin VB.Timer tmrSubEnable 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   12120
      Picture         =   "PatientInfo.frx":475DC
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   9000
      Width           =   1335
   End
   Begin MSACAL.Calendar dob_calendar 
      Height          =   2895
      Left            =   1800
      TabIndex        =   12
      Top             =   7080
      Width           =   4335
      _Version        =   524288
      _ExtentX        =   7646
      _ExtentY        =   5106
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2009
      Month           =   9
      Day             =   29
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   2895
      Left            =   11280
      TabIndex        =   14
      Top             =   1200
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
      Index           =   4
      Left            =   3720
      TabIndex        =   18
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Medical History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   4
      Left            =   10560
      TabIndex        =   34
      Top             =   4440
      Width           =   2160
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of Medical History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Index           =   3
      Left            =   12960
      TabIndex        =   33
      Top             =   4440
      Width           =   3060
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
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
      Height          =   615
      Index           =   0
      Left            =   8640
      TabIndex        =   32
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor Name"
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
      Left            =   8520
      TabIndex        =   31
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Height (cms.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   7
      Left            =   960
      TabIndex        =   30
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight (Kgs.)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   8
      Left            =   960
      TabIndex        =   29
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient ID *"
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
      Height          =   375
      Index           =   9
      Left            =   960
      TabIndex        =   28
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label First 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name *"
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
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   27
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   960
      TabIndex        =   26
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex *"
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
      Height          =   375
      Index           =   4
      Left            =   960
      TabIndex        =   25
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   24
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label First 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name *"
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
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   23
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Person ID *"
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
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   22
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth * (DD-MMM-YY)"
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
      Height          =   615
      Index           =   3
      Left            =   2040
      TabIndex        =   20
      Top             =   6600
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mandatory Fields *"
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
      Height          =   615
      Index           =   2
      Left            =   6840
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "PatientInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim appt_date, dob As String
Dim doc_id As Integer
Dim temp As String

Private Sub dob_calendar_Click()
    
    If dob_calendar.Value > Date Then
        MsgBox "Invalid"
        dob_calendar.Value = dob
    Else
        dob = dob_calendar.Day
        
        Select Case (dob_calendar.Month)
        Case 1
           dob = dob & "-Jan-"
        Case 2
           dob = dob & "-Feb-"
        Case 3
           dob = dob & "-Mar-"
        Case 4
           dob = dob & "-Apr-"
        Case 5
           dob = dob & "-May-"
        Case 6
           dob = dob & "-Jun-"
        Case 7
           dob = dob & "-Jul-"
        Case 8
           dob = dob & "-Aug-"
        Case 9
           dob = dob & "-Sep-"
        Case 10
           dob = dob & "-Oct-"
        Case 11
           dob = dob & "-Nov-"
        Case 12
           dob = dob & "-Dec-"
        End Select
        
        dob = dob & dob_calendar.Year
    End If
End Sub


Private Sub Calendar1_Click()
        
    If Calendar1.Value < Date Then
        MsgBox "Appointment Before Today is not allowed"
        Calendar1.Value = appt_date
    Else
    
        
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
    End If
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
    LoginUser.Show
End Sub

Private Sub cmdSub_Click()
    
    If Sex(0).Value = True Then
        txtInfo(4).Text = "M"
    Else
        txtInfo(4).Text = "F"
    End If
     
    Set Person = New ADODB.Recordset
    Person.Open "select * from person", Ado, adOpenKeyset, adLockOptimistic

    Set Patient = New ADODB.Recordset
    Patient.Open "select * from patient", Ado, adOpenKeyset, adLockOptimistic

    Person.AddNew
    Patient.AddNew
    
    'ACCEPTING VALUES OF PERSON TABLE
    For i = 0 To 6
    
        If i = 3 Then
            'as dob is a string not from the textbox array
            Person(i) = dob
        Else
            'IF DOB, ADDRESS, PHONE NUMBER IS NULL, INSERT NULL, ELSE THE VALUE
            If (i = 6 Or i = 5 Or i = 8) And (txtInfo(i).Text = "") Then
                    Person(i) = Null
            Else
                Person(i) = UCase(txtInfo(i).Text)
            End If
        End If
    Next i
    
    'ACCPETING VALUES FOR PATIENT TABLE
    'PATIENT ID
    Patient(0) = txtInfo(7).Text
    
    'PERSON ID
    Patient(1) = txtInfo(0).Text
    
    'HEIGHT PATIENT
    If txtInfo(7).Text = "" Then
        Patient(2) = Null
    Else
        Patient(2) = txtInfo(7).Text
    End If
    
    'WEIGHT OF PATEINT
    If txtInfo(8).Text = "" Then
        Patient(3) = Null
    Else
        Patient(3) = txtInfo(8).Text
    End If
    
    Person.Update
    Patient.Update
        
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
    
    'PATIENT_ID
    Appt(4) = txtInfo(7).Text
        
    Appt.Update
    
    If lstHistory.ListCount >= 0 Then

        Set Medical = New ADODB.Recordset
        Medical.Open "select * from medical", Ado, adOpenKeyset, adLockOptimistic

        Set History = New ADODB.Recordset
        History.Open "select * from history", Ado, adOpenKeyset, adLockOptimistic

        For i = 0 To lstHistory.ListCount - 1
                Medical.MoveFirst
                
                While Not Medical.EOF
                    If Medical(2) = lstHistory.List(i) Then
                        History.AddNew
                        History(0) = Medical(0)
                        History(1) = txtInfo(7).Text
                        History(2) = Null
                        History.Update
                        GoTo over
                    Else
                        Medical.MoveNext
                    End If
                Wend
over:
        Next i

    End If
    
    
    Unload Me
    LoginUser.Show

End Sub




Private Sub Form_Load()
    
    'INITIALIZING VALUE OF APPT CALENDAR, APPT_DATE, DOB
    Calendar1.Value = Date
    appt_date = Calendar1.Value
    dob_calendar.Value = Date
    dob = dob_calendar.Value
    
    Set Person = New ADODB.Recordset
    Person.Open "select * from person", Ado, adOpenKeyset, adLockOptimistic
    
    Set Patient = New ADODB.Recordset
    Patient.Open "select * from patient", Ado, adOpenKeyset, adLockOptimistic
    
    'SHOWING PERSON ID AND PATIENT ID
    txtInfo(0).Text = Person.RecordCount + 1
    txtInfo(7).Text = Patient.RecordCount + 1
    
    'TO ADD DOCTORS NAME IN THE DOCTOR NAME DROP DOWN LIST
    Set Doctor = New ADODB.Recordset
    Doctor.Open "select fname, lname from doctor d, person p where p.per_id = d.per_id", Ado, adOpenKeyset, adLockOptimistic
    
    While Not Doctor.EOF
            'CONCATINATING FIRST AND LAST NAME OF THE DR.
            cmbDr.AddItem (Doctor(0) & " " & Doctor(1))
            Doctor.MoveNext
    Wend
    
    'TO ADD MEDICAL HISTORY VALUES TO THE LIST FROM WHERE YOU CAN ADD FOR THE PATEINT
    Set Medical = New ADODB.Recordset
    Medical.Open "select disease, med_id from medical", Ado, adOpenKeyset, adLockOptimistic
    
    While Not Medical.EOF
            lstmedical.AddItem (Medical(0))
            Medical.MoveNext
    Wend
       
       
    'TO SET THE appt_date VALUE TO DEFAULT TODAY'S DATE
    Call Calendar1_Click
    
    'SETTNG MALE AS DEFAULT VALUE
    Sex(0).Value = True
             
End Sub

'THIS REMOVES A MEDICAL HISTORY FROM THE PATIENT AND PUTS IT BACK TO THE MEDICAL LIST
Private Sub lsthistory_Click()
        lstmedical.AddItem (lstHistory.List(lstHistory.ListIndex))
        lstHistory.RemoveItem (lstHistory.ListIndex)
End Sub

'THIS ADDS A MEDICAL HISTORY TO THE PATIENT AND REMOVES IT FROM THE MEDICAL LIST
Private Sub lstmedical_Click()
        lstHistory.AddItem (lstmedical.List(lstmedical.ListIndex))
        lstmedical.RemoveItem (lstmedical.ListIndex)
End Sub


'THIS TIMER MAKE SURES SUBMIT BUTTON IS ENABLED ONLY WHEN MANDATORY FIELDS HAVE BEEN SELECTED
Private Sub tmrSubEnable_Timer()
    If txtInfo(0).Text = "" Or txtInfo(1).Text = "" Or txtInfo(2).Text = "" Or cmbT.ListIndex < 0 Or cmbDr.ListIndex < 0 Then
        cmdSub.Enabled = False
    Else
        cmdSub.Enabled = True
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
    
    'FIRST NAME AND LAST NAME
    If Index = 1 Or Index = 2 Then
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
    'PHONE NUMBER, WEIGHT, HEIGHT
    ElseIf Index = 6 Or Index = 8 Or Index = 9 Then
        If (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

    End If
End Sub
