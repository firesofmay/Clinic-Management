VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form LoginUser 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User"
   ClientHeight    =   9450
   ClientLeft      =   3975
   ClientTop       =   1545
   ClientWidth     =   13605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LoginUser.frx":0000
   ScaleHeight     =   9450
   ScaleWidth      =   13605
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   1270
      ButtonWidth     =   4763
      ButtonHeight    =   1111
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Appointment (Ctrl + A)"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Patient (Ctrl + N)"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Patient (Ctrl + S)"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print Bill (Ctrl + P)"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit (Ctrl + X)"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2160
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   4
      Left            =   4680
      TabIndex        =   5
      Top             =   3720
      Width           =   90
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Top             =   3360
      Width           =   90
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   4680
      TabIndex        =   3
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   2640
      Width           =   90
   End
   Begin VB.Label lblinfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Index           =   0
      Left            =   4680
      TabIndex        =   1
      Top             =   2280
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome To Sapphire Clinic"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   3840
      TabIndex        =   0
      Top             =   1080
      Width           =   6135
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Begin VB.Menu mnuPatient 
         Caption         =   "Patient"
         Begin VB.Menu mnuNewpat 
            Caption         =   "New Patient"
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuSearchPat 
            Caption         =   "Search Patient"
            Shortcut        =   ^S
         End
      End
      Begin VB.Menu mnuDoctor 
         Caption         =   "Doctor"
         Begin VB.Menu mnuTodayAppt 
            Caption         =   "Today's Appointment"
            Shortcut        =   ^A
         End
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff"
         Begin VB.Menu mnuNewDoc 
            Caption         =   "New Doctor"
            Shortcut        =   ^D
         End
         Begin VB.Menu mnuAddStaff 
            Caption         =   "Add Staff"
            Shortcut        =   ^F
         End
         Begin VB.Menu mnuSearchStaff 
            Caption         =   "Update  Staff"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnuAddTest 
            Caption         =   "Add Test"
            Shortcut        =   ^T
         End
         Begin VB.Menu mnuUpdateTest 
            Caption         =   "Update Test"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuPrintBill 
            Caption         =   "Print Bill"
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuFactory 
      Caption         =   "Factory Setting"
      Begin VB.Menu mnuSetFactory 
         Caption         =   "Set Factory Settings"
      End
   End
End
Attribute VB_Name = "LoginUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Patient, Doctor, Test, Medical, Staff As ADODB.Recordset
Dim op As Integer

Dim timepass As ADODB.Recordset


Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuAddStaff_Click()
    StaffInfo.Show
    Me.Hide
End Sub

Private Sub mnuAddTest_Click()
    AddTest.Show
    Me.Hide
End Sub

Private Sub mnuNewDoc_Click()
    DoctorInfo.Show
    Me.Hide
End Sub

Private Sub mnuNewpat_Click()
    PatientInfo.Show
    Me.Hide
End Sub

Private Sub mnuPrintBill_Click()
    PrintBill.Show
    Me.Hide
End Sub


Private Sub mnuSearchPat_Click()
    CheckPatient.Show
    Me.Hide
End Sub


Private Sub mnuSearchStaff_Click()
    UpdateStaff.Show
    Me.Hide
End Sub

Private Sub mnuSetFactory_Click()
    If MsgBox("Are you sure?", vbYesNo, "Conirm Delete") = vbYes Then
        FactorySetting.Show
        Me.Hide
    End If
End Sub

Private Sub mnuTodayAppt_Click()
    Todayappt.Show
    Me.Hide
End Sub

Private Sub mnuUpdateTest_Click()
    UpdateTest.Show
    Me.Hide
End Sub

Private Sub Timer1_Timer()
    On Error GoTo dontcheck
        
    Set Doctor = New ADODB.Recordset
    Doctor.Open "select * from doctor", Ado, adOpenKeyset, adLockOptimistic

    Set Patient = New ADODB.Recordset
    Patient.Open "select * from patient", Ado, adOpenKeyset, adLockOptimistic

    Set Test = New ADODB.Recordset
    Test.Open "select * from test", Ado, adOpenKeyset, adLockOptimistic

    Set Staff = New ADODB.Recordset
    Staff.Open "select * from staff", Ado, adOpenKeyset, adLockOptimistic

    Set Medical = New ADODB.Recordset
    Medical.Open "select * from medical", Ado, adOpenKeyset, adLockOptimistic
      
    For i = 0 To 4
        lblinfo(i).Caption = " "
    Next i
      
    lblinfo(0).Caption = "Number of Patients = " & Patient.RecordCount
   
    If Doctor.RecordCount = 0 Then
        lblinfo(1).Caption = "Doctor Database Empty! Please Enter atleast One Doctor"
    Else
        lblinfo(1).Caption = "Number of Doctors = " & Doctor.RecordCount & vbCrLf
    End If
     
    If Staff.RecordCount = 0 Then
        lblinfo(2).Caption = "Staff Database Empty! Please Enter atleast One Staff"
    Else
        lblinfo(2).Caption = "Number of Staffs = " & Staff.RecordCount
    End If
    
    If Test.RecordCount = 0 Then
        lblinfo(3).Caption = "Warning : Test list Empty"
    End If
    
    If Medical.RecordCount = 0 Then
        lblinfo(4).Caption = "Warning : Medical History list Empty"
    End If
    
    Exit Sub
dontcheck:
        lblinfo(0).Caption = "Tables Missing"
        lblinfo(1).Caption = ""
        lblinfo(2).Caption = ""
        lblinfo(3).Caption = ""
        lblinfo(4).Caption = ""
        
    
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
        
    Select Case (Button)
    
        Case "New Patient (Ctrl + N)":
            mnuNewpat_Click
        
        Case "Appointment (Ctrl + A)"
            mnuTodayAppt_Click
    
        Case "Search Patient (Ctrl + S)"
            mnuSearchPat_Click
        
        Case "Print Bill (Ctrl + P)"
            mnuPrintBill_Click
        
        Case "Exit (Ctrl + X)"
            mnuExit_Click
    End Select

End Sub
