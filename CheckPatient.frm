VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CheckPatient 
   Caption         =   "Search Patient"
   ClientHeight    =   6960
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   13455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "CheckPatient.frx":0000
   ScaleHeight     =   6960
   ScaleWidth      =   13455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   9360
      Picture         =   "CheckPatient.frx":3178A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   9600
      Picture         =   "CheckPatient.frx":324D7
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   6
      Left            =   5040
      TabIndex        =   5
      Top             =   6240
      Width           =   2535
   End
   Begin VB.OptionButton OptionSex 
      BackColor       =   &H00000000&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   5040
      TabIndex        =   3
      Top             =   5640
      Width           =   1095
   End
   Begin VB.OptionButton OptionSex 
      BackColor       =   &H00000000&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   1
      Left            =   6240
      TabIndex        =   4
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtPatID 
      Height          =   495
      Left            =   11040
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
      Default         =   -1  'True
      Height          =   615
      Left            =   9120
      Picture         =   "CheckPatient.frx":32C23
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   0
      Left            =   5040
      TabIndex        =   0
      Top             =   3720
      Width           =   2535
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   1
      Top             =   4320
      Width           =   2535
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   5040
      TabIndex        =   2
      Top             =   4920
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2400
      TabIndex        =   15
      Top             =   6240
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex"
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
      Index           =   4
      Left            =   2400
      TabIndex        =   14
      Top             =   5520
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
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   13
      Top             =   3720
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
      Left            =   2400
      TabIndex        =   12
      Top             =   4320
      Width           =   1935
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
      Left            =   2400
      TabIndex        =   11
      Top             =   4920
      Width           =   1935
   End
End
Attribute VB_Name = "CheckPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Patient, zero As ADODB.Recordset
Dim temp_pat_id As String
Dim temp_fname, temp_lname, ph, Sex As String
Public patid As Integer
Dim flag As Boolean


Private Sub cmdCancel_Click()
    Unload Me
    LoginUser.Show
End Sub

Private Sub cmdSearch_Click()
    
    'REINITIALIZNG VALUES
    Sex = ""
    temp_pat_id = ""
    temp_fname = ""
    temp_lname = ""
    ph = ""
    
    If OptionSex(0).Value = True Then
        Sex = "M"
    ElseIf OptionSex(1).Value = True Then
        Sex = "F"
    End If
    
    temp_pat_id = (txtInfo(0).Text)
    temp_fname = (txtInfo(1).Text) & "%"
    temp_lname = (txtInfo(2).Text) & "%"
    ph = txtInfo(6).Text
    
    
    Set Patient = New ADODB.Recordset
        
    'CHECKING FOR PATIENT ID
    If txtInfo(0).Text <> "" Then
        DataGrid1.Visible = True
        Patient.Open "select pat.pat_id, per.per_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph from person per ,patient pat where pat.per_id = per.per_id and pat.pat_id = " & temp_pat_id, Ado, adOpenKeyset, adLockOptimistic
        
    'CHECKING BY PATIENT'S PHONE NUMBER
    ElseIf txtInfo(6).Text <> "" Then
        DataGrid1.Visible = True
        Patient.Open "select pat.pat_id, per.per_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph from person per ,patient pat where pat.per_id = per.per_id and per.ph = " & ph, Ado, adOpenKeyset, adLockOptimistic
    
    'CHECKING BY FIRST NAME AND LAST NAME
    ElseIf txtInfo(1).Text <> "" And txtInfo(2).Text <> "" Then
        DataGrid1.Visible = True
        Patient.Open "select per.per_id, pat.pat_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph from person per ,patient pat where pat.per_id = per.per_id and per.fname like '" & temp_fname & "' and per.lname like '" & temp_lname & "'", Ado, adOpenKeyset, adLockOptimistic
    
    'CHECKING BY FIRST NAME
    ElseIf txtInfo(1).Text <> "" Then
        DataGrid1.Visible = True
        Patient.Open "select pat.pat_id, per.per_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph from person per ,patient pat where pat.per_id = per.per_id and per.fname like '" & temp_fname & "'", Ado, adOpenKeyset, adLockOptimistic
    
    'CHECKING BY LAST NAME
    ElseIf txtInfo(2).Text <> "" Then
        DataGrid1.Visible = True
        Patient.Open "select pat.pat_id, per.per_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph from person per ,patient pat where pat.per_id = per.per_id and per.lname like '" & temp_lname & "'", Ado, adOpenKeyset, adLockOptimistic
    
    ElseIf Sex <> "" Then
        DataGrid1.Visible = True
        Patient.Open "select pat.pat_id, per.per_id, per.fname, per.lname, per.dob, per.sex, per.addr, per.ph from person per ,patient pat where pat.per_id = per.per_id and per.sex = '" & Sex & "'", Ado, adOpenKeyset, adLockOptimistic
    
    
    'NO VALUES ARE ENTERED
    Else
        DataGrid1.Visible = False
        MsgBox "Enter Atleast One of the fields"
        GoTo oversub
    End If
    
    ' Setup the fields - you can use any valid type of
    ' ado field type for these. I've used VarChar just
    ' for testing / demonstration purposes.
    
    
    flag = False
    
    While Not Patient.EOF
        flag = True
        
        CheckPat.AddNew
        For j = 0 To 7
            If IsNull(Patient(j)) = False Then
                CheckPat.Fields(j).Value = Patient(j)
            End If
        Next j
        Patient.MoveNext
    Wend
    
'
' Populate the datagrid
'
    If flag = True Then
        Set DataGrid1.DataSource = CheckPat
    Else
        MsgBox "No Record Found"
    End If
    
oversub:


End Sub

Private Sub cmdSub_Click()
    On Error GoTo display
    
    txtPatID.Text = DataGrid1.Columns.Item(0).Text
    If txtPatID.Text <> "" Then
        patid = txtPatID.Text
        Appointment.Show
    Else
        MsgBox "Enter Patient ID"
    End If
    Exit Sub

display:
    MsgBox "Search and Select Patient Row First"

End Sub

Private Sub Form_Load()
    'FOR CHECKPATEINT FORM
    Set CheckPat = New ADOR.Recordset
    CheckPat.Fields.Append "Pat_ID", adVarNumeric, 10
    CheckPat.Fields.Append "Person_ID", adVarNumeric, 10
    CheckPat.Fields.Append "First Name", adVarChar, 20
    CheckPat.Fields.Append "Last Name", adVarChar, 20
    CheckPat.Fields.Append "DOB", adDate, 12
    CheckPat.Fields.Append "SEX", adVarChar, 1
    CheckPat.Fields.Append "ADDRESS", adVarChar, 35
    CheckPat.Fields.Append "PHONE NUMBER", adVarNumeric, 10
    
    CheckPat.CursorType = adOpenDynamic
    CheckPat.Open

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
    ElseIf Index = 6 Or Index = 0 Then
        If (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

    End If
End Sub

