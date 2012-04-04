VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form DoctorInfo 
   Caption         =   "DoctorInfo"
   ClientHeight    =   10320
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   15225
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "DoctorInfo.frx":0000
   ScaleHeight     =   10320
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   13080
      Picture         =   "DoctorInfo.frx":4E829
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8520
      Width           =   855
   End
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   11520
      Picture         =   "DoctorInfo.frx":4EF75
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8520
      Width           =   1335
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   9
      Left            =   11520
      MaxLength       =   20
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
   End
   Begin VB.OptionButton Sex 
      BackColor       =   &H00000000&
      Caption         =   "Female"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   6480
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.OptionButton Sex 
      BackColor       =   &H00000000&
      Caption         =   "Male"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   8
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   6
      Left            =   11520
      TabIndex        =   7
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   5
      Left            =   11520
      TabIndex        =   6
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   8
      Left            =   11520
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   7
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   0
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Timer tmrSubEnable 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin MSACAL.Calendar dob_calendar 
      Height          =   2895
      Left            =   3600
      TabIndex        =   13
      Top             =   6600
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
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   4
      Left            =   5400
      TabIndex        =   14
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Specialization"
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
      Index           =   7
      Left            =   9000
      TabIndex        =   24
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Education"
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
      Index           =   8
      Left            =   9000
      TabIndex        =   23
      Top             =   3360
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff ID*"
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
      Left            =   2640
      TabIndex        =   22
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label First 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name *"
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
      Left            =   2640
      TabIndex        =   21
      Top             =   3360
      Width           =   1935
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
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   6
      Left            =   9000
      TabIndex        =   20
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sex *"
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
      Index           =   4
      Left            =   2640
      TabIndex        =   19
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Index           =   2
      Left            =   9000
      TabIndex        =   18
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label First 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name *"
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
      Left            =   2640
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Person ID *"
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
      Left            =   2640
      TabIndex        =   16
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Of Birth * (DD-MMM-YY)"
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
      Index           =   3
      Left            =   4920
      TabIndex        =   15
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mandatory Fields *"
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
      Index           =   2
      Left            =   6120
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "DoctorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Person, Doctor As ADODB.Recordset
Dim dob As String


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
        
    Person.AddNew
    Doctor.AddNew
    
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
    
    
    'ACCPETING VALUES FOR DOCTOR TABLE
    'doctor ID
    Doctor(0) = txtInfo(7).Text
    
    'person ID
    Doctor(1) = txtInfo(0).Text
    
    'SPECIALINZATION DOCTOR
    If txtInfo(8).Text = "" Then
        Doctor(2) = Null
    Else
        Doctor(2) = UCase(txtInfo(8).Text)
    End If
    
    'EDUCATION OF DOCTOR
    If txtInfo(9).Text = "" Then
        Doctor(3) = Null
    Else
        Doctor(3) = UCase(txtInfo(9).Text)
    End If
    
    Person.Update
    Doctor.Update

    Unload Me
    LoginUser.Show

End Sub

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

Private Sub Form_Load()
    dob_calendar.Value = Date
    dob = dob_calendar.Value
    
    Set Doctor = New ADODB.Recordset
    Doctor.Open "select * from doctor", Ado, adOpenKeyset, adLockOptimistic
    
    Set Person = New ADODB.Recordset
    Person.Open "select * from person", Ado, adOpenKeyset, adLockOptimistic

    
    txtInfo(0).Text = Person.RecordCount + 1
    txtInfo(7).Text = Doctor.RecordCount + 1
                 
    Sex(0).Value = True

End Sub



'THIS TIMER MAKE SURES SUBMIT BUTTON IS ENABLED ONLY WHEN MANDATORY FIELDS HAVE BEEN SELECTED
Private Sub tmrSubEnable_Timer()
    If txtInfo(0).Text = "" Or txtInfo(1).Text = "" Or txtInfo(2).Text = "" Then
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
    
    'FIRST NAME AND LAST NAME AND CATEGORY
    If Index = 1 Or Index = 2 Or Index = 9 Or Index = 8 Or Index = 9 Then
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
    'PHONE NUMBER, PAY
    ElseIf Index = 6 Then
        If (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

    End If
End Sub
