VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "mscal.ocx"
Begin VB.Form StaffInfo 
   Caption         =   "Staff Info"
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11610
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   Picture         =   "StaffInfo.frx":0000
   ScaleHeight     =   8640
   ScaleWidth      =   11610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   8040
      Picture         =   "StaffInfo.frx":25B89
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   9840
      Picture         =   "StaffInfo.frx":268D6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6720
      Width           =   855
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   0
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   7
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   2
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   3
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   8
      Left            =   8880
      TabIndex        =   9
      Top             =   2280
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   5
      Left            =   8880
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   405
      Index           =   6
      Left            =   8880
      TabIndex        =   8
      Top             =   1680
      Width           =   2175
   End
   Begin VB.OptionButton Sex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Male"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
   Begin VB.OptionButton Sex 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "Female"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   9
      Left            =   8880
      MaxLength       =   20
      TabIndex        =   10
      Top             =   2880
      Width           =   2175
   End
   Begin VB.Timer tmrSubEnable 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin MSACAL.Calendar dob_calendar 
      Height          =   2895
      Left            =   1920
      TabIndex        =   6
      Top             =   5520
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
      Left            =   3960
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
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
      Height          =   1215
      Index           =   3
      Left            =   3000
      TabIndex        =   24
      Top             =   4320
      Width           =   2655
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
      Left            =   1200
      TabIndex        =   23
      Top             =   1080
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
      Left            =   1200
      TabIndex        =   22
      Top             =   2280
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
      Height          =   375
      Index           =   2
      Left            =   6720
      TabIndex        =   21
      Top             =   1080
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
      Left            =   1200
      TabIndex        =   20
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
      Height          =   375
      Index           =   6
      Left            =   6720
      TabIndex        =   19
      Top             =   1680
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
      Left            =   1200
      TabIndex        =   18
      Top             =   2880
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
      Left            =   1200
      TabIndex        =   17
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Index           =   8
      Left            =   6720
      TabIndex        =   16
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monthly Pay *"
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
      Index           =   7
      Left            =   6720
      TabIndex        =   15
      Top             =   2280
      Width           =   1935
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
      Height          =   615
      Index           =   2
      Left            =   5040
      TabIndex        =   13
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "StaffInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Person, Staff, id_table As ADODB.Recordset
Dim dob As String

Private Sub cmdCancel_Click()
    Unload Me
    LoginUser.Show
End Sub

Private Sub cmdSub_Click()
'    On Error GoTo disablecmdsub
    
    If Sex(0).Value = True Then
        txtInfo(4).Text = "M"
    Else
        txtInfo(4).Text = "F"
    End If
    
    Set Staff = New ADODB.Recordset
    Staff.Open "select * from staff", Ado, adOpenKeyset, adLockOptimistic
    
    Set Person = New ADODB.Recordset
    Person.Open "select * from person", Ado, adOpenKeyset, adLockOptimistic
    
    Person.AddNew
    Staff.AddNew
    
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
    
    
    'ACCPETING VALUES FOR STAFF TABLE
    'staff ID
    Staff(0) = txtInfo(7).Text
    
    'person ID
    Staff(1) = txtInfo(0).Text
    
    'monthly pay of staff
    If txtInfo(8).Text = "" Then
        Staff(2) = Null
    Else
        Staff(2) = txtInfo(8).Text
    End If
    
    'category of staff
    Staff(3) = UCase(txtInfo(9).Text)
    
    Person.Update
    Staff.Update

    Unload Me
    LoginUser.Show

disablecmdsub:
    cmdSub.Enabled = False
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
    
    Set Staff = New ADODB.Recordset
    Staff.Open "select * from staff", Ado, adOpenKeyset, adLockOptimistic
    
    Set Person = New ADODB.Recordset
    Person.Open "select * from person", Ado, adOpenKeyset, adLockOptimistic

    txtInfo(0).Text = Person.RecordCount + 1
    txtInfo(7).Text = Staff.RecordCount + 1
                    
    Sex(0).Value = True

End Sub



'THIS TIMER MAKE SURES SUBMIT BUTTON IS ENABLED ONLY WHEN MANDATORY FIELDS HAVE BEEN SELECTED
Private Sub tmrSubEnable_Timer()
    If txtInfo(0).Text = "" Or txtInfo(1).Text = "" Or txtInfo(2).Text = "" Or txtInfo(8).Text = "" Then
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
    If Index = 1 Or Index = 2 Or Index = 9 Then
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
    'PHONE NUMBER, PAY
    ElseIf Index = 6 Or Index = 8 Or Index = 9 Then
        If (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

    End If
End Sub


