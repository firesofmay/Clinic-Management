VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UpdateTest 
   Caption         =   "Update Test"
   ClientHeight    =   6825
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10485
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "UpdateTest.frx":0000
   ScaleHeight     =   6825
   ScaleWidth      =   10485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   9120
      Picture         =   "UpdateTest.frx":1520F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   7200
      Picture         =   "UpdateTest.frx":1595B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtTestID 
      Height          =   495
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   2295
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   0
      Top             =   4920
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   4200
      MaxLength       =   10
      TabIndex        =   1
      Top             =   5400
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   2
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Timer cmdsubenable 
      Interval        =   500
      Left            =   240
      Top             =   4680
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   720
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
            LCID            =   16393
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
            LCID            =   16393
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
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   11
      Top             =   4080
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Result"
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
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      Top             =   5400
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Test ID Selected"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Test Name"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cost"
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
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   5880
      Width           =   1935
   End
End
Attribute VB_Name = "UpdateTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
    LoginUser.Show
End Sub

Private Sub cmdSub_Click()
    For i = 1 To 3
        Test(i) = UCase(txtInfo(i).Text)
    Next i
    
    Test.Update
    
    Unload Me
    LoginUser.Show
    
End Sub

Private Sub cmdsubenable_Timer()
    
    On Error GoTo donothing
    
    txtTestID.Text = DataGrid1.Columns.Item(0).Value
    
    Set Test = New ADODB.Recordset
    
    If txtTestID.Text <> "" Then
        Test.Open "select * from test where test_id = " & txtTestID.Text, Ado, adOpenKeyset, adLockOptimistic
    
        If Test.RecordCount > 0 Then
                Label1(4).Enabled = True
                For i = 1 To 3
                    Label1(i).Enabled = True
                    txtInfo(i).Enabled = True
                Next i
        
        ElseIf Test.RecordCount = 0 Or txtTestID.Text = "" Then
            Label1(4).Enabled = False
            For i = 1 To 3
                Label1(i).Enabled = False
                txtInfo(i).Enabled = False
            Next i
        
        End If
            
    Else
        Label1(4).Enabled = False
        For i = 1 To 3
            Label1(i).Enabled = False
            txtInfo(i).Enabled = False
        Next i
    
    End If

    If txtInfo(1).Text <> "" And txtInfo(2).Text <> "" And txtInfo(3).Text <> "" Then
        cmdSub.Enabled = True
    Else
        cmdSub.Enabled = False
    End If
    
    Exit Sub
donothing:
    
End Sub

Private Sub Form_Load()
    Set CheckTest = New ADOR.Recordset
    CheckTest.Fields.Append "Test ID", adVarNumeric, 10
    CheckTest.Fields.Append "Test Name", adVarChar, 10
    CheckTest.Fields.Append "Normal Result", adVarChar, 10
    CheckTest.Fields.Append "Cost", adVarNumeric, 10
    
    CheckTest.CursorType = adOpenDynamic
    CheckTest.Open
    
    
    Set Test = New ADODB.Recordset
    Test.Open "select * from test", Ado, adOpenKeyset, adLockOptimistic
            
    While Not Test.EOF
        CheckTest.AddNew
        For j = 0 To 3
            If IsNull(Test(j)) = False Then
                CheckTest.Fields(j).Value = Test(j)
            End If
        Next j
        Test.MoveNext
    Wend
    
'
' Populate the datagrid
'
    If CheckTest.RecordCount > 0 Then
        Set DataGrid1.DataSource = CheckTest
    Else
        MsgBox "No Record Found"
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
    If Index = 1 Then
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
    'PHONE NUMBER, PAY
    ElseIf Index = 3 Then
        If (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

    End If
End Sub

