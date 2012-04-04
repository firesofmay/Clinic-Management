VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form UpdateStaff 
   Caption         =   "Update Staff"
   ClientHeight    =   8235
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   13650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "UpdateStaff.frx":0000
   ScaleHeight     =   8235
   ScaleWidth      =   13650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   11880
      Picture         =   "UpdateStaff.frx":25B89
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdSub 
      Height          =   735
      Left            =   10080
      Picture         =   "UpdateStaff.frx":262D5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Timer cmdsubenable 
      Interval        =   500
      Left            =   360
      Top             =   5760
   End
   Begin VB.TextBox txtStaffID 
      Height          =   375
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   5
      Left            =   9360
      MaxLength       =   20
      TabIndex        =   5
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   3
      Left            =   9360
      MaxLength       =   10
      TabIndex        =   3
      Top             =   4320
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   2
      Left            =   4320
      MaxLength       =   35
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   4
      Left            =   9360
      TabIndex        =   4
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   1
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   1
      Top             =   5040
      Width           =   2175
   End
   Begin VB.TextBox txtInfo 
      Height          =   375
      Index           =   0
      Left            =   4320
      MaxLength       =   20
      TabIndex        =   0
      Top             =   4320
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Caption         =   "Mandatory Fields * "
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
      Index           =   6
      Left            =   5400
      TabIndex        =   17
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label Label1 
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
      Index           =   1
      Left            =   1680
      TabIndex        =   16
      Top             =   5040
      Width           =   2055
   End
   Begin VB.Label Label1 
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
      Index           =   0
      Left            =   1680
      TabIndex        =   15
      Top             =   4320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff ID Selected"
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
      Height          =   495
      Index           =   0
      Left            =   4440
      TabIndex        =   14
      Top             =   2880
      Width           =   2775
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
      Index           =   3
      Left            =   7080
      TabIndex        =   11
      Top             =   4320
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
      Height          =   615
      Index           =   2
      Left            =   1680
      TabIndex        =   10
      Top             =   5760
      Width           =   2295
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
      Index           =   4
      Left            =   7080
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
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
      Index           =   5
      Left            =   7080
      TabIndex        =   8
      Top             =   5760
      Width           =   2055
   End
End
Attribute VB_Name = "UpdateStaff"
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
    
    'SINCE YOU CANNOT UPDATE A TABLE IF ITS NOT A SINGLE TABLE I HAVE TO DO NESTED QUERY....
    Set temp = New ADODB.Recordset
    temp.Open "select per.per_id, staff_id from person per, staff st where st.per_id = per.per_id and st.staff_id = " & txtStaffID.Text, Ado, adOpenKeyset, adLockOptimistic
    
    'NOW THIS TEMP(0) I.E. ID FOUND WILL BE USED TO SEARCH FOR HIS PERSON RECORD AND UPDATED
    Set Person = New ADODB.Recordset
    Person.Open "select fname, lname, addr, ph from person where per_id = " & temp(0), Ado, adOpenKeyset, adLockOptimistic
    
    For i = 0 To 3
        If txtInfo(i).Text <> "" Then
            Person(i) = txtInfo(i).Text
        End If
    Next i
    
    Person.Update
    
    Set Staff = New ADODB.Recordset
    Staff.Open "select pay, category from staff st where st.staff_id = " & txtStaffID.Text, Ado, adOpenKeyset, adLockOptimistic

    For i = 0 To 1
        If txtInfo(4 + i).Text <> "" Then
            Staff(i) = txtInfo(4 + i).Text
        End If
    Next i

    Staff.Update

    Unload Me
    LoginUser.Show
    
End Sub

Private Sub cmdsubenable_Timer()
    
    On Error GoTo donothing
    
    txtStaffID.Text = DataGrid1.Columns.Item(0).Value
    
    Set Staff = New ADODB.Recordset
    
    If txtStaffID.Text <> "" Then
        Staff.Open "select * from staff where staff_id = " & txtStaffID.Text, Ado, adOpenKeyset, adLockOptimistic
    
        If Staff.RecordCount > 0 Then
                Label1(6).Enabled = True
                For i = 0 To 5
                    Label1(i).Enabled = True
                    txtInfo(i).Enabled = True
                Next i
        
        ElseIf Staff.RecordCount = 0 Or txtStaffID.Text = "" Then
            Label1(6).Enabled = False
            For i = 0 To 5
                Label1(i).Enabled = False
                txtInfo(i).Enabled = False
            Next i
        
        End If
            
    Else
        Label1(6).Enabled = False
        For i = 0 To 5
            Label1(i).Enabled = False
            txtInfo(i).Enabled = False
        Next i
    
    End If

    If txtInfo(0).Text <> "" Or txtInfo(1).Text <> "" Or txtInfo(4).Text <> "" Or txtInfo(4).Text <> "" Or txtInfo(4).Text <> "" Or txtInfo(4).Text <> "" Then
        cmdSub.Enabled = True
    Else
        cmdSub.Enabled = False
    End If
    
    Exit Sub
donothing:
    
End Sub

Private Sub Form_Load()
    Set CheckStaff = New ADOR.Recordset
    CheckStaff.Fields.Append "Staff ID", adVarChar, 20
    CheckStaff.Fields.Append "First Name", adVarChar, 20
    CheckStaff.Fields.Append "Last Name", adVarChar, 20
    CheckStaff.Fields.Append "Address", adVarChar, 35
    CheckStaff.Fields.Append "Ph Number", adVarNumeric, 10
    CheckStaff.Fields.Append "Pay", adVarNumeric, 10
    CheckStaff.Fields.Append "Category", adVarChar, 20
    
    CheckStaff.CursorType = adOpenDynamic
    CheckStaff.Open
    
    
    Set Staff = New ADODB.Recordset
    Staff.Open "select staff_id, fname, lname, addr, ph, pay, category from staff st, person per where per.per_id = st.per_id", Ado, adOpenKeyset, adLockOptimistic
            
    While Not Staff.EOF
        CheckStaff.AddNew
        For j = 0 To 6
            If IsNull(Staff(j)) = False Then
                CheckStaff.Fields(j).Value = Staff(j)
            End If
        Next j
        Staff.MoveNext
    Wend
    
'
' Populate the datagrid
'
    If CheckStaff.RecordCount > 0 Then
        Set DataGrid1.DataSource = CheckStaff
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
    If Index = 1 Or Index = 0 Or Index = 5 Then
        If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If
    
    'PHONE NUMBER, PAY
    ElseIf Index = 3 Or Index = 4 Then
        If (KeyAscii >= 48 And KeyAscii <= 58) Or KeyAscii = 8 Then
        Else
            KeyAscii = 0
        End If

    End If
End Sub

