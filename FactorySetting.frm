VERSION 5.00
Begin VB.Form FactorySetting 
   Caption         =   "Factory Setting"
   ClientHeight    =   3540
   ClientLeft      =   165
   ClientTop       =   1035
   ClientWidth     =   8580
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "FactorySetting.frx":0000
   ScaleHeight     =   3540
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtinfo 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "iagree"
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox txtinfo 
      Height          =   495
      Index           =   0
      Left            =   4200
      TabIndex        =   0
      Text            =   "master"
      Top             =   960
      Width           =   3135
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000FF&
      Height          =   420
      Left            =   4305
      TabIndex        =   4
      Top             =   2400
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Master Password"
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
      Height          =   420
      Index           =   1
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Master Username"
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
      Height          =   420
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   3630
   End
   Begin VB.Menu mnuFactoru 
      Caption         =   "Factory Setting"
      Begin VB.Menu mnuDrop 
         Caption         =   "Drop All Tables"
      End
      Begin VB.Menu mnuTables 
         Caption         =   "Create Standard Tables"
      End
      Begin VB.Menu mnuInsertValues 
         Caption         =   "Insert Standard Values"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "FactorySetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Set temp = New ADODB.Recordset
End Sub

Private Sub mnuDrop_Click()
    On Error Resume Next
    
    If txtinfo(0).Text = "master" And txtinfo(1).Text = "iagree" Then
        temp.Open "drop table test_result", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table test", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table staff", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table appt", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table history", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table medical", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table presc", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table patient", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table doctor", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table person", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "drop table for2mins", Ado, adOpenKeyset, adLockOptimistic
        lblmsg.Caption = "Database Deleted!"
    Else
        lblmsg.Caption = "Incorrect Username and Password"
    End If
                
End Sub

Private Sub mnuExit_Click()
    Unload Me
    LoginUser.Show
End Sub

Private Sub mnuInsertValues_Click()
    On Error Resume Next
    
    If txtinfo(0).Text = "master" And txtinfo(1).Text = "iagree" Then
        temp.Open "insert all into person values (1, 'MAYANK', 'JAIN', '11-May-85', 'M', 'KATRAJ', 9371744602) into person values (2, 'AMIT', 'SINGHAL', '3-Feb-79', 'M', 'SATARA', 9171234591) into person values (3, 'ATUL', 'SHARMA', '4-Mar-75', 'M', 'SATARA',  9276892390) into person values (4, 'ANITA', 'KUKREJA', '5-DEC-60', 'F', 'AUNDH', 9716446021) into person values (5, 'SHALINI', 'KULKARNI', '6-Feb-78', 'F', 'FC ROAD', 9919127891) into person values (6, 'ANIL', 'VERMA', '11-MAR-82', 'M', 'KATRAJ', 9371742302) into person values (7, 'AMEEY', 'SHINOY', '3-Feb-79', 'F', 'SATARA', 9171123591) into person values (8, 'VIKAS', 'KULKARNI', '4-Mar-72', 'M', 'AUNDH PUNE',  9274592390) into person values (9, 'PRATIMA', 'KULKARNI', '5-DEC-73', 'F', 'AUNDH PUNE', 9716946021) into person values (10, 'SHALU', 'REHMAN', '6-Feb-90', 'F', 'AUNDH', 9918997891) select * from dual", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "insert all into patient values (1, 1, 150, 67) into patient values (2, 2, 170, 100) into patient values (3, 3, 187, 89) into patient values (4, 4, 156, 56) into patient values (5, 5, 178, 82) into patient values (6, 6, 183, 72) into patient values (7, 7, 192, 83) select * from dual", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "insert all into doctor values (1, 8, 'SURGEOON' , 'PUNE UNIVERSITY') into doctor values (2, 9, 'GYNOCOLOGIST', 'PUNE UNIVERSITY') select * from dual", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "insert all into staff values (1, 10,4000 , 'RECEPTIONIST') select * from dual;", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "insert all into test values (1, 'BLOOD', '12cc' , 100) into test values (2, 'CHEST XRAY', NULL , 100) into test values (3, 'URINE', '12gm' , 250) into test values (4, 'STOOL', '13cc' , 320) select * from dual", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "insert all into medical values (1, 'OPERATION', 'HEART') into medical values (2, 'OPERATION', 'BRAIN') into medical values (3, 'OTHER', 'SMOKING') into medical values (4, 'OTHER', 'DRINKING') into medical values (5, 'CANCER', 'SKIN') into medical values (6, 'CANCER', 'LUNG') select * from dual", Ado, adOpenKeyset, adLockOptimistic
        
        lblmsg.Caption = "Standard Values Inserted In the Database"
    Else
        lblmsg.Caption = "Incorrect Username and Password"
    End If

End Sub

Private Sub mnuTables_Click()
    On Error Resume Next
    
    If txtinfo(0).Text = "master" And txtinfo(1).Text = "iagree" Then
        temp.Open "create table person(per_id number(10),fname varchar(20),lname varchar(20),dob date,sex varchar(1),addr varchar(35),ph number(10),constraint pk_person primary key (per_id))", Ado, adOpenKeyset, adLockOptimistic
        
        temp.Open "create table patient(pat_id number(10),per_id number(10),height number(3),weight number(3),constraint pk_patient primary key (pat_id),constraint fk_patient foreign key (per_id) references person(per_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table doctor(doc_id number(10),per_id number(10),spec varchar(20),edu varchar(20),constraint pk_doctor primary key (doc_id),constraint fk_doctor foreign key (per_id) references person(per_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table staff(staff_id number(10),per_id number(10),pay number(10),category varchar(20),constraint pk_staff primary key (staff_id),constraint fk_staff foreign key (per_id) references person(per_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table test(test_id number(10),tname varchar(10),norm_result varchar(10),cost number(10),constraint pk_test primary key (test_id) )", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table presc(presc_id number(10),doc_id number(10),pat_id number(10), med_recommended varchar(300),constraint pk_presc primary key (presc_id),constraint fk_presc1 foreign key (doc_id) references doctor(doc_id),constraint fk_presc2 foreign key (pat_id) references patient(pat_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table appt (doc_id number(10),appt_date date, appt_time varchar(10), presc_id number(10),pat_id number(10) Not Null,constraint pk_appt primary key (doc_id, appt_date, appt_time),constraint fk_appt1 foreign key (doc_id) references doctor(doc_id), constraint fk_appt2 foreign key (presc_id) references presc(presc_id),constraint fk_appt3 foreign key (pat_id) references patient(pat_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table test_result(test_id number(10),pat_id number(10),presc_id number(10),result varchar(10),constraint pk_test_result primary key (test_id, pat_id),constraint fk_test_result1 foreign key (test_id) references test(test_id),constraint fk_test_result2 foreign key (presc_id) references presc(presc_id),constraint fk_test_result3 foreign key (pat_id) references patient(pat_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table medical (med_id number(10),category varchar(20),disease varchar(20),constraint pk_medical primary key (med_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table history (med_id number(10),pat_id number(10),detail varchar(20),constraint pk_history primary key (med_id, pat_id),constraint fk_history1 foreign key (med_id) references medical(med_id),constraint fk_history2 foreign key (pat_id) references patient(pat_id))", Ado, adOpenKeyset, adLockOptimistic
        temp.Open "create table for2mins (dname varchar (40), pname varchar(40), testname varchar(10), cost number, temp number)", Ado, adOpenKeyset, adLockOptimistic
lblmsg.Caption = "Database Created!"
    Else
        lblmsg.Caption = "Incorrect Username and Password"
    End If

End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
        txtinfo(Index).SelStart = 0
        txtinfo(Index).SelLength = Len(txtinfo(Index).Text)
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

