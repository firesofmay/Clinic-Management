Attribute VB_Name = "Module1"
Option Explicit
Public trying As Integer
Public Conn_String As String
Public Ado As ADODB.Connection
Public CheckPat, CheckTest, CheckStaff As ADOR.Recordset
Public Patient, Person, Doctor, History, Appt, Test, Presc, Medical, Test_Result, Staff, temp As ADODB.Recordset
Public i, j As Integer
        
        
     

Public Sub modLoading()
    trying = 1
    Conn_String = "Provider=MSDASQL.1;Password=pict;Persist Security Info=True;User ID=pict;Data Source=AK"
    
    Set Ado = New ADODB.Connection
    Ado.ConnectionString = Conn_String
    Ado.Open
       
    

End Sub

