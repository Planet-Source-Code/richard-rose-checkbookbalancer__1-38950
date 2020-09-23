Attribute VB_Name = "ODBCConnect"
Public db As ADODB.Connection
Public dbr1 As ADODB.Recordset
Public connString As String

Public Sub Connect()
connString = "DSN=Checkbook;"
Set db = New ADODB.Connection
db.Open connString
Set dbr1 = New ADODB.Recordset
dbr1.Open "Transaction", db, adOpenKeyset, adLockOptimistic
End Sub

