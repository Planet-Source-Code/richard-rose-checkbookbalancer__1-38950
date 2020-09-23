Attribute VB_Name = "Connections"
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
                (ByVal hwndParent As Long, ByVal fRequest As Long, _
                ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

Public db As ADODB.Connection
Public dbr1 As ADODB.Recordset
Public connString As String

Public Sub Connect()
Call Create
connString = "DSN=Checkbook;"
Set db = New ADODB.Connection
db.Open connString
Set dbr1 = New ADODB.Recordset
dbr1.Open "Transaction", db, adOpenKeyset, adLockOptimistic
End Sub

Private Function CreateAccessDSN(DSNName As String, DatabaseFullPath As String) As Boolean
Dim sAttributes As String
    sAttributes = "DSN=" & DSNName & Chr(0)
    sAttributes = sAttributes & "DBQ=" & DatabaseFullPath & Chr(0)
    CreateAccessDSN = CreateDSN("Microsoft Access Driver (*.mdb)", sAttributes)
End Function

Private Function CreateDSN(Driver As String, Attributes As String) As Boolean
    CreateDSN = SQLConfigDataSource(0&, 1, Driver, Attributes)
End Function

Private Sub Create()
Dim blnRetVal As Boolean
        blnRetVal = CreateAccessDSN("Checkbook", App.Path & "\Checkbook.mdb")
End Sub



