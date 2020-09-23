Attribute VB_Name = "modConnection"
Public Con As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RsNAV As New ADODB.Recordset
Public InsertRecord As New ADODB.Recordset
Public UpdateRecord As New ADODB.Recordset
Public RsMisc As New ADODB.Recordset

Public Sub dbConnection()
    Set Con = New ADODB.Connection
    Con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database\IMSDB.mdb;Persist Security Info=False"
    Con.Open
End Sub

Public Sub Main()
    Call dbConnection
    frmLogin.Show
''    Load frmStart
''    frmStart.Show
End Sub




