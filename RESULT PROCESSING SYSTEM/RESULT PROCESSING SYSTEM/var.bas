Attribute VB_Name = "var"
Public con, con1, scon, PathOb As Object
Public reg, sem, dept, path As String
Public Sub openconnection()
Set con = CreateObject("adodb.Connection")
con.Open = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & var.path
End Sub
Public Sub openconnection1()
Set con1 = CreateObject("adodb.Connection")
con1.Open = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & "C:\db\SUBJECT.mdb "
End Sub
Public Sub openconnectionOLD()
Set con1 = CreateObject("adodb.Connection")
con1.Open = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & "C:\db\SUBJECT OLD.mdb "
End Sub

Public Sub pathsub()
    Set PathOb = CreateObject("adodb.Connection")
    PathOb.Open = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & "C:\db\path.mdb"
End Sub
Public Sub sopenconnection()
Set scon = CreateObject("adodb.Connection")
scon.Open = "PROVIDER=MICROSOFT.JET.OLEDB.4.0;DATA SOURCE=" & "C:\db\path.mdb"
End Sub
