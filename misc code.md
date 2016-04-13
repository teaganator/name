Sub xdfgrth()
    Dim val As String, x As Integer, y As Integer, z As Integer, i As Integer, j As Integer, k As Integer
    x = 2
    y = 18
    z = 17
    i = 1710
    j = 7
    k = 6
    Do While x < 1705
        val = Cells(x, y)
        Do While i < 2278
            If InStr(1, Cells(i, j), val) > 0 Then
                Cells(x, z) = Cells(i, k)
                i = 2278
            End If
            i = i + 1
        Loop
        If Cells(x, y) = "" Then
            Cells(x, z) = ""
        End If
        If Cells(x, y) = "" Then
            Cells(x, z) = ""
        End If
        If InStr(1, Cells(x, 7), Cells(x, 4)) > 0 Then
            Cells(x, 4) = Replace(Cells(x, 4), Cells(x, 7), "")
        End If
        If InStr(1, Cells(x, 8), Cells(x, 4)) > 0 Then
            Cells(x, 4) = Replace(Cells(x, 4), Cells(x, 8), "")
        End If
        x = x + 1
        i = 1710
    Loop
End Sub


Sub xdfgrth()
    Dim val As String, x As Integer, y As Integer, z As Integer, i As Integer
    i = 1
    x = 2
    y = 17
    z = 16
    Do While x < 10000
        If Cells(x, z) <> "" Then
            Cells(x + 1, z).EntireRow.Insert
            Do While i < 8
                Cells(x + 1, i) = Cells(x, i)
                i = i + 1
            Loop
            Cells(x + 1, i) = Cells(x, z)
            Cells(x + 1, i + 1) = Cells(x, y)
            i = 1
        End If
        x = x + 1
    Loop
End Sub		


SQL stuff that i might use later
INSERT Range("E10"),Range("E11"),Range("E12"),Range("E13"),Range("E14") INTO P/N's,Description,Package,MFG PN, MFG;

Private Sub Description_AfterGotFocus()
Dim objRec
Dim objConn
Dim cmdString

Set objRec = CreateObject("ADODB.Recordset")
Set objConn = CreateObject("ADODB.Connection")

objConn.ConnectionString = "Provider=MSDASQL;DSN=GreatPlains;Initial Catalog=TWO;User Id=sa;Password=password"
objConn.Open


cmdString = "Select ACTINDX from GL00105 where (ACTNUMST='" + Account + "')"
 
Set objRec = objConn.Execute(cmdString)

If objRec.EOF = True Then
AccountMaintenance.UserDefined1 = ""
Else
AccountMaintenance.UserDefined1 = objRec!ACTINDX
End If
objConn.Close
End Sub


conString = "Provider='manexsql';Data Source='Test_server';" & "Initial Catalor='P/Ns_Test';"
