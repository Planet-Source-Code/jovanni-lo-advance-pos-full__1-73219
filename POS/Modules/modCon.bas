Attribute VB_Name = "modCon"
Public Con As New ADODB.Connection
Public Rs As New ADODB.Recordset
Public SubRs As New ADODB.Recordset
Public StrCon As String

Public Sub OpenCon()
'open connection
Set Con = New ADODB.Connection
StrCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
        & App.Path & "\Database\POS.mdb;Persist Security Info=False"
Con.Open StrCon
dtaGroups.conDb.ConnectionString = StrCon
End Sub

Public Sub RunSql(Statement As String)
Set Rs = New ADODB.Recordset
Rs.Open Statement, Con, adOpenKeyset, adLockPessimistic
End Sub

Public Sub SubSql(Statement As String)
Set SubRs = New ADODB.Recordset
SubRs.Open Statement, Con, adOpenKeyset, adLockPessimistic
End Sub

Public Sub CloseRs(dtaRs As Variant)
If dtaRs.State = adStateOpen Then
    dtaRs.Close
End If
End Sub

Public Function CompactDB(pFileName As String) As Boolean
On Error GoTo ErrH
Dim CONN As New JRO.JetEngine
Dim ConnstringSorg As String, ConnstringDest As String

' Ensure file is not read only
SetAttr pFileName, vbNormal
ConnstringSorg = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
pFileName & ";User ID=;Password=;"
ConnstringDest = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
App.Path & "\Temp.mdb" & ";Jet OLEDB:Engine Type=5;"

Screen.MousePointer = vbHourglass
CONN.CompactDatabase ConnstringSorg, ConnstringDest
Screen.MousePointer = vbDefault

'Copy compacted file
Kill pFileName
FileCopy App.Path & "\Temp.mdb", pFileName
Kill App.Path & "\Temp.mdb"

Set CONN = Nothing
CompactDB = True
Exit Function
ErrH:
Screen.MousePointer = vbDefault
Debug.Print Err.Description
End Function



