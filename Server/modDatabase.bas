Attribute VB_Name = "modDatabase"
Option Explicit

Public conn     As ADODB.Connection
Public strSql   As String

Public Function konekToServer() As Boolean
    Dim strCon As String
    
    On Error GoTo errHandle
    
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\sampleDB.mdb"
    
    Set conn = New ADODB.Connection
    conn.ConnectionString = strCon
    conn.Open
    
    konekToServer = True
    
    Exit Function
errHandle:
    konekToServer = False
End Function

Public Sub closeRecordset(ByVal vRs As ADODB.Recordset)
    On Error Resume Next
    
    If Not (vRs Is Nothing) Then
        If vRs.State = adStateOpen Then
            vRs.Close
            Set vRs = Nothing
        End If
    End If
End Sub

Public Function dbGetValue(ByVal query As String, ByVal defValue As Variant) As Variant
    Dim rsDbGetValue  As ADODB.Recordset
    
    On Error GoTo errHandle
    
    Set rsDbGetValue = New ADODB.Recordset
    rsDbGetValue.Open query, conn, adOpenForwardOnly, adLockReadOnly
    If Not rsDbGetValue.EOF Then
        If Not IsNull(rsDbGetValue(0).Value) Then
            dbGetValue = rsDbGetValue(0).Value
        Else
            dbGetValue = defValue
        End If
    Else
        dbGetValue = defValue
    End If
        
    Call closeRecordset(rsDbGetValue)
    
    Exit Function
errHandle:
    dbGetValue = defValue
End Function

Public Function openRecordset(ByVal query As String) As ADODB.Recordset
    Dim obj As ADODB.Recordset
    
    Set obj = New ADODB.Recordset
    obj.CursorLocation = adUseClient
    obj.Open query, conn, adOpenForwardOnly, adLockReadOnly
    Set openRecordset = obj
End Function

Public Function getRecordCount(ByVal vRs As ADODB.Recordset) As Long
    On Error Resume Next
    
    vRs.MoveLast
    getRecordCount = vRs.RecordCount
    vRs.MoveFirst
End Function
