VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Server"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   3660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " [ Info Server ] "
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.Label lblHostName 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblIP 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblStatus 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   3135
      End
   End
   Begin VB.ListBox lstStatusKoneksi 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   1935
      Width           =   3375
   End
   Begin MSWinsockLib.Winsock Socket 
      Index           =   0
      Left            =   4320
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Status Koneksi"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Option Explicit

Private Const LOCAL_PORT As Long = 1007

Private Const REC_SPR As String * 1 = "|" 'separator baris
Private Const FLD_SPR As String * 1 = "#" 'separator kolom
Private Const MAX_LIMIT As Long = 1024 '1x kirim dibatasi 1 kb, kalo untuk jaringan lokal masih bisa set 4096

Private Function pembulatanKeAtas(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double
    Dim temp As Double
    
    temp = Int(X * Factor)
    pembulatanKeAtas = (temp + IIf(X = temp, 0, 1)) / Factor
End Function

Private Function getDataBarang(ByVal param As String) As String()
    Dim rs          As ADODB.Recordset
    
    Dim div         As Long
    Dim lengthData  As Long
    Dim n           As Long
    Dim i           As Long
    
    Dim tmp         As String
    Dim arrTmp()    As String
    
    strSql = "SELECT UCASE(nama), harga, stok FROM barang " & param & ""
    Set rs = openRecordset(strSql)
    If Not rs.EOF Then
        For i = 1 To getRecordCount(rs)
            tmp = tmp & rs(0).Value & FLD_SPR & rs(1).Value & FLD_SPR & rs(2).Value & REC_SPR
            
            rs.MoveNext
        Next i
        If Len(tmp) > 0 Then tmp = Left(tmp, Len(tmp) - 1)
        
        'karakter ~ sebagai penanda awal dan akhir data
        'untuk memudahkan pengecekan di klien bahwa data yg diterima sudah lengkap/belum
        'ex : ~DATA BARANG + SEPARATOR KOLOM DAN BARIS~
                        
        'contoh format data disini ada 2 :
        '1. jika data <= 1024 karakter : ~~DATA BARANG + SEPARATOR KOLOM DAN BARIS
        '2. jika data > 1024 karakter  : ~DATA BARANG + SEPARATOR KOLOM DAN BARIS~
        
        If Len(tmp) > 0 Then tmp = "~" & Left(tmp, Len(tmp) - 1) & "~"
        If Not Len(tmp) > MAX_LIMIT Then
            tmp = Left(tmp, Len(tmp) - 1)
            tmp = "~" & tmp
        End If
        
        lengthData = Len(tmp)
        If lengthData > 0 Then
            If lengthData > MAX_LIMIT Then 'data > 1024 karakter
                'data dibuat menjadi beberapa package
                'ex : jika jumlah karakter 2345
                '     package 1 -> 1024
                '     package 2 -> 1024
                '     package 3 -> 297
                '     berarti data yg dikirim ke klien sebanyak 3 x
                
                div = pembulatanKeAtas(lengthData / MAX_LIMIT)
                ReDim arrTmp(div)
                
                n = 1
                For i = 1 To div
                    arrTmp(i - 1) = Mid(tmp, n, MAX_LIMIT)
                    n = n + MAX_LIMIT
                Next i
                
            Else
                ReDim arrTmp(0)
                arrTmp(0) = tmp
            End If
            
        Else
            ReDim arrTmp(0)
            arrTmp(0) = tmp
        End If
        
    Else
        ReDim arrTmp(0)
        arrTmp(0) = "EOF" 'data barang tidak ditemukan
    End If
    Call closeRecordset(rs)
    
    getDataBarang = arrTmp
End Function

Private Function send(ByVal lngIndex As Long, ByVal strData As String) As Boolean
    If Socket(lngIndex).State = sckConnected Then
        Call Socket(lngIndex).SendData(strData)
        DoEvents
        
    Else
        send = False
        Exit Function
    End If
   
    send = True
End Function

Private Function startListening(ByVal localPort As Long) As Boolean
    'On Error GoTo errHandle
    
    If localPort > 0 Then
        'If the socket is already listening, and it's listening on the same port, don't bother restarting it.
        If (Socket(0).State <> sckListening) Or (Socket(0).localPort <> localPort) Then
            With Socket(0)
                .Close
                .localPort = localPort
                .Listen
            End With
        End If
        
        'Return true, since the server is now listening for clients.
        startListening = True
   End If
   
   Exit Function
errHandle:
   startListening = False
End Function

Private Sub startServer()
    If startListening(LOCAL_PORT) Then
        lblStatus.Caption = "Status Listening : ON"
    Else
        lblStatus.Caption = "Status Listening : OFF"
    End If
End Sub

Private Sub shutDown()
    Dim i    As Long
    
    If Socket(0).State <> sckClosed Then Socket(0).Close
   
    ' Now loop through all the clients, close the active ones and
    ' unload them all to clear the array from memory.
    For i = 1 To Socket.UBound
        If Socket(i).State <> sckClosed Then Socket(i).Close
        Call Unload(Socket(i))
    Next i
End Sub

Private Sub Form_Load()
    lblIP.Caption = "IP : " & Socket(0).LocalIP
    lblHostName.Caption = "Host Name : " & Socket(0).LocalHostName
    
    Call startServer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call shutDown
End Sub

Private Sub Socket_Close(Index As Integer)
    ' Close the socket and raise the event to the parent.
    Call Socket(Index).Close
    lstStatusKoneksi.List(Index - 1) = Socket(Index).RemoteHostIP & " on port " & Socket(Index).RemotePort & " [disconnected]"
End Sub

Private Sub Socket_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    Dim i           As Long
    Dim j           As Long
    Dim newWinsock  As Boolean
    
    'On Error GoTo errHandle
    
    ' We shouldn't get ConnectionRequests on any other socket than the listener
    ' (index 0), but check anyway. Also check that we're not going to exceed
    ' the MaxClients property.
    If (Index = 0) Then
        ' Check to see if we've got any sockets that are free.
        For i = 1 To Socket.UBound
            If Socket(i).State = sckClosed Or Socket(i).State = sckClosing Then
                j = i
                Exit For
            End If
        Next i
      
        ' If we don't have any free sockets, load another on the array.
        If (j = 0) Then
            Call Load(Socket(Socket.UBound + 1))
            j = Socket.Count - 1
            newWinsock = True
        End If
        
        ' With the selected socket, reset it and accept the new connection.
        With Socket(j)
            Call .Close
            Call .Accept(requestID)
        End With
        
        If newWinsock Then
            lstStatusKoneksi.AddItem Socket(j).RemoteHostIP & " on port " & Socket(j).RemotePort & " [connected]"
        Else
            lstStatusKoneksi.List(j - 1) = Socket(j).RemoteHostIP & " on port " & Socket(j).RemotePort & " [connected]"
        End If
    End If
    
    Exit Sub
    '
errHandle:
    ' Close the Winsock that caused the error.
    Call Socket(0).Close
End Sub

Private Sub Socket_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim i           As Long
    Dim strData     As String
    Dim ret         As Boolean
    
    Dim arrTmp()    As String
    
    'On Error GoTo errHandle
    
    ' Grab the data from the specified Winsock object, and pass it to the parent.
    Call Socket(Index).GetData(strData)
    DoEvents
    
    arrTmp = getDataBarang(strData)
    For i = LBound(arrTmp) To UBound(arrTmp)
        If Len(arrTmp(i)) > 0 Then
            ret = send(Index, arrTmp(i))
        End If
    Next i
    
    Exit Sub
errHandle:
   Call Socket(Index).Close
End Sub

Private Sub Socket_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call Socket(Index).Close
End Sub

