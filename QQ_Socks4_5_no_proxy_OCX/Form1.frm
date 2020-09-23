VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Socks Example"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Use Socks5 Proxy"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GET>>"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1800
      TabIndex        =   7
      Text            =   "1080"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "221.204.253.154"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Text            =   "80"
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Text            =   "www.google.com"
      Top             =   240
      Width           =   2055
   End
   Begin Project1.QSocks QSocks1 
      Left            =   120
      Top             =   1200
      _extentx        =   529
      _extenty        =   529
      socksversion    =   0
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Socks Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Socks Server:"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Port:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Destination Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'QSocks by: QQ
'h0_plz@hotmail.com

'if you get a connection failure its most likely due to
'an inactive proxy server or that the server is being over worked.
'So if you get this either try to connect afew more times or use a
'different proxy because the one in the example is now inactive.
'
'http://aliveproxy.com/socks5-list/

Option Explicit

Private Sub Command1_Click()
With QSocks1
  .CloseSocket
    If Check1.Value = 0 Then
      .SocksVersion = 0 'no proxy
    Else
      .SocksVersion = 5 'socks5
    End If
  .SocketType = 1 'tcp connect
      Call .Connect(Text1.Text, Text2.Text, Text3.Text, Text4.Text)
End With
End Sub

Private Sub QSocks1_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
  Call QSocks1.GetData(Buffer)
  Debug.Print Buffer
    If InStr(1, Buffer, Chr(&H30) & Chr(&HD) & Chr(&HA) & Chr(&HD) & Chr(&HA)) Then QSocks1.CloseSocket
End Sub

Private Sub QSocks1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print Description
End Sub

Private Sub QSocks1_SockClose()
Debug.Print "Socket Closed"
End Sub

Private Sub QSocks1_SockConnect()
Call QSocks1.SendData(GetUrl(Text1.Text))
End Sub

Function GetUrl(Url As String) As String
Dim Pck As String
  Pck = "GET / HTTP/1.1" & vbCrLf & _
        "Accept: */*" & vbCrLf & _
        "Accept-Language: en-us" & vbCrLf & _
        "User-Agent: Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729)" & vbCrLf & _
        "Host: " & Url & vbCrLf & _
        "Connection: Keep-Alive" & vbCrLf & vbCrLf
  GetUrl = Pck
End Function

