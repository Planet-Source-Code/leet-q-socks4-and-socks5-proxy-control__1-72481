VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.UserControl QSocks 
   BackStyle       =   0  'Transparent
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   450
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   375
   ScaleWidth      =   450
   ToolboxBitmap   =   "QSocks.ctx":0000
   Begin MSWinsockLib.Winsock Socket 
      Left            =   480
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image IMGSocks 
      Height          =   300
      Left            =   0
      Picture         =   "QSocks.ctx":0312
      Stretch         =   -1  'True
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "QSocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'QSocks by: QQ
'h0_plz@hotmail.com


'About: This is a ocx I started to allow a user to use the winsock control
'       with socks4 and 5  proxy, as well as without. You should get
'       the basic idea from looking it over, so I won't comment it a whole
'       lot just for specific things I feel are necessary. I haven't tested
'       really anything, in theory as long as I didn't mess something up due
'       to lack of sleep it should all work. If something is off and is not
'       working or you need assistance just ask. If you like it please vote.


Option Explicit

Private Local_IP As String
Private Local_Port As Long
Private Socks_Server As String
Private Socks_Port As Long
Private Dest_Server As String
Private Dest_Port As Long
Private SocketData As String
Private Socks_Version As Long '4 or 5
Private Socket_Type As Long '1 = TCP connect, 2 = TCP listen, 3 = UDP

Private Connected As Boolean
Private Assigned As Boolean
Private iSiT_host As Boolean

Public Event DataArrival(ByVal bytesTotal As Long)
Public Event Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Public Event SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
Public Event SockConnect()
Public Event SockClose()
Public Event SendComplete()
Public Event ConnectionRequest(ByVal RequestId As Long)

Public Property Get SocketType() As Long
  SocketType = Socket_Type
End Property

Public Property Let SocketType(ByVal SocketType_New As Long)
Socket.Close
  If SocketType_New = 1 Or SocketType_New = 2 Or SocketType_New = 3 Then
    SocketType_New = SocketType_New
  Else
    SocketType_New = 1
  End If
    Let Socket_Type = SocketType_New
      PropertyChanged "SocketType"
        Select Case Socket_Type
          Case 1
            Socket.Protocol = sckTCPProtocol
          Case 2
            Socket.Protocol = sckTCPProtocol
          Case 3
            Socket.Protocol = sckUDPProtocol
      End Select
End Property

Public Property Get SocksVersion() As Long
  SocksVersion = Socks_Version
End Property

Public Property Let SocksVersion(ByVal SocksVersion_New As Long)
If SocksVersion_New = 5 Or SocksVersion_New = 4 Or SocksVersion_New = 0 Then
  SocksVersion_New = SocksVersion_New
Else
  SocksVersion_New = 0
End If
  Let Socks_Version = SocksVersion_New
    PropertyChanged "SocksVersion"
End Property

Public Property Get LocalIP() As String
  LocalIP = Local_IP
End Property

Public Property Let LocalIP(ByVal IP_New As String)
  Let Local_IP = IP_New
    PropertyChanged "LocalIP"
End Property

Public Property Get LocalPort() As Long
  LocalPort = Local_Port
End Property

Public Property Let LocalPort(ByVal Port_New As Long)
  Let Local_Port = Port_New
    PropertyChanged "LocalPort"
End Property

Public Property Get SocksServer() As String
  SocksServer = Socks_Server
End Property

Public Property Let SocksServer(ByVal Server_New As String)
  Let Socks_Server = Server_New
    PropertyChanged "SocksServer"
End Property

Public Property Get SocksPort() As Long
  SocksPort = Socks_Port
End Property

Public Property Let SocksPort(ByVal Port_New As Long)
  Let Socks_Port = Port_New
    PropertyChanged "SocksPort"
End Property

Public Property Get DestServer() As String
  DestServer = Dest_Server
End Property

Public Property Let DestServer(ByVal Server_New As String)
  Let Dest_Server = Server_New
    PropertyChanged "DestServer"
End Property

Public Property Get DestPort() As Long
  DestPort = Dest_Port
End Property

Public Property Let DestPort(ByVal Port_New As Long)
  Let Dest_Port = Port_New
    PropertyChanged "DestPort"
End Property

Private Sub Socket_Close()
  RaiseEvent SockClose
End Sub

Private Sub Socket_Connect()
On Error Resume Next
  Select Case Socks_Version
    Case 0 'no proxy
      RaiseEvent SockConnect
    Case 4 'socks 4 proxy
      Call Socket.SendData(Socks4_Con(Dest_Server, Dest_Port))
    Case 5 'socks 5 proxy
      Call Socket.SendData(Socks5_Con1)
  End Select
End Sub

Private Sub Socket_ConnectionRequest(ByVal RequestId As Long)
  RaiseEvent ConnectionRequest(RequestId)
End Sub

Private Sub Socket_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Buffer As String
  Socket.GetData Buffer
    Call Handle(Buffer)
End Sub

Private Sub Socket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  RaiseEvent Error(Number, Description, Scode, Source, HelpFile, HelpContext, CancelDisplay)
End Sub

Private Sub Socket_SendComplete()
  RaiseEvent SendComplete
End Sub

Private Sub Socket_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
  RaiseEvent SendProgress(bytesSent, bytesRemaining)
End Sub

Private Sub UserControl_InitProperties()
  Let SocksVersion = 5
  Let SocketType = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Let Socks_Server = .ReadProperty("SocksServer", "")
    Let Socks_Port = .ReadProperty("SocksPort", 0)
    Let Dest_Server = .ReadProperty("DestServer", "")
    Let Dest_Port = .ReadProperty("DestPort", 0)
    Let Socks_Version = .ReadProperty("SocksVersion", 5)
    Let Socket_Type = .ReadProperty("SocketType", 1)
    Let Local_IP = .ReadProperty("LocalIP", "")
    Let Local_Port = .ReadProperty("LocalPort", 0)
  End With
End Sub

Private Sub UserControl_Resize()
  If UserControl.Width <> 300 Then UserControl.Width = 300
  If UserControl.Height <> 300 Then UserControl.Height = 300
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    Call .WriteProperty("SocksServer", Socks_Server, "")
    Call .WriteProperty("SocksPort", Socks_Port, 0)
    Call .WriteProperty("DestServer", Dest_Server, "")
    Call .WriteProperty("DestPort", Dest_Port, 0)
    Call .WriteProperty("SocksVersion", Socks_Version, 5)
    Call .WriteProperty("SocketType", Socket_Type, 1)
    Call .WriteProperty("LocalIP", Local_IP, "")
    Call .WriteProperty("LocalPort", Local_Port, 0)
  End With
End Sub

Public Sub Connect(Destination_Address As String, Destination_Port As Integer, Optional Socks_Server As String, Optional Socks_Port As Integer)
On Error Resume Next
Connected = False
Assigned = False
iSiT_host = Return_IP_Type(Destination_Address)
  Let DestServer = Destination_Address
  Let DestPort = Destination_Port
    Socket.Close
      If Socks_Version = 4 And Socket_Type <> 1 Then
        Socket_Type = 1 'socks4 don't support a non tcp socket transaction
          Let SocketType = Socket_Type
      End If
      If Socks_Server = Empty Or Socks_Port = Empty Or Socks_Version = 0 Then 'if no proxy secified and/or version set to 0 then directly connect
        Socks_Version = 0
          Call Socket.Connect(Destination_Address, Destination_Port)
      Else
          Call Socket.Connect(Socks_Server, Socks_Port)
      End If
End Sub

Public Sub AcceptConnection(ByVal RequestId As Long)
  Socket.Accept RequestId
End Sub

Public Sub Listen()
  Socket.Listen
End Sub

Public Sub Bind()
  Socket.Bind
End Sub

Public Sub CloseSocket()
  Socket.Close
End Sub

Public Sub SendData(Data)
  Call Socket.SendData(Data)
End Sub

Private Sub Handle(Data As String)
On Error Resume Next
Select Case Socks_Version
  Case 0 'no proxy
    SocketData = Data
      RaiseEvent DataArrival(Len(Data))
  Case 5 'socks5
    If Connected = False Then
      Select Case Mid(Data, 2, 1)
        Case Chr(0)
          Connected = True
            Call Socket.SendData(Socks5_Con2(Dest_Server, Dest_Port, iSiT_host))
      End Select
    ElseIf Connected = True Then
      If Assigned = False Then
        Assigned = True
          RaiseEvent SockConnect
      ElseIf Assigned = True Then
        SocketData = Data
          RaiseEvent DataArrival(Len(Data))
      End If
    End If
Case 4 'socks4
      If Assigned = False Then
        Assigned = True
          RaiseEvent SockConnect
      ElseIf Assigned = True Then
        SocketData = Data
          RaiseEvent DataArrival(Len(Data))
      End If
End Select
End Sub

Public Sub GetData(BufferString As String)
BufferString = SocketData
End Sub

Public Function State() As Long
State = Socket.State
End Function

Private Function SocketStates(SockMessage As Long) As String
Select Case SockMessage
  Case 10048
    SocketStates = "AddressInUse"
  Case 10049
    SocketStates = "AddressNotAvailable"
  Case 10037
    SocketStates = "AlreadyComplete"
  Case 10056
    SocketStates = "AlreadyConnected"
  Case 10053
    SocketStates = "ConnectAborted"
  Case 8
    SocketStates = "Closing"
  Case 0
    SocketStates = "Closed"
  Case 6
    SocketStates = "Connecting"
  Case 7
    SocketStates = "Connected"
  Case 3
    SocketStates = "ConnectionPending"
  Case 10061
    SocketStates = "ConnectionRefused"
  Case 10054
    SocketStates = "ConnectionReset"
  Case 9
    SocketStates = "Error"
  Case 394
    SocketStates = "GetNotSupported"
  Case 11001
    SocketStates = "HostNotFound"
  Case 11002
    SocketStates = "HostNotFoundTryAgain"
  Case 5
    SocketStates = "HostResolved"
  Case 10036
    SocketStates = "InProgress"
  Case 40014
    SocketStates = "InvalidArg"
  Case 10014
    SocketStates = "InvalidArgument"
  Case 40020
    SocketStates = "InvalidOp"
  Case 380
    SocketStates = "InvalidPropertyValue"
  Case 2
    SocketStates = "Listening"
  Case 10040
    SocketStates = "MsgTooBig"
  Case 10052
    SocketStates = "NetReset"
  Case 10050
    SocketStates = "NetworkSubsystemFailed"
  Case 10051
    SocketStates = "NetworkUnreachable"
  Case 10055
    SocketStates = "NoBufferSpace"
  Case 11004
    SocketStates = "NoData"
  Case 11003
    SocketStates = "NonRecoverableError"
  Case 10057
    SocketStates = "NotConnected"
  Case 10093
    SocketStates = "NotInitialized"
  Case 10038
    SocketStates = "NotSocket"
  Case 10004
    SocketStates = "OpCanceled"
  Case 1
    SocketStates = "Open"
  Case 40021
    SocketStates = "OutOfRange"
  Case 10043
    SocketStates = "PortNotSupported"
  Case 4
    SocketStates = "ResolvingHost"
  Case 383
    SocketStates = "SetNotSupported"
  Case 10058
    SocketStates = "SocketShutdown"
  Case 40017
    SocketStates = "Success"
  Case 10060
    SocketStates = "Timedout"
  Case 40018
    SocketStates = "Unsupported"
  Case 10035
    SocketStates = "WouldBlock"
  Case 40026
    SocketStates = "WrongProtocol"
End Select
End Function

Private Function Socks5_Con1() As Byte()
Dim Header(2) As Byte
  'header consists of the version, Protocol, reserve type
  Header(0) = 5
  Header(1) = 1
  Header(2) = 0
  Socks5_Con1 = Header
End Function

Private Function Socks5_Con2(Dest_Address As String, DestPort As Long, HostIp As Boolean) As String
On Error Resume Next
If HostIp = True Then 'chr(3) = host
  Socks5_Con2 = Chr(5) & Chr(Socket_Type) & Chr(0) & Chr(3) & Chr(Len(Dest_Address)) & Dest_Address & Chr(Int(DestPort / 256)) & Chr(DestPort Mod 256)
Else 'chr(1) = ip
  Socks5_Con2 = Chr(5) & Chr(Socket_Type) & Chr(0) & Chr(1) & Chr(Split(Dest_Address, ".")(0)) & Chr(Split(Dest_Address, ".")(1)) & Chr(Split(Dest_Address, ".")(2)) & Chr(Split(Dest_Address, ".")(3)) & Chr(Int(DestPort / 256)) & Chr(DestPort Mod 256)
End If
End Function

Private Function Socks4_Con(Dest_Address As String, DestPort As Long) As String
On Error Resume Next
If HostIp = True Then 'chr(3) = host
  Socks4_Con = Chr(4) & Chr(Socket_Type) & Chr(Int(DestPort / 256)) & Chr(DestPort Mod 256) & String(3, 0) & Chr(1) & Chr(0) & Dest_Address
Else 'chr(1) = ip
  Socks4_Con = Chr(4) & Chr(Socket_Type) & Chr(Int(DestPort / 256)) & Chr(DestPort Mod 256) & Chr(Split(Dest_Address, ".")(0)) & Chr(Split(Dest_Address, ".")(1)) & Chr(Split(Dest_Address, ".")(2)) & Chr(Split(Dest_Address, ".")(3)) & Chr(0)
End If
End Function

Private Function Return_IP_Type(strIp As String) As Boolean
On Error Resume Next
Dim Prts() As String
Dim X As Integer
  Prts = Split(strIp, ".")
    For X = 0 To UBound(Prts)
      If IsNumeric(Prts(X)) = False Then GoTo IsNot
    Next X
  Return_IP_Type = False
Exit Function
IsNot:
  Return_IP_Type = True
End Function
