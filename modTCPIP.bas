Attribute VB_Name = "modTCPIP"
'***************************************************************************************
'*                                 VB TCPIP Functions                                  *
'*                               Last Updated 26/03/2002                               *
'*                                                                                     *
'*  lngNewAddress   -  Converts an IP address from a string to a Long Int              *
'*  checkIP         -  Checks the validity of the format of an IP address              *
'*  Cleanup         -  Cleans up the API Winsock after use                             *
'*  WSInitialise    -  Initialises API Winsock                                         *
'*  Ping            -  Issues an ICMP Ping using API Winsock                           *
'*                                                                                     *
'*       If you have any problems or queries: martin.sidgreaves@wrigley.co.uk          *
'***************************************************************************************

Option Explicit

'Declare API functions
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
Private Declare Function IcmpSendEcho Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, ByVal RequestOptions As Long, ReplyBuffer As ICMP_ECHO_REPLY, ByVal ReplySize As Long, ByVal Timeout As Long) As Long
Private Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Private Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired As Long, lpWSADATA As WSADATA) As Long
Private Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal szHost As String, dwHostLen As Long) As Long
Private Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal szHost As String) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'Define Types
Private Type ICMP_OPTIONS
    Ttl As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Byte
    OptionsData As Long
End Type

Private Type ICMP_ECHO_REPLY
    Address As Long
    status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    DataPointer As Long
    Options As ICMP_OPTIONS
    Data As String * 250
End Type

Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type

Private Const MAX_WSADescription = 256
Private Const MAX_WSASYSStatus = 128

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type

'Define Constants
Private Const IP_STATUS_BASE = 11000
Private Const IP_SUCCESS = 0
Private Const IP_BUF_TOO_SMALL = (11000 + 1)
Private Const IP_DEST_NET_UNREACHABLE = (11000 + 2)
Private Const IP_DEST_HOST_UNREACHABLE = (11000 + 3)
Private Const IP_DEST_PROT_UNREACHABLE = (11000 + 4)
Private Const IP_DEST_PORT_UNREACHABLE = (11000 + 5)
Private Const IP_NO_RESOURCES = (11000 + 6)
Private Const IP_BAD_OPTION = (11000 + 7)
Private Const IP_HW_ERROR = (11000 + 8)
Private Const IP_PACKET_TOO_BIG = (11000 + 9)
Private Const IP_REQ_TIMED_OUT = (11000 + 10)
Private Const IP_BAD_REQ = (11000 + 11)
Private Const IP_BAD_ROUTE = (11000 + 12)
Private Const IP_TTL_EXPIRED_TRANSIT = (11000 + 13)
Private Const IP_TTL_EXPIRED_REASSEM = (11000 + 14)
Private Const IP_PARAM_PROBLEM = (11000 + 15)
Private Const IP_SOURCE_QUENCH = (11000 + 16)
Private Const IP_OPTION_TOO_BIG = (11000 + 17)
Private Const IP_BAD_DESTINATION = (11000 + 18)
Private Const IP_ADDR_DELETED = (11000 + 19)
Private Const IP_SPEC_MTU_CHANGE = (11000 + 20)
Private Const IP_MTU_CHANGE = (11000 + 21)
Private Const IP_UNLOAD = (11000 + 22)
Private Const IP_ADDR_ADDED = (11000 + 23)
Private Const IP_GENERAL_FAILURE = (11000 + 50)
Private Const MAX_IP_STATUS = 11000 + 50
Private Const IP_PENDING = (11000 + 255)
Private Const PING_TIMEOUT = 200
Private Const WS_VERSION_REQD = &H101
Private Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Private Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Private Const MIN_SOCKETS_REQD = 1
Private Const SOCKET_ERROR = -1
Private Const ERROR_SUCCESS  As Long = 0


'=========================================================================
'Function lngNewAddress(strAdd As String) As Long
'Converts the IP address from a string to a Long Integer

'Inputs:    strAdd  -  The IP address as a string

'Returns:   checkIP as boolean
'           True or false depending on validity
'=========================================================================
Function lngNewAddress(strAdd As String) As Long

  Dim strTemp As String, intCount As Integer, strOctet(1 To 4) As String

    strTemp = strAdd
    intCount = 0

    'Break up the string into Octets
    While InStr(strTemp, ".") > 0
        intCount = intCount + 1
        strOctet(intCount) = Mid$(strTemp, 1, InStr(strTemp, ".") - 1)
        strTemp = Mid$(strTemp, InStr(strTemp, ".") + 1)
    Wend

    intCount = intCount + 1
    strOctet(intCount) = strTemp
    
    'Make sure there are a valid number of Octets
    If intCount <> 4 Then
        lngNewAddress = 0
        Exit Function
    End If

    'Create the Long Int by 'Val'ing the Hex values of each Octet
    lngNewAddress = Val("&H" & Right$("00" & Hex$(strOctet(4)), 2) & Right$("00" & Hex$(strOctet(3)), 2) & Right$("00" & Hex$(strOctet(2)), 2) & Right$("00" & Hex$(strOctet(1)), 2))

End Function


'=========================================================================
'Function checkIP(ipnum As String) As boolean
'Checks the validity and range of an IP address

'Inputs:    ipnum - The IP number

'Returns:   checkIP as boolean
'           True or false depending on validity
'=========================================================================
Public Function checkIP(ipnum As String) As Boolean

  'This function checks the validity of an IP number
  
  Dim strIPstring
  Dim intScount As Integer
  Dim strTmp As String
  Dim intIPtag As Integer
  Dim bolTest As Boolean

    'Initialise the variables
    strIPstring = ipnum
    intIPtag = 0

    'First check for the correct number of octets
    Do
        intScount = InStr(1, strIPstring, ".")
        If intScount > 0 Then
            intIPtag = intIPtag + 1
            strIPstring = Mid$(strIPstring, intScount + 1)
        End If
    Loop While intScount > 0

    'Bin out if not correct number of octets
    If intIPtag = 3 Then
        'Now check number limitations
        strIPstring = ipnum
        bolTest = True
        Do
            intScount = InStr(1, strIPstring, ".")
            If intScount <> 0 Then
                strTmp = Mid$(strIPstring, 1, intScount - 1)
              Else
                strTmp = Mid$(strIPstring, 1)
            End If

            strIPstring = Mid$(strIPstring, intScount + 1)

            If Val(strTmp) > 255 Or Val(strTmp) < 0 Or strTmp = "" Then
                'Return BAD ip
                bolTest = False
            End If
        Loop While intScount <> 0

        'Return GOOD ip
        If bolTest Then
            checkIP = True
          Else
            checkIP = False
        End If
      Else
        'Return BAD ip
        checkIP = False
    End If

End Function


'=========================================================================
'Function Ping(strAddAs String) As string
'Pings an IP Address

'Inputs:    strAdd - The IP address to ping

'Returns:   Ping as long
'           Round Trip in ms
'=========================================================================
Public Function Ping(strAdd As String) As Long

  Dim lngHPort As Long, lngDAddress As Long, strMessage As String
  Dim lngResult As Long
  Dim ECHO As ICMP_ECHO_REPLY

    strMessage = "Echo This."
    lngDAddress = lngNewAddress(strAdd)

    lngHPort = IcmpCreateFile()
    lngResult = IcmpSendEcho(lngHPort, lngDAddress, strMessage, Len(strMessage), 0, ECHO, Len(ECHO), PING_TIMEOUT)
    If lngResult = 0 Then
        Ping = ECHO.status * -1
      Else
        Ping = ECHO.RoundTripTime
    End If

    lngResult = IcmpCloseHandle(lngHPort)

End Function


'=========================================================================
'Function Cleanup() As Boolean
'Closes ands frees the Winsock sockets

'Inputs:    None

'Returns:   Cleanup as boolean
'           True/False depending on whether it was a sucessful cleanup
'=========================================================================
Public Function Cleanup() As Boolean

  Dim lngResult As Long

    lngResult = WSACleanup()

    If lngResult <> 0 Then
        MsgBox "Windows Sockets error " & Trim$(Str$(lngResult)) & " occurred in Cleanup.", vbExclamation
        Cleanup = False
      Else
        Cleanup = True
    End If

End Function


'=========================================================================
'Function WSInitialise() As Boolean
'Initialises the API Winsock

'Inputs:    None

'Returns:   WSInitialise as boolean
'           True/False depending on whether it was a sucessful initilisation
'=========================================================================
Public Function WSInitialise() As Boolean

  Dim WSA_DAT As WSADATA

    'Check to make sure Winsock will talk to us!
    If WSAStartup(WS_VERSION_REQD, WSA_DAT) <> ERROR_SUCCESS Then
        MsgBox "Winsock32 failed to respond"
        WSInitialise = False
        Exit Function
    End If

    'Check for free sockets
    If WSA_DAT.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "Pinger needs at least " & CStr(MIN_SOCKETS_REQD) & " sockets."
        WSInitialise = False
        Exit Function
    End If
    
    'If we got this far it must be alright!
    WSInitialise = True

End Function
