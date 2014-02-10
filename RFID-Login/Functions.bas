Attribute VB_Name = "Functions"
Private Declare Function GetHostByName Lib "wsock32.dll" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVersionRequired&, lpWSAdata As WSAdata) As Long
Private Declare Function WSACleanup Lib "wsock32.dll" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal HANDLE As Long) As Boolean
Private Declare Function IcmpSendEcho Lib "ICMP" (ByVal IcmpHandle As Long, ByVal DestAddress As Long, ByVal RequestData As String, ByVal RequestSize As Integer, RequestOptns As IP_OPTION_INFORMATION, ReplyBuffer As IP_ECHO_REPLY, ByVal ReplySize As Long, ByVal TimeOut As Long) As Boolean
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public Type NOTIFYICONDATA                  'Minimize...
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Type LUID                           'Window-handling
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES               'Adjust Token
  PrivilegeCount As Long
  TheLuid As LUID
  Attributes As Long
End Type
Private Type WSAdata                        'Makeping...
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To 255) As Byte
    szSystemStatus(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type
Private Type Hostent                        'Makeping...
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type
Private Type IP_OPTION_INFORMATION          'Makeping...
    TTL As Byte
    Tos As Byte
    Flags As Byte
    OptionsSize As Long
    OptionsData As String * 128
End Type
Private Type IP_ECHO_REPLY                  'Makeping...
    Address(0 To 3) As Byte
    Status As Long
    RoundTripTime As Long
    DataSize As Integer
    Reserved As Integer
    data As Long
    Options As IP_OPTION_INFORMATION
End Type
Const SOCKET_ERROR = 0                      'Makeping...
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1
Private Const EWX_SHUTDOWN As Long = 1      'selbsterklärend...
Private Const EWX_FORCE As Long = 4
Private Const EWX_REBOOT = 2
Private Const EWX_LOGOFF = 0
Public Const NIM_ADD = &H0                  'Minimize...
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200       'Mausbewegung
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public nid As NOTIFYICONDATA                'Minimize...


Public Function makeping(ByVal HostName As String) As Boolean    'Ping über Netzwerk
    Dim hFile As Long, lpWSAdata As WSAdata
    Dim hHostent As Hostent, AddrList As Long
    Dim Address As Long, rIP As String
    Dim OptInfo As IP_OPTION_INFORMATION
    Dim EchoReply As IP_ECHO_REPLY
    Call WSAStartup(&H101, lpWSAdata)
    If GetHostByName(HostName + String(64 - Len(HostName), 0)) <> SOCKET_ERROR Then
        CopyMemory hHostent.h_name, ByVal GetHostByName(HostName + String(64 - Len(HostName), 0)), Len(hHostent)
        CopyMemory AddrList, ByVal hHostent.h_addr_list, 4
        CopyMemory Address, ByVal AddrList, 4
    End If
    hFile = IcmpCreateFile()
    If hFile = 0 Then
        MsgBox "Unable to Create File Handle", vbCritical + vbOKOnly
        makeping = False
        Exit Function
    End If
    OptInfo.TTL = 255
    If IcmpSendEcho(hFile, Address, String(32, "A"), 32, OptInfo, EchoReply, Len(EchoReply) + 8, 2000) Then
        rIP = CStr(EchoReply.Address(0)) + "." + CStr(EchoReply.Address(1)) + "." + CStr(EchoReply.Address(2)) + "." + CStr(EchoReply.Address(3))
    Else
        makeping = False
    End If
    If EchoReply.Status = 0 Then
        makeping = True
    Else
        makeping = False
    End If
    Call IcmpCloseHandle(hFile)
    Call WSACleanup
End Function

Private Sub AdjustToken()
   Const TOKEN_ADJUST_PRIVILEGES = &H20
   Const TOKEN_QUERY = &H8
   Const SE_PRIVILEGE_ENABLED = &H2
   Dim hdlProcessHandle As Long
   Dim hdlTokenHandle As Long
   Dim tmpLuid As LUID
   Dim tkp As TOKEN_PRIVILEGES
   Dim tkpNewButIgnored As TOKEN_PRIVILEGES
   Dim lBufferNeeded As Long
   hdlProcessHandle = GetCurrentProcess()
   OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY), hdlTokenHandle
   LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
   tkp.PrivilegeCount = 1
   tkp.TheLuid = tmpLuid
   tkp.Attributes = SE_PRIVILEGE_ENABLED
   AdjustTokenPrivileges hdlTokenHandle, False, tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
 
Public Sub ShutDown()
 AdjustToken
 ExitWindowsEx (EWX_SHUTDOWN), &HFFFF
End Sub

Public Sub ReStart()
  AdjustToken
  ExitWindowsEx (EWX_FORCE), &HFFFF
End Sub

Public Sub ReBooT()
  AdjustToken
  ExitWindowsEx (EWX_REBOOT), &HFFFF
End Sub

Public Sub LogOff()
  AdjustToken
  ExitWindowsEx (EWX_LOGOFF), &HFFFF
End Sub


