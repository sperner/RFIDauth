Attribute VB_Name = "WinLogon"
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function VirtualAllocEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReadProcessMemory Lib "kernel32.dll" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function VirtualFreeEx Lib "kernel32.dll" (ByVal hProcess As Long, lpAddress As Any, ByRef dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Type TrayData
    hwnd As Long
    id As Long
End Type
Private Type TBButton
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved(0 To 1) As Byte
    dwData As Long
    iString As Long
End Type
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeoutAndVersion As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type

Const WM_CLOSE = &H10
Const STANDARD_RIGHTS_REQUIRED As Long = &HF0000
Const SYNCHRONIZE As Long = &H100000
Const PROCESS_ALL_ACCESS As Long = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
Const MEM_COMMIT As Long = &H1000
Const PAGE_READWRITE As Long = &H4
Const WM_USER As Long = &H400
Const TB_BUTTONCOUNT As Long = (WM_USER + 24)
Const TB_GETBUTTON As Long = (WM_USER + 23)
Const NIS_HIDDEN = &H1
Const NIF_STATE = &H8
Const NIM_MODIFY As Long = &H1&
Const MEM_RELEASE As Long = &H8000
Const kLength = 255&
Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Private TIcon As NOTIFYICONDATA

Private Sub killTaskManager()
Dim hWndTask As Long
    hWndTask = FindWindow(vbNullString, "Windows Task-Manager")
    If hWndTask <> 0 Then
        PostMessage hWndTask, WM_CLOSE, ByVal 0&, ByVal 0&
    End If
End Sub

Private Sub taskManagerIconVisible(bVisible As Boolean)
Dim hWndTray As Long
Dim hProc As Long
Dim pid As Long
Dim vaPtr As Long
Dim ret As Long
Dim tbut As TBButton
Dim cButtons As Long
Dim td As TrayData
Dim i As Long
Dim Result As Long
Dim Title As String
    hWndTray = FindWindowEx(ByVal 0&, ByVal 0&, "Shell_TrayWnd", vbNullString)
    hWndTray = FindWindowEx(hWndTray, ByVal 0&, "TrayNotifyWnd", vbNullString)
    hWndTray = FindWindowEx(hWndTray, ByVal 0&, "SysPager", vbNullString)
    hWndTray = FindWindowEx(hWndTray, ByVal 0&, "ToolbarWindow32", vbNullString)
    If hWndTray = 0 Then Exit Sub
    Call GetWindowThreadProcessId(hWndTray, pid)
    If pid = 0 Then Exit Sub
    hProc = OpenProcess(PROCESS_ALL_ACCESS, 0, pid)
    If hProc = 0 Then Exit Sub
    vaPtr = VirtualAllocEx(hProc, ByVal 0&, Len(tbut), MEM_COMMIT, PAGE_READWRITE)
    If vaPtr = 0 Then GoTo cleanup
    cButtons = SendMessage(hWndTray, TB_BUTTONCOUNT, ByVal 0&, ByVal 0&)
    On Error GoTo cleanup
    For i = 0 To cButtons - 1
        Call SendMessage(hWndTray, TB_GETBUTTON, i, ByVal vaPtr)
        Call ReadProcessMemory(hProc, ByVal vaPtr, tbut, Len(tbut), ret)
        If Not tbut.dwData = 0 Then
            Call ReadProcessMemory(hProc, ByVal tbut.dwData, td, Len(td), ret)
            Result = GetWindowTextLength(td.hwnd) + 1
            Title = Space$(Result)
            Result = GetWindowText(td.hwnd, Title, Result)
            Title = Left$(Title, Len(Title) - 1)
            If Title = "Windows Task-Manager" And td.id = 0 Then ' es gibt viele versteckte TrayIcons vom Task-Manager, nur das mit der ID 0 ist für uns von Interesse
                TIcon.cbSize = Len(TIcon)
                TIcon.hwnd = td.hwnd
                TIcon.uId = td.id
                TIcon.uFlags = NIF_STATE
                TIcon.dwStateMask = NIS_HIDDEN
                If bVisible Then
                    TIcon.dwState = TIcon.dwState And Not NIS_HIDDEN
                Else
                    TIcon.dwState = TIcon.dwState Or NIS_HIDDEN
                End If
                Shell_NotifyIcon NIM_MODIFY, TIcon
            End If
        End If
    Next i
    
cleanup:
    If hProc Then
        If vaPtr Then
            Call VirtualFreeEx(hProc, ByVal vaPtr, 0&, MEM_RELEASE)
        End If
    End If
    If hProc Then CloseHandle (hProc)
End Sub

Private Function GetSysDir() As String
Dim nBuffer As String
Dim nReturn As Long
  nBuffer = Space(kLength)
  nReturn = GetSystemDirectory(nBuffer, kLength)
  If nReturn > 0 Then
      GetSysDir = Left(nBuffer, nReturn)
  End If
End Function

Public Sub disableTaskManager(bDisable As Boolean)
    If bDisable Then
        killTaskManager                             'Task-Manager zuerst schließen
        Shell GetSysDir & "\taskmgr.exe", vbHide    'Task-Manager versteckt öffnen
        Sleep 500
        taskManagerIconVisible False                'TrayIcon deaktivieren
    Else
        taskManagerIconVisible True                 'TrayIcon sichtbar machen
        killTaskManager                             'Task-Manager schließen
    End If
End Sub

Public Function PutWindowOnTop(pFrm As Form)
Dim lngWindowPosition As Long
    lngWindowPosition = SetWindowPos(pFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Public Sub RFID_Logon(Name As String, Pass As String)
Dim NameBuffer(15) As String
Dim PassBuffer(15) As String
Dim WSHShell
Dim i As Integer
    For i = 0 To 15
        NameBuffer(i) = " "
        PassBuffer(i) = " "
    Next i
    Set WSHShell = CreateObject("WScript.Shell")
    'WSHShell.AppActivate "Unbenannt - Editor"
    WSHShell.AppActivate "Windows-Anmeldung"
    WSHShell.SendKeys "{TAB}"
    WSHShell.SendKeys "{TAB}"
    WSHShell.SendKeys "{TAB}"
    WSHShell.SendKeys "{TAB}"
    For i = 1 To Len(Name)
        NameBuffer(i - 1) = Mid(Name, i, 1)
        WSHShell.SendKeys NameBuffer(i - 1)
    Next i
    WSHShell.SendKeys "{TAB}"
    For i = 1 To Len(Pass)
        PassBuffer(i - 1) = Mid(Pass, i, 1)
        WSHShell.SendKeys PassBuffer(i - 1)
    Next i
    WSHShell.SendKeys "{ENTER}"
End Sub

