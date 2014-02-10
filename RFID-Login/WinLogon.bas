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
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal uCmd As Long) As Long

Private Type TrayData                       'Taskmanager-Handling
    hwnd As Long
    id As Long
End Type
Private Type TBButton                       'Taskmanager-Handling
    iBitmap As Long
    idCommand As Long
    fsState As Byte
    fsStyle As Byte
    bReserved(0 To 1) As Byte
    dwData As Long
    iString As Long
End Type
Private Type NOTIFYICONDATA                 'Taskmanager-Handling
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
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
Const WM_SETTEXT  As Long = 12
Const BM_CLICK = &HF5

Private TIcon As NOTIFYICONDATA

Private Sub killTaskManager()                           'Taskmanager beenden
Dim hWndTask As Long
    hWndTask = FindWindow(vbNullString, "Windows Task-Manager")
    If hWndTask <> 0 Then
        PostMessage hWndTask, WM_CLOSE, ByVal 0&, ByVal 0&
    End If
End Sub

Private Sub taskManagerIconVisible(bVisible As Boolean) 'Taskmanagersymbol verstecken
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
            If Title = "Windows Task-Manager" And td.id = 0 Then
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
        If vaPtr Then Call VirtualFreeEx(hProc, ByVal vaPtr, 0&, MEM_RELEASE)
    End If
    If hProc Then CloseHandle (hProc)
End Sub

Private Function GetSysDir() As String                  'Systemverzeichnis finden
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
        killTaskManager                                 'Task-Manager zuerst schließen
        Shell GetSysDir & "\taskmgr.exe", vbHide        'Task-Manager versteckt öffnen
        Sleep 500
        taskManagerIconVisible False                    'TrayIcon deaktivieren
    Else
        taskManagerIconVisible True                     'TrayIcon sichtbar machen
        killTaskManager                                 'Task-Manager schließen
    End If
End Sub

Public Function PutWindowOnTop(pFrm As Form)            'Immer im Vordergrund
Dim lngWindowPosition As Long
    lngWindowPosition = SetWindowPos(pFrm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
End Function

Public Sub RFID_Logon(Name As String, Pass As String)   'Login via WSH
Dim NameBuffer(15) As String
Dim PassBuffer(15) As String
Dim WSHShell
Dim i As Integer
    For i = 0 To 15
        NameBuffer(i) = " "
        PassBuffer(i) = " "
    Next i
    Set WSHShell = CreateObject("WScript.Shell")        'Objekt erzeugen
    WSHShell.AppActivate "Windows-Anmeldung"            'Login-Fenster aktivieren
    WSHShell.SendKeys "{TAB}"
    WSHShell.SendKeys "{TAB}"
    WSHShell.SendKeys "{TAB}"
    WSHShell.SendKeys "{TAB}"
    For i = 1 To Len(Name)                              'Username übergeben
        NameBuffer(i - 1) = Mid(Name, i, 1)
        WSHShell.SendKeys NameBuffer(i - 1)
    Next i
    WSHShell.SendKeys "{TAB}"
    For i = 1 To Len(Pass)                              'Passwort übergeben
        PassBuffer(i - 1) = Mid(Pass, i, 1)
        WSHShell.SendKeys PassBuffer(i - 1)
    Next i
    WSHShell.SendKeys "{ENTER}"
End Sub

Public Function BreakIn(ByVal login As String, ByVal password As String, Optional locked As Boolean = False) As Boolean
    Dim hWNDuser As Long, hWNDpass As Long, hWNDok As Long, hWNDgina As Long, ret As Long, ret2 As Long, ret3 As Long
    Dim hWNDlogon As Long, hWNDErreur As Long           'Login via User32.dll
    Dim error As Boolean
    hWNDgina = FindWindow("#32770", vbNullString)
    If locked = False Then                              'jeweils Unterscheidung, ob bereits angemeldet
        hWNDuser = GetDlgItem(ByVal hWNDgina, 1502)     '...und PC gesperrt ist oder nicht
        hWNDpass = GetDlgItem(ByVal hWNDgina, 1503)
        hWNDok = GetDlgItem(ByVal hWNDgina, 1)
    Else
        hWNDuser = GetDlgItem(ByVal hWNDgina, 1953)
        hWNDpass = GetDlgItem(ByVal hWNDgina, 1954)
        hWNDok = GetDlgItem(ByVal hWNDgina, 1)
    End If
    ret = SendMessage(ByVal hWNDuser, WM_SETTEXT, 0, ByVal login)       'Benutzername übergeben
    ret2 = SendMessage(ByVal hWNDpass, WM_SETTEXT, 0, ByVal password)   'Passwort übergeben
    ret3 = SendMessage(ByVal hWNDok, BM_CLICK, 0, 0)                    '"Klick"
    hWNDgina = FindWindow("#32770", vbNullString)
    hWNDlogon = 0
    If locked = False Then
        hWNDlogon = GetDlgItem(ByVal hWNDgina, 1502)
    Else
        hWNDlogon = GetDlgItem(ByVal hWNDgina, 1953)
    End If
    While hWNDlogon <> 0
        If locked = False Then
            hWNDlogon = GetDlgItem(ByVal hWNDgina, 1502)
        Else
            hWNDlogon = GetDlgItem(ByVal hWNDgina, 1953)
        End If
        If isLoginError() = True Then
            error = True
            GoTo apr:
        End If
        xWait 100
        DoEvents
    Wend
apr:                                                    'On Error...
    If error = False Then
        BreakIn = True
    Else
debutfaux:
    hWNDgina = FindWindow("#32770", vbNullString)
         hWNDok = GetDlgItem(ByVal hWNDgina, 2)
         ret3 = SendMessage(ByVal hWNDok, BM_CLICK, 0, 0)
         DoEvents
         If isLoginError = True Then GoTo debutfaux:
        BreakIn = False
    End If
End Function

Public Sub xWait(ByVal MilsecToWait As Long)            'Pause
    Dim lngEndingTime As Long
    lngEndingTime = GetTickCount() + (MilsecToWait)
    Do While GetTickCount() < lngEndingTime
        DoEvents
    Loop
End Sub

Private Function isLoginError() As Boolean              'Fehler während Ausführung?
    Dim hWNDgina As Long
    Dim ret As String
    hWNDgina = FindWindow("#32770", vbNullString)
    ret = GetDlgItem(hWNDgina, CLng("65535"))
    If ret <> "0" Then
        isLoginError = True
    Else
        isLoginError = False
    End If
End Function


