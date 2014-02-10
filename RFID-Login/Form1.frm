VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'Kein
   Caption         =   "RFID-Login"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   13920
      TabIndex        =   5
      Top             =   9720
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogOff 
      Caption         =   "LogOff"
      Height          =   375
      Left            =   13920
      TabIndex        =   4
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Timer Timer3 
      Interval        =   1500
      Left            =   480
      Top             =   3120
   End
   Begin VB.CommandButton cmdReboot 
      Caption         =   "Reboot"
      Height          =   375
      Left            =   13920
      TabIndex        =   3
      Top             =   10440
      Width           =   1095
   End
   Begin VB.CommandButton cmdShutDown 
      Caption         =   "ShutDown"
      Height          =   375
      Left            =   13920
      TabIndex        =   2
      Top             =   10800
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   3120
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'Kein
      Height          =   1845
      Left            =   480
      Picture         =   "Form1.frx":240042
      ScaleHeight     =   7380
      ScaleMode       =   0  'Benutzerdefiniert
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   360
      Width           =   7395
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'Kein
      Height          =   1075
      Left            =   9720
      Picture         =   "Form1.frx":26C79C
      ScaleHeight     =   1080
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   0
      Width           =   6195
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim select_Type As Integer
Dim RFIDon As Boolean
Dim RFIDserial As String
Dim LEDgreen As Boolean
Dim LEDred As Boolean
Dim UserName As String
Dim PassWort As String
Dim UserLogged As String
Dim PassLogged As String
Dim EncUser As String
Dim EncPass As String
Const EncryptWith = "passwort"
Const DatFileEncWith = "kennwort"
Const AnzahlUserSets = 3
Dim LogFileName As String
Dim LogFileIndex As Integer
Const DatFileName = "c:\database.dat"

Private Sub Form_Load()             'Initialisierung
Dim Com As String
Dim Ver As String * 100
Dim Baudrate As Long
Dim Zeit
  Zeit = time
  LogFileName = "LOG-" & Mid(Date, 1, 2) & Mid(Date, 4, 2) & Mid(Date, 9, 2) & "-" & Mid(Zeit, 1, 2) & "h" & Mid(Zeit, 4, 2) & ".txt"
  Open LogFileName For Random As #1 Len = 32
  LogFileIndex = 1
  App.TaskVisible = False
  'Call disableTaskManager(True)     'Taskmanager "deaktivieren"
  Call PutWindowOnTop(Form1)        'Always on Top
    cmdLogOff.Visible = False
    cmdHide.Visible = False
    Com = "COM1"                    'Com-Port
    Baudrate = 115200               'Com-Baudrate
    If OpenComm(Com, Baudrate) = 0 Then 'Com-Port öffnen
        'lstAusgabe.AddItem ("COM: " & Com & "  Baudrate:" & Str$(Baudrate) & " Succeed")
    End If
    select_Type = 1
    If Not RF_Field(0, 1) Then      'Antenne einschalten
        RFIDon = True
    End If
    Timer1.Enabled = True           'Timer einschalten
    Timer3.Enabled = False
    If SetLED(0, 0) Then LEDgreen = False
End Sub

Public Sub ReadTAG()            'UserName & PassWort einlesen
Dim i As Integer
Dim ATQ(1) As Byte              'Request
Dim uId(3) As Byte              'Tag-Serial
Dim Collision() As Byte         'Anticoll
Dim KeyAB As Byte               'Key A oder B?
Dim Key(5) As Byte              'KeyString
Dim StartNum As Integer         'von...
Dim PageNum As Integer          '..."bis"
Dim Sector As Integer           'Sektor
Dim buffer(63) As Byte          'Puffer
Dim charbuffer(63) As String    'Puffer->String
Dim ASCIIbuffer As String       'Klartext
Dim s As String                 'Temp
   For i = 0 To 5                                                   'Key einlesen
         Key(i) = CInt("&H" & Mid("FF FF FF FF FF FF", 3 * i + 1, 2))
   Next i
   KeyAB = 96                                                       '97=keyB
   If MF_Request(0, 1, ATQ(0)) = 0 Then                             'Request All
      If MF_Anticoll(0, uId(0), Collision) = 0 Then                 'Anticollision
         If MF_Select(0, uId(0)) = 0 Then                           'Select Tag
            If MF_LoadKey(0, Key(0)) = 0 Then                       'Load Keyphrase
               StartNum = 0
               PageNum = 1
               For Sector = StartNum To StartNum + PageNum - 1
                  If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then 'Authentication
                     If MF_Read(0, Sector * 4, 4, buffer(0)) = 0 Then   'Read Tag
                        For i = 0 To 63
                            If Len(Hex(buffer(i))) = 1 Then         'bei einstelligem Wert
                               s = "0" & Hex(buffer(i))
                            Else
                               s = Hex(buffer(i))
                            End If
                            charbuffer(i) = Chr("&H" & s)           'für String-Puffer aufbereiten
                            If s = "00" Then
                                ASCIIbuffer = ASCIIbuffer & " "     'gibt sonst Fehler
                            Else
                                ASCIIbuffer = ASCIIbuffer & "" & charbuffer(i)
                            End If
                         Next i
                     End If
                  End If
               Next Sector
            End If
         End If
      End If
   Else
   End If
   EncUser = RTrim(Mid(ASCIIbuffer, 17, 16))        'Username auslesen
   EncPass = RTrim(Mid(ASCIIbuffer, 33, 16))        'Passwort auslesen
   UserName = DecodeString(EncUser, EncryptWith)    'Dekodieren
   PassWort = DecodeString(EncPass, EncryptWith)    'Dekodieren
End Sub

Private Sub Authentication(authSerial As String)    'RFID-Authentifizierung
Dim Temp As String
Dim Reason As String
Dim tmpSerial As String
Dim tmpUser As String
Dim tmpPass As String
Dim UserDataSet As Integer
  Open DatFileName For Random As #2 Len = 32        'Database öffnen
    For UserDataSet = 1 To AnzahlUserSets           'bis Anzahl Einträge
        Get #2, (UserDataSet * 4 - 3), Temp         'get...
        tmpSerial = DecodeString(Temp, DatFileEncWith)  '...Serial
        If authSerial = tmpSerial Then
            Call ReadTAG                            'Tag lesen
            Get #2, (UserDataSet * 4 - 2), Temp
            tmpUser = DecodeString(Temp, DatFileEncWith)
            Get #2, (UserDataSet * 4 - 1), Temp
            tmpPass = DecodeString(Temp, DatFileEncWith)
            If UserName = tmpUser Then              'User vergleichen
                If PassWort = tmpPass Then          'Password vergleichen
                    Put #1, LogFileIndex, vbCrLf & vbCrLf & "Tag-Serial: " & RTrim(authSerial)
                    LogFileIndex = LogFileIndex + 1
                    Put #1, LogFileIndex, vbCrLf & "Tag-User: " & RTrim(UserName)
                    LogFileIndex = LogFileIndex + 1
                    Put #1, LogFileIndex, vbCrLf & "Tag-Pass: " & RTrim(PassWort)
                    LogFileIndex = LogFileIndex + 1
                    Put #1, LogFileIndex, vbCrLf & "Access Granted " & Date
                    LogFileIndex = LogFileIndex + 1
                    Buzzer (True)
                    'lblAccess.Caption = "ACCESS GRANTED"
                    RFIDserial = authSerial
                    Call GotAccess                  'Anmelden++
                    Exit For
                Else
                    Reason = "Wrong Password "
                    If UserDataSet = AnzahlUserSets Then Call AuthError(Reason, authSerial)
                End If
            Else
                Reason = "Wrong Username "
                If UserDataSet = AnzahlUserSets Then Call AuthError(Reason, authSerial)
            End If
        Else
            Reason = "Wrong Serial "
            If UserDataSet = AnzahlUserSets Then Call AuthError(Reason, authSerial)
        End If
    Next UserDataSet
  Close #2
End Sub

Private Sub AuthError(tmpReason As String, tmpSerial)       'Error loggen
Dim Zeit
    Zeit = time
    Put #1, LogFileIndex, vbCrLf & vbCrLf & "Tag-Serial: " & RTrim(tmpSerial)
    LogFileIndex = LogFileIndex + 1
    Put #1, LogFileIndex, vbCrLf & tmpReason
    LogFileIndex = LogFileIndex + 1
    Put #1, LogFileIndex, vbCrLf & Date & " - " & Mid(Zeit, 1, 8)
    LogFileIndex = LogFileIndex + 1
    'Call ReadAll                   'Tag komplett auslesen
    Buzzer (False)
    'lblAccess.Caption = "ACCESS DENIED"
End Sub

Private Sub GotAccess()
    'Call RFID_Logon(UserName, PassWort)                    'WSH
    If BreakIn(UserName, PassWort, False) = False Then     'Gina
        MsgBox ("Login-Error")
    End If
    UserLogged = UserName
    PassLogged = PassWort
    Timer1.Enabled = False
    Timer3.Enabled = True
    cmdLogOff.Visible = True
    cmdHide.Visible = True
    'lblAccess.Caption = ""
    'Shell_NotifyIcon NIM_DELETE, nid
'    'Call disableTaskManager(False)
'    If CloseComm = 0 Then
'        Close #1
'        End
'    End If
End Sub

Public Sub Buzzer(Granted As Boolean)   'Device=0, mode: 0=off 1=on 4=pattern, pattern=Array
Dim arrPattern(4) As Byte
    If Granted = True Then
        arrPattern(0) = 1                               'Units of first on time
        arrPattern(1) = 0                               'Units of first off time
        arrPattern(2) = 0                               'Units of second on time
        arrPattern(3) = 0                               'Units of second off time
        arrPattern(4) = 1                               'cycles
    Else
        arrPattern(0) = 2                               'Units of first on time
        arrPattern(1) = 1                               'Units of first off time
        arrPattern(2) = 1                               'Units of second on time
        arrPattern(3) = 1                               'Units of second off time
        arrPattern(4) = 1                               'cycles
    End If
    If Not ActiveBuzzer(0, 4, arrPattern(0)) Then
    End If
End Sub

Private Sub Timer1_Timer()                              'Tag-Suche
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim Serial As String
'If Not lblAccess.Caption = "" Then Timer2.Enabled = True
'If Not makeping("192.168.0.9") Then                     'AD-Server erreichbar?
'    lblServerStatus.Visible = True
'Else
'    lblServerStatus.Visible = False
'End If
  Serial = ""
  If MF_Request(0, 1, ATQ(0)) = 0 Then                  'GetSerial1
     If MF_Anticoll(0, uId(0), Collision) = 0 Then      'GetSerial2
        For i = 0 To 3
            Serial = Serial + Hex(uId(i))               'Serial->String
        Next i
        Call Authentication(Serial)
     End If
  Else
  End If
End Sub

Private Sub Timer3_Timer()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim Serial As String
  Serial = ""
  If MF_Request(0, 1, ATQ(0)) = 0 Then                  'GetSerial1
     If MF_Anticoll(0, uId(0), Collision) = 0 Then      'GetSerial2
        For i = 0 To 3
            Serial = Serial + Hex(uId(i))               'Serial->String
        Next i
        If (RFIDserial = Serial) Then
            If (LEDgreen = False) Then
                If Not SetLED(0, 2) Then LEDgreen = True
            Else
                If Not SetLED(0, 0) Then LEDgreen = False
            End If
        Else
            Timer3.Enabled = False
            If Not SetLED(0, 0) Then LEDgreen = False
            If CloseComm = 0 Then Close #1
            Call LogOff
            End
        End If
     End If
  Else
    Timer3.Enabled = False
    If Not SetLED(0, 0) Then LEDgreen = False
    If CloseComm = 0 Then Close #1
    Call LogOff
    End
  End If
End Sub

Private Sub cmdReboot_Click()                           'Neustart
    ReBooT
End Sub

Private Sub cmdShutDown_Click()                         'Herunterfahren
    ShutDown
End Sub

Private Sub cmdLogOff_Click()                           'Abmelden
    LogOff
End Sub

Private Sub cmdHide_Click()
Dim WSHShell
    Set WSHShell = CreateObject("WScript.Shell")        'Objekt erzeugen
    WSHShell.SendKeys "%{TAB}"
    WSHShell.SendKeys "^(%({DEL}))"
    WSHShell.SendKeys "{ESC}"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    Timer3.Enabled = False
    'Call disableTaskManager(False)
    If CloseComm = 0 Then
        Close #1
        End
    End If
End Sub

