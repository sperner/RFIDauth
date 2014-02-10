VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "RFID-Login"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   8325
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdStoreKey 
      Caption         =   "StoreKey"
      Height          =   375
      Left            =   120
      TabIndex        =   61
      Top             =   5880
      Width           =   855
   End
   Begin VB.CommandButton cmdLoadKey 
      Caption         =   "LoadKey"
      Height          =   375
      Left            =   1080
      TabIndex        =   60
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txtData2 
      Height          =   375
      Left            =   0
      TabIndex        =   59
      Text            =   "11 22 33 44 55 66 77 88 99 10 11 12 13 14 15 16"
      Top             =   7320
      Width           =   4095
   End
   Begin VB.CommandButton cmdDatGet 
      Caption         =   "GetSet"
      Height          =   375
      Left            =   4200
      TabIndex        =   52
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton cmdDatPut 
      Caption         =   "PutSet"
      Height          =   375
      Left            =   4200
      TabIndex        =   51
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox txtDatFileName 
      Height          =   375
      Left            =   6000
      TabIndex        =   50
      Text            =   "database"
      Top             =   7800
      Width           =   2175
   End
   Begin VB.TextBox txtDatEncPass 
      Height          =   375
      Left            =   6000
      TabIndex        =   49
      Text            =   "kennwort"
      Top             =   8280
      Width           =   2175
   End
   Begin VB.TextBox txtDatSet 
      Height          =   375
      Left            =   4200
      TabIndex        =   48
      Text            =   "1"
      Top             =   6600
      Width           =   735
   End
   Begin VB.TextBox txtDatSerial 
      Height          =   375
      Left            =   6000
      TabIndex        =   47
      Text            =   "E610DCAA"
      Top             =   6360
      Width           =   2175
   End
   Begin VB.TextBox txtDatPass 
      Height          =   375
      Left            =   6000
      TabIndex        =   46
      Text            =   "rfid"
      Top             =   7320
      Width           =   2175
   End
   Begin VB.TextBox txtDatUser 
      Height          =   375
      Left            =   6000
      TabIndex        =   45
      Text            =   "administrator"
      Top             =   6840
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetEnc 
      Caption         =   "GetEnc"
      Height          =   375
      Left            =   600
      TabIndex        =   44
      Top             =   8280
      Width           =   735
   End
   Begin VB.CommandButton cmdPutEnc 
      Caption         =   "PutEnc"
      Height          =   375
      Left            =   600
      TabIndex        =   43
      Top             =   7800
      Width           =   735
   End
   Begin VB.TextBox txtNumber 
      Height          =   375
      Left            =   1440
      TabIndex        =   42
      Text            =   "1"
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox txtGet 
      Height          =   375
      Left            =   1920
      TabIndex        =   41
      Text            =   "decoded..."
      Top             =   8280
      Width           =   1575
   End
   Begin VB.TextBox txtPut 
      Height          =   375
      Left            =   1920
      TabIndex        =   40
      Text            =   "...to encode"
      Top             =   7800
      Width           =   1575
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   8280
      Width           =   495
   End
   Begin VB.CommandButton cmdPut 
      Caption         =   "Put"
      Height          =   375
      Left            =   120
      TabIndex        =   38
      Top             =   7800
      Width           =   495
   End
   Begin VB.TextBox txtEncPass 
      Height          =   375
      Left            =   1440
      TabIndex        =   30
      Text            =   "passwort"
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton cmdDecrypt 
      Caption         =   "Decrypt Read"
      Height          =   615
      Left            =   2520
      TabIndex        =   17
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdEncrypt 
      Caption         =   "Encrypt Write"
      Height          =   615
      Left            =   2520
      TabIndex        =   14
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      Left            =   1440
      TabIndex        =   29
      Text            =   "rfid"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CommandButton cmdTimer1 
      Caption         =   "T1"
      Height          =   375
      Left            =   2760
      TabIndex        =   26
      Top             =   6480
      Width           =   375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3840
      TabIndex        =   18
      Top             =   8160
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3120
      Top             =   6480
   End
   Begin VB.ComboBox cmbBlockNr 
      Height          =   315
      Left            =   2760
      TabIndex        =   24
      Text            =   "1"
      Top             =   5880
      Width           =   735
   End
   Begin VB.CommandButton cmdWriteASCII 
      Caption         =   "Write  ASCII"
      Height          =   615
      Left            =   1200
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1440
      TabIndex        =   28
      Text            =   "Administrator"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CommandButton cmdReadASCII 
      Caption         =   "Read  ASCII"
      Height          =   615
      Left            =   1200
      TabIndex        =   13
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdTagSerial 
      Caption         =   "RFID-Tag Serial"
      Height          =   615
      Left            =   2520
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdWiegandMode 
      Caption         =   "Wiegand Mode"
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox txtData 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Text            =   "11 22 33 44 55 66 77 88 99 10 11 12 13 14 15 16"
      Top             =   6960
      Width           =   4095
   End
   Begin VB.OptionButton optKeyB 
      Caption         =   "KeyB"
      Height          =   315
      Left            =   1080
      TabIndex        =   21
      Top             =   5520
      Width           =   735
   End
   Begin VB.OptionButton optKeyA 
      Caption         =   "KeyA"
      Height          =   315
      Left            =   240
      TabIndex        =   20
      Top             =   5520
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.ComboBox cmbKey 
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Text            =   "FF FF FF FF FF FF"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.ComboBox cmbBlocks 
      Height          =   315
      Left            =   2760
      TabIndex        =   23
      Text            =   "1"
      Top             =   5520
      Width           =   735
   End
   Begin VB.ComboBox cmbSector 
      Height          =   315
      Left            =   2760
      TabIndex        =   22
      Text            =   "0"
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "Write"
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdGetUser 
      Caption         =   "GetUser"
      Height          =   615
      Left            =   1200
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdGetSerial 
      Caption         =   "GetSerial"
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.CommandButton cmdWSoff 
      Caption         =   "Wiegand Start-Off"
      Height          =   615
      Left            =   960
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdWSon 
      Caption         =   "Wiegand Start-On"
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdWiegand 
      Caption         =   "Wiegand Active"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdRFID 
      Caption         =   "RFID On/Off"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdBuzzer 
      Caption         =   "Buzzer"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read"
      Height          =   615
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdLEDpower 
      Caption         =   "LEDs On/Off"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtStatus 
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Text            =   "OK"
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton cmdLEDblink 
      Caption         =   "LEDs blink"
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.ListBox lstAusgabe 
      Height          =   6105
      ItemData        =   "Form1.frx":0000
      Left            =   3600
      List            =   "Form1.frx":0002
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Label lblDatSet 
      Caption         =   "Satz#:"
      Height          =   255
      Left            =   4200
      TabIndex        =   58
      Top             =   6360
      Width           =   495
   End
   Begin VB.Label lblDatEncPass 
      Caption         =   "Encryption Passwort:"
      Height          =   375
      Left            =   5040
      TabIndex        =   57
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label lblDatFileName 
      Caption         =   "Dateiname:"
      Height          =   255
      Left            =   5040
      TabIndex        =   56
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label lblDatPass 
      Caption         =   "Passwort:"
      Height          =   255
      Left            =   5040
      TabIndex        =   55
      Top             =   7320
      Width           =   855
   End
   Begin VB.Label lblDatUser 
      Caption         =   "Username:"
      Height          =   255
      Left            =   5040
      TabIndex        =   54
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label lblDatSerial 
      Caption         =   "Tag-Serial:"
      Height          =   255
      Left            =   5040
      TabIndex        =   53
      Top             =   6360
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "De- / Encryption:"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblMaxPass 
      Caption         =   "Pass eingeben:"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   35
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Block-Nr"
      Height          =   255
      Left            =   2040
      TabIndex        =   34
      Top             =   5880
      Width           =   615
   End
   Begin VB.Label lblMaxUser 
      Caption         =   "User eingeben:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label lblBlocks 
      Caption         =   "Blocks"
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label lblSector 
      Caption         =   "Sector"
      Height          =   255
      Left            =   2040
      TabIndex        =   31
      Top             =   5160
      Width           =   495
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
Dim LEDgreen As Boolean
Dim LEDred As Boolean
Dim BuzzerOn As Boolean
Dim WiegandOn As Boolean
Dim UserName As String
Dim PassWort As String
Dim EncUser As String
Dim EncPass As String
Dim LogFileName As String
Dim DatFileName As String

Private Sub cmdDatPut_Click()
Dim DateTime As String
  DatFileName = txtDatFileName.Text & ".dat"
  Open DatFileName For Random As #2 Len = 32
    DateTime = Date + time
    Put #2, (txtDatSet.Text * 4 - 3), EncodeString(txtDatSerial.Text, txtDatEncPass.Text)
    Put #2, (txtDatSet.Text * 4 - 2), EncodeString(txtDatUser.Text, txtDatEncPass.Text)
    Put #2, (txtDatSet.Text * 4 - 1), EncodeString(txtDatPass.Text, txtDatEncPass.Text)
    Put #2, (txtDatSet.Text * 4), EncodeString(DateTime, txtDatEncPass.Text)
  Close #2
End Sub

Private Sub cmdDatGet_Click()
Dim Temp As String
  DatFileName = txtDatFileName.Text & ".dat"
  Open DatFileName For Random As #2 Len = 32
    Get #2, (txtDatSet.Text * 4 - 3), Temp
    txtDatSerial.Text = DecodeString(Temp, txtDatEncPass.Text)
    Get #2, (txtDatSet.Text * 4 - 2), Temp
    txtDatUser.Text = DecodeString(Temp, txtDatEncPass.Text)
    Get #2, (txtDatSet.Text * 4 - 1), Temp
    txtDatPass.Text = DecodeString(Temp, txtDatEncPass.Text)
    'Get #2, (txtDatSet.Text * 4), EncodeString(DateTime, txtDatEncPass.Text)
  Close #2
End Sub

Private Sub cmdPut_Click()
    Put #1, txtNumber.Text, txtPut.Text
End Sub

Private Sub cmdGet_Click()
Dim Temp As String
    Get #1, txtNumber.Text, Temp
    txtGet.Text = Temp
End Sub

Private Sub cmdPutEnc_Click()
    Put #1, txtNumber.Text, EncodeString(txtPut.Text, txtEncPass.Text)
End Sub

Private Sub cmdGetEnc_Click()
Dim Temp As String
    Get #1, txtNumber.Text, Temp
    Temp = DecodeString(Temp, txtEncPass.Text)
    txtGet.Text = Temp
End Sub

Private Sub cmdEncrypt_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Sector As Integer
Dim Key(5) As Byte
Dim buffer(31) As Byte '63?
Dim block As Byte
  lstAusgabe.AddItem ("-")
  txtStatus.Text = "Write ASCII"
  Sector = CStr(cmbSector.Text)
  If optKeyA.value = True Then
     KeyAB = 96
  Else
     KeyAB = 97
  End If
  block = CInt("&H0" & Mid(cmbBlocks.Text, 1, 1))
  For i = 0 To 5
         Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
  Next i
  For i = 0 To 31
      buffer(i) = 0
  Next i
  EncUser = EncodeString(txtUser.Text, txtEncPass.Text)
  EncPass = EncodeString(txtPass.Text, txtEncPass.Text)
  For i = 0 To ((Len(EncUser)) - 1)
      buffer(i) = Asc(Mid(EncUser, i + 1, 1))
  Next i
  For i = 0 To ((Len(EncPass)) - 1)
      buffer(i + 16) = Asc(Mid(EncPass, i + 1, 1))
  Next i
  If MF_Request(0, 1, ATQ(0)) = 0 Then
     If MF_Anticoll(0, uId(0), Collision) = 0 Then
        If MF_Select(0, uId(0)) = 0 Then
           If MF_LoadKey(0, Key(0)) = 0 Then
              If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                 If MF_Write(0, Sector * 4 + 1, 2, buffer(0)) = 0 Then
                    lstAusgabe.AddItem ("Sector " & CStr(Sector) & " data write:")
                    lstAusgabe.AddItem ("-")
                    lstAusgabe.AddItem "Decoded Username: " & txtUser.Text
                    lstAusgabe.AddItem "Decoded Password:  " & txtPass.Text
                 Else
                    txtStatus.Text = (MF_Write(0, 1, 1, buffer(0)))
                 End If
              End If
           End If
        End If
     End If
   Else
     txtStatus.Text = "No Tag-Response"
  End If
  lstAusgabe.AddItem "Encoded Username: " & EncUser
  lstAusgabe.AddItem "Encoded Password:  " & EncPass
End Sub

Private Sub cmdDecrypt_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Key(5) As Byte
Dim StartNum As Integer
Dim PageNum As Integer
Dim Sector As Integer
Dim buffer(63) As Byte
Dim Temp As Integer
Dim charbuffer(63) As String
Dim ASCIIbuffer As String
Dim s As String
Dim PWlen As Integer
Dim UNlen As Integer
   lstAusgabe.AddItem ("-")
   txtStatus.Text = "Read ASCII"
   For i = 0 To 5
         Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
   Next i
   If optKeyA.value = True Then
      KeyAB = 96
   Else
      KeyAB = 97
   End If
   If MF_Request(0, 1, ATQ(0)) = 0 Then
      If MF_Anticoll(0, uId(0), Collision) = 0 Then
         If MF_Select(0, uId(0)) = 0 Then
            If MF_LoadKey(0, Key(0)) = 0 Then
               StartNum = CStr(cmbSector.Text)
               PageNum = CStr(cmbBlocks.Text)
               For Sector = StartNum To StartNum + PageNum - 1
                  If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                     If MF_Read(0, Sector * 4, 4, buffer(0)) = 0 Then
                        lstAusgabe.AddItem ("Sector " & CStr(Sector) & " data read:")
                        lstAusgabe.AddItem ("-")
                        For i = 0 To 63
                            If Len(Hex(buffer(i))) = 1 Then
                               s = "0" & Hex(buffer(i))
                            Else
                               s = Hex(buffer(i))
                            End If
                            Temp = "&H" & s
                            charbuffer(i) = Chr("&H" & s)
                            If s = "00" Then
                                ASCIIbuffer = ASCIIbuffer & " "
                            Else
                                ASCIIbuffer = ASCIIbuffer & "" & charbuffer(i)
                            End If
                         Next i
                     Else
                        txtStatus.Text = MF_Read(0, Sector * 4, 4, buffer(0))
                     End If
                  End If
               Next Sector
            End If
         End If
      End If
   Else
     txtStatus.Text = "No Tag-Response"
   End If
   EncUser = RTrim(Mid(ASCIIbuffer, 17, 16))
   EncPass = RTrim(Mid(ASCIIbuffer, 33, 16))
   lstAusgabe.AddItem "Encoded Username: " & EncUser
   lstAusgabe.AddItem "Encoded Password:  " & EncPass
   UserName = DecodeString(RTrim(EncUser), txtEncPass.Text)
   PassWort = DecodeString(RTrim(EncPass), txtEncPass.Text)
   lstAusgabe.AddItem "Decoded Username: " & UserName
   lstAusgabe.AddItem "Decoded Password:  " & PassWort
End Sub

Private Sub cmdGetSerial_Click()
Dim x As Integer
Dim DeviceAddress As Long
Dim SerialNumber(7) As Byte
Dim SerialString As String
SerialString = ""
DeviceAddress = 0
lstAusgabe.AddItem ("-")
    txtStatus.Text = GetSerialNum(0, DeviceAddress, SerialNumber(0))   'Device=0, NumLED: 1=red 2=green 3=both, on-time in 100ms, cycles
        For x = 0 To 7
            SerialString = SerialString + Chr(SerialNumber(x))
        Next x
        lstAusgabe.AddItem ("RFID-ReaderSerial:  -> " & SerialString & " <-")
    txtStatus.Text = "Reader Serial"
End Sub

Private Sub cmdGetUser_Click()
Dim x As Integer
Dim SerialNumber(31) As Byte
Dim SerialString As String
lstAusgabe.AddItem ("-")
    txtStatus.Text = GetUserInfo(0, SerialNumber(0))
        For x = 0 To 31
            SerialString = SerialString + Chr(SerialNumber(x))   'SN nur Pointer!
        Next x
        lstAusgabe.AddItem (SerialString)
    txtStatus.Text = "User-Info"
End Sub

Private Sub cmdRFID_Click()
lstAusgabe.AddItem ("-")
    If (RFIDon = False) Then
        If Not RF_Field(0, 1) Then        'Device=0, off-time in 100us: 1=on after 100us
            txtStatus.Text = "RFID On"
            lstAusgabe.AddItem ("RFID On")
            RFIDon = True
        End If
    Else
        If Not RF_Field(0, 0) Then        'Device=0, off-time in 100us: 0=static off
            txtStatus.Text = "RFID Off"
            lstAusgabe.AddItem ("RFID Off")
            RFIDon = False
        End If
    End If
End Sub

Private Sub cmdLEDblink_Click()
lstAusgabe.AddItem ("-")
    If Not ActiveLED(0, 3, 4, 3) Then   'Device=0, NumLED: 1=red 2=green 3=both, on-time in 100ms, cycles
        txtStatus.Text = "LEDs blink"
        lstAusgabe.AddItem ("LEDs blink")
    End If
End Sub

Private Sub cmdLEDpower_Click()
lstAusgabe.AddItem ("-")
    If (LEDgreen = False) Then
        If Not SetLED(0, 3) Then        'Device=0, NumLED: 0=off 1=red 2=green 3=both
            txtStatus.Text = "LEDs On"
            lstAusgabe.AddItem ("LEDs On")
            LEDgreen = True
        End If
    Else
        If Not SetLED(0, 0) Then        'Device=0, NumLED: 0=off 1=red 2=green 3=both
            txtStatus.Text = "LEDs Off"
            lstAusgabe.AddItem ("LEDs Off")
            LEDgreen = False
        End If
    End If
End Sub

Private Sub cmdBuzzer_Click()
Dim arrPattern(4) As Byte
arrPattern(0) = 3   'Units of first on time
arrPattern(1) = 1   'Units of first off time
arrPattern(2) = 1   'Units of second on time
arrPattern(3) = 2   'Units of second off time
arrPattern(4) = 2   'cycles
lstAusgabe.AddItem ("-")
    If (BuzzerOn = False) Then
        If Not ActiveBuzzer(0, 4, arrPattern(0)) Then 'Device=0, mode: 0=off 1=on 4=pattern, pattern=Array
            txtStatus.Text = "Buzzer On"
            lstAusgabe.AddItem ("Buzzer On")
            BuzzerOn = True
        End If
    End If
End Sub

Private Sub cmdWiegand_Click()                  'nicht in DLL! ???
lstAusgabe.AddItem ("-")
    If (WiegandOn = False) Then
        If Not ActiveWiegandMode(0, 16) Then    'Device=0, status 1=indicate L&B 2=extPin L&B 16=WiegandMode
            txtStatus.Text = "Wiegand On"
            lstAusgabe.AddItem ("Wiegand On")
            WiegandOn = True
        End If
    Else
        If Not ActiveWiegandMode(0, 0) Then     'Device=0, status 1=indicate L&B 2=extPin L&B 16=WiegandMode
            txtStatus.Text = "Wiegand Off"
            lstAusgabe.AddItem ("Wiegand Off")
            WiegandOn = False
        End If
    End If
End Sub

Private Sub cmdWiegandMode_Click()  'not tested!        kein Einsprungpunkt in .dll!?
Dim DString(10) As Byte
DString(0) = 0      '0 = wiegand_26, 1 = wiegand_32, 2 = wiegand_40, 3=wiegand_34
DString(1) = 0      '0 = basic mode, output card ID-Nr; 1~63 = secure mode, key is needed. Output first 4Byte of block on Wiegand format
DString(2) = &H26   '0x26 = IDLE, 0x52 = All
DString(3) = 19     '1 = L&B by ext I/O, 2 = L&B auto alarm, 16 = Enable Wiegand
DString(4) = 0      '0 = same Wiegand output as card data, 1 = 4-bit keypad output format, 2 = 8-bit kof
DString(5) = &HFF   '5-10 = KeyA of User Card
DString(6) = &HFF
DString(7) = &HFF
DString(8) = &HFF
DString(9) = &HFF
DString(10) = &HFF
lstAusgabe.AddItem ("-")
        If Not WiegandMode(0, DString(0)) Then     'Device=0, status 1=indicate L&B 2=extPin L&B 16=WiegandMode
            txtStatus.Text = "WiegandMode Off"
            lstAusgabe.AddItem ("WiegandMode Off")
            WiegandOn = False
        Else
            txtStatus.Text = "WiegandMode On"
            lstAusgabe.AddItem ("WiegandMode On")
            WiegandOn = True
        End If
End Sub

Private Sub cmdWSoff_Click()
    If Not SetWiegandStatus(0, 0) Then      'Device=0, status 1=indicate L&B 2=extPin L&B 16=WiegandMode
        txtStatus.Text = "Wiegand StartUp-Off"
        lstAusgabe.AddItem ("-")
        lstAusgabe.AddItem ("Wiegand StartUp-Off")
    End If
End Sub

Private Sub cmdWSon_Click()
    If Not SetWiegandStatus(0, 18) Then     'Device=0, status 1=auto-ind L&B 2=extPin L&B 16=WiegandMode
        txtStatus.Text = "Wiegand StartUp-On"
        lstAusgabe.AddItem ("-")
        lstAusgabe.AddItem ("Wiegand StartUp-On")
    End If
End Sub

Private Sub cmdRead_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Key(5) As Byte
Dim StartNum As Integer
Dim PageNum As Integer
Dim Sector As Integer
Dim buffer(63) As Byte
Dim Readdate As String
Dim s As String
   lstAusgabe.AddItem ("-")
   For i = 0 To 5
         Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
   Next i
   If optKeyA.value = True Then
      KeyAB = 96
   Else
      KeyAB = 97
   End If
   If MF_Request(0, 1, ATQ(0)) = 0 Then
      If MF_Anticoll(0, uId(0), Collision) = 0 Then
         If MF_Select(0, uId(0)) = 0 Then
            If MF_LoadKey(0, Key(0)) = 0 Then
               StartNum = CStr(cmbSector.Text)
               PageNum = CStr(cmbBlocks.Text)
               For Sector = StartNum To StartNum + PageNum - 1
                  If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                     If MF_Read(0, Sector * 4, 4, buffer(0)) = 0 Then
                        lstAusgabe.AddItem ("Sector " & CStr(Sector) & " data read:")
                        lstAusgabe.AddItem ("-")
                        txtStatus.Text = "Read Hex"
                        For i = 0 To 63
                            If Len(Hex(buffer(i))) = 1 Then
                               s = "0" & Hex(buffer(i))
                            Else
                               s = Hex(buffer(i))
                            End If
                            Readdate = Readdate + " " + s
                            If i = 15 Then
                               Readdate = Mid(Readdate, 2, Len(Readdate) - 1)
                               lstAusgabe.AddItem Hex(Sector * 4) + "," + Readdate
                               Readdate = ""
                            End If
                            If i = 31 Then
                               Readdate = Mid(Readdate, 2, Len(Readdate) - 1)
                               lstAusgabe.AddItem Hex(Sector * 4 + 1) + "," + Readdate
                               Readdate = ""
                            End If
                            If i = 47 Then
                               Readdate = Mid(Readdate, 2, Len(Readdate) - 1)
                               lstAusgabe.AddItem Hex(Sector * 4 + 2) + "," + Readdate
                               Readdate = ""
                            End If
                            If i = 63 Then
                               Readdate = Mid(Readdate, 2, Len(Readdate) - 1)
                               lstAusgabe.AddItem Hex(Sector * 4 + 3) + "," + Readdate
                               Readdate = ""
                            End If
                        Next i
                     Else
                        txtStatus.Text = MF_Read(0, Sector * 4, 4, buffer(0))
                     End If
                  Else
                    txtStatus.Text = "No Authentication"
                  End If
               Next Sector
            End If
         End If
      End If
   Else
     txtStatus.Text = "No Tag-Response"
   End If
End Sub

Private Sub cmdReadASCII_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Key(5) As Byte
Dim StartNum As Integer
Dim PageNum As Integer
Dim Sector As Integer
Dim buffer(63) As Byte
Dim Temp As Integer
Dim charbuffer(63) As String
Dim ASCIIbuffer As String
Dim s As String
Dim PWlen As Integer
Dim UNlen As Integer
   lstAusgabe.AddItem ("-")
   txtStatus.Text = "Read ASCII"
   For i = 0 To 5
         Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
   Next i
   If optKeyA.value = True Then
      KeyAB = 96
   Else
      KeyAB = 97
   End If
   If MF_Request(0, 1, ATQ(0)) = 0 Then
      If MF_Anticoll(0, uId(0), Collision) = 0 Then
         If MF_Select(0, uId(0)) = 0 Then
            If MF_LoadKey(0, Key(0)) = 0 Then
               StartNum = CStr(cmbSector.Text)
               PageNum = CStr(cmbBlocks.Text)
               For Sector = StartNum To StartNum + PageNum - 1
                  If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                     If MF_Read(0, Sector * 4, 4, buffer(0)) = 0 Then
                        lstAusgabe.AddItem ("Sector " & CStr(Sector) & " data read:")
                        lstAusgabe.AddItem ("-")
                        For i = 0 To 63
                            If Len(Hex(buffer(i))) = 1 Then
                               s = "0" & Hex(buffer(i))
                            Else
                               s = Hex(buffer(i))
                            End If
                            Temp = "&H" & s
                            charbuffer(i) = Chr("&H" & s)
                            If s = "00" Then
                                ASCIIbuffer = ASCIIbuffer & " "
                            Else
                                ASCIIbuffer = ASCIIbuffer & "" & charbuffer(i)
                            End If
                         Next i
                     Else
                        txtStatus.Text = MF_Read(0, Sector * 4, 4, buffer(0))
                     End If
                  End If
               Next Sector
            End If
         End If
      End If
   Else
     txtStatus.Text = "No Tag-Response"
   End If
   UserName = RTrim(Mid(ASCIIbuffer, 17, 16))
   PassWort = RTrim(Mid(ASCIIbuffer, 33, 16))
   UNlen = Len(UserName)
   PWlen = Len(PassWort)
   lstAusgabe.AddItem UserName
   lstAusgabe.AddItem UNlen
   lstAusgabe.AddItem PassWort
   lstAusgabe.AddItem PWlen
End Sub

Private Sub cmdWrite_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Sector As Integer
Dim Key(5) As Byte
Dim buffer(31) As Byte  '63?
Dim strBuffer As String
Dim block As Byte
  lstAusgabe.AddItem ("-")
  txtStatus.Text = "Write Hex"
  Sector = CStr(cmbSector.Text)
  If optKeyA.value = True Then
     KeyAB = 96
  Else
     KeyAB = 97
  End If
  block = CInt("&H0" & Mid(cmbBlocks.Text, 1, 1))
  For i = 0 To 5
         Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
  Next i
  For i = 0 To 15
      buffer(i) = CInt("&H" & Mid(txtData.Text, 3 * i + 1, 2))
      strBuffer = strBuffer + " " + CStr(Mid(txtData.Text, 3 * i + 1, 2))
  Next i
  For i = 0 To 15
      buffer(i + 16) = CInt("&H" & Mid(txtData2.Text, 3 * i + 1, 2))
      strBuffer = strBuffer + " " + CStr(Mid(txtData2.Text, 3 * i + 1, 2))
  Next i
  If MF_Request(0, 1, ATQ(0)) = 0 Then
     If MF_Anticoll(0, uId(0), Collision) = 0 Then
        If MF_Select(0, uId(0)) = 0 Then
           If MF_LoadKey(0, Key(0)) = 0 Then
              If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                 If MF_Write(0, Sector * 4 + 1, 2, buffer(0)) = 0 Then
                    lstAusgabe.AddItem ("Sector " & CStr(Sector) & " data write:")
                    lstAusgabe.AddItem ("-")
                    lstAusgabe.AddItem txtData.Text
                    lstAusgabe.AddItem txtData2.Text
                 Else
                    txtStatus.Text = (MF_Write(0, Sector * 4 + 1, 1, buffer(0)))
                 End If
              End If
           End If
        End If
     End If
   Else
     txtStatus.Text = "No Tag-Response"
  End If
End Sub

Private Sub cmdWriteASCII_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Sector As Integer
Dim Key(5) As Byte
Dim buffer(31) As Byte '63?
Dim block As Byte
  lstAusgabe.AddItem ("-")
  txtStatus.Text = "Write ASCII"
  Sector = CStr(cmbSector.Text)
  If optKeyA.value = True Then
     KeyAB = 96
  Else
     KeyAB = 97
  End If
  block = CInt("&H0" & Mid(cmbBlocks.Text, 1, 1))
  For i = 0 To 5
         Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
  Next i
  For i = 0 To 31
      buffer(i) = 0
  Next i
  For i = 0 To ((Len(txtUser.Text)) - 1)
      buffer(i) = Asc(Mid(txtUser.Text, i + 1, 1))
  Next i
  For i = 0 To ((Len(txtPass.Text)) - 1)
      buffer(i + 16) = Asc(Mid(txtPass.Text, i + 1, 1))
  Next i
  If MF_Request(0, 1, ATQ(0)) = 0 Then
     If MF_Anticoll(0, uId(0), Collision) = 0 Then
        If MF_Select(0, uId(0)) = 0 Then
           If MF_LoadKey(0, Key(0)) = 0 Then
              If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                 If MF_Write(0, Sector * 4 + 1, 2, buffer(0)) = 0 Then
                    lstAusgabe.AddItem ("Sector " & CStr(Sector) & " data write:")
                    lstAusgabe.AddItem ("-")
                    lstAusgabe.AddItem txtUser.Text
                    lstAusgabe.AddItem txtPass.Text
                 Else
                    txtStatus.Text = (MF_Write(0, 1, 1, buffer(0)))
                 End If
              End If
           End If
        End If
     End If
   Else
     txtStatus.Text = "No Tag-Response"
  End If
End Sub

Private Sub cmdTagSerial_Click()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim Serial As String
  lstAusgabe.AddItem ("-")
  Serial = ""
  If MF_Request(0, 1, ATQ(0)) = 0 Then
     If MF_Anticoll(0, uId(0), Collision) = 0 Then
        For i = 0 To 3
            Serial = Serial + Hex(uId(i))
        Next i
        lstAusgabe.AddItem "RFID-Tag Serial: " + Serial
     End If
   Else
     txtStatus.Text = "No Tag-Response"
  End If
End Sub

Private Sub Form_Load()             'INIT
Dim Com As String
Dim Ver As String * 100
Dim Baudrate As Long
  LogFileName = "Log-" & Date & "-.txt"
  Open LogFileName For Random As #1 Len = 32
    Com = "COM1"
    Baudrate = 115200
    If GetVersionAPI(Ver) = 0 Then
        lstAusgabe.AddItem ("API " & Ver)
    End If
    If OpenComm(Com, Baudrate) = 0 Then
        lstAusgabe.AddItem ("COM: " & Com & "  Baudrate:" & Str$(Baudrate) & " Succeed")
    End If
    lstAusgabe.AddItem "------------------------------------------------------------------"
    select_Type = 1
    If Not RF_Field(0, 1) Then
        RFIDon = True
    End If
    LEDgreen = False
    LEDred = False
    BuzzerOn = False
    WiegandOn = False
End Sub

Private Sub cmdExit_Click()
    If CloseComm = 0 Then
        Close #1
        Timer1.Enabled = False
        End
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call cmdExit_Click
End Sub

Private Sub cmdTimer1_Click()
    If Timer1.Enabled = True Then
        Timer1.Enabled = False
    Else
        Timer1.Enabled = True
    End If
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim Serial As String
  lstAusgabe.AddItem ("-")
  Serial = ""
  If MF_Request(0, 1, ATQ(0)) = 0 Then
     If MF_Anticoll(0, uId(0), Collision) = 0 Then
        For i = 0 To 3
            Serial = Serial + Hex(uId(i))
        Next i
        If Serial = "E610DCAA" Then
            Call cmdReadASCII_Click
            Call RFID_Logon(UserName, PassWort)
            Timer1.Enabled = False
            If CloseComm = 0 Then
                End
            End If
        Else
            txtStatus.Text = Serial
        End If
     End If
   Else
     txtStatus.Text = "No Tag-Response"
  End If
End Sub

Private Sub ChangeKey()
Dim ATQ(1) As Byte
Dim uId(3) As Byte
Dim Collision() As Byte
Dim KeyAB As Byte
Dim Key(5) As Byte
Dim buffer(63) As Byte
Dim keyA(5) As Byte
Dim keyB(5) As Byte
Dim Sector As Integer
Dim i As Integer
  Sector = CStr(cmbSector.Text)
  For i = 0 To 5
    Key(i) = CInt("&H" & Mid(cmbKey.Text, 3 * i + 1, 2))
  Next i
  If optKeyA.value = True Then
    KeyAB = 96
  Else
    KeyAB = 97
  End If
  For i = 0 To 5
    keyA(i) = CInt("&H" & Mid("FF FF FF FF FF FF", 3 * i + 1, 2))    'cmbKey.Text
    keyB(i) = CInt("&H" & Mid("FF FF FF FF FF FF", 3 * i + 1, 2))    'sollte zweite textbox
  Next i
  If MF_Request(0, 1, ATQ(0)) = 0 Then
     lstAusgabe.AddItem ("Request Succeed")
     If MF_Anticoll(0, uId(0), Collision) = 0 Then
        lstAusgabe.AddItem ("Anticoll Succeed")
        If MF_Select(0, uId(0)) = 0 Then
           lstAusgabe.AddItem ("Select Succeed")
           If MF_LoadKey(0, Key(0)) = 0 Then
              lstAusgabe.AddItem ("LoadKey Succeed")
              If MF_Auth(0, KeyAB, uId(0), Sector * 4) = 0 Then
                 lstAusgabe.AddItem ("Auth Succeed")
                 If MF_Read(0, Sector * 4, 4, buffer(0)) = 0 Then  'sector*4+3
                    lstAusgabe.AddItem ("Read Succeed")
                    For i = 0 To 5
                        buffer(i) = keyA(i)
                        buffer(i + 58) = keyB(i)
                    Next i
                    If MF_Write(0, Sector * 4 + 3, 1, buffer(0)) = 0 Then
                       lstAusgabe.AddItem ("Modify Key Succeed")
                    End If
                 Else
                    txtStatus.Text = MF_Read(0, Sector * 4, 4, buffer(0))
                 End If
              End If
           End If
        End If
     End If
  End If
End Sub

Private Sub cmdLoadKey_Click()              'Load Key from Reader-EEPROM
Dim KeyAB As Byte                           'macht keinen Sinn als Button!
Dim Sector As Byte
    Sector = CStr(cmbSector.Text)
    If optKeyA.value = True Then
      KeyAB = 96
    Else
      KeyAB = 97
    End If
    If MF_LoadKeyFromEF(0, KeyAB, Sector) = 0 Then
        lstAusgabe.AddItem "Key" & Str(KeyAB)
    End If
End Sub

Private Sub cmdStoreKey_Click()
    Call ChangeKey 'MF_StoreKeyToEE
End Sub

Private Sub txtPass_Change()        'Info, wenn maxLänge (abhängig von block/write) überschritten
    If Len(txtPass.Text) = 16 Then  'anzBlocks x 16
        lblMaxPass.Caption = "Maximum erreicht"
    Else
        If Len(txtPass.Text) > 16 Then
            lblMaxPass.Caption = "max. 16 Zeichen!"
        Else
            lblMaxPass.Caption = "Pass Eingeben:"
        End If
    End If
    txtStatus.Text = "Pass-In"
End Sub

Private Sub txtUser_Change()        'Info, wenn maxLänge (abhängig von block/write) überschritten
    If Len(txtUser.Text) = 16 Then  'anzBlocks x 16
        lblMaxUser.Caption = "Maximum erreicht"
    Else
        If Len(txtUser.Text) > 16 Then
            lblMaxUser.Caption = "max. 16 Zeichen!"
        Else
            lblMaxUser.Caption = "User Eingeben:"
        End If
    End If
    txtStatus.Text = "User-In"
End Sub

