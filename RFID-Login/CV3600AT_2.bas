Attribute VB_Name = "CV3600AT_2"
'RFID-Reader SDK / DLL-Aufrufe
Public Declare Function GetVersionAPI Lib "CVAPIV01.dll " (ByVal aver As String) As Integer
Public Declare Function OpenComm Lib "CVAPIV01.dll " (ByVal Com As String, ByVal Baudrate As Long) As Integer
Public Declare Function CloseComm Lib "CVAPIV01.dll " () As Integer
Public Declare Function ActiveLED Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal NumLED As Byte, ByVal ontime As Byte, ByVal cycle As Byte) As Integer
Public Declare Function SetLED Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal LEDState As Byte) As Integer
Public Declare Function ActiveBuzzer Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByRef pattern As Byte) As Integer
Public Declare Function SetPort Lib "cvapiv01.dll" (ByVal DeviceAddress As Long, ByVal PortState As Byte) As Integer
Public Declare Function GetPort Lib "cvapiv01.dll" (ByVal DeviceAddress As Long, ByVal Status As Byte) As Integer
Public Declare Function RF_Field Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal time As Byte) As Integer
Public Declare Function MF_Request Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByRef ATQ As Byte) As Integer
Public Declare Function MF_Anticoll Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte, ByRef Collision() As Byte) As Integer
Public Declare Function MF_Select Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte) As Integer
Public Declare Function MF_Auth Lib "CVAPIV01.dll  " (ByVal DeviceAddress As Long, ByVal KeyAB As Byte, ByRef uId As Byte, ByVal add_blk As Byte) As Integer    'auth
Public Declare Function MF_Read Lib "cvapiv01.dll" (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal num_blk As Byte, ByRef buffer As Byte) As Integer 'read
Public Declare Function MF_Write Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal num_blk As Byte, ByRef buffer As Byte) As Integer 'write
Public Declare Function MF_LoadKey Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef value As Byte) As Integer

'System Error/Status Codes (0x00-0x0F)
'OK 0x00 Command OK. ( success)
'PARA_ERR 0x01 Parameter value out of range error
'TMO_ERR 0x04 Reader reply time out error
'SEQ_ERR 0x05 Communication Sequence Number out of order
'CMD_ERR 0x06 Reader received unknown command
'CHKSUM_ERR 0x07 Communication Check Sum Error
'INTR_ERR 0x08 Unknown Internal Error

'Card Error/Status Codes (0x10-0x1F)
'NOTAG_ERR 0x11 No card detected
'CRC_ERR 0x12 Wrong CRC received from card
'PARITY_ERR 0x13 Wrong Parity Received from card
'BITCNT_ERR 0x14 Wrong number of bits received from the card
'BYTECNT_ERR 0x15 Wrong number of bytes received from the card
'CRD_ERR 0x16 Any other error happened when communicate with card

'MIFARE Error/Status Codes (0x20-0x2F)
'MF_AUTHERR 0x20 No Authentication Possible
'MF_SERNRERR 0x21 Wrong Serial Number read during Anti-collision.
'MF_NOAUTHERR 0x22 Card is not authenticated
'MF_VALFMT 0x23 Not value block format
'MF_VAL 0x24 Any problem with the VALUE related function

'ISO15693 Error/Status Codes (0x20-0x2F)
'BlockLocked_ERR 0x17 the block has been locked and cannot write the value and cannot lock
'again
'Command_unsurport 0x18 the command to the card do not support
'Commandformat_err 0x19 the command format is wrong
'Option_unsurport 0x1A the Option flag do not support
'Unknown_err 0x1B unknown error
'Block_notexist 0x1C the masked block do not exist
'Block_lockunsucess 0x1D block_lock option is not succeed
'Flag_WRONGPARAM 0x1F the input value of the flag and other parameters is wrong.

'Type-B Card Error/Status Codes (0x30-0x3F)
'<To be defined>

'SAM Error/Status Codes (0x40-0x4F)
'<To be defined>
