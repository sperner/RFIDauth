Attribute VB_Name = "CV3600AT_2"
'                   Int GetVersionAPI                     (char *VersionAPI)
Public Declare Function GetVersionAPI Lib "CVAPIV01.dll " (ByVal aver As String) As Integer
'                   int OpenComm                     (char *Com, int Baudrate)
Public Declare Function OpenComm Lib "CVAPIV01.dll " (ByVal Com As String, ByVal Baudrate As Long) As Integer
'                   int CloseComm                     (Void)
Public Declare Function CloseComm Lib "CVAPIV01.dll " () As Integer
'                   int CreateCommPort                (int SerialNUM,unsigned char *CommID,int Baudrate )
'                   int SetPort                       (int DeviceAddress, unsigned char PortState)
Public Declare Function SetPort Lib "cvapiv01.dll" (ByVal DeviceAddress As Long, ByVal PortState As Byte) As Integer
'                   int GetPort                       (int DeviceAddress, unsigned char *status)
Public Declare Function GetPort Lib "cvapiv01.dll" (ByVal DeviceAddress As Long, ByVal Status As Byte) As Integer
'                   int ActiveLED                     (int DeviceAddress, unsigned char NumLED, unsigned char ontime, ussigned char cycle)
Public Declare Function ActiveLED Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal NumLED As Byte, ByVal ontime As Byte, ByVal cycle As Byte) As Integer
'                   int SetLED                     (int DeviceAddress, unsigned char LEDState)
Public Declare Function SetLED Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal LEDState As Byte) As Integer
'                   int ActiveBuzzer                     (int DeviceAddress, unsigned char mode, unsigned char *pattern)
Public Declare Function ActiveBuzzer Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByRef pattern As Byte) As Integer
'                   int RF_Field                     (int DeviceAddress, unsigned char time)
Public Declare Function RF_Field Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal time As Byte) As Integer
'                   int SetWiegandStatus                     (int DeviceAddress,unsigned char status)
Public Declare Function SetWiegandStatus Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal Status As Byte) As Integer
'                   int ActiveWiegandMode                     (int DeviceAddress,unsigned char status)
Public Declare Function ActiveWiegandMode Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal Status As Byte) As Integer
'                   int WiegandMode                     (int DeviceAddress,unsigned char*data)
Public Declare Function WiegandMode Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef data As Byte) As Integer
'                   int SLE_Generic                    (int DeviceAddress,unsigned char CRC_Flag,unsigned char &length,unsigned char *buffer)
'                   int MF_Request                     (int DeviceAddress, unsigned char mode, unsigned char *ATQ)
Public Declare Function MF_Request Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByRef ATQ As Byte) As Integer
'                   int MF_Anticoll                     (int DeviceAddress, unsigned char *UID, unsigned char &Collision)
Public Declare Function MF_Anticoll Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte, ByRef Collision() As Byte) As Integer
Public Declare Function MF_Anticoll2 Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte, ByRef Collision() As Byte) As Integer
Public Declare Function MF_Anticoll3 Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte, ByRef Collision() As Byte) As Integer
'                   int MF_Select                     (int Device Address, unsigned char *UID)
Public Declare Function MF_Select Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte) As Integer
Public Declare Function MF_Select2 Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte) As Integer
Public Declare Function MF_Select3 Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef uId As Byte) As Integer
'                   int MF_Auth                       (int DeviceAddress, unsigned char KeyAB, unsigned char *snr, unsigned char add_blk)
Public Declare Function MF_Auth Lib "CVAPIV01.dll  " (ByVal DeviceAddress As Long, ByVal KeyAB As Byte, ByRef uId As Byte, ByVal add_blk As Byte) As Integer    'auth
'                   int MF_Read                    (int DeviceAddress,unsigned char add_blk, unsigned char num_blk, unsigned char *buffer)
Public Declare Function MF_Read Lib "cvapiv01.dll" (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal num_blk As Byte, ByRef buffer As Byte) As Integer 'read
'                   int MF_Write                     (int DeviceAddress,unsigned char add_blk, unsigned char num_blk, unsigned char *buffer)
Public Declare Function MF_Write Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal num_blk As Byte, ByRef buffer As Byte) As Integer 'write
'                   int MF_Halt                     (int Device Address, unsigned char mode)
Public Declare Function MF_Halt Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte) As Integer
'                   int MF_Incremnet                     (int DeviceAddress, unsigned char add_blk, int value)
Public Declare Function MF_Increment Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal val As Long) As Integer
'                   int MF_Decrement                     (int DeviceAddress, unsigned char add_blk, int value)
Public Declare Function MF_Decrement Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal val As Long) As Integer
'                   int MF_Transfer                     (int DeviceAddress, unsigned char add_blk )
Public Declare Function MF_Transfer Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte) As Integer
'                   int MF_Restore                     (int DeviceAddress, unsigned char add_blk )
Public Declare Function MF_Restore Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte) As Integer
'                   int MF_InitValue                     (int DeviceAddress, unsigned char add_blk, int value)
Public Declare Function MF_InitValue Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByVal value As Long) As Integer
'                   int MF_ReadValue                     (int DeviceAddress, unsigned char add_blk, int *value)
Public Declare Function MF_ReadValue Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal add_blk As Byte, ByRef value As Long) As Integer
'                   int MF_LoadKey                     (int DeviceAddress, unsigned char *Key)
Public Declare Function MF_LoadKey Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef value As Byte) As Integer
'                   int MF_StoreKeyToEE                     (int DeviceAddress, unsigned char KeyAB, unsigned char Sector, unsigned char *Key)
Public Declare Function MF_StoreKeyToEE Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal KeyAB As Byte, ByVal Sector As Byte, ByRef Key As Long) As Integer
'                   int MF_LoadKeyFromEF                     (int DeviceAddress, unsigned char KeyAB, unsigned char Sector)
Public Declare Function MF_LoadKeyFromEF Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal KeyAB As Byte, ByVal Sector As Byte) As Integer
'                   int MF_HLRead                     (int DeviceAddress, unsigned char mode, unsigned char add_blk, unsigned char num_blk, unsigned char *snr, unsigned char *buffer)
Public Declare Function MF_HLRead Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByVal add_blk As Byte, ByVal num_blk As Byte, ByRef snr As Byte, ByRef buffer As Byte) As Integer 'hlr
'                   int MF_HLWrite                    (int DeviceAddress, unsigned char mode, unsigned char add_blk, unsigned char num_blk, unsigned char *snr, unsigned char *buffer)
Public Declare Function MF_HLWrite Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByVal add_blk As Byte, ByVal num_blk As Byte, ByRef snr As Byte, ByRef buffer As Byte) As Integer
'                   int MF_HLInitVal                     (int DeviceAddress, unsigned char mode, unsigned char sect_num, unsigned char *snr, int value)
Public Declare Function MF_HLInitVal Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByVal sect_num As Byte, ByRef snr As Byte, ByVal value As Long) As Integer
'                   int MF_HLInc                     (int DeviceAddress, unsigned char mode, unsigned char sect_num, unsigned char *snr, int *value)
Public Declare Function MF_HLInc Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByVal sect_num As Byte, ByRef snr As Byte, ByRef value As Long) As Integer
'                   int MF_HLDec                     (int DeviceAddress, unsigned char mode, unsigned char sect_num, unsigned char *snr, int *value)
Public Declare Function MF_HLDec Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByVal sect_num As Byte, ByRef snr As Byte, ByRef value As Long) As Integer
'                   int MF_HLRequest                     (int DeviceAddress, unsigned char mode, int &length, unsigned char *UID)
Public Declare Function MF_HLRequest Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal mode As Byte, ByRef Length As Long, ByRef uId As Byte) As Integer
'                   int SetFirmwareBaudrate                     (int DeviceAddress, unsigned char Baudrate)
Public Declare Function SetFirmwareBaudrate Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal Baudrate As Byte) As Integer
'                   int SetDeviceAddress                     (int DeviceAddress, unsigned char &newAddress)
Public Declare Function SetDeviceAddress Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal newAddress As Byte) As Integer
'                   int GetVersionlNum                     (int DeviceAddress, char *VersionNUM)
Public Declare Function GetVersionlNum Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef VersionNUM As Byte) As Integer
'                   int GetSerialNum                     (int DeviceAddress, int &CurrentAddress, char *SerialNUM)
Public Declare Function GetSerialNum Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef CurrentAddress As Long, ByRef SerialNUM As Byte) As Integer
'                   int GetUserInfo                     (int DeviceAddress, char *UserInfo)
Public Declare Function GetUserInfo Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef UserInfo As Byte) As Integer
'                   int SetUserInfo                     (int DeviceAddress, char *UserInfo)
Public Declare Function SetUserInfo Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef UserInfo As Byte) As Integer
'                   int LcdDisplayLogo                     (int DeviceAddress,unsigned char*data)
Public Declare Function LcdDisplayLogo Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef data As Byte) As Integer
'                   int LcdDisplay                     (int DeviceAddress,unsigned char address,unsigned char length, char *Dstring)
Public Declare Function LcdDisplay Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByVal Address As Byte, ByRef Length As Byte, ByRef DString As Byte) As Integer
'                       GetKey                     (int DeviceAddress,unsigned char *keybuffer)
Public Declare Function GetKey Lib "CVAPIV01.dll " (ByVal DeviceAddress As Long, ByRef keybuffer As Byte) As Integer
'                   int ReadChar                     (unsigned char *byte)
Public Declare Function ReadChar Lib "CVAPIV01.dll " (ByRef keybuffer As Byte) As Integer

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

