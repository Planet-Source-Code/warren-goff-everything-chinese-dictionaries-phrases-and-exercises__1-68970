Attribute VB_Name = "modNetDetect"
Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean

Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8
