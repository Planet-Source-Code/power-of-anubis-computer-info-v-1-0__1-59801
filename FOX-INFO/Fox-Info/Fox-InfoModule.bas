Attribute VB_Name = "InfoModule"
'*******************************************************
'*  _ __  (_) _ _  ___  _ __      _ _  ___  / _|_____  *
'* | '_ ` | || '_\/ _ \| '_ `    / __|/ _ \| |_ _   _| *
'* | | | || || |   (_) | | | |   \__ \ (_) |  _| | |   *
'* |_| |_||_||_|  \___/|_| |_|   |___/\___/|_|   |_|   *
'*                                                     *
'*******************************************************
'
'Fox-Info v 1.0
'Copyright © : 2005
'
'Thank's to http://www.planet-source-code.com form where I used few codes.
'I worked hard to make this program, so please vote.


Option Explicit
Enum PERF_DETAIL
PERF_DETAIL_NOVICE = 100      ' The uninformed can understand it
PERF_DETAIL_ADVANCED = 200    ' For the advanced user
PERF_DETAIL_EXPERT = 300      ' For the expert user
PERF_DETAIL_WIZARD = 400      ' For the system designer
End Enum

Enum PDH_STATUS
PDH_CSTATUS_VALID_DATA = &H0
PDH_CSTATUS_NEW_DATA = &H1
PDH_CSTATUS_NO_MACHINE = &H800007D0
PDH_CSTATUS_NO_INSTANCE = &H800007D1
PDH_MORE_DATA = &H800007D2
PDH_CSTATUS_ITEM_NOT_VALIDATED = &H800007D3
PDH_RETRY = &H800007D4
PDH_NO_DATA = &H800007D5
PDH_CALC_NEGATIVE_DENOMINATOR = &H800007D6
PDH_CALC_NEGATIVE_TIMEBASE = &H800007D7
PDH_CALC_NEGATIVE_VALUE = &H800007D8
PDH_DIALOG_CANCELLED = &H800007D9
PDH_CSTATUS_NO_OBJECT = &HC0000BB8
PDH_CSTATUS_NO_COUNTER = &HC0000BB9
PDH_CSTATUS_INVALID_DATA = &HC0000BBA
PDH_MEMORY_ALLOCATION_FAILURE = &HC0000BBB
PDH_INVALID_HANDLE = &HC0000BBC
PDH_INVALID_ARGUMENT = &HC0000BBD
PDH_FUNCTION_NOT_FOUND = &HC0000BBE
PDH_CSTATUS_NO_COUNTERNAME = &HC0000BBF
PDH_CSTATUS_BAD_COUNTERNAME = &HC0000BC0
PDH_INVALID_BUFFER = &HC0000BC1
PDH_INSUFFICIENT_BUFFER = &HC0000BC2
PDH_CANNOT_CONNECT_MACHINE = &HC0000BC3
PDH_INVALID_PATH = &HC0000BC4
PDH_INVALID_INSTANCE = &HC0000BC5
PDH_INVALID_DATA = &HC0000BC6
PDH_NO_DIALOG_DATA = &HC0000BC7
PDH_CANNOT_READ_NAME_STRINGS = &HC0000BC8
End Enum

Public Declare Sub GlobalMemoryStatus Lib "kernel32" _
   (lpBuffer As MEMORYSTATUS)


Declare Function PdhVbGetOneCounterPath _
    Lib "PDH.DLL" _
    (ByVal PathString As String, _
    ByVal PathLength As Long, _
    ByVal DetailLevel As Long, _
    ByVal CaptionString As String) _
    As Long
    
Declare Function PdhVbCreateCounterPathList _
        Lib "PDH.DLL" _
        (ByVal PERF_DETAIL As Long, _
         ByVal CaptionString As String) _
        As Long

Declare Function PdhVbGetCounterPathFromList _
        Lib "PDH.DLL" _
        (ByVal Index As Long, _
         ByVal Buffer As String, _
         ByVal BufferLength As Long) _
        As Long

Declare Function PdhOpenQuery _
    Lib "PDH.DLL" _
    (ByVal Reserved As Long, _
    ByVal dwUserData As Long, _
    ByRef hQuery As Long) _
    As PDH_STATUS

Declare Function PdhCloseQuery _
    Lib "PDH.DLL" _
    (ByVal hQuery As Long) _
    As PDH_STATUS

Declare Function PdhVbAddCounter _
    Lib "PDH.DLL" _
    (ByVal QueryHandle As Long, _
    ByVal CounterPath As String, _
    ByRef CounterHandle As Long) _
    As PDH_STATUS

Declare Function PdhCollectQueryData _
    Lib "PDH.DLL" _
    (ByVal QueryHandle As Long) _
    As PDH_STATUS
    
Declare Function PdhVbIsGoodStatus _
    Lib "PDH.DLL" _
    (ByVal StatusValue As Long) _
    As Long
    
Declare Function PdhVbGetDoubleCounterValue _
    Lib "PDH.DLL" _
    (ByVal CounterHandle As Long, _
    ByRef CounterStatus As Long) _
    As Double
    
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Public memInfo As MEMORYSTATUS

'DRIVES SPACE
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2


'PROCESSES
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function terminateprocess Lib "kernel32" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPheaplist = &H1
Public Const TH32CS_SNAPthread = &H4
Public Const TH32CS_SNAPmodule = &H8
Public Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long
 th32DefaultHeapID As Long
 th32ModuleID As Long
 cntThreads As Long
 th32ParentProcessID As Long
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * MAX_PATH
End Type

Public Const MAX_WSADescription As Long = 256
Public Const MAX_WSASYSStatus As Long = 128
Public Const ERROR_SUCCESS       As Long = 0
Public Const WS_VERSION_REQD     As Long = &H101
Public Const WS_VERSION_MAJOR    As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR    As Long = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD    As Long = 1
Public Const SOCKET_ERROR        As Long = -1

Public Type HOSTENT
   hName      As Long
   hAliases   As Long
   hAddrType  As Integer
   hLen       As Integer
   hAddrList  As Long
End Type

Public Type WSADATA
   wVersion      As Integer
   wHighVersion  As Integer
   szDescription(0 To MAX_WSADescription)   As Byte
   szSystemStatus(0 To MAX_WSASYSStatus)    As Byte
   wMaxSockets   As Integer
   wMaxUDPDG     As Integer
   dwVendorInfo  As Long
End Type

Public Declare Function WSAGetLastError Lib "wsock32" () As Long

Public Declare Function WSAStartup Lib "wsock32" _
  (ByVal wVersionRequired As Long, _
   lpWSADATA As WSADATA) As Long
   
Public Declare Function WSACleanup Lib "wsock32" () As Long

Public Declare Function gethostname Lib "wsock32" _
  (ByVal szHost As String, _
   ByVal dwHostLen As Long) As Long
   
Public Declare Function gethostbyname Lib "wsock32" _
  (ByVal szHost As String) As Long
   
Public Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
  (hpvDest As Any, _
   ByVal hpvSource As Long, _
   ByVal cbCopy As Long)


Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
   
Private Declare Function DeleteUrlCacheEntry Lib "Wininet.dll" _
   Alias "DeleteUrlCacheEntryA" _
  (ByVal lpszUrlName As String) As Long
      
Public Function GetPublicIP()
   Dim sSourceUrl As String
   Dim sLocalFile As String
   Dim hfile As Long
   Dim buff As String
   Dim pos1 As Long
   Dim pos2 As Long
   sSourceUrl = "http://vbnet.mvps.org/resources/tools/getpublicip.shtml"
   sLocalFile = "c:\ip.txt"
   Call DeleteUrlCacheEntry(sSourceUrl)
   If DownloadFile(sSourceUrl, sLocalFile) Then
      hfile = FreeFile
      Open sLocalFile For Input As #hfile
         buff = Input$(LOF(hfile), hfile)
      Close #hfile
      pos1 = InStr(buff, "var ip =")
      If pos1 Then
         pos1 = InStr(pos1 + 1, buff, "'", vbTextCompare) + 1
         pos2 = InStr(pos1 + 1, buff, "'", vbTextCompare) '- 1
         GetPublicIP = Mid$(buff, pos1, pos2 - pos1)
      Else
         GetPublicIP = "Fox-Info can't get you wan ip address"
      End If
      Kill sLocalFile
   Else
      GetPublicIP = "Fox-Info can't get you wan ip address"
   End If
End Function

Private Function DownloadFile(ByVal sURL As String, _
                             ByVal sLocalFile As String) As Boolean
   
  DownloadFile = URLDownloadToFile(0, sURL, sLocalFile, 0, 0) = ERROR_SUCCESS
   
End Function





Public Function SocketsInitialize() As Boolean

   Dim WSAD As WSADATA
   Dim sLoByte As String
   Dim sHiByte As String
   
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
      MsgBox "The 32-bit Windows Socket is not responding."
      SocketsInitialize = False
      Exit Function
   End If
   
   
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        MsgBox "This application requires a minimum of " & _
                CStr(MIN_SOCKETS_REQD) & " supported sockets."
        
        SocketsInitialize = False
        Exit Function
    End If
   
   
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or _
     (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And _
      HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
      
      sHiByte = CStr(HiByte(WSAD.wVersion))
      sLoByte = CStr(LoByte(WSAD.wVersion))
      
      MsgBox "Sockets version " & sLoByte & "." & sHiByte & _
             " is not supported by 32-bit Windows Sockets."
      
      SocketsInitialize = False
      Exit Function
      
   End If
    
   SocketsInitialize = True
        
End Function

Public Function GetIPAddress() As String

   Dim sHostName    As String * 256
   Dim lpHost    As Long
   Dim HOST      As HOSTENT
   Dim dwIPAddr  As Long
   Dim tmpIPAddr() As Byte
   Dim i         As Integer
   Dim sIPAddr  As String
   
   If Not SocketsInitialize() Then
      GetIPAddress = ""
      Exit Function
   End If

   If gethostname(sHostName, 256) = SOCKET_ERROR Then
      GetIPAddress = ""
      MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & _
              " has occurred. Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If
   
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
    
   If lpHost = 0 Then
      GetIPAddress = ""
      MsgBox "Windows Sockets are not responding. " & _
              "Unable to successfully get Host Name."
      SocketsCleanup
      Exit Function
   End If

   CopyMemory HOST, lpHost, Len(HOST)
   CopyMemory dwIPAddr, HOST.hAddrList, 4
   
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemory tmpIPAddr(1), dwIPAddr, HOST.hLen
   
   For i = 1 To HOST.hLen
      sIPAddr = sIPAddr & tmpIPAddr(i) & "."
   Next
  
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   
   SocketsCleanup
    
End Function

Public Function HiByte(ByVal wParam As Integer) As Byte

   HiByte = (wParam And &HFF00&) \ (&H100)
 
End Function

Public Function LoByte(ByVal wParam As Integer) As Byte

   LoByte = wParam And &HFF&

End Function


Public Sub SocketsCleanup()

    If WSACleanup() <> ERROR_SUCCESS Then
        MsgBox "Socket error occurred in Cleanup."
    End If
    
End Sub



    

