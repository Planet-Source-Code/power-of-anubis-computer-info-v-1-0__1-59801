Attribute VB_Name = "RegEdit"
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
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Const LANG_NEUTRAL = &H0
Const SUBLANG_DEFAULT = &H1

Global Const SPIF_SENDWININICHANGE = &H2
Global Const SPIF_UPDATEINIFILE = &H1
Global Const SPI_GETSCREENSAVETIMEOUT = 14
Global Const SPI_SETSCREENSAVETIMEOUT = 15
Global Const SPI_GETSCREENSAVEACTIVE = 16
Global Const SPI_SETSCREENSAVEACTIVE = 17
Global Const SPI_SETDESKWALLPAPER = 20
Public Const READ_CONTROL = &H20000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_READ = READ_CONTROL
Public Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
   KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
   KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
   KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const KEY_EXECUTE = ((KEY_READ) And (Not SYNCHRONIZE))
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE _
   Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
' Possible registry data types
Public Enum InTypes
   ValNull = 0
   ValString = 1
   ValXString = 2
   ValBinary = 3
   ValDWord = 4
   ValLink = 6
   ValMultiString = 7
   ValResList = 8
End Enum
' Registry value type definitions
Public Const REG_NONE As Long = 0
Public Const REG_SZ As Long = 1
Public Const REG_EXPAND_SZ As Long = 2
Public Const REG_BINARY As Long = 3
Public Const REG_DWORD As Long = 4
Public Const REG_LINK As Long = 6
Public Const REG_MULTI_SZ As Long = 7
Public Const REG_RESOURCE_LIST As Long = 8
' Registry section definitions
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006
' Codes returned by Reg API calls
Private Const ERROR_NONE = 0
Private Const ERROR_BADDB = 1
Private Const ERROR_BADKEY = 2
Private Const ERROR_CANTOPEN = 3
Private Const ERROR_CANTREAD = 4
Private Const ERROR_CANTWRITE = 5
Private Const ERROR_OUTOFMEMORY = 6
Private Const ERROR_INVALID_PARAMETER = 7
Private Const ERROR_ACCESS_DENIED = 8
Private Const ERROR_INVALID_PARAMETERS = 87
Private Const ERROR_NO_MORE_ITEMS = 259

'Public Const HKEY_CURRENT_USER = &H80000001
'Global Const SPIF_SENDWININICHANGE = &H2
'Global Const SPIF_UPDATEINIFILE = &H1
'Global Const SPI_GETSCREENSAVETIMEOUT = 14
'Global Const SPI_SETSCREENSAVETIMEOUT = 15
'Global Const SPI_GETSCREENSAVEACTIVE = 16
'Global Const SPI_SETSCREENSAVEACTIVE = 17
'Global Const SPI_SETDESKWALLPAPER = 20
' Registry API functions used in this module (there are more of them)
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long
Public Declare Function RegFlushKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal numBytes As Long)
'TOP MOST DECLARATION
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long

Public Sub CreateKey(Folder As String, Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value

End Sub

Public Sub CreateIntegerKey(Folder As String, Value As Integer)

Dim b As Object
On Error Resume Next
Set b = CreateObject("wscript.shell")
b.RegWrite Folder, Value, "REG_DWORD"


End Sub

Public Function ReadKey(Value As String) As String

Dim b As Object
Dim R
On Error Resume Next
Set b = CreateObject("wscript.shell")
R = b.regread(Value)
ReadKey = R
End Function


Public Sub DeleteKey(Value As String)

Dim b As Object
On Error Resume Next
Set b = CreateObject("Wscript.Shell")
b.RegDelete Value
End Sub

Public Function ReadRegistryGetAll(ByVal Group As Long, ByVal Section As String, Idx As Long) As Variant
Dim lResult As Long, lKeyValue As Long, lDataTypeValue As Long
Dim lValueLength As Long, lValueNameLength As Long
Dim sValueName As String, sValue As String
Dim td As Double
On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
sValueName = Space$(2048)
lValueLength = Len(sValue)
lValueNameLength = Len(sValueName)
lResult = RegEnumValue(lKeyValue, Idx, sValueName, lValueNameLength, 0&, lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid$(sValue, 1, 1)) + &H100& * Asc(Mid$(sValue, 2, 1)) + &H10000 * Asc(Mid$(sValue, 3, 1)) + &H1000000 * CDbl(Asc(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
   End If
   sValue = Left$(sValue, lValueLength - 1)
   sValueName = Left$(sValueName, lValueNameLength)
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
' Return the datatype, value name and value as an array
ReadRegistryGetAll = Array(lDataTypeValue, sValueName, sValue)
End Function


Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub






















'PROCESS

Public Sub KillProcessById(p_lngProcessId As Long)
  Dim lnghProcess As Long
  Dim lngReturn As Long
    
    lnghProcess = OpenProcess(1&, -1&, p_lngProcessId)
    lngReturn = terminateprocess(lnghProcess, 0&)
    
    If lngReturn = 0 Then
        RetrieveError
    End If
End Sub




'PROCESS

Private Sub RetrieveError()
  Dim strBuffer As String
    strBuffer = Space(200)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, GetLastError, LANG_NEUTRAL, strBuffer, 200, ByVal 0&
    MsgBox strBuffer
End Sub



'TOP MOST FUNCTION

Public Function SetTopMostWindow(hwnd As Long, TopMost As Boolean) _
   As Long

   If TopMost = True Then
      SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, _
         0, FLAGS)
   Else
      SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, _
         0, 0, FLAGS)
      SetTopMostWindow = False
   End If
End Function


'--------------DXVERSION
  
