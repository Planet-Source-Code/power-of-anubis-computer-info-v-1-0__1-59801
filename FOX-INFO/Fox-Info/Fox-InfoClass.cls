VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GetInfomation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Private Const LOCALE_USER_DEFAULT = &H400
Private Const LOCALE_SENGLANGUAGE = &H1001      '  English name of language
Private Const LOCALE_SENGCOUNTRY = &H1002       '  English name of country
Private Const LOCALE_SCURRENCY = &H14           '  local monetary symbol
Private Const LOCALE_SLONGDATE = &H20           '  long date format string
Private Const LOCALE_SDAYNAME1 = &H2A           '  long name for Monday
Private Const LOCALE_SDAYNAME2 = &H2B           '  long name for Tuesday
Private Const LOCALE_SDAYNAME3 = &H2C           '  long name for Wednesday
Private Const LOCALE_SDAYNAME4 = &H2D           '  long name for Thursday
Private Const LOCALE_SDAYNAME5 = &H2E           '  long name for Friday
Private Const LOCALE_SDAYNAME6 = &H2F           '  long name for Saturday
Private Const LOCALE_SDAYNAME7 = &H30           '  long name for Sunday
Private Const LOCALE_SABBREVDAYNAME1 = &H31     '  abbreviated name for Monday
Private Const LOCALE_SABBREVDAYNAME2 = &H32     '  abbreviated name for Tuesday
Private Const LOCALE_SABBREVDAYNAME3 = &H33     '  abbreviated name for Wednesday
Private Const LOCALE_SABBREVDAYNAME4 = &H34     '  abbreviated name for Thursday
Private Const LOCALE_SABBREVDAYNAME5 = &H35     '  abbreviated name for Friday
Private Const LOCALE_SABBREVDAYNAME6 = &H36     '  abbreviated name for Saturday
Private Const LOCALE_SABBREVDAYNAME7 = &H37     '  abbreviated name for Sunday
Private Const LOCALE_SDECIMAL = &HE             '  decimal separator
Private Const LOCALE_STHOUSAND = &HF            '  thousand separator
Private Const LOCALE_SGROUPING = &H10           '  digit grouping
Private Const LOCALE_IDIGITS = &H11             '  number of fractional digits
Private Const LOCALE_ILZERO = &H12              '  leading zeros for decimal
Private Const LOCALE_SDATE = &H1D               '  date separator
Private Const LOCALE_STIME = &H1E               '  time separator
Private Const LOCALE_SSHORTDATE = &H1F          '  short date format string
Private Const LOCALE_STIMEFORMAT = &H1003       '  time format string
Private Const LOCALE_SMONTHNAME1 = &H38         '  long name for January
Private Const LOCALE_SMONTHNAME2 = &H39         '  long name for February
Private Const LOCALE_SMONTHNAME3 = &H3A         '  long name for March
Private Const LOCALE_SMONTHNAME4 = &H3B         '  long name for April
Private Const LOCALE_SMONTHNAME5 = &H3C         '  long name for May
Private Const LOCALE_SMONTHNAME6 = &H3D         '  long name for June
Private Const LOCALE_SMONTHNAME7 = &H3E         '  long name for July
Private Const LOCALE_SMONTHNAME8 = &H3F         '  long name for August
Private Const LOCALE_SMONTHNAME9 = &H40         '  long name for September
Private Const LOCALE_SMONTHNAME10 = &H41        '  long name for October
Private Const LOCALE_SMONTHNAME11 = &H42        '  long name for November
Private Const LOCALE_SMONTHNAME12 = &H43        '  long name for December
Private Const LOCALE_SABBREVMONTHNAME1 = &H44   '  abbreviated name for January
Private Const LOCALE_SABBREVMONTHNAME2 = &H45   '  abbreviated name for February
Private Const LOCALE_SABBREVMONTHNAME3 = &H46   '  abbreviated name for March
Private Const LOCALE_SABBREVMONTHNAME4 = &H47   '  abbreviated name for April
Private Const LOCALE_SABBREVMONTHNAME5 = &H48   '  abbreviated name for May
Private Const LOCALE_SABBREVMONTHNAME6 = &H49   '  abbreviated name for June
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A   '  abbreviated name for July
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B   '  abbreviated name for August
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C   '  abbreviated name for September
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D  '  abbreviated name for October
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E  '  abbreviated name for November
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F  '  abbreviated name for December
Private Const LOCALE_SPOSITIVESIGN = &H50       '  positive sign
Private Const LOCALE_SNEGATIVESIGN = &H51       '  negative sign
Private Const LOCALE_IPOSSIGNPOSN = &H52        '  positive sign position
Private Const LOCALE_INEGSIGNPOSN = &H53        '  negative sign position
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Function GetPositiveSignPosition() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IPOSSIGNPOSN, Buffer, 99)
   GetPositiveSignPosition = LPSTRToVBString(Buffer)
End Function
Public Function GetNegativeSignPosition() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_INEGSIGNPOSN, Buffer, 99)
   GetNegativeSignPosition = LPSTRToVBString(Buffer)
End Function
Public Function GetNegativeSign() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SNEGATIVESIGN, Buffer, 99)
   GetNegativeSign = LPSTRToVBString(Buffer)
End Function
Public Function GetPositiveSign() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SPOSITIVESIGN, Buffer, 99)
   GetPositiveSign = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName12() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME12, Buffer, 99)
   GetShortMonthName12 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName11() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME11, Buffer, 99)
   GetShortMonthName11 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName10() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME10, Buffer, 99)
   GetShortMonthName10 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName9() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME9, Buffer, 99)
   GetShortMonthName9 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName8() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME8, Buffer, 99)
   GetShortMonthName8 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName7() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME7, Buffer, 99)
   GetShortMonthName7 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName6() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME6, Buffer, 99)
   GetShortMonthName6 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName5() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME5, Buffer, 99)
   GetShortMonthName5 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName4() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME4, Buffer, 99)
   GetShortMonthName4 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName3() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME3, Buffer, 99)
   GetShortMonthName3 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName2() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME2, Buffer, 99)
   GetShortMonthName2 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortMonthName1() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVMONTHNAME1, Buffer, 99)
   GetShortMonthName1 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName12() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME12, Buffer, 99)
   GetLongMonthName12 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName11() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME11, Buffer, 99)
   GetLongMonthName11 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName10() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME10, Buffer, 99)
   GetLongMonthName10 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName9() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME9, Buffer, 99)
   GetLongMonthName9 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName8() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME8, Buffer, 99)
   GetLongMonthName8 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName7() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME7, Buffer, 99)
   GetLongMonthName7 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName6() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME6, Buffer, 99)
   GetLongMonthName6 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName5() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME5, Buffer, 99)
   GetLongMonthName5 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName4() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME4, Buffer, 99)
   GetLongMonthName4 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName3() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME3, Buffer, 99)
   GetLongMonthName3 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName2() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME2, Buffer, 99)
   GetLongMonthName2 = LPSTRToVBString(Buffer)
End Function

Public Function GetLongMonthName1() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SMONTHNAME1, Buffer, 99)
   GetLongMonthName1 = LPSTRToVBString(Buffer)
End Function

Public Function GetTimeFormat() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIMEFORMAT, Buffer, 99)
   GetTimeFormat = LPSTRToVBString(Buffer)
End Function
Public Function GetShortDateFormat() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SSHORTDATE, Buffer, 99)
   GetShortDateFormat = LPSTRToVBString(Buffer)
End Function
Public Function GetTimeSeparator() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STIME, Buffer, 99)
   GetTimeSeparator = LPSTRToVBString(Buffer)
End Function
Public Function GetDateSeparator() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDATE, Buffer, 99)
   GetDateSeparator = LPSTRToVBString(Buffer)
End Function
Public Function GetLeadingZerosForDecimal() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_ILZERO, Buffer, 99)
   GetLeadingZerosForDecimal = LPSTRToVBString(Buffer)
End Function
Public Function GetNumberOfFractionalDigits() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_IDIGITS, Buffer, 99)
   GetNumberOfFractionalDigits = LPSTRToVBString(Buffer)
End Function
Public Function GetDigitGrouping() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SGROUPING, Buffer, 99)
   GetDigitGrouping = LPSTRToVBString(Buffer)
End Function
Public Function GetThousandSeparator() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_STHOUSAND, Buffer, 99)
   GetThousandSeparator = LPSTRToVBString(Buffer)
End Function
Public Function GetCurrencySymbol() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SCURRENCY, Buffer, 99)
   GetCurrencySymbol = LPSTRToVBString(Buffer)
End Function
Public Function GetLanguage() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGLANGUAGE, Buffer, 99)
   GetLanguage = LPSTRToVBString(Buffer)
End Function
Public Function GetCountry() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SENGCOUNTRY, Buffer, 99)
   GetCountry = LPSTRToVBString(Buffer)
End Function
Public Function GetLongDateFormat() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SLONGDATE, Buffer, 99)
   GetLongDateFormat = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay1() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME1, Buffer, 99)
   GetLongNameDay1 = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay2() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME2, Buffer, 99)
   GetLongNameDay2 = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay3() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME3, Buffer, 99)
   GetLongNameDay3 = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay4() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME4, Buffer, 99)
   GetLongNameDay4 = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay5() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME5, Buffer, 99)
   GetLongNameDay5 = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay6() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME6, Buffer, 99)
   GetLongNameDay6 = LPSTRToVBString(Buffer)
End Function
Public Function GetLongNameDay7() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDAYNAME7, Buffer, 99)
   GetLongNameDay7 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay1() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME1, Buffer, 99)
   GetShortNameDay1 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay2() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME2, Buffer, 99)
   GetShortNameDay2 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay3() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME3, Buffer, 99)
   GetShortNameDay3 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay4() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME4, Buffer, 99)
   GetShortNameDay4 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay5() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME5, Buffer, 99)
   GetShortNameDay5 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay6() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME6, Buffer, 99)
   GetShortNameDay6 = LPSTRToVBString(Buffer)
End Function
Public Function GetShortNameDay7() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SABBREVDAYNAME7, Buffer, 99)
   GetShortNameDay7 = LPSTRToVBString(Buffer)
End Function
Public Function GetDecimalSeparator() As String
   Dim Buffer As String * 100
   Dim dl&
   dl& = GetLocaleInfo(LOCALE_USER_DEFAULT, LOCALE_SDECIMAL, Buffer, 99)
   GetDecimalSeparator = LPSTRToVBString(Buffer)
End Function
Private Function LPSTRToVBString$(ByVal s$)
   Dim nullpos&
   nullpos& = InStr(s$, Chr$(0))
   If nullpos > 0 Then
      LPSTRToVBString = Left$(s$, nullpos - 1)
   Else
      LPSTRToVBString = ""
   End If
End Function
Public Function TempDir() As String
    Dim sTmp1$
    sTmp1 = Environ$("temp")
    While Right$(sTmp1, 1) = "\"
        sTmp1 = Left$(sTmp1, Len(sTmp1) - 1)
    Wend
    If sTmp1 <> "" Then
        On Error Resume Next
        MkDir sTmp1
        On Error GoTo 0
    End If
    TempDir = sTmp1
End Function
Public Function GetOS() As String
    GetOS = Environ$("os")
End Function
Public Function GetUSERNAME() As String
    GetUSERNAME = Environ$("username")
End Function
Public Function GetSystemDrive() As String
    GetSystemDrive = Environ$("systemdrive")
End Function
Public Function GetWinDir() As String
    GetWinDir = Environ$("windir")
End Function
Public Function SystemDir() As String
Dim result
Dim SystemDirectory As String
SystemDirectory = Space(144)
result = GetSystemDirectory(SystemDirectory, 144)
If result = 0 Then
    MsgBox "Cannot Get the Windows System Directory", vbCritical, "Warning"
Else
    SystemDir = Trim(SystemDirectory)
End If
End Function

