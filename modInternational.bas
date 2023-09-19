Attribute VB_Name = "modInternational"
Option Explicit

Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

Private Const LOCALE_SYSTEM_DEFAULT = &H400
Private Const LOCALE_USER_DEFAULT = &H800
Private Const LOCALE_NOUSEROVERRIDE = &H80000000

Private Const LOCALE_STHOUSAND As Long = &HF&
Private Const LOCALE_SDECIMAL As Long = &HE
Private Const LOCALE_SLIST As Long = &HC

Private distTest As Object 'PMDist.DistParser

Private Function GetInfo(ByVal lInfo As Long) As String
    Dim buffer As String
    Dim bufferLength As Integer
    
    buffer = Space(255)
    bufferLength = GetLocaleInfo(LOCALE_SYSTEM_DEFAULT, lInfo, buffer, Len(buffer))
    GetInfo = Left$(buffer, bufferLength - 1)
End Function

Public Function GetSystemListSeparator() As String
    GetSystemListSeparator = GetInfo(LOCALE_SLIST)
End Function

Public Function GetDecimalSeparator() As String
    GetDecimalSeparator = GetInfo(LOCALE_SDECIMAL)
End Function

Public Function FormatEnglishToSystemLocale(ByVal numberString As String) As String
    FormatEnglishToSystemLocale = Replace(numberString, ".", GetInfo(LOCALE_SDECIMAL))
End Function

Public Function FormatSystemLocaleToEnglish(ByVal numberString As String) As String
    FormatSystemLocaleToEnglish = Replace(CDbl(numberString), GetInfo(LOCALE_SDECIMAL), ".")
End Function

Public Function ConvertSystemToEnglish(ByVal Value As String) As String
    Set distTest = CreateObject("PMDist.DistParser")
    'is this a number?
    If IsNumeric(Value) Then
        ConvertSystemToEnglish = FormatSystemLocaleToEnglish(Value)
    Else
        
        'is this a distribution?
        distTest.ParamSeparator = GetSystemListSeparator
        distTest.Parse Value
        If distTest.IsValid And distTest.DistType <> Constant Then
            distTest.ParamSeparator = ","
            ConvertSystemToEnglish = distTest.ToStringEnglish
        Else
            'must be attribute/variable
            ConvertSystemToEnglish = Value
        End If
    End If

End Function

Public Function ConvertEnglishToSystem(ByVal Value As String) As String
    Set distTest = CreateObject("PMDist.DistParser")
    'is this a number?
    If IsNumeric(Value) Then
        ConvertEnglishToSystem = FormatEnglishToSystemLocale(Value)
    Else
        
        'is this a distribution?
        distTest.ParamSeparator = ","
        distTest.Parse Value, True
        If distTest.IsValid And distTest.DistType <> Constant Then
            distTest.ParamSeparator = GetSystemListSeparator
            ConvertEnglishToSystem = distTest.ToString
        Else
            'must be attribute/variable
            ConvertEnglishToSystem = Value
        End If
    End If

End Function
