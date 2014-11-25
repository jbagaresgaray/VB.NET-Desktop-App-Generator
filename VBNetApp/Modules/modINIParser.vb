'======================================================================================
'Filename:  modINIParser.vb
'Type:      Module
'Author:    Philip Cesar B.Garay
'Date:      12.Oct.2012
'Email:     philipgaray2@gmail.com
'
'Purpose:   This module parses the values in a particular Section->Key
'           inside the .INI file
'
'Usage:     ReadINI(whichSection, whichKey, the_ini_file_to_parse)
'
'
'           Example:  This will read the value of the Server under the DATABASE section
'
'                     ReadINI("DATABASE", "Server", "config.ini")
'
'=======================================================================================

Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Text

Module modINIParser

    Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer
    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Int32

    Function ReadINI(ByVal pSection As String, ByVal pKey As String, ByVal pIniFilename As String) As String

        Const DEFAULT_VALUE As String = ""          'default value or an empty string

        Dim lngReturnValue As Long                 'return value of the API call
        Dim strResult As String                     'the resulting string
        Dim lngBuffer As Long
        'length of the resulting string


        If (Trim$(pSection) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pSection") : Exit Function
        If (Trim$(pKey) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pKey") : Exit Function
        If (Trim$(pIniFilename) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pIniFilename") : Exit Function

        strResult = StrDup(1000, vbNullChar)         'pad the resulting string with NULL chars
        lngBuffer = System.Text.Encoding.Unicode.GetByteCount(strResult)                 'get the length of the resulting string

        lngReturnValue = GetPrivateProfileString(pSection, pKey, DEFAULT_VALUE, strResult, lngBuffer, pIniFilename)

        'remove comment
        If InStr(strResult, vbTab & ";", CompareMethod.Text) > 0 Then
            strResult = Trim$(Left$(strResult, InStr(strResult, vbTab & ";", CompareMethod.Text)))
        End If

        strResult = Replace(strResult, vbNullChar, "", , , CompareMethod.Text)       'strip-off all NULL characters

        ReadINI = Trim$(Replace(strResult, vbTab, "", , , CompareMethod.Text))       'strip-off all TAB characters

    End Function


    '==========================================================================
    ' This procedure will handle the saving of Key->Values to the .ini file
    '===========================================================================

    Function SaveINI(ByVal pSection As String, ByVal pKey As String, ByVal pValue As String, ByVal pIniFilename As String) As Long


        Dim lngReturnValue As Long

        If (Trim$(pSection) = "") Then MsgBox("DEBUG::mod_INIParser::SaveINI() - Bad or missing parameter pSection") : Exit Function
        If (Trim$(pKey) = "") Then MsgBox("DEBUG::mod_INIParser::SaveINI() - Bad or missing parameter pKey") : Exit Function
        If (Trim$(pIniFilename) = "") Then MsgBox("DEBUG::mod_INIParser::ReadINI() - Bad or missing parameter pIniFilename") : Exit Function

        'Comment = ReadComment(pSection, pKey, pValue)

        lngReturnValue = WritePrivateProfileString(pSection, pKey, pValue, pIniFilename)

        SaveINI = lngReturnValue

    End Function

    Function ReadComment(ByVal pSection As String, ByVal pKey As String, ByVal pIniFilename As String) As String

        Const DEFAULT_VALUE As String = ""          'default value or an empty string

        Dim lngReturnValue As Long                  'return value of the API call
        Dim strResult As String                     'the resulting string
        Dim lngBuffer As Long                       'length of the resulting string


        If (Trim$(pSection) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pSection") : Exit Function
        If (Trim$(pKey) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pKey") : Exit Function
        If (Trim$(pIniFilename) = "") Then MsgBox("DEBUG::mod_IniParser::ReadINI() - Bad or missing parameter pIniFilename") : Exit Function

        strResult = StrDup(1000, vbNullChar)         'pad the resulting string with NULL chars
        lngBuffer = Len(strResult)                 'get the length of the resulting string

        lngReturnValue = GetPrivateProfileString(pSection, pKey, DEFAULT_VALUE, strResult, lngBuffer, pIniFilename)

        'remove comment
        '    If VBA.InStr(strResult, vbTab & ";", , vbTextCompare) > 0 Then
        '        strResult = Trim$(VBA.Left$(strResult, VBA.InStr(strResult, vbTab & ";", , vbTextCompare)))
        '    End If

        strResult = Replace(strResult, vbNullChar, "", , , vbTextCompare)       'strip-off all NULL characters

        ReadComment = Trim$(Replace(strResult, vbTab, "", , , vbTextCompare))       'strip-off all TAB characters

    End Function

End Module
