Attribute VB_Name = "modIniFile"
Rem // --- This mod created an INI File, verry, verry Simple.

Option Explicit

Rem // --- I will use the INI Functions in 'KERNEL32' to Do The Work. :)
Public Declare Function getprivateprofilestring Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function writeprivateprofilestring Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Rem // --- OK all im doing is Creating a Function that can be accessed anywhere in this APP.
Rem // --- Then getting 'KERNEL32' To Do the Work.
Function ReadINI(Section As String, KeyName As String, FileName As String) As String
On Error Resume Next
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, getprivateprofilestring(Section, ByVal KeyName$, "", sRet, Len(sRet), FileName))
End Function

Rem // --- OK all im doing here is the same as above but Writeing the INI File Now.
Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
On Error Resume Next
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function
