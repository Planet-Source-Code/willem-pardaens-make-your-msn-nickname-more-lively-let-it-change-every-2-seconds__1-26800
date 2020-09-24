Attribute VB_Name = "ModReg"
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal HKEY As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal HKEY As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1

Private strLocation As String
Private strKey As String
Private strSubKey As String
Private strValue As String
Private KeyHand As Long
Private lResult As Long
Private lValueType As Long
Private strBuf As String
Private lDataBufSize As Long
Private intZeroPos As Integer

'**I didn't write this code, it found it on PSC**

Private Function RegQueryStringValue(ByVal HKEY As Long, ByVal strValueName As String)

On Error GoTo 0
lResult = RegQueryValueEx(HKEY, strValueName, 0&, lValueType, ByVal 0&, lDataBufSize)
If lResult = ERROR_SUCCESS Then
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(HKEY, strValueName, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            RegQueryStringValue = StripTerminator(strBuf)
        End If
    End If
End If

End Function

Private Function GetSettingEx(HKEY As Long, sPath As String, sValue As String)
    
Call RegOpenKey(HKEY, sPath, KeyHand&)
GetSettingEx = RegQueryStringValue(KeyHand&, sValue)
Call RegCloseKey(KeyHand&)

End Function

Private Function StripTerminator(ByVal strString As String) As String

intZeroPos = InStr(strString, Chr$(0))
If intZeroPos > 0 Then
    StripTerminator = Mid(strString, 1, intZeroPos - 1)
Else
    StripTerminator = strString
End If

End Function

Private Sub SaveSettingEx(HKEY As Long, sPath As String, sValue As String, sData As String)

Call RegCreateKey(HKEY, sPath, KeyHand)
Call RegSetValueEx(KeyHand&, sValue, 0, REG_SZ, ByVal sData, Len(sData))
Call RegCloseKey(KeyHand&)

End Sub

Public Sub RegPut(strLocation As String, strKey As String, strValue As String)

Call SaveSettingEx(HKEY_CURRENT_USER, strLocation, strKey, strValue)

End Sub

Public Function RegGet(strLocation As String, strKey As String)

RegGet = GetSettingEx(HKEY_CURRENT_USER, strLocation, strKey)

End Function

Public Sub DelRegKey(strKey, strSubKey)

Call RegDeleteKey(strKey, strSubKey)

End Sub
