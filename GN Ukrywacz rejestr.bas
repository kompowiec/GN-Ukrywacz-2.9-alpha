Attribute VB_Name = "Rejestr"
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&

Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_DWORD = 4 ' 32-bit number

Public Function GetString(Hkey As Long, strPath As String, strValue As String) As String
   Dim keyhand As Long
   Dim datatype As Long
   Dim lResult As Long
   Dim strBuf As String
   Dim lDataBufSize As Long
   Dim intZeroPos As Integer
   r = RegOpenKey(Hkey, strPath, keyhand)
   lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
   If lValueType = REG_SZ Then
       strBuf = String(lDataBufSize, " ")
       lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
       If lResult = ERROR_SUCCESS Then
           intZeroPos = InStr(strBuf, Chr$(0))
           If intZeroPos > 0 Then
               GetString = Left$(strBuf, intZeroPos - 1)
           Else
               GetString = strBuf
           End If
       End If
   End If
End Function

Public Sub SaveString(Hkey As Long, strPath As String, strValue As String, strdata As String)
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
End Sub
