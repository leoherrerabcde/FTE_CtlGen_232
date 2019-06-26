Attribute VB_Name = "Procedimiento_ConfigRegional"
Option Explicit

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
'Declare Function RegGetValue Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long



'Declare Function RegCloseKey& Lib "advapi32.dll" (ByVal hKey&)
Declare Function RegOpenKeyExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszSubKey$, dwOptions&, ByVal samDesired&, lpHKey&)
Declare Function RegQueryValueExA& Lib "advapi32.dll" (ByVal hKey&, ByVal lpszValueName$, ByVal lpdwRes&, lpdwType&, ByVal lpDataBuff$, nSize&)

'Const HKEY_CLASSES_ROOT = &H80000000
'Const HKEY_CURRENT_USER = &H80000001
'Const HKEY_LOCAL_MACHINE = &H80000002
'Const HKEY_USERS = &H80000003

Const ERROR_SUCCESS = 0&
'Const REG_SZ = 1&                          ' Unicode nul terminated string
'Const REG_DWORD = 4&                       ' 32-bit number

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ


'****************************************************************************************************************
'Creado por          : http://www.arcatapet.net/vbregget.cfm
'Fecha Creación      : 22/09/2004
'Descripción         : Lee la clave del registro de Windows.
'Fecha Modificación  :
'Modificado por      :
'Descripción         :
'****************************************************************************************************************
Function RegGetValue$(MainKey&, SubKey$, value$)
   
    ' MainKey must be one of the Publicly declared HKEY constants.
    Dim sKeyType&       'to return the key type.  This function expects REG_SZ or REG_DWORD
    Dim ret&            'returned by registry functions, should be 0&
    Dim lpHKey&         'return handle to opened key
    Dim lpcbData&       'length of data in returned string
    Dim ReturnedString$ 'returned string value
    Dim ReturnedLong    As Long
    Dim fTempDbl!
    
    If MainKey >= &H80000000 And MainKey <= &H80000006 Then
        ' Open key
        ret = RegOpenKeyExA(MainKey, SubKey, 0&, KEY_READ, lpHKey)
        
        If ret <> ERROR_SUCCESS Then
           RegGetValue = ""
           Exit Function     'No key open, so leave
        End If
        
        ' Set up buffer for data to be returned in.
        ' Adjust next value for larger buffers.
        lpcbData = 255
        ReturnedString = Space$(lpcbData)
        
        ' Read key
        ret& = RegQueryValueExA(lpHKey, value, ByVal 0&, sKeyType, ReturnedString, lpcbData)
        If ret <> ERROR_SUCCESS Then
            RegGetValue = ""   'Value probably doesn't exist
        Else
            If sKeyType = REG_DWORD Then
                ret = RegQueryValueExA(lpHKey, value, ByVal 0&, sKeyType, ReturnedLong, 4)
                If ret = ERROR_SUCCESS Then RegGetValue = CStr(ReturnedLong)
            Else
                RegGetValue = Left$(ReturnedString, lpcbData - 1)
            End If
        End If
        ' Always close opened keys.
        ret = RegCloseKey(lpHKey)
    End If
   
End Function



'****************************************************************************************************************
'****************************************************************************************************************
Public Sub savestring(hKey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand                     As Long
Dim r                           As Long
Dim lvValue
Dim lv_Guardar_Config_Reg       As Integer

    lvValue = RegGetValue$(hKey, strPath, strValue)
    
    lv_Guardar_Config_Reg = GetSetting(App.Title, "Control Panel", "Guardar Control Panel", 1)
    SaveSetting App.Title, "Control Panel", "Guardar Control Panel", lv_Guardar_Config_Reg
    If lv_Guardar_Config_Reg Then
        SaveSetting App.Title, "Control Panel" & "\" & strPath, strValue, lvValue
    End If
    
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)

End Sub


Sub RestaurarRegistro(hKey As Long, strPath As String, strValue As String, strdata As String)

Dim keyhand                     As Long
Dim r                           As Long
Dim lvValue

    
    strdata = GetSetting(App.Title, "Control Panel" & "\" & strPath, strValue, strdata)
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)

End Sub
