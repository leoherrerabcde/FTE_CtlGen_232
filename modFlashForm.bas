Attribute VB_Name = "modFlashForm"
Option Explicit


'Declaraciones de la API necesarias
Private Declare Function FlashWindow Lib "user32" (ByVal hWnd As Long, ByVal bInvert As Long) As Long
Private Declare Function FlashWindowEx Lib "user32" (FWInfo As FLASHWINFO) As Boolean

Public Enum FlashWindowFlags
vbOnlyTitle = 1
vbOnlyBar = 2
vbTitleAndBar = 3
End Enum

Private Type FLASHWINFO
  cbSize As Long
  hWnd As Long
  dwFlags As Long
  uCount As Long
  dwTimeout As Long
End Type

Private Const FLASHW_TRAY = 2

'Funciones necesarias para saber si una funcion API está presente
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Sub Flash(Form As Form, Optional NumberOfFlashes As Integer = 1, Optional Flags As FlashWindowFlags = 3, Optional Timeout As Long = 500)
       
'Notas: W98 o superior necesario

'Prevenimos errores mirando si la función
'Está activa en el SO actual
If Not APIFunctionPresent("FlashWindowEx", "user32") Then
'Si no esta activa llamamos a una función mas sencilla
'El problema es que con esta funcion
'Solo podemos decirle que haga un flash al "Title" y a la "Bar"
    Call FlashWindow(Form.hWnd, True)
    Exit Sub
End If

Dim bRet As Boolean
Dim udtFWInfo As FLASHWINFO

With udtFWInfo
   .cbSize = 20
   .hWnd = Form.hWnd
   .dwFlags = Flags
   .uCount = NumberOfFlashes
   .dwTimeout = Timeout
End With

bRet = FlashWindowEx(udtFWInfo)
End Sub

'Función para saber si una función de la API está disponible
Private Function APIFunctionPresent(ByVal FunctionName As String, ByVal DllName As String) As Boolean
    Dim lHandle As Long, lAddr  As Long
    lHandle = LoadLibrary(DllName)
    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        FreeLibrary lHandle
    End If
    APIFunctionPresent = (lAddr <> 0)
End Function

