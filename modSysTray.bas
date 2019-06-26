Attribute VB_Name = "modSysTray"
Option Explicit
' Esta es la Estructura que necesita InitCommonControlsEx

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub btnCambiar_Click()
    
    ' Mensajillo Personalizado
'    Notify.dwInfoFlags = NIIF_WARNING
'    Notify.szInfoTitle = Me.txtTitulo.Text + Chr$(0)
'    Notify.szInfo = "Usted lo que escribio es :" + vbCr + Me.txtMensaje.Text + Chr$(0)
'    ' llamamos a NIM_MODIFY para mostrar de nuevo el ballon
'    Shell_NotifyIcon NIM_MODIFY, Notify
End Sub

Private Sub Close_Click()
'    Unload Me
End Sub

'Private Sub Form_Initialize()
''    Ini.dwSize = Len(Ini)
''    Ini.dwICC = ICC_COOL_CLASSES
''    ' Verifica si se inicializan correctamente los controles
''    If Not InitCommonControlsEx(Ini) Then
''        MsgBox "no se inicializo", vbCritical, "Error al inicializarse"
''    End If
'End Sub

'Private Sub Form_Load()
'    Me.Hide                                                                             ' Oculto el Form
'    Notify.cbSize = Len(Notify)                                                                 ' Tamaño de la estructura
'    Notify.hIcon = Me.Icon                      ' Notify mostrado en la barra
'    Notify.hwnd = Me.hwnd                       ' Ventana que manipula el proceso
'    Notify.uCallbackMessage = WM_MOUSEMOVE                      ' Procedimiento que maneja los eventos
'    Notify.szTip = "Notify con Ballon tool tip" & Chr$(0)       ' tool tip clasico
'    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_TIP ' los eventos que pueden hacerse
'    ' Mensaje que se mostrara en el ballon tool tip
'    Notify.szInfo = "Esto solo es una prueba" + vbCr + "Aprete aqui para ver el" + vbCr + "El formulario" + Chr$(0)
'    ' Titulo del ballon tool tip
'    Notify.szInfoTitle = "Prueba" & Chr$(0)
'    ' Tiempo en milisegundos (Aunque no responde)
'    Notify.uTimeout = 10 'Or NOTIFYICON_VERSION
'    ' Hacer que se muestre el ballon tool tip al crearse
'    Notify.dwInfoFlags = NIIF_INFO
'    'Notify.uVersion = NOTIFYICON_VERSION (Si es que se quiere saber la version del Notify)
'    Notify.uID = 1& ' un identificador del Notify
'    Shell_NotifyIcon NIM_ADD, Notify ' llamamos a la funcion para añadirlo
'End Sub

'Private Sub Form_Unload(Cancel As Integer)
'    ' al cerrar quitamos el Notyfi
'    Shell_NotifyIcon NIM_DELETE, Notify
'End Sub

'Private Sub form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
'    Static rec As Boolean
'    Dim msg As Long
'    msg = x / Screen.TwipsPerPixelX
'    If rec = False Then
'        rec = True
'        Select Case msg
'            Case WM_LBUTTONDBLCLK:                      ' doble click con el boton izquierdo del raton
'                Me.Show                                                                 ' mostramos el formulario
'            Case WM_RBUTTONUP:
'                Me.PopupMenu Menu                       ' click con el boton secundario, mostramos el menu correspondiente
'            Case NIN_BALLOONUSERCLICK:  'Click al ballon Tool Tip
'                MsgBox "hizo click al ballon", vbExclamation, "Mensaje"
'                Me.Show
'        End Select
'        rec = False
'    End If
'End Sub
'
'Private Sub Show_Click()
'    Me.Show
'End Sub



Sub Form_On_Top(hwnd As Long, lv_OnTop As Boolean)
   
   If lv_OnTop = True Then
      SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
   Else
      SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
   End If

End Sub


