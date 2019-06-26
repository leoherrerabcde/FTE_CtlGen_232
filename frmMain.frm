VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.MDIForm MdiMain 
   BackColor       =   &H8000000C&
   Caption         =   "Medición de Componentes de RF Automatizadas - MCRFA"
   ClientHeight    =   5805
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11595
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLevantarMdi 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   3600
   End
   Begin VB.Timer tmrDataInstrument 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4560
      Top             =   2880
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   3120
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "1035"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "1036"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "1037"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "1038"
            ImageKey        =   "Print"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "1039"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "1040"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "1041"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Bold"
            Object.ToolTipText     =   "1042"
            ImageKey        =   "Bold"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Italic"
            Object.ToolTipText     =   "1043"
            ImageKey        =   "Italic"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Underline"
            Object.ToolTipText     =   "1044"
            ImageKey        =   "Underline"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Left"
            Object.ToolTipText     =   "1045"
            ImageKey        =   "Align Left"
            Style           =   2
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Center"
            Object.ToolTipText     =   "1046"
            ImageKey        =   "Center"
            Style           =   2
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Align Right"
            Object.ToolTipText     =   "1047"
            ImageKey        =   "Align Right"
            Style           =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5535
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Status"
            TextSave        =   "Status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "08/07/2017"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "17:04"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Duración:"
            TextSave        =   "Duración:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Tpo Restante"
            TextSave        =   "Tpo Restante"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Total de Medidas"
            TextSave        =   "Total de Medidas"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Medida "
            TextSave        =   "Medida "
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2725
            Text            =   "Medidas Restantes :"
            TextSave        =   "Medidas Restantes :"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   5040
      Top             =   1305
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4920
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":041C
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":052E
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0752
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0864
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0976
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0A88
            Key             =   "Bold"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B9A
            Key             =   "Italic"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CAC
            Key             =   "Underline"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0DBE
            Key             =   "Align Left"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0ED0
            Key             =   "Center"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FE2
            Key             =   "Align Right"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "1000"
      Begin VB.Menu mnuFileNew 
         Caption         =   "1001"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "1002"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "1003"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "1004"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "1005"
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "1006"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "1007"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "1008"
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "1009"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "1010"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "1011"
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileMRU 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "1012"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "1013"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "1014"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "1015"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "1016"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "1017"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "1018"
      End
   End
   Begin VB.Menu mnuProyect 
      Caption         =   "1078"
      Begin VB.Menu mnuNewProject 
         Caption         =   "1079"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpenProyect 
         Caption         =   "1080"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuInstrumentos 
      Caption         =   "1081"
      Begin VB.Menu mnuAgregarNewInstrument 
         Caption         =   "1082"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "1019"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "1020"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "1021"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "1022"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "1023"
      End
      Begin VB.Menu mnuViewWebBrowser 
         Caption         =   "1024"
      End
      Begin VB.Menu mnuConfigurarCOM 
         Caption         =   "Configurar COM"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "1025"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "1026"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "1027"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "1028"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "1029"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "1030"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "1031"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "1032"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "1033"
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "1034"
      End
   End
   Begin VB.Menu mnuCerrar 
      Caption         =   "Cerrar"
      Begin VB.Menu mnuCerrarApp 
         Caption         =   "&Cerrar"
      End
      Begin VB.Menu mnuContinuar 
         Caption         =   "Con&tinuar"
      End
   End
End
Attribute VB_Name = "MdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


  ' the subclass procedure is in the cSubclass class module
  Private WithEvents m_oSubclass As cSubclass
Attribute m_oSubclass.VB_VarHelpID = -1
  

Private Type tagINITCOMMONCONTROLSEX
        dwSize As Long
        dwICC As Long
End Type
' Aqui estan los tipos de inicializacion de temas
Private Const ICC_LISTVIEW_CLASSES = &H1          ' listview, header
Private Const ICC_TREEVIEW_CLASSES = &H2          ' treeview, tooltips
Private Const ICC_BAR_CLASSES = &H4               ' toolbar, statusbar, trackbar, tooltips
Private Const ICC_TAB_CLASSES = &H8               ' tab, tooltips
Private Const ICC_UPDOWN_CLASS = &H10             ' updown
Private Const ICC_PROGRESS_CLASS = &H20           ' progress
Private Const ICC_HOTKEY_CLASS = &H40             ' hotkey
Private Const ICC_ANIMATE_CLASS = &H80            ' animate
Private Const ICC_WIN95_CLASSES = &HFF
Private Const ICC_DATE_CLASSES = &H100            ' month picker, date picker, time picker, updown
Private Const ICC_USEREX_CLASSES = &H200          ' comboex
Private Const ICC_COOL_CLASSES = &H400            ' rebar (coolbar) control


'' Mostrar el formulario indicado, dentro de picDock
'Private Sub mostrarForm(ByVal formhWnd As Long, Optional ByVal ajustar As Boolean = True)
'    ' Hacer el formulario indicado, un hijo del picDock
'    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
'    ' si Ajustar es False, se quedará con el tamaño actual.
'    Call SetParent(formhWnd, picDock.hWnd)
'    posicionarForm formhWnd, ajustar
'    Call ShowWindow(formhWnd, NORMAL_eSW)
'End Sub
'
'' Posicionar el formulario indicado dentro de picDock
'Private Sub posicionarForm(ByVal formhWnd As Long, Optional ByVal ajustar As Boolean = True)
'    ' Posicionar el formulario indicado en las coordenadas del picDock
'    ' Si Ajustar es True, se ajustará al tamaño del contenedor,
'    ' si Ajustar es False, se quedará con el tamaño actual.
'    Dim nWidth As Long, nHeight As Long
'    Dim wndPl As WINDOWPLACEMENT
'    '
'    If ajustar Then
'        nWidth = picDock.ScaleWidth \ Screen.TwipsPerPixelX
'        nHeight = picDock.ScaleHeight \ Screen.TwipsPerPixelY
'    Else
'        ' el tamaño del formulario que se va a posicionar
'        Call GetWindowPlacement(formhWnd, wndPl)
'        With wndPl.rcNormalPosition
'            nWidth = .Right - .Left
'            nHeight = .Bottom - .Top
'        End With
'    End If
'    Call MoveWindow(formhWnd, 0, 0, nWidth, nHeight, True)
'End Sub


Sub Habilitar_Mnu_Instrumentos()

    Me.mnuInstrumentos.Visible = False
    Me.mnuAgregarNewInstrument.Enabled = False

End Sub

'Sub SetNotifyIcon()        '(LV_Title As String, LV_Msg As String)
'
'    Notify.cbSize = Len(Notify)
'    'Notify.dwInfoFlags = NIIF_INFO      ' NIIF_WARNING
'    Notify.szTip = "COM Desconectada." & Chr$(0)       ' tool tip clasico
'    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_TIP  'NIF_ICON Or NIF_TIP Or NIF_MESSAGE
'    Notify.szInfoTitle = App.Title + Chr$(0)
'    Notify.szInfo = "COM desconectada." + vbCr '+ "Doble Click para" + Chr$(0)
'    Notify.hIcon = Me.Icon
'    Notify.hWnd = Me.hWnd
'    Notify.uCallbackMessage = WM_MOUSEMOVE
'    Notify.uID = 1& ' un identificador del Notify
'    ' llamamos a NIM_MODIFY para mostrar de nuevo el ballon
'    Shell_NotifyIcon NIM_ADD, Notify
'
'End Sub

Function IniciarCommPort() As Boolean

    With Me
        IniciarCommPort = False
        If .MSComm1.PortOpen = False Then
            On Error Resume Next
            '.MSComm.CommPort = 2
            .MSComm1.PortOpen = True
            If Err Then
                MsgBox Error$, vbOKOnly
                IniciarCommPort = False
            Else
                IniciarCommPort = True
            End If
        Else
            IniciarCommPort = True
        End If
    End With

End Function

Function Refresca_Contador_Tpo(LV_Tpo_Ini As Single, LV_Ptos_Total As Long, LV_Ptos As Long)

Dim LV_Tpo_Enlazado     As Single
Dim LV_Tpo_Rest         As Single

    LV_Tpo_Enlazado = Timer - LV_Tpo_Ini
    LV_Tpo_Rest = LV_Tpo_Enlazado
    If LV_Ptos And LV_Ptos_Total Then
        LV_Tpo_Rest = LV_Tpo_Rest / LV_Ptos * LV_Ptos_Total
    Else
        LV_Tpo_Rest = -1
    End If
    LV_Tpo_Rest = LV_Tpo_Rest - LV_Tpo_Enlazado
    
    With Me.sbStatusBar
        .Panels(4).Text = "Transcurrido: " & Format(LV_Tpo_Enlazado / 24 / 3600, "hh:mm:ss")
        .Panels(5).Text = "Tpo Restante: " & Format(LV_Tpo_Rest / 24 / 3600, "hh:mm:ss")
        .Panels(6).Text = "Total Mediciones: " & LV_Ptos_Total
        .Panels(7).Text = "Medición: " & LV_Ptos
        .Panels(8).Text = "Mediciones Restantes: " & LV_Ptos_Total - LV_Ptos
    End With
    
End Function

Sub Levantar_Menu_Cerrar()

    Me.PopupMenu mnuCerrar

End Sub
'Sub Running_in_Background()
'
'    With Me
'        .Hide
'        SetNotifyIcon
'        'Load frmCtlGenMin
'        'frmCtlGenMin.Show vbModal
'    End With
'
'End Sub

Function SendRS232(LV_Cmd As String)

    With Me.MSComm1
        GV_Data_Instrument_Ok = False
        GV_Data_Instrument = ""
        With Me.tmrDataInstrument
            .Enabled = False
            .Interval = 400
            .Enabled = True
        End With
        If .PortOpen = True Then
            .Output = vbCrLf & LV_Cmd
        End If
    End With
    
End Function

Sub Cerrar_Formularios()

    Dim tForm As Form
    Dim s As String
    '
    s = Me.Name
    For Each tForm In Forms
        'tForm.ActiveControl
        If tForm.Name <> s And _
           tForm.Name <> frmToolBoxBottom.Name And _
           tForm.Name <> frmToolBoxLeft.Name And _
           tForm.Name <> frmToolBoxRight.Name _
           Then
            Unload tForm
        End If
    Next

End Sub

Sub CloseMsComm()

    If Me.MSComm1.PortOpen = True Then
        Me.MSComm1.PortOpen = False
    End If
    
End Sub

Sub Deshabilitar_Mnu_Instrumentos()

    Me.mnuInstrumentos.Visible = False
    Me.mnuAgregarNewInstrument.Enabled = False

End Sub

Sub LoadFormNewInstrument()

'Dim frmD                    As frmNewInstrument
'Dim LV_Cod_Instrument       As Integer
'Dim LV_Caption              As String
'
'    Set frmD = New frmNewInstrument
'
'    'frmD.Caption = LV_Caption
'    frmD.WindowState = vbMaximized
'    frmD.Show

End Sub

Sub CloseFormCtlGen()

'Dim frmD As frmCtlGen
Dim tForm               As Form
    'Set frmD = New frmCtlGen
    
    For Each tForm In Forms
        'tForm.ActiveControl
        If tForm.Name = "frmCtlGen" Then
            Unload tForm
        End If
    Next
    
'    If Not (ActiveForm Is Nothing) Then
'        With Me.ActiveForm
'            If .Name = "frmDialogNewProject" Then
'                ActiveForm.GotFocus
'            End If
'        End With
'    End If

End Sub

Sub LoadFormAbrirCtlGen()

    Dim frmD As frmCtlGen
    
    Set frmD = New frmCtlGen
    
    frmD.WindowState = vbMaximized
    
    frmD.Show
    
    Deshabilitar_Mnu_Instrumentos

End Sub

Sub LoadFormAbrirProyectos()

'    Dim frmD As frmAbrirProyecto
'
'    'lDocumentCount = lDocumentCount + 1
'
'    Set frmD = New frmAbrirProyecto
'
'    frmD.WindowState = vbMaximized
'
'    frmD.Show
'
'    Deshabilitar_Mnu_Instrumentos

End Sub

Sub LoadFormListaInstrumentos()

'Dim frmD                    As frmListaInstrumentos
'Dim LV_Cod_Instrument       As Integer
'Dim LV_Caption              As String
'
'    'LV_Caption = Me.ActiveForm.Caption
'
'    Set frmD = New frmListaInstrumentos
'
'    'frmD.Caption = LV_Caption
'    frmD.WindowState = vbMaximized
'    frmD.Show
'
'    frmD.SetFocus
'
''    LV_Cod_Instrument = ConvertLongToInt(lParam)
''
''    Select Case wParam
''
''        Case PARAM_NEW_INSTRUMENT
''
''        Case PARAM_MODIFI_INSTRUMENT
''
''            frmD.Leer_Instrumento_From_BD GV_Actual_Project.Cod_Project, LV_Cod_Instrument
''
''    End Select
    
End Sub

Sub LoadFormNewProject()

'    Dim frmD As frmDialogNewProject
'
'    Set frmD = New frmDialogNewProject
'
'    frmD.Show vbModal
    
End Sub

Sub LoadFormProject()

    'Static lDocumentCount As Long
    
    Dim frmD As frmTabsProject
    
    'lDocumentCount = lDocumentCount + 1
    
    Set frmD = New frmTabsProject
    
    'frmD.Caption = "Proyecto " & lDocumentCount
    
    frmD.WindowState = vbMaximized
    
    frmD.Show
    
    Deshabilitar_Mnu_Instrumentos


End Sub

Sub LoadFormRs232Props()

    Dim frmD As frmProps
    
    Set frmD = New frmProps
    
    frmD.Show vbModal

End Sub


'Sub Cargar_ToolBox_Left()
'
'    'dockForm frmToolBoxLeft.hwnd, Me.PictureLeft, True
'
'    With frmToolBoxLeft
'        MakeRoundRect .hwnd, _
'                      .Width \ Screen.TwipsPerPixelX, _
'                    .Height \ Screen.TwipsPerPixelY, _
'                     25
'
'    End With
'
'End Sub
'
'
'Sub Cargar_ToolBox_Right()
'
'    dockForm frmToolBoxRight.hwnd, Me.PictureRight, True
'
'    With frmToolBoxRight
'        MakeRoundRect .hwnd, _
'                      .Width \ Screen.TwipsPerPixelX, _
'                    .Height \ Screen.TwipsPerPixelY, _
'                     25
'
'    End With
'
'End Sub
'
'
'Sub Cargar_ToolBox_Bottom()
'
'    dockForm frmToolBoxBottom.hwnd, Me.PictureBottom, True
'
'    With frmToolBoxBottom
'        MakeRoundRect .hwnd, _
'                      .Width \ Screen.TwipsPerPixelX, _
'                    .Height \ Screen.TwipsPerPixelY, _
'                     25
'
'    End With
'
'End Sub




Private Sub MDIForm_Activate()

    If Not (ActiveForm Is Nothing) Then
        With Me.ActiveForm
            If .Name = "frmDialogNewProject" Then
                ActiveForm.GotFocus
            End If
        End With
    End If
    
End Sub

Private Sub MDIForm_Load()
    
    GV_Actual_Project.Cod_Project = 0
    
    LoadResStrings Me
    Me.Hide
    'Me.WindowState = vbMaximized
    
'    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
'    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
'    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
'    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    
    
    Load_RS232_Param
    
    'LoadFormAbrirCtlGen
    
    
End Sub


Private Sub LoadNewDoc()
    
'    Static lDocumentCount As Long
'
'    Dim frmD As frmDocument
'
'    lDocumentCount = lDocumentCount + 1
'
'    Set frmD = New frmDocument
'
'    frmD.Caption = "Document " & lDocumentCount
'
'    frmD.Show

End Sub

Sub Load_RS232_Param()

Dim Settings            As String
Dim CommPort            As Integer
Dim Handshaking         As Integer
Dim Echo                As String

    With Me
        On Error Resume Next
        ' Carga la configuración del registro
        Settings = GetSetting(App.Title, "Properties", "Settings", "") ' frmTerminal.MSComm1.Settings]\
        If Settings <> "" Then
            MSComm1.Settings = Settings
            If Err Then
                MsgBox Error$, 48
                On Error GoTo 0
                Exit Sub
            End If
        End If
        
        CommPort = GetSetting(App.Title, "Properties", "CommPort", "0") ' frmTerminal.MSComm1.CommPort
        If CommPort <> 0 Then
            MSComm1.CommPort = CommPort
        End If
        
        Handshaking = GetSetting(App.Title, "Properties", "Handshaking", MSComm1.Handshaking)  'frmTerminal.MSComm1.Handshaking
        'If Handshaking <> "" Then
            MSComm1.Handshaking = Handshaking
            If Err Then
                MsgBox Error$, 48
                On Error GoTo 0
                Exit Sub
            End If
        'End If
        
        Echo = GetSetting(App.Title, "Properties", "Echo", "") ' Echo
        On Error GoTo 0

    End With
    
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Static rec As Boolean
'Dim Msg As Long
'
'    Msg = X / Screen.TwipsPerPixelX
'    If rec = False Then
'        rec = True
'        Select Case Msg
'            Case WM_LBUTTONDBLCLK:                      ' doble click con el boton izquierdo del raton
'                IniciarCommPort
'                frmCtlGenMin.Show                                                                 ' mostramos el formulario
'                Shell_NotifyIcon NIM_DELETE, Notify
'            Case WM_RBUTTONUP:
'                Me.PopupMenu mnuCerrar                       ' click con el boton secundario, mostramos el menu correspondiente
'            Case NIN_BALLOONUSERCLICK:  'Click al ballon Tool Tip
'                'MsgBox "hizo click al ballon", vbExclamation, "Mensaje"
'                'Me.Show
'        End Select
'        rec = False
'    End If
'
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim tForm As Form
    Dim s As String
    '
    s = Me.Name
    For Each tForm In Forms
        'tForm.ActiveControl
        If tForm.Name <> s Then
            Unload tForm
        End If
    Next

End Sub

Private Sub MDIForm_Resize()

'    With Me
'        If .WindowState = vbMinimized Then
'            .Hide
'            SetNotifyIcon
'            Load frmCtlGenMin
'            frmCtlGenMin.Show vbModal
'        End If
'    End With
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
  ' destroy the object so we don't crash since the
  ' subclass is terminated in the Class_Terminate event
  Set m_oSubclass = Nothing
  
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

Private Sub m_oSubclass_ProjectChange(wParam As Long, lParam As Long)

'    Select Case wParam
'        Case Project_Name
'            Me.ActiveForm.Caption = frmNewProject.Caption
'    End Select
    
End Sub

Private Sub m_oSubclass_InstrumentMessage(wParam As Long, lParam As Long)

'Static lDocumentCount As Long

'Dim frmD                    As frmNewInstrument
'Dim LV_Cod_Instrument       As Integer
'Dim LV_Caption              As String
'
'    LV_Caption = Me.ActiveForm.Caption
'
'    Set frmD = New frmNewInstrument
'
'    frmD.Caption = LV_Caption
'    frmD.WindowState = vbMaximized
'    frmD.Show
'    LV_Cod_Instrument = ConvertLongToInt(lParam)
'
'    Select Case wParam
'
'        Case PARAM_NEW_INSTRUMENT
'
'        Case PARAM_MODIFI_INSTRUMENT
'
'            frmD.Leer_Instrumento_From_BD GV_Actual_Project.Cod_Project, LV_Cod_Instrument
'
'    End Select
    
End Sub

Private Sub m_oSubclass_CustomMessage()
  ' event raised from the class object when a custom message is recieved

  Beep
  
  MsgBox "Custom message recieved at " & Format$(Now, "hh:mm:ss") & " !"

End Sub

Private Sub m_oSubclass_FormMove()
  ' event raised from the class object when the form is moved
  
  'MsgBox "You moved the window"
  
End Sub

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Picture1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)

End Sub

Private Sub mnuAgregarNewInstrument_Click()

    LoadFormListaInstrumentos
    
    ActiveForm.cmdNewInstrument.value = True

End Sub

Private Sub mnuCerrarApp_Click()

    Unload Me
    End
    
End Sub

Private Sub mnuConfigurarCOM_Click()

    LoadFormRs232Props
    
End Sub

Private Sub mnuNewProject_Click()

    Cerrar_Formularios
    
    LoadFormNewProject
    
    Me.Deshabilitar_Mnu_Instrumentos
        
End Sub

Private Sub mnuOpenProyect_Click()

    Cerrar_Formularios
    
    LoadFormAbrirProyectos
    
End Sub

Private Sub MSComm1_OnComm()

    With Me
        Select Case Me.MSComm1.CommEvent
            Case Is = comEvReceive
                GV_Data_Instrument = GV_Data_Instrument & Me.MSComm1.Input
                GV_fMainForm.ShowData GV_Data_Instrument
                With Me.tmrDataInstrument
                    .Enabled = False
                    .Interval = 100
                    .Enabled = True
                End With
        End Select
    End With
    
End Sub

Private Sub PictureBottom_Resize()

    'Cargar_ToolBox_Bottom
    
End Sub

Private Sub PictureLeft_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'    With Me.PictureLeft
'
'        If x > .Width - 25 Then
'
'            x = x
'
'
'
'        End If
'
'    End With
    
End Sub

Private Sub PictureLeft_Resize()

    'Cargar_ToolBox_Left
    
End Sub

Private Sub PictureRight_Resize()

    'Cargar_ToolBox_Right

End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Bold"
            
            LoadFormNewProject
            ActiveForm.SSTabProject.Tab = 5
            'frmPrueba.cmdComenzar.value = True
            
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hWnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub

Private Sub mnuWindowNewWindow_Click()
    LoadNewDoc
End Sub

Private Sub mnuViewWebBrowser_Click()
'    Dim frmB As New frmBrowser
'    frmB.StartingAddress = "http://www.microsoft.com"
'    frmB.Show
End Sub

Private Sub mnuViewOptions_Click()
'    frmOptions.Show vbModal, Me
End Sub

Private Sub mnuViewRefresh_Click()
    'ToDo: Add 'mnuViewRefresh_Click' code.
    MsgBox "Add 'mnuViewRefresh_Click' code."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'ToDo: Add 'mnuEditPasteSpecial_Click' code.
    MsgBox "Add 'mnuEditPasteSpecial_Click' code."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
    'ToDo: Add 'mnuEditUndo_Click' code.
    MsgBox "Add 'mnuEditUndo_Click' code."
End Sub


Private Sub mnuFileExit_Click()
    'unload the form
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'ToDo: Add 'mnuFileSend_Click' code.
    MsgBox "Add 'mnuFileSend_Click' code."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Print"
        .CancelError = True
        .flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .flags = .flags + cdlPDAllPages
        Else
            .flags = .flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'ToDo: Add 'mnuFilePrintPreview_Click' code.
    MsgBox "Add 'mnuFilePrintPreview_Click' code."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Page Setup"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'ToDo: Add 'mnuFileProperties_Click' code.
    MsgBox "Add 'mnuFileProperties_Click' code."
End Sub

Private Sub mnuFileSaveAll_Click()
    'ToDo: Add 'mnuFileSaveAll_Click' code.
    MsgBox "Add 'mnuFileSaveAll_Click' code."
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Save As"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "All Files (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub

Private Sub mnuFileClose_Click()
    'ToDo: Add 'mnuFileClose_Click' code.
    MsgBox "Add 'mnuFileClose_Click' code."
End Sub

Private Sub mnuFileOpen_Click()
    Dim sFile As String


    If ActiveForm Is Nothing Then LoadNewDoc
    

    With dlgCommonDialog
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.rtfText.LoadFile sFile
    ActiveForm.Caption = sFile

End Sub

Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub

Private Sub tmrDataInstrument_Timer()

    Me.tmrDataInstrument.Enabled = False
    GV_Data_Instrument_Ok = True
    
End Sub

Private Sub tmrLevantarMdi_Timer()

    With Me
        .WindowState = vbMaximized
    End With
    
End Sub
