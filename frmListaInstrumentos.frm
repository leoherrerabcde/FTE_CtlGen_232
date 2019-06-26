VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListaInstrumentos 
   BorderStyle     =   0  'None
   Caption         =   "Listado de Instrumentos y Componentes"
   ClientHeight    =   6240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10515
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameListadoInstrumentos 
      Caption         =   "Listado de Instrumentos y Componentes"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdNewInstrument 
         Caption         =   "&Nuevo"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   5880
         TabIndex        =   2
         Top             =   4080
         Width           =   1335
      End
      Begin MSComctlLib.ListView LstVwInstrumentos 
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
End
Attribute VB_Name = "frmListaInstrumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PV_Column_Header()      As String

Sub Habilitar_Commands()

End Sub

Sub Deshabilitar_Commands()

End Sub

Sub Verificar_Habilitacion_Commands()

    If GV_Flag_Add_Medicion = False _
       And GV_Flag_Add_Exita = False _
       And GV_Flag_Add_Prueba = False _
       And GV_Flag_Add_Accesorio = False Then
    
        Me.cmdAgregar.Enabled = False

    Else
        
        Me.cmdAgregar.Enabled = True
        
    End If
    
End Sub

Sub SetWindowsState()

    With Me
        If .WindowState <> vbMaximized Then
            .WindowState = vbMaximized
        End If
    End With
    
End Sub

Sub Refresh_Lista_Instrumentos()

    With Me
        .LstVwInstrumentos.ListItems.Clear
        
        BD_Fill_LstVw_Instruments .LstVwInstrumentos
        If .LstVwInstrumentos.ListItems.count Then
            .LstVwInstrumentos.ListItems(1).EnsureVisible
        Else
        End If
        .LstVwInstrumentos.Refresh
        '.LstVwInstrumentos.SetFocus
    End With
    
End Sub

Private Sub cmdAgregar_Click()

Dim i                   As Long
Dim Index               As Integer
Dim LV_Flag_Add         As Boolean
Dim LV_Checked_Count    As Integer

    LV_Checked_Count = Contar_Items_Chequeados(Me.LstVwInstrumentos)
    
    If GV_Flag_Add_Prueba = True Then
        If LV_Checked_Count > 1 Then
            MsgBox "Seleccione sólo un Dispositivo", vbOKOnly
            Exit Sub
        End If
    End If
    
    With Me.LstVwInstrumentos
        For i = 1 To .ListItems.count
        
            If LVItemChecked(Me.LstVwInstrumentos, i) = True Then
                BD_Agregar_Instru_To_Project GV_Actual_Project.Cod_Project, .ListItems(i).Tag
            Else
            End If
        Next
    End With
    
    Unload Me
    
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
    
End Sub

Private Sub cmdModificar_Click()

Dim Index           As Integer
    
    With Me
        Index = LstVwFindItemChecked(.LstVwInstrumentos)
        If Index Then
            GV_Cod_Instru = Me.LstVwInstrumentos.ListItems(Index).Tag
        Else
            GV_Cod_Instru = 0
        End If
        .cmdNewInstrument.value = True
    End With
    
End Sub

Private Sub cmdNewInstrument_Click()


    Dim frmNew      As frmNewInstrument

    Set frmNew = New frmNewInstrument

    frmNew.WindowState = vbMaximized
    frmNew.Show
    
    frmNew.cboInfoDispo(1).Enabled = True

End Sub

Private Sub Form_GotFocus()

    Me.SetWindowsState
    Me.Refresh_Lista_Instrumentos
    
End Sub

Private Sub Form_Load()

    Set_Column_Lista_Instrumentos Me.LstVwInstrumentos, PV_Column_Header
    
'    ReDim PV_Column_Header(7)
'
'    PV_Column_Header(0) = "Dispositivo"
'    PV_Column_Header(1) = "Funcion"
'    PV_Column_Header(2) = "Nombre"
'    PV_Column_Header(3) = "Fabricante"
'    PV_Column_Header(4) = "Modelo"
'    PV_Column_Header(5) = "Numero de Parte"
'    PV_Column_Header(6) = "Numero de Serie"
'    PV_Column_Header(7) = "Comunicacion"
'
'    AddColumListView Me.LstVwInstrumentos, PV_Column_Header
    
    Refresh_Lista_Instrumentos
    
    'Call LVSetStyleEx(Me.LstVwInstrumentos, FullRowSelect, True)
    'Call LVSetStyleEx(Me.LstVwInstrumentos, GridLines, True)
    'Call LVSetStyleEx(Me.LstVwInstrumentos, Checkboxes, True)
    
    'Me.SetFocus
    Me.Verificar_Habilitacion_Commands
    
    Me.Show

End Sub

Private Sub Form_Resize()

    With Me
        
        .frameListadoInstrumentos.Width = .Width - 2 * .frameListadoInstrumentos.Left
        .frameListadoInstrumentos.Height = .Height - 2 * .frameListadoInstrumentos.Top
        
        .LstVwInstrumentos.Width = .frameListadoInstrumentos.Width _
                                   - 2 * .LstVwInstrumentos.Left
        
        .cmdAgregar.Left = .LstVwInstrumentos.Width _
                           + .LstVwInstrumentos.Left _
                           - .cmdCancelar.Left _
                           - .cmdAgregar.Width
        .cmdAgregar.Top = .frameListadoInstrumentos.Height _
                          - .cmdAgregar.Height _
                          - .LstVwInstrumentos.Top
                          
        .cmdModificar.Top = .cmdAgregar.Top
        .cmdCancelar.Top = .cmdAgregar.Top
        .cmdNewInstrument.Top = .cmdAgregar.Top
        
        .cmdNewInstrument.Left = (.cmdAgregar.Left + .cmdAgregar.Width - .cmdCancelar.Left) / 3
        .cmdModificar.Left = 2 * .cmdNewInstrument.Left
        
        .cmdNewInstrument.Left = .cmdNewInstrument.Left - .cmdNewInstrument.Width / 2
        .cmdModificar.Left = .cmdModificar.Left - .cmdNewInstrument.Width / 2
        
        
        .LstVwInstrumentos.Height = .cmdAgregar.Top _
                                   - 2 * .LstVwInstrumentos.Top
                                   
    End With
    
End Sub

