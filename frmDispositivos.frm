VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDispositivos 
   BorderStyle     =   0  'None
   Caption         =   "Dispositivos"
   ClientHeight    =   5775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8715
   LinkTopic       =   "Form2"
   ScaleHeight     =   5775
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameDispositivo 
      Caption         =   "Accesorios"
      Height          =   1335
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   3960
      Width           =   6135
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Index           =   3
         Left            =   4920
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin MSComctlLib.ListView LstViewDispositivo 
         Height          =   615
         Index           =   3
         Left            =   75
         TabIndex        =   19
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dispositivo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Función"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Componente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fabricante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Modelo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Part Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Comunicación"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame frameDispositivo 
      Caption         =   "Instrumentos de Medición"
      Height          =   1335
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Top             =   2640
      Width           =   6135
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   13
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Index           =   2
         Left            =   4920
         TabIndex        =   12
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin MSComctlLib.ListView LstViewDispositivo 
         Height          =   615
         Index           =   2
         Left            =   75
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dispositivo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Función"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Componente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fabricante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Modelo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Part Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Comunicación"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame frameDispositivo 
      Caption         =   "Instrumentos de Generación"
      Height          =   1335
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   6135
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Index           =   1
         Left            =   4920
         TabIndex        =   7
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin MSComctlLib.ListView LstViewDispositivo 
         Height          =   615
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dispositivo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Función"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Componente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fabricante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Modelo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Part Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Comunicación"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame frameDispositivo 
      Caption         =   "Dispositivo en Prueba"
      Height          =   1335
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   975
      End
      Begin MSComctlLib.ListView LstViewDispositivo 
         Height          =   615
         Index           =   0
         Left            =   75
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Dispositivo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Función"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Componente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fabricante"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Modelo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Part Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Serial Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Comunicación"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDispositivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub UpDateDispositivos()

Dim LV_Column()         As String
Dim i                   As Integer

    With Me
        For i = 0 To .LstViewDispositivo.UBound
            Set_Column_Lista_Instrumentos .LstViewDispositivo(i), LV_Column
            Select Case i
            Case 0
                BD_Fill_LstVw_With_Instru_From_Pjt .LstViewDispositivo(i), GV_Actual_Project.Cod_Project, 3
            Case 1
                BD_Fill_LstVw_With_Instru_From_Pjt .LstViewDispositivo(i), GV_Actual_Project.Cod_Project, 1
            Case 2
                BD_Fill_LstVw_With_Instru_From_Pjt .LstViewDispositivo(i), GV_Actual_Project.Cod_Project, 2
            Case 3
                BD_Fill_LstVw_With_Instru_From_Pjt .LstViewDispositivo(i), GV_Actual_Project.Cod_Project, 4
            End Select
        Next
    End With
    
End Sub

Private Sub cmdAgregar_Click(Index As Integer)


    'Call PostMessage(GV_hWnd_Mdi, NV_CUSTOM_MESSAGE, 0&, 0&)
    'Call PostMessage(GV_hWnd_Mdi, NV_INSTRUMENT_MESSAGE, 0&, 0&)
    
'    Dim frmNew      As frmNewInstrument
'
'    Set frmNew = New frmNewInstrument
'
'    frmNew.WindowState = vbMaximized
'    frmNew.Show
'
'    Select Case Index
'
'    Case 0
'        frmNew.cboInfoDispo(1).ListIndex = 3
'    Case 1
'        frmNew.cboInfoDispo(1).ListIndex = 1
'    Case 2
'        frmNew.cboInfoDispo(1).ListIndex = 2
'    Case 3
'        frmNew.cboInfoDispo(1).ListIndex = 3
'
'    End Select

    GV_Flag_Add_Prueba = False
    GV_Flag_Add_Exita = False
    GV_Flag_Add_Medicion = False
    GV_Flag_Add_Accesorio = False

    
    Select Case Index
    Case 0
        GV_Flag_Add_Prueba = True
    Case 1
        GV_Flag_Add_Exita = True
    Case 2
        GV_Flag_Add_Medicion = True
    Case 3
        GV_Flag_Add_Accesorio = True
    End Select
    
    Dim frmNew      As frmListaInstrumentos
    
    Set frmNew = New frmListaInstrumentos
    
    frmNew.WindowState = vbMaximized
    frmNew.Show
    
    
End Sub

Private Sub Form_Activate()

    UpDateDispositivos
    
End Sub

Private Sub Form_GotFocus()

     UpDateDispositivos
     
End Sub

Private Sub Form_Load()

    UpDateDispositivos
    
End Sub

Private Sub Form_Resize()

Dim i           As Integer
Dim lHeight     As Long

    With Me
        
        i = 0
        
        .frameDispositivo(i).Width = .Width
        
        .cmdAgregar(i).Left = .Width - .cmdEliminar(i).Left - .cmdAgregar(i).Width
        .cmdModificar(i).Left = (.Width - .cmdModificar(i).Width) / 2
        
        .LstViewDispositivo(i).Width = .frameDispositivo(i).Width - 2 * .LstViewDispositivo(i).Left
        .LstViewDispositivo(i).Height = .cmdAgregar(i).Top _
                                        - .LstViewDispositivo(i).Top _
                                        - 120
        
        lHeight = (.Height - .frameDispositivo(i).Height) / .frameDispositivo.UBound
        
        For i = 1 To .frameDispositivo.UBound
            
            .frameDispositivo(i).Width = .Width
            .frameDispositivo(i).Height = lHeight
            .frameDispositivo(i).Top = .frameDispositivo(i - 1).Height _
                                        + .frameDispositivo(i - 1).Top
            
            .cmdAgregar(i).Left = .cmdAgregar(i - 1).Left
            .cmdModificar(i).Left = .cmdModificar(i - 1).Left
            .cmdAgregar(i).Top = .frameDispositivo(i).Height _
                                - (.frameDispositivo(i - 1).Height _
                                - .cmdAgregar(i - 1).Top)
            .cmdEliminar(i).Top = .cmdAgregar(i).Top
            .cmdModificar(i).Top = .cmdAgregar(i).Top
            
            .LstViewDispositivo(i).Width = .LstViewDispositivo(i - 1).Width
            .LstViewDispositivo(i).Height = .cmdAgregar(i).Top _
                                            - .LstViewDispositivo(i).Top _
                                            - 120
        
        Next
    
    End With
    
End Sub
