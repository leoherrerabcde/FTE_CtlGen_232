VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRangoControl 
   BorderStyle     =   0  'None
   Caption         =   "Rangos de Control"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11925
   LinkTopic       =   "Form2"
   ScaleHeight     =   7650
   ScaleWidth      =   11925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frameRangoControl 
      Caption         =   "Rangos de Control"
      Height          =   4095
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   9375
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   255
         Left            =   4200
         TabIndex        =   24
         Top             =   3720
         Width           =   855
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   855
      End
      Begin MSComctlLib.ListView LstVwRangos 
         Height          =   1575
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   2778
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
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
   Begin VB.Frame frameEdicionRango 
      Caption         =   "Edición de Rango de Control"
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   4200
      Width           =   10815
      Begin VB.TextBox txtPWInc 
         Height          =   285
         Left            =   240
         TabIndex        =   42
         Text            =   "txtPWInc"
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtPRIInc 
         Height          =   285
         Left            =   240
         TabIndex        =   41
         Text            =   "txtPRIInc"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox txtFrecInc 
         Height          =   285
         Left            =   240
         TabIndex        =   40
         Text            =   "txtFrecInc"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox chkPW_Incremental 
         Caption         =   "PW"
         Height          =   255
         Left            =   1320
         TabIndex        =   39
         Top             =   2520
         Width           =   735
      End
      Begin VB.CheckBox chkPRI_Incremental 
         Caption         =   "PRI"
         Height          =   255
         Left            =   1320
         TabIndex        =   37
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CheckBox chkFrecIncremental 
         Caption         =   "Frecuencia"
         Height          =   255
         Left            =   1320
         TabIndex        =   36
         Top             =   1800
         Width           =   1575
      End
      Begin VB.CommandButton cmdEditarEtapa 
         Caption         =   "Editar"
         Height          =   195
         Left            =   6360
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox txtPRI 
         Height          =   285
         Left            =   9720
         TabIndex        =   32
         Text            =   "txtPRI"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtPW 
         Height          =   285
         Left            =   9720
         TabIndex        =   31
         Text            =   "txtPW"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox chk50Ohm 
         Caption         =   "50 Ohm"
         Height          =   255
         Left            =   7560
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox cboOscVoltDiv 
         Height          =   315
         Left            =   8280
         TabIndex        =   28
         Text            =   "cboOscVoltDiv"
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton cmdSelCurvaPot 
         Caption         =   "Command1"
         Height          =   195
         Left            =   10200
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txtCurvaVideoPot 
         Height          =   285
         Left            =   7560
         TabIndex        =   26
         Text            =   "txtCurvaVideoPot"
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CheckBox chkCurvaVideoPot 
         Caption         =   "Aplicar Curva Video Pot"
         Height          =   195
         Left            =   7560
         TabIndex        =   25
         Top             =   240
         Width           =   2055
      End
      Begin VB.CommandButton cmdInsertar 
         Caption         =   "&Insertar"
         Height          =   255
         Left            =   6360
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cboUnidad 
         Height          =   315
         Index           =   1
         Left            =   3960
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1455
      End
      Begin VB.ComboBox cboUnidad 
         Height          =   315
         Index           =   0
         Left            =   3960
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtParamControl 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Text            =   "txtParamControl"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtParamControl 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   2
         Text            =   "txtParamControl"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtParamControl 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Text            =   "txtParamControl"
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox txtParamControl 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   5
         Text            =   "txtParamControl"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtParamControl 
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   6
         Text            =   "txtParamControl"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtParamControl 
         Height          =   285
         Index           =   5
         Left            =   3000
         TabIndex        =   7
         Text            =   "txtParamControl"
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Incremental"
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "PRI:"
         Height          =   255
         Left            =   9360
         TabIndex        =   34
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "PW:"
         Height          =   255
         Left            =   9360
         TabIndex        =   33
         Top             =   1080
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Volt/Div"
         Height          =   255
         Left            =   7560
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblRangoControl 
         Alignment       =   1  'Right Justify
         Caption         =   "Frecuencia"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   480
         Width           =   885
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Mínimo"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Máximo"
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Paso"
         Height          =   255
         Index           =   3
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   16
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblRangoControl 
         Alignment       =   1  'Right Justify
         Caption         =   "Potencia"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Mínimo"
         Height          =   255
         Index           =   6
         Left            =   1200
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Máximo"
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Paso"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   12
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblRangoControl 
         Caption         =   "Unidad"
         Height          =   255
         Index           =   9
         Left            =   4320
         TabIndex        =   11
         Top             =   840
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmRangoControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'   GV_Actual_Project.EtapasDeControl
'   GV_Actual_Project

Dim PV_Column_Header()          As String
Dim PV_Etapa_Edit               As Integer

Sub Verificar_Incremental()

    With Me
        If .chkFrecIncremental.value = 1 Then
            If IsNumeric(.txtFrecInc.Text) = True Then
                .txtParamControl(0).Text = Val(.txtParamControl(0).Text) + Val(.txtFrecInc.Text)
                .txtParamControl(1).Text = Val(.txtParamControl(1).Text) + Val(.txtFrecInc.Text)
            End If
        End If
        If .chkPRI_Incremental.value = 1 Then
            If IsNumeric(.txtPRIInc.Text) = True Then
                .txtPRI.Text = Val(.txtPRI.Text) + Val(.txtPRIInc.Text)
            End If
        End If
        If .chkPW_Incremental.value = 1 Then
            If IsNumeric(.txtPWInc.Text) = True Then
                .txtPW.Text = Val(.txtPW.Text) + Val(.txtPWInc.Text)
            End If
        End If
    End With
    
End Sub

Sub Activar_Boton_Editar()
    
    With Me
        .cmdEditarEtapa.Visible = True
    End With
    
End Sub

Sub Desactivar_Boton_Editar()
    
    With Me
        .cmdEditarEtapa.Visible = False
    End With
    
End Sub

Sub Copiar_Etapa(Index As Integer)

    Dim i           As Integer
    
    With GV_Actual_Project
        For i = 0 To 1
            Me.txtParamControl(3 * i).Text = .Rango(Index + i).ValorMin
            Me.txtParamControl(3 * i + 1).Text = .Rango(Index + i).ValorMax
            Me.txtParamControl(3 * i + 2).Text = .Rango(Index + i).Paso
        Next
        Me.txtPRI.Text = .Rango(Index).PRI
        Me.txtPW.Text = .Rango(Index).PW
    End With
    
End Sub

Sub Refresh_Column_Header()

    ReDim PV_Column_Header(5)
    
    PV_Column_Header(0) = "Etapa"
    PV_Column_Header(1) = "Parámetro"
    PV_Column_Header(2) = "Valor Mín"
    PV_Column_Header(3) = "Valor Máx"
    PV_Column_Header(4) = "Paso"
    PV_Column_Header(5) = "Unidad"
    
    AddColumListView Me.LstVwRangos, PV_Column_Header
    
End Sub

Sub RellenarComboBox()

Dim i       As Integer

    With Me
        .cboUnidad(0).Clear
        .cboUnidad(1).Clear
        
        With .cboUnidad(0)
            i = INDEX_DIMENSION_GHZ
            .AddItem GV_Lbls_Dimensiones(i + 0)
            .AddItem GV_Lbls_Dimensiones(i + 1)
            .AddItem GV_Lbls_Dimensiones(i + 2)
            .AddItem GV_Lbls_Dimensiones(i + 3)
            .ListIndex = 1
        End With
        
        With .cboUnidad(1)
            i = INDEX_DIMENSION_DBM
            .AddItem GV_Lbls_Dimensiones(i)
            .AddItem GV_Lbls_Dimensiones(i + 1)
            .AddItem GV_Lbls_Dimensiones(i + 2)
            .ListIndex = 0
        End With
    End With
    
End Sub

Sub Update_BD_Rangos()

End Sub

Sub UpDate_LstVw_Rangos()

Dim LV_Etapas           As Integer
Dim LV_Campos()         As String
Dim i                   As Integer

    Me.LstVwRangos.ListItems.Clear
    
    LV_Etapas = GV_Actual_Project.EtapasDeControl
    
    If LV_Etapas Then
    
        ReDim LV_Campos(UBound(PV_Column_Header))
        
        For i = 1 To LV_Etapas / 2
        
            With GV_Actual_Project.Rango(2 * (i - 1))
                LV_Campos(0) = i
                LV_Campos(1) = .Parametro
                LV_Campos(2) = .ValorMin
                LV_Campos(3) = .ValorMax
                LV_Campos(4) = .Paso
                LV_Campos(5) = .Unidad
            End With
            
            AddItemListView Me.LstVwRangos, LV_Campos
            
            With GV_Actual_Project.Rango(2 * i - 1)
                LV_Campos(0) = ""
                LV_Campos(1) = .Parametro
                LV_Campos(2) = .ValorMin
                LV_Campos(3) = .ValorMax
                LV_Campos(4) = .Paso
                LV_Campos(5) = .Unidad
            End With
            
            AddItemListView Me.LstVwRangos, LV_Campos
            
        Next
        
    End If
    
End Sub

Sub LoadRangos()

Dim LV_Etapas           As Integer
Dim LV_Campos()         As Integer
Dim i                   As Integer

    Me.LstVwRangos.ListItems.Clear
    
    LV_Etapas = BD_Get_Rangos_Control(GV_Actual_Project.Cod_Project, GV_Actual_Project.Rango)

    GV_Actual_Project.EtapasDeControl = LV_Etapas
    
    Me.UpDate_LstVw_Rangos
    
End Sub

Function Verificar_Integridad_Rango(LV_Min, LV_Max, LV_Paso) As Boolean

Dim LV_Msg          As String

    Verificar_Integridad_Rango = False
    
    If LV_Min > LV_Max Then
        LV_Msg = "El Valor Mínimo y Máximo no corresponde."
    ElseIf LV_Paso <= 0 And LV_Min <> LV_Max Then
        If LV_Msg <> "" Then
            LV_Msg = LV_Msg & " Además, el Paso debe ser un valor superior a Cero."
        Else
            LV_Msg = "El Paso debe ser un valor superior a Cero."
        End If
    Else
        Verificar_Integridad_Rango = True
        Exit Function
    End If
    
    MsgBox LV_Msg, vbOKOnly
    
End Function

Function Verificar_Rangos() As Boolean

Dim i           As Integer
Dim LV_Num      As Double
Dim LV_Min      As Double
Dim LV_Max      As Double
Dim LV_Paso     As Double

    With Me
        Verificar_Rangos = True
        i = 0
        If ConvTextToNumeric(.txtParamControl(i).Text, LV_Num) = True Then
            LV_Min = LV_Num
            ConvTextToNumeric .txtParamControl(i + 1).Text, LV_Num
            LV_Max = LV_Num
            If ConvTextToNumeric(.txtParamControl(i + 2).Text, LV_Num) = False Then
                MsgBox "Paso de Frecuencia Inválido", vbOKOnly
                Verificar_Rangos = False
            Else
                LV_Paso = LV_Num
                Verificar_Rangos = Verificar_Integridad_Rango(LV_Min, LV_Max, LV_Paso)
            End If
        Else
            MsgBox "Frecuencia Mínima Inválida", vbOKOnly
            Verificar_Rangos = False
        End If
    
        If Verificar_Rangos = False Then
            Exit Function
        End If
        
        i = 3
        If ConvTextToNumeric(.txtParamControl(i).Text, LV_Num) = True Then
            LV_Min = LV_Num
            ConvTextToNumeric .txtParamControl(i + 1).Text, LV_Num
            LV_Max = LV_Num
            If ConvTextToNumeric(.txtParamControl(i + 2).Text, LV_Num) = False Then
                MsgBox "Paso de Potencia Inválido", vbOKOnly
                Verificar_Rangos = False
            Else
                LV_Paso = LV_Num
                Verificar_Rangos = Verificar_Integridad_Rango(LV_Min, LV_Max, LV_Paso)
            End If
        Else
            MsgBox "Potencia Mínima Inválida", vbOKOnly
            Verificar_Rangos = False
        End If
    End With
    
End Function

Private Sub chkCurvaVideoPot_Click()

    With Me
        If .chkCurvaVideoPot.value Then
            .txtCurvaVideoPot.Visible = True
            .cmdSelCurvaPot.Visible = True
            .cmdSelCurvaPot.value = True
        Else
            .txtCurvaVideoPot.Visible = False
            .cmdSelCurvaPot.Visible = False
        End If
    End With
    
End Sub

Private Sub cmdAgregar_Click()

Dim LV_Etapa        As Integer
Dim LV_Campos()     As Integer
Dim i               As Integer
Dim LV_Num          As Double

    If Verificar_Rangos = False Then
        MsgBox "Lo siento, no se pudo agregar este Rango de Control", vbOKOnly
        Exit Sub
    End If
    
    LV_Etapa = GV_Actual_Project.EtapasDeControl
    
    ReDim Preserve GV_Actual_Project.Rango(LV_Etapa + 1)
    
    GV_Actual_Project.Rango(LV_Etapa).Cod_Dimension = Me.cboUnidad(0).ListIndex + INDEX_DIMENSION_GHZ + 1
    GV_Actual_Project.Rango(LV_Etapa).Cod_Parametro = 5
    GV_Actual_Project.Rango(LV_Etapa).Etapa = LV_Etapa + 1
    GV_Actual_Project.Rango(LV_Etapa).Parametro = "Frecuencia"
    ConvTextToNumeric Me.txtParamControl(2).Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa).Paso = LV_Num
    GV_Actual_Project.Rango(LV_Etapa).Unidad = Me.cboUnidad(0).Text
    ConvTextToNumeric Me.txtParamControl(1).Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa).ValorMax = LV_Num
    ConvTextToNumeric Me.txtParamControl(0).Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa).ValorMin = LV_Num
    ConvTextToNumeric Me.txtPRI.Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa).PRI = LV_Num
    ConvTextToNumeric Me.txtPW.Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa).PW = LV_Num
    
    With GV_Actual_Project.Rango(LV_Etapa)
        If Me.chkCurvaVideoPot.value Then
            .AplicarPV = 1
            If Me.chk50Ohm.value Then
                .b50Ohms = 1
            Else
                .b50Ohms = 0
            End If
            .CurvaPV = Me.txtCurvaVideoPot
            .VoltDiv = Me.cboOscVoltDiv.Text
        Else
            .AplicarPV = 0
            .b50Ohms = 0
            .CurvaPV = "No"
            .VoltDiv = 0
        End If
    End With
    
    GV_Actual_Project.Rango(LV_Etapa + 1).Cod_Dimension = Me.cboUnidad(1).ListIndex + INDEX_DIMENSION_DBM + 1
    GV_Actual_Project.Rango(LV_Etapa + 1).Cod_Parametro = 6
    GV_Actual_Project.Rango(LV_Etapa + 1).Etapa = LV_Etapa + 1
    GV_Actual_Project.Rango(LV_Etapa + 1).Parametro = "Potencia"
    ConvTextToNumeric Me.txtParamControl(5).Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa + 1).Paso = LV_Num
    GV_Actual_Project.Rango(LV_Etapa + 1).Unidad = Me.cboUnidad(1).Text
    ConvTextToNumeric Me.txtParamControl(4).Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa + 1).ValorMax = LV_Num
    ConvTextToNumeric Me.txtParamControl(3).Text, LV_Num
    GV_Actual_Project.Rango(LV_Etapa + 1).ValorMin = LV_Num
    
    With GV_Actual_Project.Rango(LV_Etapa + 1)
        .AplicarPV = 0
        .b50Ohms = 0
        .CurvaPV = "No"
        .VoltDiv = 0
    End With
    
    GV_Actual_Project.EtapasDeControl = GV_Actual_Project.EtapasDeControl + 2
    LV_Etapa = LV_Etapa + 2
    
    For i = LV_Etapa - 2 To LV_Etapa - 1
    
        BD_Update_Rango_Control GV_Actual_Project.Cod_Project, i
        
    Next
    

    Me.UpDate_LstVw_Rangos
    
    Verificar_Incremental
    
End Sub

Private Sub cmdInsertar_Click()

    'Me.LstVwRangos.SelectedItem.Index
    
End Sub

Private Sub cmdSelCurvaPot_Click()

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    lPath = GV_Actual_Project.Path_Project
    With Me.CommonDialog
        .Filter = "*.csv"
        .InitDir = lPath
        .CancelError = False
        .DialogTitle = "Archivo de Curva Video Potencia"
        .ShowOpen
        sFile = .fileName
    End With
    'sFile = BrowseForFile(lPath, "*.csv", "Archivo de COmpensación de Salida")
    Me.txtCurvaVideoPot.Text = sFile


End Sub

Private Sub Form_Load()

Dim i       As Integer

    With Me
    
        For i = 0 To .txtParamControl.UBound
        
            .txtParamControl(i).Text = ""
            
        Next
        
        .txtPRI.Text = 1000
        .txtPW.Text = 1
'        .txtParamControl(0).Text = 2000
'        .txtParamControl(1).Text = 2000
'        .txtParamControl(2).Text = 1100
'        .txtParamControl(4).Text = -70
'        .txtParamControl(5).Text = -40
'        .txtParamControl(6).Text = 0.2

        
        Refresh_Column_Header
        .LoadRangos
        .UpDate_LstVw_Rangos
        
        RellenarComboBox
        
        .cboOscVoltDiv.Clear
        .cboOscVoltDiv.AddItem "2"
        .cboOscVoltDiv.AddItem "5"
        .cboOscVoltDiv.AddItem "10"
        .cboOscVoltDiv.AddItem "20"
        .cboOscVoltDiv.AddItem "50"
        .cboOscVoltDiv.AddItem "100"
        .cboOscVoltDiv.AddItem "200"
        .cboOscVoltDiv.AddItem "500"
        .cboOscVoltDiv.AddItem "1000"
        .cboOscVoltDiv.AddItem "2000"
        .cboOscVoltDiv.AddItem "5000"
        .cboOscVoltDiv.ListIndex = 0
        
    End With

End Sub

Private Sub Form_Resize()

    With Me
        .frameEdicionRango.Width = .Width - 2 * .frameEdicionRango.Left
        .frameRangoControl.Width = .frameEdicionRango.Width
        
        .frameEdicionRango.Top = .Height - .frameEdicionRango.Height _
                                - 2 * .frameRangoControl.Top
        .frameRangoControl.Height = .frameEdicionRango.Top - .frameRangoControl.Top
        
        .cmdEliminar.Top = .frameRangoControl.Height - .cmdEliminar.Height _
                            - .LstVwRangos.Top
        .cmdModificar.Top = .cmdEliminar.Top
        
        
        .LstVwRangos.Width = .frameRangoControl.Width - 2 * .LstVwRangos.Left
        .LstVwRangos.Height = .cmdEliminar.Top - 2 * .LstVwRangos.Top
        
    End With
    
End Sub

Sub EditarItem()

Dim i           As Integer

    With Me
        i = .LstVwRangos.SelectedItem.Index
        i = Int((i - 1) / 2)
        PV_Etapa_Edit = 2 * i
        Editar_Etapa PV_Etapa_Edit
        Actualizar_BD_Etapas_Control
    End With

End Sub

Sub SeleccionarItemLstVw()

Dim i           As Integer

    With Me
        i = .LstVwRangos.SelectedItem.Index
        i = Int((i - 1) / 2)
        PV_Etapa_Edit = 2 * i
        Copiar_Etapa PV_Etapa_Edit
        Activar_Boton_Editar
    End With

End Sub

Private Sub LstVwRangos_Click()

    SeleccionarItemLstVw
    
End Sub

Private Sub LstVwRangos_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    SeleccionarItemLstVw
    
End Sub

Private Sub LstVwRangos_ItemClick(ByVal Item As MSComctlLib.ListItem)

    SeleccionarItemLstVw
    
End Sub

Private Sub LstVwRangos_KeyDown(KeyCode As Integer, Shift As Integer)

    SeleccionarItemLstVw
    
End Sub

Private Sub LstVwRangos_KeyUp(KeyCode As Integer, Shift As Integer)

    SeleccionarItemLstVw
    
End Sub

Private Sub txtParamControl_GotFocus(Index As Integer)

    SeleccionarText Me.txtParamControl(Index)
    
End Sub

Private Sub txtParamControl_KeyPress(Index As Integer, KeyAscii As Integer)

    If Verify_Valid_Digit(KeyAscii) = False Then
        Select Case KeyAscii
        Case 8
        Case 13
            If Index < 5 Then
                Me.txtParamControl(Index + 1).SetFocus
            Else
                Me.cmdAgregar.SetFocus
            End If
        Case Else
            KeyAscii = 0
        End Select
    End If
    
End Sub
