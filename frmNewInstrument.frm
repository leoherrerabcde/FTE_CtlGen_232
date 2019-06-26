VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNewInstrument 
   BorderStyle     =   0  'None
   Caption         =   "Nuevo Instrumento"
   ClientHeight    =   9075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9510
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frameTablaComandos 
      Caption         =   "Tabla de Comandos"
      Height          =   5055
      Left            =   120
      TabIndex        =   33
      Top             =   3960
      Width           =   9255
      Begin VB.CommandButton cmdEjecutarCmd 
         Caption         =   "&Probar"
         Height          =   255
         Left            =   7800
         TabIndex        =   47
         Top             =   4200
         Width           =   1455
      End
      Begin VB.ComboBox cboParametro 
         Height          =   315
         Left            =   5640
         TabIndex        =   45
         Text            =   "cboParametro"
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   255
         Left            =   7800
         TabIndex        =   44
         Top             =   4560
         Width           =   1455
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   255
         Left            =   7200
         TabIndex        =   43
         Top             =   4560
         Width           =   1455
      End
      Begin VB.TextBox txtCommand 
         Height          =   375
         Left            =   1800
         TabIndex        =   41
         Text            =   "txtCommand"
         Top             =   4560
         Width           =   3735
      End
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmNewInstrument.frx":0000
         Left            =   120
         List            =   "frmNewInstrument.frx":0002
         TabIndex        =   39
         Text            =   "cboTipo"
         Top             =   4560
         Width           =   1575
      End
      Begin MSComctlLib.ListView LstVwTblComandos 
         Height          =   3735
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   6588
         View            =   3
         LabelEdit       =   1
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
      Begin VB.Label lblParametro 
         Caption         =   "Parámetro"
         Height          =   255
         Left            =   5640
         TabIndex        =   46
         Top             =   4320
         Width           =   975
      End
      Begin VB.Label lblComando 
         Caption         =   "Comando"
         Height          =   255
         Left            =   1800
         TabIndex        =   42
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   4320
         Width           =   1455
      End
      Begin VB.Label lblComandos 
         Caption         =   "Comandos de Configuración :"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame FrameParamComponente 
      Caption         =   "Parámetros"
      Height          =   3975
      Left            =   4800
      TabIndex        =   30
      Top             =   0
      Width           =   4575
      Begin MSComctlLib.ListView LstVwParametros 
         Height          =   3255
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
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
   Begin VB.ComboBox cboInfoDispo 
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Frame frameInfoDispositivos 
      Caption         =   "Información Dispositivo"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton cmdSelArchivo 
         Caption         =   "Command1"
         Height          =   195
         Left            =   4200
         TabIndex        =   17
         Top             =   3360
         Width           =   255
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   9
         Left            =   1680
         TabIndex        =   36
         Text            =   "txtInfoDispo"
         Top             =   3240
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   9
         Left            =   2160
         TabIndex        =   35
         Text            =   "Combo1"
         Top             =   3240
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   255
         Left            =   3360
         TabIndex        =   18
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   8
         Left            =   1440
         TabIndex        =   28
         Text            =   "txtInfoDispo"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   8
         Left            =   2160
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   15
         Text            =   "txtInfoDispo"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   14
         Text            =   "txtInfoDispo"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   13
         Text            =   "txtInfoDispo"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   7
         Left            =   2160
         TabIndex        =   16
         Text            =   "Combo1"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   6
         Left            =   2160
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   5
         Left            =   2160
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   4
         Left            =   2160
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   3
         Left            =   2160
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cboInfoDispo 
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   23
         Text            =   "txtInfoDispo"
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   24
         Text            =   "txtInfoDispo"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   25
         Text            =   "txtInfoDispo"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   26
         Text            =   "txtInfoDispo"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDispo 
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   27
         Text            =   "txtInfoDispo"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Caracterización :"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   37
         Top             =   3240
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Componente :"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dispositivo :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Funcion :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fabricante :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Instrumento :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Modelo :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Part Number"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   3
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Comunicación :"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   2
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Number :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   1
         Top             =   2520
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmNewInstrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




'-----------------------------------------
Private PV_Otro_Enable          As Boolean
Private PV_Enable_Edit_Cbo      As Boolean
Private PV_ColHeader_Param()    As String
Private PV_ColHeader_Cmmds()    As String
Private PV_Comandos()           As Type_Comando
Private PV_Updated              As Boolean
'
'


Public Sub Leer_Instrumento_From_BD(Cod_Proyecto As Integer, Cod_Instrument As Integer)

End Sub


Sub Agregar_Nuevo_Instru_Comp_en_BD()

Dim Cod_Compo       As Integer
Dim Cod_Comu        As Integer
Dim Cod_Dispo       As Integer
Dim Cod_Fabri       As Integer
Dim Cod_Func        As Integer
Dim Modelo          As String
Dim NumParte        As String
Dim NumSerie        As String
Dim Cod_TabParam    As Integer
Dim Cod_Tab_Command As Integer
Dim Cod_Instru      As Integer
Dim LV_Address      As Integer

    With Me
        
        Cod_Dispo = .cboInfoDispo(Enum_Orden_Tabla.Tipo_Disp).ListIndex + 1
        
        Select Case Cod_Dispo
            
            Case Is = 1
            
                Cod_Compo = .cboInfoDispo(Enum_Orden_Tabla.Clase_Componen).ListIndex + 1
                Cod_Instru = 0
                Cod_Comu = 0 ' .cboInfoDispo(Enum_Orden_Tabla.Comunicacion).ListIndex + 1
        
            Case Is = 2
            
                Cod_Instru = .cboInfoDispo(Enum_Orden_Tabla.Clase_Instru).ListIndex + 1
                Cod_Compo = 0
                Cod_Comu = .cboInfoDispo(Enum_Orden_Tabla.Comunicacion).ListIndex + 1
                
                If Cod_Comu = 1 Then
                    LV_Address = .txtInfoDispo(9).Text
                End If
        
        End Select
        
        Cod_Fabri = .cboInfoDispo(Enum_Orden_Tabla.Fabricante).ListIndex + 1
        Cod_Func = .cboInfoDispo(Enum_Orden_Tabla.Funcion).ListIndex + 1
        Modelo = .txtInfoDispo(Enum_Orden_Tabla.Modelo).Text
        NumParte = .txtInfoDispo(Enum_Orden_Tabla.PartNumber).Text
        NumSerie = .txtInfoDispo(Enum_Orden_Tabla.SerialNumber).Text
        Cod_TabParam = 0 '.cboInfoDispo(Enum_Orden_Tabla.).ListIndex + 1
        Cod_Tab_Command = 0 ' .cboInfoDispo(Enum_Orden_Tabla).ListIndex + 1
        
        If GV_Cod_Instru Then
            GV_Cod_Instru = BD_Agregar_Instrumento(Cod_Compo, _
                    Cod_Instru, _
                    Cod_Comu, _
                    Cod_Dispo, _
                    Cod_Fabri, _
                    Cod_Func, _
                    Modelo, _
                    NumParte, _
                    NumSerie, _
                    Cod_TabParam, _
                    Cod_Tab_Command, _
                    LV_Address)
        Else
            GV_Cod_Instru = BD_Agregar_Instrumento(Cod_Compo, _
                    Cod_Instru, _
                    Cod_Comu, _
                    Cod_Dispo, _
                    Cod_Fabri, _
                    Cod_Func, _
                    Modelo, _
                    NumParte, _
                    NumSerie, _
                    Cod_TabParam, _
                    Cod_Tab_Command, _
                    LV_Address)
        End If
        
        If Me.LstVwTblComandos.ListItems.Count Then
        Put_Cod_Instru_To_Comandos GV_Cod_Instru, PV_Comandos
        
        BD_Agregar_Comandos PV_Comandos
        End If
        
    End With
    
End Sub
'-----------------------------------------



Sub Agregar_Nuevos_Elemento_CboInfo_en_BD()

Dim i           As Integer

    With Me
        
        For i = 0 To .cboInfoDispo.UBound
        
            If .cboInfoDispo(i).ListCount > .cboInfoDispo(i).Tag + 1 Then
            
                Agregar_NuevoTipo_en_BD i, .cboInfoDispo(i).Text
                
            End If
        
        Next
    
    End With
    
End Sub


Sub Agregar_New_Info_a_CboInfo(Index As Integer)

Dim i           As Integer

    With Me.cboInfoDispo(Index)
    
        ' ------------
        ' Excepciones:
        
        ' 1: Texto Nulo
        If .Text = "" Then
        
            .List(.ListCount - 1) = "Otro"
            .ListIndex = .ListCount - 1
            Exit Sub
            
        End If
        
        i = BD_Index_Elemento(Index, .Text)
        
        If i > 0 Then
        
            .List(.ListCount - 1) = "Otro"
            .ListIndex = i - 1
            
            MsgBox "Nombre Existente", vbExclamation
        
            Exit Sub
            
        End If
        ' ------------
        '
        '
        If .List(.Tag) = "Ingrese Aquí" Then
            
            .AddItem "Otro"
        
        End If
        
        .List(.Tag) = .Text
            
        .ListIndex = .Tag
            
    End With
    
End Sub

Sub Cargar_Instrumento(LV_Cod_Instru As Integer)

Dim LV_Campos()         As String
Dim LV_Indexs()         As Integer
Dim LV_Cmds()           As Type_Comando

    If LV_Cod_Instru = 0 Then
        Exit Sub
    End If
    
    If BD_Read_Instrument(LV_Cod_Instru, LV_Campos, LV_Indexs) = True Then
        With Me
            .cboInfoDispo(0).ListIndex = LV_Indexs(0) - 1
            .cboInfoDispo(1).ListIndex = LV_Indexs(1) - 1
            If LV_Indexs(2) Then
                .cboInfoDispo(2).ListIndex = LV_Indexs(2) - 1
            Else
                .cboInfoDispo(8).ListIndex = LV_Indexs(3) - 1
            End If
            .cboInfoDispo(3).ListIndex = LV_Indexs(4) - 1
            
            .txtInfoDispo(4) = LV_Campos(0)
            .txtInfoDispo(5) = LV_Campos(1)
            .txtInfoDispo(6) = LV_Campos(2)
            .txtInfoDispo(9) = LV_Campos(3)
            
            If LV_Indexs(5) Then
                .cboInfoDispo(7).ListIndex = LV_Indexs(5) - 1
                .txtInfoDispo(9).Text = LV_Campos(4)
                BD_Read_Commands_Instrument LV_Cod_Instru, LV_Cmds
            End If
        End With
    Else
    End If
    
End Sub

Sub Inicializar_Prueba_Comando()

Dim i           As Integer

    i = 0
    ReDim GV_Instrumentos(i)
    
    With GV_Instrumentos(i)
    
        .Address = Me.txtInfoDispo(9).Text
        
        Select Case Me.cboTipo.ListIndex
        Case 0
            .Cmd_Config = Me.txtCommand.Text
        Case 1
            ReDim .Cmd_Set_Var(0)
            .Cmd_Set_Var(0) = Me.txtCommand.Text
        Case 2
            ReDim .Cmd_Consult(0)
            .Cmd_Consult(0) = Me.txtCommand.Text
        
        End Select
        
    End With
 
 End Sub

 Sub Iniciar_LstView_Parametros()

    ReDim PV_ColHeader_Param(4)
    
    PV_ColHeader_Param(0) = "Parámetro"
    PV_ColHeader_Param(1) = "Mín"
    PV_ColHeader_Param(2) = "Typ"
    PV_ColHeader_Param(3) = "Máx"
    PV_ColHeader_Param(4) = "Unidad"
    
    AddColumListView Me.LstVwParametros, PV_ColHeader_Param
    
End Sub

Sub Iniciar_LstVw_Caracterizacion()

    ReDim PV_ColHeader_Cmmds(3)
    
    PV_ColHeader_Cmmds(0) = "Pot In"
    PV_ColHeader_Cmmds(1) = "Frecuencia"
    PV_ColHeader_Cmmds(2) = "Pot Sal"
    PV_ColHeader_Cmmds(3) = "Ganancia/Pérdida"
    'PV_ColHeader_Cmmds(4) = "Descripción"
    
    AddColumListView Me.LstVwTblComandos, PV_ColHeader_Cmmds
    
End Sub

Sub Iniciar_LstVw_Comandos()

    ReDim PV_ColHeader_Cmmds(4)
    
    PV_ColHeader_Cmmds(0) = "Tipo"
    PV_ColHeader_Cmmds(1) = "Comando"
    PV_ColHeader_Cmmds(2) = "Parámetro"
    PV_ColHeader_Cmmds(3) = "Descripción"
    PV_ColHeader_Cmmds(4) = ""
    
    AddColumListView Me.LstVwTblComandos, PV_ColHeader_Cmmds
    
End Sub

Sub Refresh_LstVw_Comandos()

Dim LV_Campos()         As String
Dim i                   As Integer

    With Me.LstVwTblComandos
        .ListItems.Clear
        
        ReDim LV_Campos(2)
        
        For i = 0 To UBound(PV_Comandos)
            With PV_Comandos(i)
                LV_Campos(0) = GV_Lbls_Fn_Comunica(.Cod_Funcion)
                LV_Campos(1) = .Comando
                LV_Campos(2) = GV_Lbls_Parametros(.Cod_Parametro)
                
                AddItemListView Me.LstVwTblComandos, LV_Campos, True
            End With
        Next
        
    End With
    
End Sub

Function Send_Command(Index As Integer, Str_Cmd As String)

'Dim Str_Cmd         As String
Dim LV_Str_Val      As String

    'Str_Cmd = GV_Instrumentos(Index).Cmd_Set_Var(0) '& " " & LV_Str_Val
    
    Call Send(GPIB0, GV_Result_List(Index), Str_Cmd, NLend)
    
    If (ibsta And EERR) Then
        Error_GPIB
'        MsgBox "Error sending " & Str_Cmd, vbOKOnly
'        End
    End If
    

End Function

Function Listen_Command(LV_Index As Integer) As String

Dim Str_Cmd         As String
Dim LV_Reading      As String

    Str_Cmd = GV_Instrumentos(LV_Index).Cmd_Consult(0)
    
    If Str_Cmd <> "" Then
        Call Send(GPIB0, GV_Result_List(LV_Index), Str_Cmd, NLend)
    End If
    
    If (ibsta And EERR) Then
        Error_GPIB
'        MsgBox "Error sending '*IDN?'. "
'        End
    End If
    
    
    LV_Reading = Space$(&H32)
    Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
    If (ibsta And EERR) Then
         MsgBox "Error in receiving response  "
    End If
    
    If ibcntl > 0 Then
        LV_Reading = Left$(LV_Reading, ibcntl - 1)
        Listen_Command = LV_Reading
    End If
    
End Function

Sub Set_Tabla_Caracterizacion()

    Iniciar_LstVw_Caracterizacion
    
    With Me
        .frameTablaComandos.Caption = "Tabla de Caracterización"
        .lblComandos.Caption = "Puntos de Caracterización :"
        .cboParametro.Visible = False
        .cboTipo.Visible = False
        .txtCommand.Visible = False
        .lblComando.Visible = False
        .lblParametro.Visible = False
        .lblTipo.Visible = False
        .cmdAgregar.Visible = False
    End With
    
End Sub

Sub Set_Tabla_Comandos()

Dim LV_Ptos()       As Type_Ptos_Charac

    With Me
        
        Iniciar_LstVw_Comandos
        
        .frameTablaComandos.Caption = "Tabla de Comandos"
        .lblComandos.Caption = "Comandos de Configuración :"
        .cboTipo.Visible = True
        .txtCommand.Visible = True
        .lblComando.Visible = True
        .lblTipo.Visible = True
        .cmdAgregar.Visible = True
        
        If .cboTipo.ListIndex Then
            .cboParametro.Visible = True
            .lblParametro.Visible = True
        End If
        
        If Open_Characteriz_File(.txtInfoDispo(9).Text, LV_Ptos) = False Then
            .LstVwTblComandos.ListItems.Clear
        Else
            Fill_LstVw_Ptos_Carac .LstVwTblComandos, LV_Ptos
        End If
    End With
    
End Sub

Sub Verificar_Parametros_Correctos()

Dim i           As Integer

    With Me
    
        For i = 0 To .cboInfoDispo.UBound
        
            If i <> 0 And i <> 1 Then
                If .cboInfoDispo(i).ListIndex = .cboInfoDispo(i).ListCount - 1 Then
                    
                    MsgBox "El parámetro " & Replace(.Label1(i), ":", "") & " no ha sido correctamente seleccionado.", vbExclamation
                    
                    Exit Sub
                    
                End If
            End If
        Next
        
        For i = 4 To 6
        
            If .txtInfoDispo(i).Text = "" Then
            
                MsgBox "El parámetro " & Replace(.Label1(i), ":", "") & " no ha sido correctamente ingresado.", vbExclamation
            
            End If
            
        Next
        
    End With
    
End Sub


Private Sub cboInfoDispo_Click(Index As Integer)

    Select Case Index
    
        Case Is = Enum_Orden_Tabla.Clase_Componen
        
        Case Is = Enum_Orden_Tabla.Clase_Instru
        
        Case Is = Enum_Orden_Tabla.Comunicacion
        
            If Me.cboInfoDispo(Enum_Orden_Tabla.Tipo_Disp).ListIndex = 1 Then
                ' Es instrumento
                If Me.cboInfoDispo(Index).ListIndex = 0 Then
                    Me.Label1(9).Visible = True
                    Me.txtInfoDispo(9).Visible = True
                Else
                    Me.Label1(9).Visible = False
                    Me.txtInfoDispo(9).Visible = False
                End If
            End If
        
        Case Is = Enum_Orden_Tabla.Fabricante
        
        Case Is = Enum_Orden_Tabla.Funcion
        
        Case Is = Enum_Orden_Tabla.Modelo
        
        Case Is = Enum_Orden_Tabla.PartNumber
        
        Case Is = Enum_Orden_Tabla.SerialNumber
        
        Case Is = Enum_Orden_Tabla.Tipo_Disp
    
            With Me
            
                Select Case .cboInfoDispo(Index).ListIndex
                
                    Case Is = 0
                    
                        Set_Tabla_Caracterizacion
                        .cboInfoDispo(2).Visible = False
                        .cboInfoDispo(8).Visible = True
                        .cboInfoDispo(7).Visible = False
                        .Label1(2).Visible = False
                        .Label1(8).Visible = True
                        .Label1(7).Visible = False
                        .txtInfoDispo(7).Visible = False
'                        .txtInfoDispo(2).Visible = False
                        .txtInfoDispo(9).Visible = True
                        .Label1(9).Visible = True
                        .Label1(9).Caption = "Caracterización :"
                        .cmdSelArchivo.Visible = True
                        
                        .txtInfoDispo(9).Tag = .txtInfoDispo(9).Top
                        .Label1(9).Tag = .Label1(9).Top
                        .cmdSelArchivo.Tag = .cmdSelArchivo.Top
                        
                        .txtInfoDispo(9).Top = .txtInfoDispo(7).Top
                        .Label1(9).Top = .Label1(7).Top
                        .cmdSelArchivo.Top = .Label1(9).Top
                        
                        
                    Case Is = 1
                        
                        Set_Tabla_Comandos
                        .cboInfoDispo(2).Visible = True
                        .cboInfoDispo(8).Visible = False
                        .cboInfoDispo(7).Visible = True
                        .Label1(2).Visible = True
                        .Label1(8).Visible = False
                        .Label1(7).Visible = True
                        .txtInfoDispo(7).Visible = False
'                        .txtInfoDispo(2).Visible = False
                        If Me.cboInfoDispo(Enum_Orden_Tabla.Comunicacion).ListIndex = 0 Then
                            Me.Label1(9).Visible = True
                            Me.txtInfoDispo(9).Visible = True
                        Else
                            Me.Label1(9).Visible = False
                            Me.txtInfoDispo(9).Visible = False
                        End If
                        .Label1(9).Caption = "GPIB Address :"
                        .cmdSelArchivo.Visible = False
                        
                        .txtInfoDispo(9).Top = .txtInfoDispo(9).Tag
                        .Label1(9).Top = .Label1(9).Tag
                        .cmdSelArchivo.Top = .cmdSelArchivo.Tag
                        
                        
                End Select
                
            End With
    
    End Select
    
    ' Agregar Nuevo Elemento
    If PV_Otro_Enable = True Then
    
        With Me
            
            If .cboInfoDispo(Index).Text = "Otro" Then
            
                PV_Enable_Edit_Cbo = True
                
                .cboInfoDispo(Index).List(.cboInfoDispo(Index).ListIndex) = "Ingrese Aquí"
                
            End If
        
        End With
        
    End If
        
End Sub

Private Sub cboInfoDispo_KeyPress(Index As Integer, KeyAscii As Integer)

    If PV_Enable_Edit_Cbo = False Then
    
        KeyAscii = 0
        
    Else
    
        If Index < 4 Or Index > 3 Then
        
            If KeyAscii = 10 Or KeyAscii = 13 Then
            
                Agregar_New_Info_a_CboInfo (Index)
                
                PV_Enable_Edit_Cbo = False
                
            End If
                
        End If
        
    End If

End Sub


Private Sub cboInfoDispo_LostFocus(Index As Integer)

    If PV_Enable_Edit_Cbo = True Then
        
        If Index < 4 Or Index > 3 Then
        
            Agregar_New_Info_a_CboInfo (Index)
            
        End If
        
        PV_Enable_Edit_Cbo = False

    End If
    
End Sub

Private Sub cboTipo_Click()

    With Me
        If .cboTipo.ListIndex = 0 Then
            .lblParametro.Visible = False
            .cboParametro.Visible = False
        Else
            .lblParametro.Visible = True
            .cboParametro.Visible = True
        End If
    End With

End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)

    With Me
        If .cboTipo.ListIndex = 0 Then
            .lblParametro.Visible = False
            .cboParametro.Visible = False
        Else
            .lblParametro.Visible = True
            .cboParametro.Visible = True
        End If
    End With

End Sub

Private Sub cmdAgregar_Click()

    Dim i           As Integer
    
'        If PV_Comandos Is Nothing Then
'            Exit Sub
'        End If
        
        i = UBound(PV_Comandos)
        
        i = i + 1
        
        ReDim Preserve PV_Comandos(i)
        
        With Me
        
            PV_Comandos(i).Cod_Comunicacion = .cboInfoDispo(7).ListIndex + 1
            PV_Comandos(i).Cod_Funcion = .cboTipo.ListIndex + 1
            PV_Comandos(i).Cod_Instrumento = GV_Cod_Instru
            PV_Comandos(i).Cod_Parametro = .cboParametro.ListIndex + 1 + INDEX_PARAMETER
            PV_Comandos(i).Comando = .txtCommand.Text
            
        End With
        
        Me.Refresh_LstVw_Comandos
        
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
    
End Sub

Private Sub cmdEjecutarCmd_Click()

    With Me
'        If SetIO(.txtInfoDispo(9).Text) = True Then
'        Else
'            MsgBox "No fue posible conectarse al Instrumento"
'        End If
        Inicializar_Prueba_Comando
        
        Gen_Lists
        
        IniciarCommInstrumento
        'Iniciar_Instrumentos_RS232
        
        Select Case .cboTipo.ListIndex
        Case 0
            '.Send_Command 0, GV_Instrumentos(0).Cmd_Config
        Case 1
            .Send_Command 0, GV_Instrumentos(0).Cmd_Set_Var(0)
        Case 2
            MsgBox .Listen_Command(0), vbOKOnly
        End Select
    End With
    
End Sub

Private Sub cmdGuardar_Click()

    Verificar_Parametros_Correctos
    
    Agregar_Nuevos_Elemento_CboInfo_en_BD
    
    Agregar_Nuevo_Instru_Comp_en_BD
    
    Unload Me
    
End Sub

Private Sub cmdSelArchivo_Click()

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    
    With Me.CommonDialog
        .DialogTitle = "Archivo de Caracterización"
        .CancelError = False
        .filename = ""
        'ToDo: set the flags and attributes of the common dialog control
        '.Filter = "All Files (*.*)|*.*"
        .Filter = "*.csv"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
        Me.txtInfoDispo(9).Text = sFile
    End With
    
    
    
'    lFlags = BIF_RETURNONLYFSDIRS Or BIF_BROWSEINCLUDEFILES
'    lPath = ""
'
'    sDir = BrowseForFolder(Me.hWnd, "Seleccionar Directorio", lPath, lFlags)
'
'    If Err = 0 Then
'        Me.txtProyecto(1).Text = sDir
'    Else
'        'MsgBox "Se ha cancelado la operación, el error devuelto es:" & vbCrLf & _
'               "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
'        Err = 0
'    End If

End Sub

Private Sub Form_Load()

Dim i           As Integer
Dim j           As Integer

Dim lLista()    As String

    PV_Otro_Enable = False
    
    PV_Enable_Edit_Cbo = False
    
    With Me
        
        ' -----------------
        ' Cargar Combobox's
        ' -----------------
        For j = 0 To .cboInfoDispo.UBound
            
            .cboInfoDispo(j).Clear
            '.cboInfoDispo(j).Locked = True
            
            DB_Leer_Tipos lLista, j
            
            For i = 0 To UBound(lLista)
            
                .cboInfoDispo(j).AddItem lLista(i)
                
            Next
            
            .cboInfoDispo(j).Tag = i        '   Indica Count ó "Otro"
            .cboInfoDispo(j).ListIndex = 0
            
            If j = 2 Or j = 8 Or j = 3 Then
            
                .cboInfoDispo(j).AddItem "Otro"
            
            End If
            
        Next
        
        ' -----------------
        ' Borrar Textos
        ' -----------------
        For j = .txtInfoDispo.LBound To .txtInfoDispo.UBound
            
            .txtInfoDispo(j).Text = ""
            
            If j < 4 Or j > 6 And j < 9 Then
            
                .txtInfoDispo(j).Left = .cboInfoDispo(j).Left
                .txtInfoDispo(j).Visible = False
                
            End If
            
        Next
        
        
        .txtCommand.Text = ""
        
        With .cboTipo
            .Clear
            .AddItem "Configuración"
            .AddItem "Comando"
            .AddItem "Consulta"
            .ListIndex = 0
        End With
        
        If BD_Fill_Cbo_With_Parameters(.cboParametro) = True Then
        
            .cboParametro.ListIndex = 0
        
        End If
        
        Me.Iniciar_LstView_Parametros
        'Me.Iniciar_LstVw_Comandos
        
    End With
    
    PV_Otro_Enable = True
    
    Cargar_Instrumento GV_Cod_Instru
    
End Sub

Private Sub Form_Resize()

    With Me
        .FrameParamComponente.Width = .Width - .FrameParamComponente.Left _
                                      - 2 * .frameInfoDispositivos.Left
        .frameTablaComandos.Width = .FrameParamComponente.Width _
                                    + .FrameParamComponente.Left _
                                    - .frameInfoDispositivos.Left
        .frameTablaComandos.Height = .Height - .frameTablaComandos.Top _
                                     - 60
                                     
        .txtCommand.Top = .frameTablaComandos.Height - .txtCommand.Height _
                        - 60
        .cboTipo.Top = .txtCommand.Top
        .cboParametro.Top = .cboTipo.Top
        .cmdAgregar.Top = .txtCommand.Top
        .cmdEliminar.Top = .txtCommand.Top
        .cmdEjecutarCmd.Top = .txtCommand.Top
        
        
        .cmdEliminar.Left = .frameTablaComandos.Width _
                            - .cmdEliminar.Width _
                            - .cboTipo.Left
        
        .cmdAgregar.Left = .cmdEliminar.Left - .cmdAgregar.Width _
                            - .cboTipo.Left
        
        .cmdEjecutarCmd.Left = .cmdAgregar.Left - .cmdEjecutarCmd.Width _
                            - .cboTipo.Left
        
        .lblComando.Top = .txtCommand.Top - .lblComando.Height
        .lblTipo.Top = .lblComando.Top
        .lblParametro.Top = .lblComando.Top
        
        .LstVwParametros.Width = .FrameParamComponente.Width _
                                 - 2 * .LstVwParametros.Left
        .LstVwTblComandos.Width = .frameTablaComandos.Width _
                                  - 2 * .LstVwTblComandos.Left
        .LstVwTblComandos.Height = .lblComando.Top _
                                   - .LstVwTblComandos.Top - 120
                                  
    End With
    
End Sub

Private Sub txtInfoDispo_GotFocus(Index As Integer)

    SeleccionarText Me.txtInfoDispo(Index)
    
End Sub
