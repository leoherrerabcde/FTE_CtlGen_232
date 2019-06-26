VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPrueba 
   BorderStyle     =   0  'None
   Caption         =   "Prueba"
   ClientHeight    =   9060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10905
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   0
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   0
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Frame framePrueba 
      Caption         =   "Visualizaci�n"
      Height          =   2295
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   10695
      Begin MSComctlLib.ListView LstVwVisualTest 
         Height          =   1215
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2143
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame framePrueba 
      Height          =   3855
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   4920
      Width           =   10695
      Begin VB.Frame frameControlGral 
         Caption         =   "General"
         Height          =   1815
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   10695
         Begin VB.CheckBox chkRepetirPrueba 
            Caption         =   "Repetir Prueba"
            Height          =   195
            Left            =   7560
            TabIndex        =   70
            Top             =   1440
            Width           =   2055
         End
         Begin VB.CheckBox chkPisarArchivoSalida 
            Caption         =   "Pisar Archivo Salida"
            Height          =   195
            Left            =   1920
            TabIndex        =   69
            Top             =   1560
            Width           =   1875
         End
         Begin VB.CheckBox chkCortarRFalTerminar 
            Caption         =   "RF OFF al Terminar"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   1560
            Width           =   1815
         End
         Begin VB.CommandButton cmdSelTablaParam 
            Caption         =   "Command1"
            Height          =   195
            Left            =   6840
            TabIndex        =   49
            Top             =   1440
            Width           =   255
         End
         Begin VB.TextBox txtTablaParam 
            Height          =   285
            Left            =   3840
            TabIndex        =   48
            Text            =   "txtTablaParam"
            Top             =   1320
            Width           =   2895
         End
         Begin VB.CheckBox chkUsarTablaParam 
            Caption         =   "Usar Tabla de Par�metros"
            Height          =   195
            Left            =   3840
            TabIndex        =   47
            Top             =   1080
            Width           =   2295
         End
         Begin VB.CommandButton cmdSelArchiCompSal 
            Caption         =   "Command1"
            Height          =   195
            Left            =   6840
            TabIndex        =   44
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cance&lar"
            Enabled         =   0   'False
            Height          =   255
            Left            =   0
            TabIndex        =   43
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdComenzar 
            Caption         =   "&Comenzar"
            Height          =   255
            Left            =   6360
            TabIndex        =   42
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox chkEsperaEstabi 
            Caption         =   "Esperar Estabilizaci�n"
            Height          =   255
            Left            =   1440
            TabIndex        =   41
            Top             =   360
            Width           =   2055
         End
         Begin VB.TextBox txtTpoEspera 
            Height          =   285
            Left            =   1680
            TabIndex        =   39
            Text            =   "txtTpoEspera"
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chkManual 
            Caption         =   "Manual"
            Height          =   195
            Left            =   4080
            TabIndex        =   38
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox txtArchCompSal 
            Height          =   285
            Left            =   3840
            TabIndex        =   37
            Text            =   "txtArchCompSal"
            Top             =   720
            Width           =   2535
         End
         Begin VB.CheckBox chkCurvaVideoPot 
            Caption         =   "Aplicar Curva Video Pot"
            Height          =   195
            Left            =   1440
            TabIndex        =   36
            Top             =   120
            Width           =   2055
         End
         Begin VB.TextBox txtCurvaVideoPot 
            Height          =   285
            Left            =   3840
            TabIndex        =   35
            Text            =   "txtCurvaVideoPot"
            Top             =   240
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.CommandButton cmdSelCurvaPot 
            Caption         =   "Command1"
            Height          =   195
            Left            =   6840
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CheckBox chkCapPotGen 
            Caption         =   "Capturar Pot Gen"
            Height          =   255
            Left            =   1440
            TabIndex        =   33
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkAdquirir 
            Caption         =   "Adquirir"
            Height          =   195
            Left            =   1440
            TabIndex        =   40
            Top             =   960
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "mseg"
            Height          =   255
            Left            =   3000
            TabIndex        =   77
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.Frame frameParamGPIB 
         Caption         =   "GPIB"
         Height          =   2055
         Left            =   0
         TabIndex        =   11
         Top             =   1800
         Width           =   7695
         Begin VB.CheckBox chkGeneradorEn 
            Caption         =   "Generador"
            Height          =   195
            Left            =   120
            TabIndex        =   78
            Top             =   1440
            Width           =   1215
         End
         Begin VB.CheckBox chkRS232 
            Caption         =   "RS-232"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   1800
            Width           =   1095
         End
         Begin VB.Frame frmAnaliEsp 
            Caption         =   "Analizador de Espectro"
            Height          =   2055
            Left            =   3600
            TabIndex        =   25
            Top             =   0
            Width           =   3855
            Begin VB.TextBox txtAtt 
               Height          =   285
               Left            =   1200
               TabIndex        =   52
               Text            =   "txtAtt"
               Top             =   1320
               Width           =   1095
            End
            Begin VB.ComboBox cboPeakGraph 
               Height          =   315
               Left            =   120
               TabIndex        =   46
               Text            =   "cboPeakGraph"
               Top             =   1680
               Width           =   1455
            End
            Begin VB.ComboBox cboCenterSpan 
               Height          =   315
               Left            =   120
               TabIndex        =   45
               Text            =   "cboCenterSpan"
               Top             =   240
               Width           =   1455
            End
            Begin VB.TextBox txtRefLvl 
               Height          =   285
               Left            =   1200
               TabIndex        =   30
               Text            =   "txtRefLvl"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtSpan 
               Height          =   285
               Left            =   1200
               TabIndex        =   29
               Text            =   "txtSpan"
               Top             =   840
               Width           =   1095
            End
            Begin VB.TextBox txtCenterFreq 
               Height          =   285
               Left            =   1200
               TabIndex        =   28
               Text            =   "txtCenterFreq"
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label4 
               Caption         =   "Att:"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label5 
               Caption         =   "Ref Level:"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   1080
               Width           =   975
            End
            Begin VB.Label lblSpanFreq 
               Caption         =   "Span:"
               Height          =   255
               Left            =   120
               TabIndex        =   27
               Top             =   840
               Width           =   855
            End
            Begin VB.Label lblCenterFreq 
               Caption         =   "Center Freq:"
               Height          =   255
               Left            =   120
               TabIndex        =   26
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.Frame frameOsciloscopio 
            Caption         =   "Osciloscopio"
            Height          =   2055
            Left            =   1440
            TabIndex        =   17
            Top             =   0
            Width           =   2175
            Begin VB.ComboBox cboOscVoltDiv 
               Height          =   315
               Left            =   1080
               TabIndex        =   24
               Text            =   "cboOscVoltDiv"
               Top             =   960
               Width           =   975
            End
            Begin VB.OptionButton optOscNivel 
               Caption         =   "Pulso"
               Height          =   195
               Index           =   1
               Left            =   1080
               TabIndex        =   22
               Top             =   720
               Width           =   855
            End
            Begin VB.OptionButton optOscNivel 
               Caption         =   "Nivel"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   21
               Top             =   720
               Width           =   1095
            End
            Begin VB.CheckBox chkInvertir 
               Caption         =   "Invertir"
               Height          =   255
               Left            =   1080
               TabIndex        =   20
               Top             =   480
               Width           =   975
            End
            Begin VB.CheckBox chk50Ohm 
               Caption         =   "50 Ohm"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Width           =   855
            End
            Begin VB.CheckBox chkOscCh 
               Caption         =   "Canal 1"
               Height          =   195
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label2 
               Caption         =   "Volt/Div"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   960
               Width           =   735
            End
         End
         Begin VB.CheckBox chkMedirAnalizador 
            Caption         =   "Analizador"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox chkMedirPowerMeter 
            Caption         =   "Power Meter"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   1215
         End
         Begin VB.CheckBox chkMedirOsciloscopio 
            Caption         =   "Osciloscopio"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtAddressGen 
            Height          =   285
            Left            =   600
            TabIndex        =   13
            Text            =   "txtAddressGen"
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label1 
            Caption         =   "Address Gen:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
      End
   End
   Begin VB.Frame framePrueba 
      Caption         =   "Estado de Control"
      Height          =   2535
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.TextBox txtAgregarPot 
         Height          =   285
         Left            =   1200
         TabIndex        =   76
         Text            =   "txtAgregarPot"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkIncPot 
         Caption         =   "Agregar"
         Height          =   195
         Left            =   240
         TabIndex        =   75
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton cmdBajarPot 
         Caption         =   "-.1"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   74
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdSubirPot 
         Caption         =   "+.1"
         Height          =   195
         Index           =   2
         Left            =   3000
         TabIndex        =   73
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdBajarPot 
         Caption         =   "-5"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   72
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdSubirPot 
         Caption         =   "+5"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   71
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtCronometro 
         Height          =   285
         Left            =   6360
         TabIndex        =   67
         Text            =   "txtCronometro"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtPW 
         Height          =   285
         Left            =   5760
         TabIndex        =   66
         Text            =   "txtPW"
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtPRI 
         Height          =   285
         Left            =   5760
         TabIndex        =   65
         Text            =   "txtPRI"
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox chkALCOn 
         Caption         =   "ALC ON"
         Height          =   255
         Left            =   5400
         TabIndex        =   64
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkRFOn 
         Caption         =   "RF ON"
         Height          =   255
         Left            =   5400
         TabIndex        =   63
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtStepFreq 
         Height          =   285
         Left            =   3960
         TabIndex        =   58
         Text            =   "txtStepFreq"
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdDecFreq 
         Caption         =   "-"
         Height          =   195
         Left            =   3000
         TabIndex        =   57
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton cmdIncFreq 
         Caption         =   "+"
         Height          =   195
         Left            =   3000
         TabIndex        =   56
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton cmdSubirPot 
         Caption         =   "+"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   55
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdBajarPot 
         Caption         =   "-"
         Height          =   195
         Index           =   0
         Left            =   2400
         TabIndex        =   54
         Top             =   600
         Width           =   255
      End
      Begin MSComctlLib.ListView LstVwRangoControl 
         Height          =   1095
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   1931
         View            =   3
         MultiSelect     =   -1  'True
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
      Begin VB.TextBox txtPrueba 
         Height          =   285
         Index           =   2
         Left            =   3480
         TabIndex        =   5
         Text            =   "txtPrueba"
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtPrueba 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   3
         Text            =   "txtPrueba"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPrueba 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   1
         Text            =   "txtPrueba"
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "PW:"
         Height          =   255
         Left            =   5400
         TabIndex        =   62
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "PRI:"
         Height          =   255
         Left            =   5400
         TabIndex        =   61
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "[KHz]"
         Height          =   255
         Left            =   4800
         TabIndex        =   60
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Step:"
         Height          =   255
         Left            =   3480
         TabIndex        =   59
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblPrueba 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado :"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   6
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPrueba 
         Alignment       =   1  'Right Justify
         Caption         =   "Potencia Actual :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblPrueba 
         Alignment       =   1  'Right Justify
         Caption         =   "Frecuencia Actual :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Timer tmrPrueba 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   4800
      TabIndex        =   50
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PV_Ini_Tmr          As Long
Dim PV_End_Tmr          As Long

Dim PV_TckCnt_E1        As Long

Dim PV_Ptos_Prueba      As Long
Dim PV_Ptos_Now         As Long
Dim PV_Tpo_Ini          As Single
Dim PV_Estado           As Integer
Dim PV_Frec_Min         As Long
Dim PV_Frec_Max         As Long
Dim PV_Frec_Paso        As Long
Dim PV_Pot_Min          As Double
Dim PV_Pot_Max          As Double
Dim PV_Pot_Paso         As Double
Dim PV_Factor_Pot       As Double
Dim PV_Etapa            As Integer
Dim PV_EmpezarEtapa     As Boolean
Dim PV_Frec_Now         As Long
Dim PV_Pot_Now          As Double
Dim PV_Index_Tabla      As Integer

Dim PV_Flag_RF_On       As Boolean

Dim PV_Estabiliza_Counter       As Integer

Dim PV_Handle_Instruments()    As Long

Dim PV_Address_List()       As Long
Dim PC_Config_Commands()    As String


Dim PV_Flag_Freq            As Boolean
Dim PV_Flag_Pow             As Boolean

Dim PV_File_Hdl             As Integer

Dim PV_Column_Header()           As String

Dim PV_Compensa()           As Type_Correccion
Dim PV_Correccion()         As Type_Correccion
Dim PV_TablaParam()         As Type_ParamControl

Dim PV_Index_Correc         As Integer
Dim PV_CommPort             As Boolean

Function Calc_Factor_Pot(ByVal LV_Paso As Double) As Double

Dim i           As Integer
Dim LV_Factor   As Double
Dim lv

    LV_Factor = 1
    
    i = 0
    Do
        If LV_Paso = Int(LV_Paso * LV_Factor) / LV_Factor Then
            
            Calc_Factor_Pot = LV_Factor
        
            Exit Do
            
        End If
        
        i = i + 1
        
        LV_Factor = LV_Factor * 10
        
    Loop Until i = 2
    
    Calc_Factor_Pot = LV_Factor
    
End Function

Function Calcular_Compensa(LV_Frec As Long, LV_Pot As Double) As Double

Dim i       As Integer

    Calcular_Compensa = LV_Pot

    For i = 0 To UBound(PV_Compensa)
        With PV_Compensa(i)
            If .Freq = LV_Frec Then
                Calcular_Compensa = LV_Pot - .Correccion
                Exit Function
            End If
        End With
    Next

End Function

Function Calcular_Correccion(LV_Frec As Long, LV_Pot As Double) As Double

Dim i       As Integer

    Calcular_Correccion = LV_Pot

    For i = 0 To UBound(PV_Correccion)
        With PV_Correccion(i)
            If .Freq = LV_Frec Then
                Calcular_Correccion = LV_Pot - .Correccion
                Exit Function
            End If
        End With
    Next
    
End Function

Function Calcular_Correccion_Inv(LV_Frec As Long, LV_Pot As Double) As Double

Dim i       As Integer

    Calcular_Correccion_Inv = LV_Pot

    For i = 0 To UBound(PV_Correccion)
        With PV_Correccion(i)
            If .Freq = LV_Frec Then
                Calcular_Correccion_Inv = LV_Pot + .Correccion
                Exit Function
            End If
        End With
    Next
    
End Function

Sub Cargar_Curva_Video_Pot(Optional LV_File As String)

Dim h               As Integer
Dim i               As Integer
Dim LV_Line         As String
Dim LV_Campos()     As String
'Dim LV_File         As String
Dim j               As Integer

    ReDim GV_Lista_Frec(99)
    ReDim GV_Tabla_Vid_Pot(99)
    i = 1
    h = FreeFile
    
    If LV_File = "" Then
        LV_File = Me.txtCurvaVideoPot.Text
    Else
        Me.txtCurvaVideoPot.Text = LV_File
    End If
    
    If VerificarExiste(LV_File) = True Then
        
        Open LV_File For Input As h
        
        For i = 1 To 2
            If EOF(h) = True Then
                Exit Sub
            End If
            Line Input #h, LV_Line
        Next
        
        If InStr(LV_Line, ";") Then
            ' Leer Potencias
            LV_Campos = Split(LV_Line, ";")
            ReDim GV_Lista_Pot(UBound(LV_Campos) - 1)
            For i = 1 To UBound(LV_Campos)
                GV_Lista_Pot(i - 1) = LV_Campos(i)
            Next
            
            ' Leer Frecuencias
            i = 0
            Do
                If EOF(h) = False Then
                    Line Input #h, LV_Line
                    If LV_Line <> "" Then
                        LV_Campos = Split(LV_Line, ";")
                        If UBound(GV_Lista_Frec) < i Then
                            ReDim Preserve GV_Lista_Frec(99 + i)
                            ReDim Preserve GV_Tabla_Vid_Pot(99 + i)
                        End If
                        GV_Lista_Frec(i) = LV_Campos(0)
                        ReDim GV_Tabla_Vid_Pot(i).Filas(UBound(GV_Lista_Pot))
                        For j = 0 To UBound(GV_Lista_Pot)
                            GV_Tabla_Vid_Pot(i).Filas(j) = LV_Campos(j + 1)
                        Next
                        i = i + 1
                    Else
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
        End If
        
        Close h
    
    End If
    
    ReDim Preserve GV_Tabla_Vid_Pot(i - 1)
    
    ReDim Preserve GV_Lista_Frec(i - 1)
    
End Sub

Sub Cargar_Compensa()

Dim h               As Integer
Dim i               As Integer
Dim LV_Line         As String
Dim LV_Campos()     As String
Dim LV_File         As String

    h = FreeFile
    'PV_Index_Correc = 0
    
    ReDim PV_Compensa(0)
    
    LV_File = Me.txtArchCompSal
    
    If VerificarExiste(LV_File) = True Then
        
        Open LV_File For Input As h
        
        For i = 1 To 5
            If EOF(h) = True Then
                Exit Sub
            End If
            Line Input #h, LV_Line
        Next
        
        i = 0
        Do
            If EOF(h) = False Then
                Line Input #h, LV_Line
                If LV_Line <> "" Then
                    LV_Campos = Split(LV_Line, ";")
                    ReDim Preserve PV_Compensa(i)
                    PV_Compensa(i).Correccion = LV_Campos(3)
                    PV_Compensa(i).Freq = LV_Campos(0)
                    i = i + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        
        Close h
    
    End If
    
    Exit Sub
    
End Sub

Sub CargarTablaParam()

Dim h               As Integer
Dim i               As Integer
Dim LV_Line         As String
Dim LV_Campos()     As String
Dim LV_File         As String
Dim j               As Integer

    ReDim PV_TablaParam(99)
    i = 1
    h = FreeFile
    
    LV_File = Me.txtTablaParam.Text
    
    If VerificarExiste(LV_File) = True Then
        
        Open LV_File For Input As h
        
        For i = 1 To 5
            If EOF(h) = True Then
                Exit Sub
            End If
            Line Input #h, LV_Line
        Next
        
        ' Leer Frecuencias
        i = 0
        Do
            If EOF(h) = False Then
                Line Input #h, LV_Line
                If LV_Line <> "" Then
                    LV_Campos = Split(LV_Line, ";")
                    If UBound(PV_TablaParam) < i Then
                        ReDim Preserve PV_TablaParam(99 + i)
                    End If
                    PV_TablaParam(i).Frec = LV_Campos(0)
                    PV_TablaParam(i).Pot_Gen = LV_Campos(1)
                    i = i + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
        
        Close h
    
    End If
    
    ReDim Preserve PV_TablaParam(i - 1)
    PV_Index_Tabla = 0

End Sub

Sub Cargar_Correccion()

Dim h               As Integer
Dim i               As Integer
Dim LV_Line         As String
Dim LV_Campos()     As String

    h = FreeFile
    PV_Index_Correc = 0
    
    ReDim PV_Correccion(0)
    
    'Set PV_Correccion = vbNull

    On Error GoTo NoHayArchivo
    
    Open GetFileCorreccion For Input As h
    
    On Error GoTo 0
    
    For i = 1 To 5
        If EOF(h) = True Then
            Exit Sub
        End If
        Line Input #h, LV_Line
    Next
    
    i = 0
    Do
        If EOF(h) = False Then
            Line Input #h, LV_Line
            If LV_Line <> "" Then
                LV_Campos = Split(LV_Line, ";")
                ReDim Preserve PV_Correccion(i)
                PV_Correccion(i).Correccion = LV_Campos(3)
                PV_Correccion(i).Freq = LV_Campos(0)
                i = i + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    Close h
    
    Exit Sub
    
NoHayArchivo:

    MsgBox "No se ha encontrado archivo de Correcci�n.", vbInformation
    
    On Error GoTo 0
    

End Sub

Sub Cargar_Valores_Etapa(LV_Etapa As Integer)

    PV_EmpezarEtapa = True
    
    With GV_Actual_Project.Rango(LV_Etapa)
    
        PV_Frec_Min = .ValorMin
        PV_Frec_Max = .ValorMax
        PV_Frec_Paso = .Paso
    
    End With
    
    With GV_Actual_Project.Rango(LV_Etapa + 1)
        
        PV_Pot_Min = .ValorMin
        PV_Pot_Max = .ValorMax
        PV_Pot_Paso = .Paso
        
        PV_Factor_Pot = Calc_Factor_Pot(PV_Pot_Paso)
        
    End With
    
End Sub

Sub Cargar_Valores()

    PV_Etapa = 0
    
    If GV_Actual_Project.EtapasDeControl = 0 Then
        Me.cmdComenzar.Enabled = False
        Exit Sub
    Else
        Me.cmdComenzar.Enabled = True
    End If
    
    Me.Cargar_Valores_Etapa PV_Etapa
    
    With Me
        
        PV_Frec_Now = PV_Frec_Min
        PV_Pot_Now = PV_Pot_Min
        
        If Me.chkUsarTablaParam.value Then
            CargarValoresFromTablaParam
        End If
        
        PV_Flag_Freq = True
        PV_Flag_Pow = True
        
        .txtPrueba(0).Text = PV_Frec_Min
        .txtPrueba(1).Text = PV_Pot_Min
        .txtPrueba(2).Text = ""
        
    End With
    
End Sub

Sub CargarValoresFromTablaParam()

    If PV_Index_Tabla <= UBound(PV_TablaParam) Then
        PV_Frec_Now = PV_TablaParam(PV_Index_Tabla).Frec
        PV_Pot_Now = PV_TablaParam(PV_Index_Tabla).Pot_Gen
        PV_Index_Tabla = PV_Index_Tabla + 1
    End If
    
End Sub

Sub Repetir_Prueba()

    With Me
        .tmrPrueba.Enabled = False
        .cmdComenzar.Caption = "&Reanudar"
        .cmdComenzar.Refresh
        PV_Index_Tabla = 0
        PV_EmpezarEtapa = True
        PV_Etapa = 0
        PV_Estado = 0
        PV_Ptos_Now = 0
        PV_Tpo_Ini = Timer
        .Cargar_Valores_Etapa PV_Etapa
        Cargar_Valores
        If .chkCortarRFalTerminar.value = 1 Then
            .chkRFOn.value = 0
            SendRFPowerOff 0
        End If
    End With
    
End Sub

Sub Cancelar_Prueba()

    With Me
        If .chkCortarRFalTerminar.value = 1 Then
            .chkRFOn.value = 0
            SendRFPowerOff 0
        End If

        .chkRS232.Enabled = True
        .tmrPrueba.Enabled = False
        .cmdCancelar.Enabled = False
        .cmdComenzar.Caption = "&Comenzar"
        .Close_Devices
        .Cerrar_Archivo_Salida
        .txtPrueba(2).Text = ""
    End With

End Sub

Function Capturar_Valores() As Boolean

Dim LV_Data         As String
Dim LV_Value        As Double
Dim i               As Integer

    If Me.chkCapPotGen.value Then
        If Me.chkRS232.value = 0 Then
            LV_Value = Capturar_Pot_Gen(0)
        Else
            LV_Value = Capturar_Pot_Gen_RS232(0)
            LV_Value = Me.txtPrueba(1)
        End If
        LV_Value = Calcular_Correccion_Inv(PV_Frec_Now, LV_Value)
        GV_Data_Captur.Data(1) = LV_Value
    End If
    
    Capturar_Valores = True

For i = 1 To UBound(GV_Instrumentos)
    LV_Data = Get_Data_From_Instr(i)
    
    If LV_Data <> "" Then
        If IsNumeric(LV_Data) = False Then
            LV_Value = ExtraerNumeric(LV_Data)
        End If
        LV_Value = Val(LV_Data)
        ' Obtiene Compensaci�n de Salidas
        LV_Value = Calcular_Compensa(PV_Frec_Now, LV_Value)
        GV_Data_Captur.Data(i + 1) = LV_Value
        If Me.chkCurvaVideoPot.value Then
            LV_Value = Convertir_Video_en_Pot(LV_Value, PV_Frec_Now)
        End If
        If UBound(GV_Data_Captur.Data) > i + 1 Then
            GV_Data_Captur.Data(2 + i) = LV_Value
        End If
    End If
    Capturar_Valores = True
Next

End Function

Sub Cerrar_Archivo_Salida()

    If PV_File_Hdl <> 0 Then
    
        Print #PV_File_Hdl, vbCrLf
        Print #PV_File_Hdl, "Hora T�rmino:" & GV_Ch_Decimal & format(Time(), "hh:mm:ss") & vbCrLf
    
        Close #PV_File_Hdl
        
        PV_File_Hdl = 0
        
    End If
    
End Sub

Sub Close_Devices()

Dim i           As Integer

    If Me.chkRS232.value = 0 Then
        If GPIBglobalsRegistered = 1 Then
             For i = 0 To UBound(GV_Instrumentos)
                With GV_Instrumentos(i)
                    If .Cmd_End <> "" Then
                        SendCmd .Cmd_Config, i
                    End If
                End With
            Next
        
           Call DevClearList(GPIB0, GV_Result_List)
    
            ilonl GPIB0, 0
        
            GPIBglobalsRegistered = 0
            
        End If
    Else
        fMainForm.CloseMsComm
    End If

End Sub

Sub Crear_Estructuras()

Dim i           As Integer

    ReDim GV_Data_Captur.Data(3)
    
    ReDim GV_Data_Captur.NombreCampo(3)
    
    With GV_Data_Captur
    
        For i = 0 To 3
        
                    
            .NombreCampo(i) = ""
            
        Next
        
        .NombreCampo(0) = "Frecuencia Gen"
        .NombreCampo(1) = "Potencia Gen"
        .NombreCampo(2) = "Voltaje"
        .NombreCampo(3) = "Potencia"
        
    End With
    
End Sub

Function CommandFrec(ByVal Index As Integer, _
                    ByVal LV_Freq As Double, _
                    Optional LV_Str As String) As String

Dim Str_Cmd         As String
Dim LV_Str_Val      As String
Dim LV_Cmds()       As String
Dim i               As Integer

    CommandFrec = ""
    
    Str_Cmd = GV_Instrumentos(Index).Cmd_Set_Var(0) '& " " & LV_Freq
    
    If Str_Cmd = "" Then
        Exit Function
    End If
    
    LV_Str_Val = LV_Freq
    Str_Cmd = Replace_Value(Str_Cmd, LV_Str_Val, "%F")
    
    If LV_Str <> "" Then
        Str_Cmd = Str_Cmd & LV_Str
    End If
    
    LV_Cmds = Split(Str_Cmd, ";")
    For i = 0 To UBound(LV_Cmds)
        ParseCommand LV_Cmds(i)
        
        CommandFrec = LV_Cmds(i) & vbCrLf
    Next
    
End Function

Sub Enviar_Frecuencia()

Dim LV_F        As Double
Dim LV_F_I      As Double
Dim i           As Integer

    If PV_Flag_Freq = True Then
            
        For i = 0 To UBound(GV_Instrumentos)
            If i Or Me.chkGeneradorEn.value = 1 Then
                With GV_Instrumentos(i)
                    LV_F = PV_Frec_Now * 1000000#
                    
                    'SendFrec Index, LV_F
                    'If i Then
                    If Me.chkRS232.value = 0 Then
                        SendFrec i, PV_Frec_Now
                    Else
                        fMainForm.SendRS232 CommandFrec(i, PV_Frec_Now)
                    End If
                    GV_Data_Captur.Data(0) = PV_Frec_Now
                    
                    If Me.chkRS232.value = 0 Then
                        'Verify_Frec_State 0, LV_F_I
                    Else
                    End If
                    
                    If LV_F_I <> LV_F Then
                        
                        'Me.cmdCancelar.value = True
                    
                    End If
                    
                    PV_Flag_Freq = False
                        
                    'SendFrec 2, PV_Frec_Now, "MZ"
                End With
            End If
        Next
    End If
    
End Sub

Function CommandPow(ByVal Index As Integer, _
                    ByVal LV_Pow As Double) As String
                    
Dim Str_Cmd         As String
Dim LV_Str_Val      As String

    LV_Str_Val = LV_Pow
    LV_Str_Val = Replace(LV_Str_Val, ",", ".")
    
    Str_Cmd = Replace_Value(GV_Instrumentos(Index).Cmd_Set_Var(1), LV_Str_Val, "%P")
    
    CommandPow = Str_Cmd & vbCrLf
    
End Function

Sub Enviar_Potencia()

Dim LV_Pow          As Double
Dim LV_Pow_Cor      As Double

    If Me.chkGeneradorEn.value = 0 Then
        Exit Sub
    End If
    If PV_Flag_Pow = True Then
        
        If Me.chkIncPot.value And IsNumeric(Me.txtAgregarPot.Text) = True Then
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, PV_Pot_Now + Me.txtAgregarPot.Text)
        Else
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, PV_Pot_Now)
        End If
                
        If Me.chkRS232.value = 0 Then
            SendPow 0, LV_Pow_Cor   'PV_Pot_Now - PV_Correccion(PV_Index_Correc)
        Else
            fMainForm.SendRS232 CommandPow(0, LV_Pow_Cor)
        End If
        
        GV_Data_Captur.Data(1) = PV_Pot_Now
        
        If Me.chkRS232.value = 0 Then
            'Verify_Pow_State 0, LV_Pow
        End If
        
        If LV_Pow <> PV_Pot_Now Then
            
            'Me.cmdCancelar.value = True
        
        End If
        
        PV_Flag_Pow = False
        
    End If
    
End Sub

Sub Enviar_Frecuencia_Potencia()

    Enviar_Frecuencia
    
    Enviar_Potencia
    
    If PV_Flag_RF_On = True Then
        PV_Flag_RF_On = False
        SendRFPowerOn 0
        Me.chkRFOn.value = 1
    End If
    
End Sub
    
Function GetFileCorreccion() As String

    'GetFileCorreccion = GV_Actual_Project.Path_Project & "\" & "Correccion Setup.csv"
    GetFileCorreccion = GV_Actual_Project.CompensacionSetup
    
End Function

Sub Guardar_Valores()

Dim LV_Data()         As String
Dim i               As Integer

    If PV_File_Hdl = 0 Then
    
        Crear_Archivo_Salida PV_File_Hdl, GV_Archivo_Salida
        
        Iniciar_List_View
        
    End If
        
    'GV_Data_Captur.Data(3) = GV_Data_Captur.Data(2)
    
    Guardar_Data PV_File_Hdl, GV_Data_Captur.Data
        
    ReDim LV_Data(UBound(GV_Data_Captur.Data))
    
    For i = 0 To UBound(LV_Data)
        LV_Data(i) = GV_Data_Captur.Data(i)
    Next i
    
    AddItemListView Me.LstVwVisualTest, LV_Data
    
End Sub

Sub Incrementar_Paso()

Dim LV_Frec     As Double
Dim LV_Pow      As Double
Dim LV_Pow_L    As Long

    With Me
    
        LV_Frec = PV_Frec_Now
        LV_Pow = PV_Pot_Now
        
        If Me.chkUsarTablaParam.value Then
            If PV_Index_Tabla <= UBound(PV_TablaParam) Then
                CargarValoresFromTablaParam
            Else
                ' Prueba Terminada
                .tmrPrueba.Enabled = False
                Beep
                Beep
                Beep
                If .chkRepetirPrueba.value Then
                    Repetir_Prueba
                Else
                    Cancelar_Prueba
                End If
            End If
        Else
            ' Incrementar Frecuencia
            If PV_Frec_Paso > 0 Then
                PV_Frec_Now = PV_Frec_Now + PV_Frec_Paso
                PV_Index_Correc = PV_Index_Correc + 1
                PV_Flag_Pow = True
                If PV_Frec_Now > PV_Frec_Max Then
                    PV_Frec_Now = PV_Frec_Min
                    ' Incrementar Potencia
                    If PV_Pot_Paso > 0 Then
                        LV_Pow_L = PV_Pot_Now * PV_Factor_Pot + PV_Pot_Paso * PV_Factor_Pot
                        PV_Pot_Now = LV_Pow_L / PV_Factor_Pot
                        PV_Index_Correc = 0
                        PV_Flag_Pow = True
                        ' Verificar Condici�n Final de Rango
                        If PV_Pot_Now > PV_Pot_Max Then
                            ProcesaEventoFinEtapa
                            PV_Index_Correc = 0
                            PV_Flag_Pow = True
                            PV_Etapa = PV_Etapa + 2
                            'PV_EmpezarEtapa = True
                            If PV_Etapa >= GV_Actual_Project.EtapasDeControl Then
                                ' Prueba Terminada
                                .tmrPrueba.Enabled = False
                                Beep
                                Beep
                                Beep
                                If .chkRepetirPrueba.value Then
                                    Repetir_Prueba
                                Else
                                    Cancelar_Prueba
                                End If
                            Else
                                .Cargar_Valores_Etapa PV_Etapa
                                PV_Frec_Now = PV_Frec_Min
                                PV_Pot_Now = PV_Pot_Min
                        
                                PV_Flag_Freq = True
                                PV_Flag_Pow = True
                                
                            End If
                        End If
                    Else
                        ProcesaEventoFinEtapa
                        PV_Etapa = PV_Etapa + 2
                        If PV_Etapa >= GV_Actual_Project.EtapasDeControl Then
                            ' Prueba Terminada
                            .tmrPrueba.Enabled = False
                            If .chkRepetirPrueba.value Then
                                Repetir_Prueba
                            Else
                                Cancelar_Prueba
                            End If
                        Else
                            .Cargar_Valores_Etapa PV_Etapa
                            PV_Frec_Now = PV_Frec_Min
                            PV_Pot_Now = PV_Pot_Min
                    
                            PV_Flag_Freq = True
                            PV_Flag_Pow = True
                        End If
                    End If
                End If
            Else
                If PV_Pot_Paso > 0 Then
                    ' Incrementar Potencia
                    LV_Pow_L = PV_Pot_Now * PV_Factor_Pot + PV_Pot_Paso * PV_Factor_Pot
                    PV_Pot_Now = LV_Pow_L / PV_Factor_Pot
                    ' Verificar Condici�n Final de Rango
                    If PV_Pot_Now > PV_Pot_Max Then
                        ProcesaEventoFinEtapa
                        PV_Index_Correc = 0
                        PV_Flag_Pow = True
                        PV_Etapa = PV_Etapa + 2
                        If PV_Etapa >= GV_Actual_Project.EtapasDeControl Then
                            ' Prueba Terminada
                            .tmrPrueba.Enabled = False
                            If .chkRepetirPrueba.value Then
                                Repetir_Prueba
                            Else
                                Cancelar_Prueba
                            End If
                        Else
                            .Cargar_Valores_Etapa PV_Etapa
                            PV_Frec_Now = PV_Frec_Min
                            PV_Pot_Now = PV_Pot_Min
                    
                            PV_Flag_Freq = True
                            PV_Flag_Pow = True
                        End If
                    End If
                Else
                    ' Fin de Etapa
                    ProcesaEventoFinEtapa
                    PV_Index_Correc = 0
                    PV_Flag_Pow = True
                    PV_Etapa = PV_Etapa + 2
                    If PV_Etapa >= GV_Actual_Project.EtapasDeControl Then
                        ' Prueba Terminada
                        .tmrPrueba.Enabled = False
                        If .chkRepetirPrueba.value Then
                            Repetir_Prueba
                        Else
                            Cancelar_Prueba
                        End If
                    Else
                        .Cargar_Valores_Etapa PV_Etapa
                        PV_Frec_Now = PV_Frec_Min
                        PV_Pot_Now = PV_Pot_Min
                
                        PV_Flag_Freq = True
                        PV_Flag_Pow = True
                    End If
                End If
            End If
        End If

        ' Verificar Frecuencia
        If PV_Frec_Now <> LV_Frec Then
        
            PV_Flag_Freq = True
            
        End If
        
        If PV_Pot_Now <> LV_Pow Then
        
            PV_Flag_Pow = True
            
        End If
        
    End With
    
End Sub

Sub Inicializar_Comandos_Instrumentos()

Dim LV_Cod_Instru()     As Integer
Dim LV_Qty              As Integer
Dim i                   As Integer
Dim LV_Volt_Div         As String
Dim LV_50Ohms           As String
Dim LV_Inverter         As String

    'i = 1
    'ReDim GV_Instrumentos(i)
    
    i = 0
    ReDim GV_Instrumentos(i)
    
    LV_Qty = BD_Get_Cod_Instruments(LV_Cod_Instru, GV_Actual_Project.Cod_Project, 1)
    
    If LV_Qty Then
        ReDim GV_Instrumentos(LV_Qty)
        For i = 0 To LV_Qty
            BD_Fill_Comandos_Instru GV_Instrumentos(1), LV_Cod_Instru(i)
        Next
    End If
    
    'BD_Fill_LstVw_With_Instru_From_Pjt  .LstViewDispositivo(i), GV_Actual_Project.Cod_Project, 1
    
    BD_Fill_Comandos_Instru GV_Instrumentos(i), GV_Actual_Project.Cod_Project
    
    'LV_Cod_Instru = BD_Get_Cod_Instru
    'BD_Get_Comandos_From_
    
    With GV_Instrumentos(i)
    
        .Comunicacion = Type_Communica.COMM_GPIB
        .address = 19
        .address = Me.txtAddressGen.Text
        '.Address = 28
        
        '.Cmd_Config = "UNIT:POW:DBM"
        '.Cmd_Config = "OUTPUT ON"
        .Cmd_Config = ""
        
        ReDim .Cmd_Consult(1)
        ReDim .Cmd_Set_Var(1)
        
        '.Cmd_Set_Var(0) = "FREQ:CW"
        '.Cmd_Set_Var(0) = "FREQ %F MHz"
        .Cmd_Set_Var(0) = "FREQ %FMHz"
        .Cmd_Consult(0) = "FREQ:CW?"
        
        '.Cmd_Set_Var(1) = "POW:LEV"
        '.Cmd_Set_Var(1) = "POW %P dBm"
        .Cmd_Set_Var(1) = "POW %PdBm"
        .Cmd_Consult(1) = "POW:LEV?"
        
        .Cmd_End = ":OUTput1:PON OFF"
    End With
    
    If Me.chkMedirOsciloscopio.value Then
        i = i + 1
        ReDim Preserve GV_Instrumentos(i)
        '----------------------
        ' Osciloscopio
        With GV_Instrumentos(i)
        
            .Name = "Osciloscopio"
            .Comunicacion = COMM_GPIB
            .address = 2
            
            LV_Volt_Div = Trim$(Me.cboOscVoltDiv.Text / 1000#)
            LV_50Ohms = "OFF"
            LV_Inverter = "OFF"
            If Me.chk50Ohm.value Then
                LV_50Ohms = "ON"
            End If
            If Me.chkInvertir.value Then
                LV_Inverter = "ON"
            End If
            
            .Cmd_Config = "DAT:SOU CH1" _
                        & ";:DAT:ENC ASCII" _
                        & ";:DAT:WID 1" _
                        & ";:DAT:STAR 1" _
                        & ";:DAT:STOP 1" _
                        & ";CH1 VOLts:" & LV_Volt_Div _
                        & ";CH1 POSition:-4" _
                        & ";CH1 FIFty:" & LV_50Ohms _
                        & ";CH1 INVERT:" & LV_Inverter _
                        & ";:HOR:MAIN:SCALE 5e-4"
    
            '.Cmd_Config = .Cmd_Config _
                        & ";CH2 VOLts:" & LV_Volt_Div _
                        & ";CH2 POSition:-4" _
                        & ";CH2 FIFty:" & LV_50Ohms _
                        & ";CH2 INVERT:" & LV_Inverter _
                        & ";:HOR:MAIN:SCALE 5e-4"
    
    
            'GV_Volt_Div = 500 / 25
            GV_Volt_Div = Me.cboOscVoltDiv.Text / 25
            GV_Offset = -4
            
            ReDim .Cmd_Consult(0)
            ReDim .Cmd_Set_Var(0)
            '.Cmd_Set_Var(0) = "DAT:SOU CH2"
            'ReDim .Cmd_Set_Var(1)
            
            If Me.optOscNivel(0).value Then
                .Cmd_Consult(0) = "AVG?"
            Else
                .Cmd_Consult(0) = "CURVE?"
            End If
            
        End With
    End If
    
    'Exit Sub
    'i = 2
    If Me.chkMedirAnalizador.value Then
        i = i + 1
        ReDim Preserve GV_Instrumentos(i)
        'i = 1
        '----------------------
        ' Analizador de Espectro
        With GV_Instrumentos(i)
        
            .Comunicacion = COMM_GPIB
            .address = 8
            .Name = "Analizador de Espectro"
            .Cmd_Config = "RL1" _
                        & ";RE " & Trim$(Me.txtRefLvl) & "DB" _
                        & ";AT " & Trim$(Me.txtAtt) & "DB" _
                        & ";SP " & Trim$(Me.txtSpan) & "MZ" _
                        & ";PS"
                        
            '.Cmd_Config = "RL1" _
                        & ";RE 0DBM" _
                        & ";AT 0DB" _
                        & ";SP 20MZ" _
                        & ";PS"
    
    
            '.Cmd_Config = "CF 1900MZ" & vbCrLf
            '.Cmd_Config = "CH1 FIFty:ON"
            
            ReDim .Cmd_Consult(0)
            ReDim .Cmd_Set_Var(0)
            
            '.Cmd_Set_Var(0) = "SP 20MZ ; CF %FMZ ; PS "
            .Cmd_Set_Var(0) = "CF %FMZ ; PS "
            If Me.cboPeakGraph.ListIndex = 0 Then
                .Cmd_Consult(0) = "PS; MFL?"
            Else
                .Cmd_Consult(0) = "PS; PLOT"
            End If
            .TpoEspera = 0
            
            ReDim GV_Analizador_Sp(0)
            
            With GV_Analizador_Sp(0)
                .RefLvl = Me.txtRefLvl
                .CenterFreq = Me.txtCenterFreq
                .SPAN = Me.txtSpan
                
            End With
            
        End With
        '----------------------
        'Exit Sub
        'i = i
    End If
    
    'If i <= UBound(GV_Instrumentos) Then
    If Me.chkMedirPowerMeter.value Then
        i = i + 1
        ReDim Preserve GV_Instrumentos(i)
        ' Power Meter
        With GV_Instrumentos(i)
        
            .Comunicacion = COMM_GPIB
            .address = 13
            
            .Cmd_Config = "LG;OC1"
    '
            ReDim .Cmd_Consult(0)
    '        ReDim .Cmd_Set_Var(1)
    '
    '        .Cmd_Set_Var(0) = "FREQ:CW"
            '.Cmd_Consult(0) = "MEAS1?"
            .Cmd_Consult(0) = "LG"
    '
    '        .Cmd_Set_Var(1) = "POW:LEV"
    '        .Cmd_Consult(1) = "POW:LEV?"
            ReDim .Cmd_Set_Var(0)
            
            .Cmd_Set_Var(0) = "FREQ %F MHz"
        .TpoEspera = 0
            
        End With
    End If

End Sub

Function Iniciar_Instrumentos_RS232() As Boolean

Dim LV_i            As Integer

    For LV_i = 0 To UBound(GV_Instrumentos)
        
        With GV_Instrumentos(LV_i)
            'If .Comunicacion = COMM_RS232 Then
                'IniciarCommPort
                PV_CommPort = fMainForm.IniciarCommPort
                If .Cmd_Config <> "" Then
                    'SendRS232 .Cmd_Config
                    fMainForm.SendRS232 .Cmd_Config
                End If
            'End If
        End With
        
    Next

End Function

Function Inicializar_Comm_GPIB() As Boolean

Dim i           As Integer

    'ReDim PV_Handle_Instruments(2)
    
    Inicializar_Comandos_Instrumentos
    
    If Me.chkRS232.value = 0 Then
        IniciarCommInstrumento
    Else
        Iniciar_Instrumentos_RS232
    End If
    'For i = 0 To 2
    
        'PV_Handle_Instruments(i) =
        
     '   IniciarCommInstrumento (i)
        
'        If PV_Handle_Instruments(i) = 0 Then
'
'            Inicializar_Comm_GPIB = False
'
'            'Exit Function
'
'        End If
        
    'Next
    
    Inicializar_Comm_GPIB = True

End Function

Function IniciarCommPort() As Boolean

    With Me
        If .MSComm.PortOpen = False Then
            .MSComm.CommPort = 2
            .MSComm.PortOpen = True
            IniciarCommPort = True
        Else
            IniciarCommPort = True
        End If
    End With

End Function

Function SendRS232(LV_Cmd As String)

    With Me.MSComm
        If .PortOpen = True Then
            .Output = LV_Cmd
        End If
    End With
    
End Function

Function LeerRS2323() As String

    With Me.MSComm
        If .PortOpen = True Then
            LeerRS2323 = .Input
        End If
    End With
    
End Function

Sub Iniciar_List_View()

    With Me
        With .LstVwVisualTest
            .ColumnHeaders.Clear
            .ListItems.Clear
        End With
        
        AddColumListView .LstVwVisualTest, GV_Data_Captur.NombreCampo
    
    End With
    
End Sub
        
Sub LoadControles()

    BD_Get_Controles_Proyecto
    
    With Me
        .txtAddressGen = GV_Actual_Project.Controles.AddressGPIB
        .chkAdquirir.value = GV_Actual_Project.Controles.Adquirir
        .chkCurvaVideoPot = GV_Actual_Project.Controles.AplicarCurvaVideoPot
        .txtArchCompSal = GV_Actual_Project.Controles.ArchivoCompensaSalida
        .chkCapPotGen = GV_Actual_Project.Controles.CapturarPot
        .chkMedirAnalizador = GV_Actual_Project.Controles.ControlAnalizaEspec
        .chkMedirOsciloscopio = GV_Actual_Project.Controles.ControlOscilos
        .chkMedirPowerMeter = GV_Actual_Project.Controles.ControlPowerMeter
        .chkEsperaEstabi = GV_Actual_Project.Controles.EsperarEstabiliza
        .txtCurvaVideoPot = GV_Actual_Project.Controles.FileCurvaVideoPot
        .txtTablaParam = GV_Actual_Project.Controles.FileTablaParam
        .chkManual = GV_Actual_Project.Controles.OperacionManual
        .txtTpoEspera = GV_Actual_Project.Controles.TpoEspera
        .chkUsarTablaParam = GV_Actual_Project.Controles.UsarTablaParam
    End With
    
End Sub

Sub LoadRangos()

Dim LV_Etapas           As Integer
Dim LV_Campos()         As Integer
Dim i                   As Integer

    Me.LstVwRangoControl.ListItems.Clear
    
    LV_Etapas = BD_Get_Rangos_Control(GV_Actual_Project.Cod_Project, GV_Actual_Project.Rango)

    GV_Actual_Project.EtapasDeControl = LV_Etapas
    
    Me.UpDate_LstVw_Rangos
    
End Sub

Sub ProcesaEventoIniEtapa()

Dim LV_Cmd          As String
Dim LV_Inv          As Integer
Dim i               As Integer

    Me.txtPRI.Text = GV_Actual_Project.Rango(PV_Etapa).PRI
    Me.txtPW.Text = GV_Actual_Project.Rango(PV_Etapa).PW
    
    If Me.chkMedirOsciloscopio.value Then
        If GV_Actual_Project.Rango(PV_Etapa).AplicarPV Then
            Me.chkCurvaVideoPot.value = 1
            LV_Inv = 0
            If Me.chkInvertir.value Then
                LV_Inv = 1
            End If
            With GV_Actual_Project.Rango(PV_Etapa)
                LV_Cmd = ComandoOsciloscopio(.b50Ohms, LV_Inv, .VoltDiv)
                '.CurvaPV
                Cargar_Curva_Video_Pot .CurvaPV
            End With
            For i = 0 To UBound(GV_Instrumentos)
                If GV_Instrumentos(i).Name = "Osciloscopio" Then
                    Call Send(GPIB0, GV_Result_List(i), LV_Cmd, NLend)
                    Exit For
                End If
            Next
        End If
    End If
    
End Sub

Sub ProcesaEventoFinEtapa()

    With GV_Actual_Project.Rango(PV_Etapa)
        If .AccionFinEtapa = "1" Then
            Me.chkRFOn.value = 1
            SendRFPowerOff 0
            Me.tmrPrueba.Enabled = False
            Me.cmdComenzar.Caption = "&Reanudar"
            Me.cmdComenzar.Refresh
        End If
    End With
    
End Sub


Sub Refresh_Column_Header()

    ReDim PV_Column_Header(7)
    
    PV_Column_Header(0) = "Etapa"
    PV_Column_Header(1) = "Par�metro"
    PV_Column_Header(2) = "Valor M�n"
    PV_Column_Header(3) = "Valor M�x"
    PV_Column_Header(4) = "Paso"
    PV_Column_Header(5) = "Unidad"
    PV_Column_Header(6) = "Puntos"
    PV_Column_Header(7) = "Duraci�n"
    
    
    AddColumListView Me.LstVwRangoControl, PV_Column_Header
    
End Sub

Sub Refresh_Estado()

    With Me.txtPrueba(2)
        Select Case PV_Estado
        
            Case Is = 0
                .Text = "Enviando Potencia y Frecuencia..."
                
            Case Is = 1
                .Text = "Esperando Estabilizaci�n de la Medici�n..."
                
            Case Is = 2
                .Text = "Adquiriendo Medidas..."
                
            Case Is = 3
                .Text = "Guardando Medidas..."
                
            Case Is = 4
                .Text = "Incrementando Potencia y Frecuencia..."
                
        End Select
        .Refresh
    End With
    
End Sub
        

Sub Refresh_Valores()

    With Me
        
        .txtPrueba(0).Text = PV_Frec_Now
        .txtPrueba(1).Text = PV_Pot_Now
        
    End With
    
End Sub

Sub UpDateControlesProyecto()
    
    BD_Update_Controles_Proyecto
    
End Sub

Sub UpDate_LstVw_Rangos()

Dim LV_Etapas           As Integer
Dim LV_Campos()         As String
Dim i                   As Integer
Dim LV_Ptos             As Double
Dim LV_Ptos_2           As Double

    PV_Ptos_Prueba = 0
    
    Me.LstVwRangoControl.ListItems.Clear
    
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
                LV_Campos(6) = ""
                If .Paso <> 0 Then
                    LV_Ptos = (.ValorMax - .ValorMin) / .Paso + 1
                Else
                    LV_Ptos = 1
                End If
            End With
            
            AddItemListView Me.LstVwRangoControl, LV_Campos
            
            With GV_Actual_Project.Rango(2 * i - 1)
                LV_Campos(0) = ""
                LV_Campos(1) = .Parametro
                LV_Campos(2) = .ValorMin
                LV_Campos(3) = .ValorMax
                LV_Campos(4) = .Paso
                LV_Campos(5) = .Unidad
                If .Paso <> 0 Then
                    LV_Ptos_2 = (.ValorMax - .ValorMin) / .Paso + 1
                Else
                    LV_Ptos_2 = 1
                End If
                LV_Campos(6) = LV_Ptos_2 * LV_Ptos
                PV_Ptos_Prueba = PV_Ptos_Prueba + LV_Ptos_2 * LV_Ptos
            End With
            
            
            AddItemListView Me.LstVwRangoControl, LV_Campos
            
        Next
        
    End If
    
End Sub

Private Sub cboCenterSpan_Change()

    With Me.cboCenterSpan
        SaveSetting App.Title, "GPIB Config", "Analizador Center-Span", .ListIndex
        Select Case .ListIndex
            Case Is = 0
                Me.lblCenterFreq = "Center Freq:"
                Me.lblSpanFreq = "Span Freq:"
            Case Is = 1
                Me.lblCenterFreq = "Start Freq:"
                Me.lblSpanFreq = "Stop Freq:"
        End Select
    End With
    
End Sub

Private Sub cboOscVoltDiv_Click()

    With Me.cboOscVoltDiv
        SaveSetting App.Title, "GPIB Config", "Osciloscopio Index Volt Div", .ListIndex
    End With
    
End Sub

Private Sub cboPeakGraph_Change()

    With Me.cboPeakGraph
        SaveSetting App.Title, "GPIB Config", "Analizador Peak vs Graph", .ListIndex
    End With
    
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chk50Ohm_Click()

    With Me.chk50Ohm
        SaveSetting App.Title, "GPIB Config", "Osciloscopio 50 Ohms", .value
    End With
    
End Sub

Private Sub chkAdquirir_Click()

    With Me.chkAdquirir
        GV_Actual_Project.Controles.Adquirir = .value
        UpDateControlesProyecto
    End With
    
End Sub

Private Sub chkALCOn_Click()

    With Me.chkALCOn
        If .value Then
            SendALCState 0, True
        Else
            SendALCState 0, False
        End If
    End With
    
End Sub

Private Sub chkCapPotGen_Click()

    With Me.chkCapPotGen
        GV_Actual_Project.Controles.CapturarPot = .value
        UpDateControlesProyecto
    End With

End Sub

Private Sub chkCortarRFalTerminar_Click()

    With Me.chkCortarRFalTerminar
        GV_Actual_Project.Controles.CortarRFalTerminar = .value
        UpDateControlesProyecto
    End With

End Sub

Private Sub chkCurvaVideoPot_Click()

    With Me
        If .chkCurvaVideoPot.value Then
            .txtCurvaVideoPot.Visible = True
            .cmdSelCurvaPot.Visible = True
            If .cmdComenzar.Caption = "&Comenzar" Then
                .cmdSelCurvaPot.value = True
            End If
        Else
            .txtCurvaVideoPot.Visible = False
            .cmdSelCurvaPot.Visible = False
        End If
        GV_Actual_Project.Controles.AplicarCurvaVideoPot = .chkCurvaVideoPot.value
        UpDateControlesProyecto
    End With
    
End Sub

Private Sub chkEsperaEstabi_Click()

    With Me.chkEsperaEstabi
        GV_Actual_Project.Controles.EsperarEstabiliza = .value
        UpDateControlesProyecto
    End With

End Sub

Private Sub chkInvertir_Click()

    With Me.chkInvertir
        SaveSetting App.Title, "GPIB Config", "Osciloscopio Invertir", .value
    End With
    
End Sub

Private Sub chkManual_Click()

    With Me.chkManual
        GV_Actual_Project.Controles.OperacionManual = .value
        UpDateControlesProyecto
    End With

End Sub

Private Sub chkMedirAnalizador_Click()

    With Me.chkMedirAnalizador
        SaveSetting App.Title, "GPIB Config", "Medir Analizador", .value
        GV_Actual_Project.Controles.ControlAnalizaEspec = .value
        UpDateControlesProyecto
    End With
        
End Sub

Private Sub chkMedirOsciloscopio_Click()

    With Me.chkMedirOsciloscopio
        SaveSetting App.Title, "GPIB Config", "Medir Osciloscopio", .value
        GV_Actual_Project.Controles.ControlOscilos = .value
        UpDateControlesProyecto
    End With
    
End Sub

Private Sub chkMedirPowerMeter_Click()

    With Me.chkMedirPowerMeter
        SaveSetting App.Title, "GPIB Config", "Medir Power Meter", .value
        GV_Actual_Project.Controles.ControlPowerMeter = .value
        UpDateControlesProyecto
    End With
    
End Sub

Private Sub chkOscCh_Click()

    With Me.chkOscCh
        SaveSetting App.Title, "GPIB Config", "Osciloscopio Canal 1", .value
    End With
    
End Sub

Private Sub chkPisarArchivoSalida_Click()

    With Me.chkPisarArchivoSalida
        GV_Actual_Project.Controles.PisarArchivo = .value
        UpDateControlesProyecto
    End With

End Sub

Private Sub chkRFOn_Click()

    With Me.chkRFOn
        If .value Then
            PV_Ini_Tmr = GetTickCount
            SendRFPowerOn 0
        Else
            PV_End_Tmr = GetTickCount
            Me.txtCronometro.Text = PV_End_Tmr - PV_Ini_Tmr
            SendRFPowerOff 0
            
        End If
    End With
    
End Sub

Private Sub chkRS232_Click()

    With Me.chkRS232
        If GetSetting(App.Title, "GPIB Config", "Comunicacion RS 232", .value) <> .value Then
            If .value Then
                'Load frmProps
                fMainForm.LoadFormRs232Props
            End If
        End If
        SaveSetting App.Title, "GPIB Config", "Comunicacion RS 232", .value
    End With

End Sub

Private Sub chkUsarTablaParam_Click()

    If Me.chkUsarTablaParam.value Then
        Me.cmdSelTablaParam.value = True
    End If
    With Me.chkUsarTablaParam
        GV_Actual_Project.Controles.UsarTablaParam = .value
        UpDateControlesProyecto
    End With
    
End Sub


Private Sub cmdBajarPot_Click(Index As Integer)

Dim LV_Pow_Cor      As Double
Dim LV_Inc          As Double

    Select Case Index
        Case Is = 0
            LV_Inc = 1
        Case Is = 1
            LV_Inc = 5
        Case Is = 2
            LV_Inc = 0.1
    End Select
    
    With Me
        If .cmdComenzar.Caption <> "&Comenzar" Then
            LV_Pow_Cor = .txtPrueba(1).Text
            LV_Pow_Cor = LV_Pow_Cor - LV_Inc
            .txtPrueba(1).Text = LV_Pow_Cor
            
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, LV_Pow_Cor)
            
            If Me.chkRS232.value = 0 Then
                SendPow 0, LV_Pow_Cor   'PV_Pot_Now - PV_Correccion(PV_Index_Correc)
            Else
                fMainForm.SendRS232 CommandPow(0, LV_Pow_Cor)
            End If
        End If
    End With

End Sub

'-----------------------------------------------------------------------
'-----------------------------------------------------------------------
'-----------------------------------------------------------------------

Private Sub cmdCancelar_Click()

    With Me
    
        If MsgBox("�Est� seguro que desea Cancelar la Medici�n?", vbYesNo) = vbYes Then
        
            .Cancelar_Prueba
            
        End If
        
    End With
    
End Sub

Private Sub cmdComenzar_Click()

    With Me
    
        If .cmdComenzar.Caption = "&Comenzar" Then
        
            .chkRS232.Enabled = False
            PV_Flag_RF_On = True
            'ReDim PV_Correccion(0)
            'PV_Correccion(0) = 0
            PV_Index_Tabla = 0
            PV_EmpezarEtapa = True
            GV_Ch_Decimal = Get_Decimal_From_Regional_Config
            
            .LstVwVisualTest.ListItems.Clear
            
            PV_Estado = 0
            PV_Ptos_Now = 0
            
            If Me.chkUsarTablaParam.value Then
                CargarTablaParam
            Else
                ReDim PV_TablaParam(0)
            End If
            
            Cargar_Valores
            
            ' Carga Curva Correcci�n Setup
            Cargar_Correccion
            ' Compensacion de la Salida
            Cargar_Compensa
            ' Medici�n con Diodo Pin
            Cargar_Curva_Video_Pot
            
            Crear_Estructuras
            
            .cmdCancelar.Enabled = True
            
            .cmdComenzar.Caption = "&Pausar"
            
            .txtPRI.Text = ""
            .txtPW.Text = ""
            
            Inicializar_Comm_GPIB
            
            If Me.chkRS232.value And PV_CommPort = False Then
                MsgBox "Existe un Problema con el Puerto Serial. La prueba Finalizar�.", vbYes
                .cmdCancelar.value = True
                Exit Sub
            End If
            
            GV_Archivo_Salida = Obtener_Archivo_Salida
            
            If Me.chkPisarArchivoSalida.value = 0 And VerificarExiste(GV_Archivo_Salida) = True Then
                If MsgBox("El archivo existe. �Desea continuar?", vbYesNo) = vbNo Then
                    .cmdCancelar.value = True
                    Exit Sub
                End If
            End If
            
            PV_Tpo_Ini = Timer
            
            .tmrPrueba.Enabled = True
            
        ElseIf .cmdComenzar.Caption = "&Pausar" Then
        
            .cmdComenzar.Caption = "&Reanudar"
            
            .tmrPrueba.Enabled = False
        
        Else
            
            .cmdComenzar.Caption = "&Pausar"
            
            .tmrPrueba.Enabled = True
            PV_Flag_RF_On = True
            'If .chkRFOn.value = 0 Then
            '    .chkRFOn.value = 1
            '    SendRFPowerOn 0
            'End If
            
        End If
        
    End With
    
End Sub

Private Sub cmdDecFreq_Click()

Dim LV_F        As Double
Dim LV_F_I      As Double
Dim i           As Integer
    
        LV_F = Me.txtPrueba(0).Text
        LV_F = LV_F - Me.txtStepFreq.Text / 1000#
        Me.txtPrueba(0).Text = LV_F
        
        'LV_F = LV_F * 1000000#
        
        If Me.chkRS232.value = 0 Then
            SendFrec 0, LV_F
        Else
            fMainForm.SendRS232 CommandFrec(0, LV_F)
        End If
        
        If Me.chkRS232.value = 0 Then
            Verify_Frec_State 0, LV_F_I
        Else
        End If
        
        If LV_F_I <> LV_F Then
            
            'Me.cmdCancelar.value = True
        
        End If
        

End Sub

Private Sub cmdIncFreq_Click()

Dim LV_F        As Double
Dim LV_F_I      As Double
Dim i           As Integer
    
        LV_F = Me.txtPrueba(0).Text
        LV_F = LV_F + Me.txtStepFreq.Text / 1000#
        Me.txtPrueba(0).Text = LV_F
        
        'LV_F = LV_F * 1000000#
        
        If Me.chkRS232.value = 0 Then
            SendFrec 0, LV_F
        Else
            fMainForm.SendRS232 CommandFrec(0, LV_F)
        End If
        
        If Me.chkRS232.value = 0 Then
            Verify_Frec_State 0, LV_F_I
        Else
        End If
        
        If LV_F_I <> LV_F Then
            
            'Me.cmdCancelar.value = True
        
        End If
        

End Sub

Private Sub cmdSelArchiCompSal_Click()

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    lPath = GV_Actual_Project.Path_Project
    With Me.CommonDialog
        .Filter = "*.csv"
        .InitDir = lPath
        .CancelError = False
        .DialogTitle = "Archivo de COmpensaci�n de Salida"
        .ShowOpen
        sFile = .fileName
    End With
    'sFile = BrowseForFile(lPath, "*.csv", "Archivo de COmpensaci�n de Salida")
    Me.txtArchCompSal.Text = sFile
    
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
    'sFile = BrowseForFile(lPath, "*.csv", "Archivo de COmpensaci�n de Salida")
    Me.txtCurvaVideoPot.Text = sFile

End Sub


Private Sub cmdSelTablaParam_Click()

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    lPath = GV_Actual_Project.Path_Project
    With Me.CommonDialog
        .Filter = "*.csv"
        .InitDir = lPath
        .CancelError = False
        .DialogTitle = "Selecci�n Tabla de Par�metros"
        .ShowOpen
        sFile = .fileName
    End With
    'sFile = BrowseForFile(lPath, "*.csv", "Archivo de COmpensaci�n de Salida")
    Me.txtTablaParam.Text = sFile
    
End Sub


Private Sub cmdSubirPot_Click(Index As Integer)

Dim LV_Pow_Cor      As Double
Dim LV_Inc          As Double

    Select Case Index
        Case Is = 0
            LV_Inc = 1
        Case Is = 1
            LV_Inc = 5
        Case Is = 2
            LV_Inc = 0.1
    End Select
    With Me
        If .cmdComenzar.Caption <> "&Comenzar" Then
            LV_Pow_Cor = .txtPrueba(1).Text
            LV_Pow_Cor = LV_Pow_Cor + LV_Inc
            If Calcular_Correccion(PV_Frec_Now, LV_Pow_Cor) > 15 Then
                Exit Sub
            End If
            .txtPrueba(1).Text = LV_Pow_Cor
            
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, LV_Pow_Cor)
            
            If Me.chkRS232.value = 0 Then
                SendPow 0, LV_Pow_Cor   'PV_Pot_Now - PV_Correccion(PV_Index_Correc)
            Else
                fMainForm.SendRS232 CommandPow(0, LV_Pow_Cor)
            End If
        End If
    End With

End Sub

Private Sub Form_GotFocus()

    If GV_Actual_Project.Flag_NewMeasure = True Then
        GV_Actual_Project.Flag_NewMeasure = False
        With Me
            .LstVwRangoControl.ListItems.Clear
            .Refresh_Column_Header
            .LoadRangos
            .UpDate_LstVw_Rangos
            .LoadControles
        End With
    End If
    
End Sub

Private Sub Form_Load()

Dim i           As Integer

    With Me
        PV_Estado = 0
        .txtStepFreq.Text = "250"
        .Refresh_Column_Header
        .LoadRangos
        .UpDate_LstVw_Rangos
        .LoadControles
        .Cargar_Valores
        '.chkEsperaEstabi.value = 1
        '.chkAdquirir.value = 1
        '.txtTpoEspera = GetSetting(App.Title, "Properties", "Tiempo Espera Estabilizacion", 10)
        '.txtAddressGen.Text = GetSetting(App.Title, "GPIB Config", "Address Generator", 19)
        '.chkMedirAnalizador.value = GetSetting(App.Title, "GPIB Config", "Medir Analizador", 0)
        '.chkMedirOsciloscopio.value = GetSetting(App.Title, "GPIB Config", "Medir Osciloscopio", 0)
        '.chkMedirPowerMeter.value = GetSetting(App.Title, "GPIB Config", "Medir Power Meter", 0)
        .chk50Ohm = GetSetting(App.Title, "GPIB Config", "Osciloscopio 50 Ohms", 0)
        .chkInvertir = GetSetting(App.Title, "GPIB Config", "Osciloscopio Invertir", 0)
        .chkOscCh = GetSetting(App.Title, "GPIB Config", "Osciloscopio Canal 1", 0)
        .chkRS232.value = GetSetting(App.Title, "GPIB Config", "Comunicacion RS 232", 0)
        .txtAgregarPot.Text = GetSetting(App.Title, "Properties", .txtAgregarPot.Name, 0)
        
        For i = 0 To .optOscNivel.UBound
            .optOscNivel(i).value = GetSetting(App.Title, "GPIB Config", "Osciloscopio Opci�n " & i, 0)
        Next
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
        .cboOscVoltDiv.ListIndex = GetSetting(App.Title, "GPIB Config", "Osciloscopio Index Volt Div", 0)
        .txtCenterFreq = GetSetting(App.Title, "GPIB Config", "SA Center Freq", 1000000000)
        .txtSpan = GetSetting(App.Title, "GPIB Config", "SA Span", 100000000)
        .txtRefLvl = GetSetting(App.Title, "GPIB Config", "SA Ref Level", -10)
        .txtAtt = GetSetting(App.Title, "GPIB Config", "SA Att", 20)
        
        With .cboCenterSpan
            .AddItem "Center-Span"
            .AddItem "Start-Stop"
            .ListIndex = GetSetting(App.Title, "GPIB Config", "Analizador Center-Span", 0)
        End With
        
        With .cboPeakGraph
            .AddItem "Peak Measure"
            .AddItem "Graph Measure"
            .ListIndex = GetSetting(App.Title, "GPIB Config", "Analizador Peak vs Graph", 0)
        End With
        
        With .txtPRI
            .Text = GetSetting(App.Title, "Generador Config", "PRI", 1000)
        End With
        
        With .txtPW
            .Text = GetSetting(App.Title, "Generador Config", "PW", 1)
        End With
        .chkGeneradorEn.value = 1
    End With
        
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Close_Devices
    
    Me.tmrPrueba.Enabled = False
    
End Sub

Private Sub Form_Resize()

Dim i           As Integer
Dim lHeight     As Long

    With Me
        
        For i = 0 To .framePrueba.UBound
        
            .framePrueba(i).Width = .Width - 2 * .framePrueba(i).Left
            
        Next
        
        i = 2
        
        .framePrueba(i).Top = .Height - .framePrueba(i).Height
        .cmdComenzar.Left = .Width - .cmdCancelar.Left - .cmdComenzar.Width
        '.cmdModificar(i).Left = (.Width - .cmdModificar(i).Width) / 2
        
        
        i = 1
        
        .framePrueba(i).Top = .framePrueba(i - 1).Height
        .framePrueba(i).Height = .framePrueba(i + 1).Top - .framePrueba(i - 1).Height
        
        .LstVwVisualTest.Width = .framePrueba(i).Width - 2 * .LstVwVisualTest.Left
        .LstVwVisualTest.Height = .framePrueba(i).Height _
                                     - 2 * .LstVwVisualTest.Top
        
        .LstVwRangoControl.Width = .framePrueba(0).Width - 2 * .LstVwRangoControl.Left
        'lHeight = (.Height - .framePrueba(i).Height) / .framePrueba.UBound
        
        .frameControlGral.Width = .framePrueba(i).Width - 2 * .frameControlGral.Left
        .frameParamGPIB.Width = .frameControlGral.Width
    End With
    
End Sub

Private Sub Text1_Change()

End Sub

Private Sub optOscNivel_Click(Index As Integer)

Dim i           As Integer

    For i = 0 To Me.optOscNivel.UBound
        With Me.optOscNivel(i)
            SaveSetting App.Title, "GPIB Config", "Osciloscopio Opci�n " & i, .value
        End With
    Next
    
End Sub

Private Sub tmrPrueba_Timer()

Dim LV_Enable           As Boolean

    With Me
    
        If PV_EmpezarEtapa = True Then
            PV_EmpezarEtapa = False
            ProcesaEventoIniEtapa
        End If
        
        Refresh_Estado
        
        Select Case PV_Estado
        
            Case Is = 0
                Enviar_Frecuencia_Potencia
                
                PV_Ptos_Now = PV_Ptos_Now + 1
                
                If IsNumeric(Me.txtTpoEspera.Text) = True Then
                    PV_Estabiliza_Counter = Me.txtTpoEspera.Text
                Else
                    PV_Estabiliza_Counter = 10
                End If
                
                PV_Estado = PV_Estado + 1
                
                If Me.chkManual.value = 1 Then
                    Me.cmdComenzar.value = True
                End If
                
                PV_TckCnt_E1 = GetTickCount
                LV_Enable = False
                
            Case Is = 1
                'Estabilizando Medidas
                PV_Estabiliza_Counter = PV_Estabiliza_Counter - 1
                'If PV_Estabiliza_Counter <= 0 Or Me.chkEsperaEstabi.value = 0 Then
                '    PV_Estado = PV_Estado + 1
                'End If
                PV_Estado = PV_Estado + 1
                
            Case Is = 2
            
                If Me.chkEsperaEstabi.value = 0 Then
                    LV_Enable = True
                ElseIf IsNumeric(Me.txtTpoEspera.Text) Then
                    If GetTickCount - PV_TckCnt_E1 >= Me.txtTpoEspera.Text Then
                        LV_Enable = True
                    End If
                Else
                    DoEvents
                End If
                If LV_Enable = True Then
                    If Me.chkAdquirir.value Then
                        If Capturar_Valores = True Then
                            PV_Estado = PV_Estado + 1
                        End If
                    Else
                        PV_Estado = PV_Estado + 1
                    End If
                    LV_Enable = False
                End If
                
            Case Is = 3
                Guardar_Valores
            
                PV_Estado = PV_Estado + 1
            
            Case Is = 4
                Incrementar_Paso
            
                Refresh_Valores
                
                fMainForm.Refresca_Contador_Tpo PV_Tpo_Ini, PV_Ptos_Prueba, PV_Ptos_Now
                
                PV_Estado = 0
        End Select
    End With
    
End Sub

Private Sub txtAddressGen_Change()

    With Me.txtAddressGen
        SaveSetting App.Title, "GPIB Config", "Address Generator", .Text
        GV_Actual_Project.Controles.AddressGPIB = .Text
        UpDateControlesProyecto
    End With
    
End Sub

Private Sub txtAgregarPot_Change()

    With Me.txtAgregarPot
        If IsNumeric(.Text) = True Then
            SaveSetting App.Title, "Properties", .Name, .Text
            'GV_Actual_Project.Controles.TpoEspera = .Text
            'UpDateControlesProyecto
        End If
    End With

End Sub

Private Sub txtArchCompSal_Change()

    With Me.txtArchCompSal
        GV_Actual_Project.Controles.ArchivoCompensaSalida = .Text
        UpDateControlesProyecto
    End With

End Sub

Private Sub txtAtt_Change()

    With Me.txtAtt
        SaveSetting App.Title, "GPIB Config", "SA Att", .Text
    End With
End Sub

Private Sub txtCenterFreq_Change()

    With Me.txtCenterFreq
        SaveSetting App.Title, "GPIB Config", "SA Center Freq", .Text
    End With
    
End Sub


Private Sub txtCurvaVideoPot_Change()

    With Me.txtCurvaVideoPot
        GV_Actual_Project.Controles.FileCurvaVideoPot = .Text
        UpDateControlesProyecto
    End With

End Sub

Private Sub txtPRI_Change()

    If GPIBglobalsRegistered = 0 Then
        Exit Sub
    End If
    
    With Me.txtPRI
        If IsNumeric(.Text) = True Then
            If Val(.Text) > 0 Then
                SendPRI 0, .Text
                SaveSetting App.Title, "Generador Config", "PRI", .Text
            End If
        End If
    End With
    
End Sub

Private Sub txtPW_Change()

    If GPIBglobalsRegistered = 0 Then
        Exit Sub
    End If
    
    With Me.txtPW
        If IsNumeric(.Text) = True Then
            If Val(.Text) > 0 Then
                If Val(.Text) <= 0.22 Then
                    Me.chkALCOn.value = 0
                    SendALCState 0, False
                Else
                    Me.chkALCOn.value = 1
                    SendALCState 0, True
                End If
                SendPW 0, .Text
                SaveSetting App.Title, "Generador Config", "PW", .Text
            End If
        End If
    End With

End Sub

Private Sub txtRefLvl_Change()

    With Me.txtRefLvl
        SaveSetting App.Title, "GPIB Config", "SA Ref Level", .Text
    End With

End Sub

Private Sub txtSpan_Change()

    With Me.txtSpan
        SaveSetting App.Title, "GPIB Config", "SA Span", .Text
    End With
    
End Sub

Private Sub txtTablaParam_Change()

    With Me.txtTablaParam
        GV_Actual_Project.Controles.FileTablaParam = .Text
        UpDateControlesProyecto
    End With

End Sub

Private Sub txtTpoEspera_Change()

    With Me.txtTpoEspera
        If IsNumeric(.Text) = True Then
            SaveSetting App.Title, "Properties", "Tiempo Espera Estabilizacion", .Text
            GV_Actual_Project.Controles.TpoEspera = .Text
            UpDateControlesProyecto
        End If
    End With
    
End Sub