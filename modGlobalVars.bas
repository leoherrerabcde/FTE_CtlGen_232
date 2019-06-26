Attribute VB_Name = "modGlobalVars"
Option Explicit

Type TypeFilas
    Filas()             As Double
End Type

Global GV_hWnd_Mdi              As Long

Type Type_ParamControl
    Frec                As Long
    Pot_Gen             As Double
End Type

Type Type_Correccion
    Correccion          As Double
    Freq                As Long
End Type

Type Type_Comando
    
    Cod_Instrumento     As Integer
    Cod_Parametro       As Integer
    Cod_Comunicacion    As Integer
    Cod_Funcion         As Integer
    Comando             As String
    Valor               As Double
    
End Type

Type Type_RangoControl

    Parametro           As String
    Unidad              As String
    Cod_Parametro       As Integer
    Cod_Dimension       As Integer
    ValorMin            As Double
    ValorMax            As Double
    Paso                As Double
    Etapa               As Integer
    AplicarPV           As Integer
    CurvaPV             As String
    b50Ohms             As Integer
    VoltDiv             As Integer
    PRI                 As Double
    PW                  As Double
    AccionFinEtapa      As String
    
End Type

Type type_Propiedades_Control
    AplicarCurvaVideoPot    As Integer
    EsperarEstabiliza       As Integer
    TpoEspera               As Integer
    Adquirir                As Integer
    CapturarPot             As Integer
    AddressGPIB             As Integer
    ControlOscilos          As Integer
    ControlPowerMeter       As Integer
    ControlAnalizaEspec     As Integer
    FileCurvaVideoPot       As String
    OperacionManual         As Integer
    ArchivoCompensaSalida   As String
    UsarTablaParam          As Integer
    FileTablaParam          As String
    CortarRFalTerminar      As Integer
    PisarArchivo            As Integer
End Type

Type Type_Project_Struct

    Flag_NewMeasure             As Boolean
    Flag_UpDate                 As Boolean
    Cod_Project                 As Integer
    Cod_Lista_Instrumentos      As Integer
    Cod_Lista_Rangos            As Integer
    
    Project_Name                As String
    Path_Project                As String
    Dispositivo                 As String
    Num_Parte                   As String
    Num_Serie                   As String
    Fecha                       As String
    Result_File                 As String
    Encargado                   As String
    OCompra                     As String
    CompensacionSetup           As String
    
    Rango()                     As Type_RangoControl
    EtapasDeControl             As Integer
    Controles                   As type_Propiedades_Control
    
    Contro_GPIB                 As Boolean
    
End Type

Type Type_Labels
    
    Etiqueta()                  As String
    
End Type


Type Type_Ptos_Charac

    Freq            As Double
    Pot_In          As Double
    Pot_Out         As Double
    Gain            As Double
    
End Type


Global GV_Actual_Project        As Type_Project_Struct
Global GV_Lbl()                 As Type_Labels
Global GV_Cod_Instru            As Integer
Global GV_Cod_Compo             As Integer
Global GV_Ptos_Comp()           As Type_Ptos_Charac

' Nombres Generales
Global GV_Lbls_Componentes()    As String
Global GV_Lbls_Comunicacion()   As String
Global GV_Lbls_Dimensiones()    As String
Global GV_Lbls_Dispositivos()   As String
Global GV_Lbls_Fabricantes()    As String
Global GV_Lbls_Fn_Comunica()    As String
Global GV_Lbls_Funciones()      As String
Global GV_Lbls_Instrumentos()   As String
Global GV_Lbls_Parametros()     As String

' Señalizar Agregar Instrumentos
Global GV_Flag_Add_Prueba       As Boolean
Global GV_Flag_Add_Exita        As Boolean
Global GV_Flag_Add_Medicion     As Boolean
Global GV_Flag_Add_Accesorio    As Boolean


Global Const PARAM_NEW_INSTRUMENT = 1
Global Const PARAM_MODIFI_INSTRUMENT = 2
'Global Const PARAM_NEW_INSTRUMENT = 1
'Global Const PARAM_NEW_INSTRUMENT = 1
'Global Const PARAM_NEW_INSTRUMENT = 1

Global Const Project_Name = 1
Global Const PROJECT_ = 2

Global Const INDEX_PARAMETER = 4


Global GV_Ch_Decimal            As String

Global GV_Volt_Div              As Double
Global GV_Offset                As Double

Global GV_Tabla_Vid_Pot()       As TypeFilas
Global GV_Lista_Frec()          As Long
Global GV_Lista_Pot()           As Double

Global GV_Data_Instrument_Ok    As Boolean
Global GV_Data_Instrument       As String

Global GV_fMainForm             As formTextDisplay

