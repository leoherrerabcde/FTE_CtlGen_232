Attribute VB_Name = "modManejoBaseDatos"
Option Explicit

Enum E_NUM_TABLA
    
    LISTA_COMPONENTES
    LISTA_INSTRUMENTOS
    LISTA_PARAMCONTROL
    LISTA_PROYECTOS
    TABLA_COMANDOS
    TABLA_COMANDOS_GPIB
    Tabla_Instrumentos
    TABLA_PARAM_COMPONENTES
    TABLA_PRECISIONES
    TIPO_COMPONENTE
    Tipo_Comunicacion
    TIPO_DIMENSION
    TIPO_DISPOSITIVOS
    TIPO_FABRICANTES
    TIPO_FUNCIONES
    TIPO_INSTRUMENTOS
    TIPO_PARAMETROS
    
End Enum

Enum E_NUM_CAMPO

    CAMPO_CODIGO
    CAMPO_NOMBRE
    CAMPO_UBICACION
    CAMPO_DISPOSITIVO
    CAMPO_PJT_NUM_PARTE
    CAMPO_PJT_NUM_SERIE
    CAMPO_FECHA
    CAMPO_RESULT_FILE
    CAMPO_OCOMPRA
    CAMPO_ENCARGADO
    CAMPO_COD_LISTAINSTRU
    CAMPO_COD_LISTACOMPON
    CAMPO_COD_LISTAPARAM
    'CAMPO_
    'CAMPO_
    
End Enum

Enum E_NUM_CAMPO_TBL_INSTRUMENT

    CAMPO_CODIGO_INSTRUMENT
    CAMPO_COD_COMPONENTE
    CAMPO_COD_INSTRUMENTO
    CAMPO_COD_COMUNICACION
    CAMPO_COD_DISPOSITIVO
    CAMPO_COD_FABRICANTE
    CAMPO_COD_FUNCION
    CAMPO_MODELO
    CAMPO_NUM_PARTE
    CAMPO_NUM_SERIE
    CAMPO_COD_TBL_PARAM
    CAMPO_COD_TBL_CMDS
    
End Enum

Global Const INDEX_DIMENSION_DBM = 5
Global Const INDEX_DIMENSION_GHZ = 0
'Global GV_Index_Proyecto            As Integer
'
'Enum Enum_Tabla
'
'    Tipo_Componentes
'    Tipo_Comunicacion
'    Tipo_Dimensiones
'    Tipo_Dispositivos
'    Tipo_Fabricantes
'    Tipo_Funciones
'    Tipo_Instrumentos
'    Tipo_Parametros
'
'End Enum
'
'
'
'


'Sub BD_Add_Rango_Control()
'
'End Sub

Function BD_Agregar_Instrumento(Cod_Compo As Integer, _
            Cod_Instrumento As Integer, _
            Cod_Comu As Integer, _
            Cod_Dispo As Integer, _
            Cod_Fabri As Integer, _
            Cod_Func As Integer, _
            Modelo As String, _
            NumParte As String, _
            NumSerie As String, _
            Cod_TabParam As Integer, _
            Cod_Tab_Command As Integer, _
            address As Integer _
            ) As Integer

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Instrum      As Integer

    LV_Nombre_Tabla = "Tabla_Instrumentos"
    
    

    lv_Sql = " INSERT INTO " & LV_Nombre_Tabla _
            & " (Cod_Componente,Cod_Instrumento,Cod_Comunicacion,Cod_Dispositivo" _
            & ",Cod_Fabricante,Cod_Funcion,Modelo,Numero_Parte" _
            & ",Numero_Serie,GPIB_Address,Cod_Tabla_Param,Cod_Tabla_Commands) " _
            & "  VALUES ((" & Cod_Compo & ")" _
            & ",(" & Cod_Instrumento & ")" _
            & ",(" & Cod_Comu & ")" _
            & ",(" & Cod_Dispo & ")" _
            & ",(" & Cod_Fabri & ")" _
            & ",(" & Cod_Func & ")" _
            & ",('" & Modelo & "')" _
            & ",('" & NumParte & "')" _
            & ",('" & NumSerie & "')" _
            & ",('" & address & "')" _
            & ",(" & Cod_TabParam & ")" _
            & ",(" & Cod_Tab_Command & ") );"
        
    GV_BD_INSTRUMENT.Execute (lv_Sql)

    LV_Cod_Instrum = BD_Get_Cod_Instrument(Cod_Compo, _
                        Cod_Instrumento, _
                        Cod_Comu, _
                        Cod_Dispo, _
                        Cod_Fabri, _
                        Cod_Func, _
                        Modelo, _
                        NumParte, _
                        NumSerie)
    
    BD_Agregar_Instrumento = LV_Cod_Instrum
    
    BD_Fill_Labels
    
    'BD_Agregar_Instru_To_Project GV_Actual_Project.Cod_Project, LV_Cod_Instrum
    
End Function

'Sub BD_Add_Cod_Intrument_To_Project(LV_Cod_Project As Integer, LV_Cod_Instru As Integer)
'
'Dim lv_Sql          As String
'Dim LV_Query        As Recordset
'Dim LV_Nombre_Tabla     As String
'Dim LV_Cod_Instrum      As Integer
'
'    If LV_Cod_Project = 0 Then
'        Exit Sub
'    End If
'
'    LV_Nombre_Tabla = "Lista_Instrumentos"
'
'    lv_Sql = " INSERT INTO " & LV_Nombre_Tabla _
'            & " (Cod_Proyecto,Cod_Instrumento) " _
'            & "  VALUES ((" & LV_Cod_Project & ")" _
'            & ",(" & LV_Cod_Instru & ") );"
'
'    GV_BD_INSTRUMENT.Execute (lv_Sql)
'
'End Sub
Sub BD_Agregar_Comandos(LV_Cmds() As Type_Comando)

Dim i       As Integer
Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = "Tabla_Comandos"
    
    For i = 0 To UBound(LV_Cmds)
        With LV_Cmds(i)
            lv_Sql = " INSERT INTO " & LV_Nombre_Tabla _
                    & " (Cod_Instrumento,Cod_Parametro,Cod_Funcion " _
                    & ",Cod_Comunicacion,Comando) " _
                    & "  VALUES ((" & .Cod_Instrumento & ")" _
                    & ",(" & .Cod_Parametro & ")" _
                    & ",(" & .Cod_Funcion & ")" _
                    & ",(" & .Cod_Comunicacion & ")" _
                    & ",(" & .Comando & ") );"
                
            GV_BD_INSTRUMENT.Execute (lv_Sql)
        End With
    Next
    
End Sub

Sub BD_Agregar_Instru_To_Project(LV_Cod_Project As Integer, LV_Cod_Instru As Integer)

Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Funcion      As Integer

    If LV_Cod_Project = 0 Then
        Exit Sub
    End If
    
    LV_Nombre_Tabla = "Lista_Instrumentos"
    
    If GV_Flag_Add_Prueba = True Then
        LV_Cod_Funcion = 3
    ElseIf GV_Flag_Add_Exita = True Then
        LV_Cod_Funcion = 1
    ElseIf GV_Flag_Add_Medicion = True Then
        LV_Cod_Funcion = 2
    ElseIf GV_Flag_Add_Accesorio = True Then
        LV_Cod_Funcion = 4
    Else
        Exit Sub
    End If
    
    lv_Sql = " INSERT INTO " & LV_Nombre_Tabla _
            & " (Cod_Proyecto,Cod_Instrumento,Cod_Funcion) " _
            & "  VALUES ((" & LV_Cod_Project & ")" _
            & ",(" & LV_Cod_Instru & ")" _
            & ",(" & LV_Cod_Funcion & ") );"
        
    GV_BD_INSTRUMENT.Execute (lv_Sql)

End Sub

Sub BD_Agregar_Proyecto( _
            Nombre As String, _
            Ubicacion As String, _
            Disp_Prueba As String, _
            Num_Parte As String, _
            Num_Serie As String, _
            Fecha As String, _
            Result_File As String, _
            OCompra As String, _
            Encargado As String, _
            LV_CompSetup As String)

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String


    LV_Nombre_Tabla = "Proyectos"
    
    

    lv_Sql = " INSERT INTO " & LV_Nombre_Tabla _
            & " (Nombre,Ubicacion,Disp_Prueba" _
            & ",Num_Parte,Num_Serie,Fecha,Result_File" _
            & ",OrdenCompra,Encargado) " _
            & "  VALUES (('" & Nombre & "')" _
            & ",('" & Ubicacion & "')" _
            & ",('" & Disp_Prueba & "')" _
            & ",('" & Num_Parte & "')" _
            & ",('" & Num_Serie & "')" _
            & ",('" & Fecha & "')" _
            & ",('" & Result_File & "')" _
            & ",('" & OCompra & "')" _
            & ",('" & Encargado & "'));" '_
            '& ",('" & LV_CompSetup & "') );"
            
    GV_BD_INSTRUMENT.Execute (lv_Sql)

End Sub

Sub Agregar_NuevoTipo_en_BD(ByVal IndexTipo As Integer, lsData As String)

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = BD_Nombre_Tabla(IndexTipo)
    
    If LV_Nombre_Tabla = "" Then
        
        Exit Sub
    
    End If
    
    If BD_Index_Elemento(IndexTipo, lsData) <> -1 Then
    
        ' Elemento Existente
        Exit Sub
        
    End If
    
    lv_Sql = " INSERT INTO " & LV_Nombre_Tabla & " (Nombre) " _
            & "  VALUES (('" & lsData _
            & "') );"
    
    GV_BD_INSTRUMENT.Execute (lv_Sql)

End Sub

Function BD_Buscar_Instrumentos_Project(IndexPjt As Integer) As Integer

End Function

Function BD_Fill_Cbo_With_Parameters(LV_Cbo As ComboBox) As Boolean

Dim i               As Integer
Dim LV_Cod          As Integer
Dim LV_Cod2         As Integer
Dim LV_Campos()     As String
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Campo        As String
Dim LV_Query        As Recordset

    LV_Nombre_Tabla = "Tipo_Parametros"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla & ";"
    
    LV_Cbo.Clear
    BD_Fill_Cbo_With_Parameters = False
    
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            ReDim LV_Campos(7)
            
            .MoveFirst
            
            Do
                If !Codigo >= INDEX_PARAMETER + 1 Then
                    
                    LV_Cbo.AddItem !Nombre
                    BD_Fill_Cbo_With_Parameters = True
                    
                End If
                
                .MoveNext
            
            Loop Until .EOF Or .BOF
            
        End If
        
    End With
    
    
    
End Function

Function BD_Fill_Labels()

Dim Index               As Integer
Dim LV_Nombre_Tabla     As String
    

    ReDim GV_Lbl(8)
    
    ' Dispositivos
    Index = Enum_Orden_Tabla.Tipo_Disp
    LV_Nombre_Tabla = BD_Nombre_Tabla(Index)
    BD_Leer_Tipos LV_Nombre_Tabla, GV_Lbl(Index).Etiqueta
    
    ' Funciones
    Index = Enum_Orden_Tabla.Funcion
    LV_Nombre_Tabla = BD_Nombre_Tabla(Index)
    BD_Leer_Tipos LV_Nombre_Tabla, GV_Lbl(Index).Etiqueta
    
    ' Componentes
    Index = Enum_Orden_Tabla.Clase_Componen
    LV_Nombre_Tabla = BD_Nombre_Tabla(Index)
    BD_Leer_Tipos LV_Nombre_Tabla, GV_Lbl(Index).Etiqueta
    
    ' Fabricantes
    Index = Enum_Orden_Tabla.Fabricante
    LV_Nombre_Tabla = BD_Nombre_Tabla(Index)
    BD_Leer_Tipos LV_Nombre_Tabla, GV_Lbl(Index).Etiqueta
    
    ' Comunicacion
    Index = Enum_Orden_Tabla.Comunicacion
    LV_Nombre_Tabla = BD_Nombre_Tabla(Index)
    BD_Leer_Tipos LV_Nombre_Tabla, GV_Lbl(Index).Etiqueta
    
    ' Intrumentos
    Index = Enum_Orden_Tabla.Clase_Instru
    LV_Nombre_Tabla = BD_Nombre_Tabla(Index)
    BD_Leer_Tipos LV_Nombre_Tabla, GV_Lbl(Index).Etiqueta
    
    
End Function

Function BD_Fill_LstVw_Instruments(LV_LstVw As ListView)

Dim i               As Integer
Dim LV_Cod          As Integer
Dim LV_Cod2         As Integer
Dim LV_Campos()     As String
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Campo        As String
Dim LV_Query        As Recordset

    LV_Nombre_Tabla = "Tabla_Instrumentos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla & ";"
    LV_Cod2 = 0
            
    If GV_Flag_Add_Medicion = True Then
        LV_Cod = 2
'        lv_Sql = lv_Sql & " WHERE " & LV_Nombre_Tabla & ".Cod_Funcion" _
                 & " = (" & 2 & ");"
    ElseIf GV_Flag_Add_Exita = True Then
        LV_Cod = 1
'        lv_Sql = lv_Sql & " WHERE " & LV_Nombre_Tabla & ".Cod_Funcion" _
                 & " = (" & 1 & ");"
    ElseIf GV_Flag_Add_Prueba = True Then
        LV_Cod = 3
        LV_Cod2 = 4
'        lv_Sql = lv_Sql & " WHERE " & LV_Nombre_Tabla & ".Cod_Funcion" _
                 & " = (" & 3 & ")"
'        lv_Sql = lv_Sql & " OR " & LV_Nombre_Tabla & ".Cod_Funcion" _
                 & " = (" & 4 & ");"
    ElseIf GV_Flag_Add_Accesorio = True Then
        LV_Cod = 3
        LV_Cod2 = 4
'        lv_Sql = lv_Sql & " WHERE " & LV_Nombre_Tabla & ".Cod_Funcion" _
                 & " = (" & 3 & ")"
'        lv_Sql = lv_Sql & " OR " & LV_Nombre_Tabla & ".Cod_Funcion" _
                 & " = (" & 4 & ");"
    Else
'        lv_Sql = lv_Sql & ";"
        LV_Cod = 0
    End If

    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
'    GV_Flag_Add_Medicion = False
'    GV_Flag_Add_Exita = False
'    GV_Flag_Add_Prueba = False
'    GV_Flag_Add_Accesorio = False
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            ReDim LV_Campos(7)
            
            .MoveFirst
            
            Do
                If LV_Cod = 0 Or _
                   LV_Cod = !Cod_Funcion Or _
                   LV_Cod2 = !Cod_Funcion _
                Then

                    LV_Campos(0) = GV_Lbl(Enum_Orden_Tabla.Tipo_Disp).Etiqueta(!Cod_Dispositivo - 1)
                    LV_Campos(1) = GV_Lbl(Enum_Orden_Tabla.Funcion).Etiqueta(!Cod_Funcion - 1)
                    If !Cod_Componente Then
                        LV_Campos(2) = GV_Lbl(Enum_Orden_Tabla.Clase_Componen).Etiqueta(!Cod_Componente - 1)
                    Else
                        LV_Campos(2) = GV_Lbl(Enum_Orden_Tabla.Clase_Instru).Etiqueta(!Cod_Instrumento - 1)
                    End If
                    LV_Campos(3) = GV_Lbl(Enum_Orden_Tabla.Fabricante).Etiqueta(!Cod_Fabricante - 1)
                    LV_Campos(4) = !Modelo
                    LV_Campos(5) = !Numero_Parte
                    LV_Campos(6) = !Numero_Serie
                    If !Cod_Comunicacion Then
                        LV_Campos(7) = GV_Lbl(Enum_Orden_Tabla.Comunicacion).Etiqueta(!Cod_Comunicacion - 1)
                    End If
                    AddItemListView LV_LstVw, LV_Campos, True, !Codigo
                
                End If
                    
                .MoveNext
            
            Loop Until .EOF Or .BOF
            
        End If
        
    End With
    
End Function

'Function BD_Get_Cod_Funcion(LV_Medicion As Boolean, _
'                             LV_Exitacion As Boolean, _
'                             LV_Prueba As Boolean, _
'                             LV_Accesorio As Boolean)
'
'Dim lv_Sql          As String
'Dim LV_Query        As Recordset
'Dim i               As Integer
'Dim LV_Nombre_Tabla     As String
'Dim LV_Cod_Campo    As Integer
'
'    LV_Nombre_Tabla = "Tipo_Funciones"
'
'    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
'            & ".* From " & LV_Nombre_Tabla _
'            & " WHERE (("
'
'    If LV_Medicion = True Then
'        LV_Campo = "Exitación"
'    ElseIf LV_Exitacion = True Then
'        LV_Campo = "Medición"
'    ElseIf LV_Prueba = True Then
'        LV_Campo = "Prueba"
'    ElseIf LV_Accesorio = True Then
'        LV_Campo = "Accesorio"
'
'    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
'             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
'             & ")= (" _
'             & Cod_Compo & ")"
'
'    lv_Sql = lv_Sql & ");"
'
'    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
'
'    With LV_Query
'
'        If Not .EOF And Not .BOF Then
'
'            .MoveFirst
'
'            BD_Get_Cod_Funcion = !Codigo
'        Else
'
'            BD_Get_Cod_Funcion = 0
'
'        End If
'
'    End With
'
'End Function

Function BD_Fill_LstVw_Projects(LV_LstVw As ListView)

Dim i               As Integer
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim LV_Campos()     As String
    
    LV_Nombre_Tabla = "Proyectos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla & ";"

    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            ReDim LV_Campos(7)
            
            .MoveFirst
            
            Do
                LV_Campos(0) = !Nombre
                LV_Campos(1) = !Ubicacion
                LV_Campos(2) = !Disp_Prueba
                LV_Campos(3) = !Num_Parte
                LV_Campos(4) = !Num_Serie
                LV_Campos(5) = !Fecha
                LV_Campos(6) = !Encargado
                LV_Campos(7) = !Result_File
                
                AddItemListView LV_LstVw, LV_Campos, True, !Codigo
                
                .MoveNext
            
            Loop Until .EOF Or .BOF
            
        End If
        
    End With
    
End Function

Function BD_Fill_LstVw_With_Array(LV_LstVw As ListView, LV_Cod_Instru() As Integer)

Dim i               As Integer
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim LV_Campos()     As String
    
    LV_Nombre_Tabla = "Tabla_Instrumentos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla & ";"

    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            ReDim LV_Campos(7)
            
            .MoveFirst
            
            Do
                If If_Index_Meet(!Codigo, LV_Cod_Instru) Then

                    LV_Campos(0) = GV_Lbl(Enum_Orden_Tabla.Tipo_Disp).Etiqueta(!Cod_Dispositivo - 1)
                    LV_Campos(1) = GV_Lbl(Enum_Orden_Tabla.Funcion).Etiqueta(!Cod_Funcion - 1)
                    If !Cod_Componente Then
                        LV_Campos(2) = GV_Lbl(Enum_Orden_Tabla.Clase_Componen).Etiqueta(!Cod_Componente - 1)
                    Else
                        LV_Campos(2) = GV_Lbl(Enum_Orden_Tabla.Clase_Instru).Etiqueta(!Cod_Instrumento - 1)
                    End If
                    LV_Campos(3) = GV_Lbl(Enum_Orden_Tabla.Fabricante).Etiqueta(!Cod_Fabricante - 1)
                    LV_Campos(4) = !Modelo
                    LV_Campos(5) = !Numero_Parte
                    LV_Campos(6) = !Numero_Serie
                    If !Cod_Comunicacion Then
                        LV_Campos(7) = GV_Lbl(Enum_Orden_Tabla.Comunicacion).Etiqueta(!Cod_Comunicacion - 1)
                    End If
                    AddItemListView LV_LstVw, LV_Campos, True, !Codigo
                
                End If
                    
                .MoveNext
            
            Loop Until .EOF Or .BOF
            
        End If
        
    End With
    
End Function

Function BD_Fill_Comandos_Instru(LV_Intrumento As Type_Instrumento, _
                                LV_Cod_Instru As Integer) As Boolean

Dim i               As Integer
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim LV_Qty          As Integer
Dim LV_Cmds()       As Type_Comando
Dim LV_Q_Var        As Integer
Dim LV_Q_Consult    As Integer

    LV_Q_Var = 0
    LV_Q_Consult = 0
    
    LV_Qty = BD_Read_Commands_Instrument(LV_Cod_Instru, LV_Cmds)
    
    If LV_Qty Then
        For i = 0 To LV_Qty - 1
            With LV_Intrumento
                Select Case LV_Cmds(i).Cod_Funcion
                Case 1
                    ' Configuracion
                    If .Cmd_Config = "" Then
                        .Cmd_Config = LV_Cmds(i).Comando
                    Else
                        .Cmd_Config = .Cmd_Config & ";" & LV_Cmds(i).Comando
                    End If
                Case 2
                    ' Comando
                    ReDim Preserve .Cmd_Set_Var(LV_Q_Var)
                    ReDim Preserve .Cmd_Set_Param(LV_Q_Var)
                    .Cmd_Set_Var(LV_Q_Var) = LV_Cmds(i).Comando
                    .Cmd_Set_Param(LV_Q_Var) = LV_Cmds(i).Cod_Parametro
                    LV_Q_Var = LV_Q_Var + 1
                Case 3
                    ' Consulta
                    ReDim Preserve .Cmd_Consult(LV_Q_Consult)
                    ReDim Preserve .Cmd_Consu_Param(LV_Q_Consult)
                    .Cmd_Consult(LV_Q_Consult) = LV_Cmds(i).Comando
                    .Cmd_Consu_Param(LV_Q_Consult) = LV_Cmds(i).Cod_Parametro
                    LV_Q_Consult = LV_Q_Consult + 1
                End Select
            End With
        Next
    End If

End Function




Function BD_Fill_LstVw_With_Instru_From_Pjt(LV_LstVw As ListView, _
                                            LV_Cod_Pjt As Integer, _
                                            LV_Cod_Funcion)

Dim i               As Integer
Dim LV_Cod_Instru() As Integer
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Query        As Recordset

    ' Obtener Lista de Instrumentos del Proyecto

    LV_Nombre_Tabla = "Lista_Instrumentos"

    lv_Sql = " SELECT " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE "

    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
             & "Cod_Proyecto" _
             & "= (" _
             & LV_Cod_Pjt & ");"
             
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
            
            Do
            
                If LV_Cod_Funcion = !Cod_Funcion Then
                
                    ReDim Preserve LV_Cod_Instru(i)
                    
                    LV_Cod_Instru(i) = !Cod_Instrumento
                    
                    i = i + 1
                
                End If
                
                .MoveNext
                
            Loop Until (.EOF Or .BOF)
            
        End If
        
    End With
    
    If i Then
        BD_Fill_LstVw_With_Array LV_LstVw, LV_Cod_Instru
    End If
    
End Function

Function BD_Get_Cod_Instruments(LV_Cod_Instru() As Integer, _
                                LV_Cod_Pjt As Integer, _
                                LV_Cod_Funcion) As Integer

Dim i               As Integer
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Query        As Recordset
    
    LV_Nombre_Tabla = "Lista_Instrumentos"

    lv_Sql = " SELECT " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE "
    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
             & "Cod_Proyecto" _
             & "= (" _
             & LV_Cod_Pjt & ");"
             
    i = 0
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    With LV_Query
        If Not .EOF And Not .BOF Then
            .MoveFirst
            Do
                If LV_Cod_Funcion = !Cod_Funcion Then
                    ReDim Preserve LV_Cod_Instru(i)
                    LV_Cod_Instru(i) = !Cod_Instrumento
                    i = i + 1
                End If
                .MoveNext
            Loop Until (.EOF Or .BOF)
        End If
    End With
    
    'BD_Get_Cod_Instru = i
    
End Function

Function BD_Get_Cod_Instrument(Cod_Compo As Integer, _
            Cod_Instrumento As Integer, _
            Cod_Comu As Integer, _
            Cod_Dispo As Integer, _
            Cod_Fabri As Integer, _
            Cod_Func As Integer, _
            Modelo As String, _
            NumParte As String, _
            NumSerie As String _
            ) As Integer

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Campo    As Integer

    LV_Nombre_Tabla = "Tabla_Instrumentos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE (("
    
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_COD_COMPONENTE
    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= (" _
             & Cod_Compo & ")"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_COD_INSTRUMENTO
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= (" _
             & Cod_Instrumento & ")"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_COD_DISPOSITIVO
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= (" _
             & Cod_Dispo & ")"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_COD_FABRICANTE
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= (" _
             & Cod_Fabri & ")"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_COD_FUNCION
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= (" _
             & Cod_Func & ")"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_MODELO
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= ('" _
             & Modelo & "')"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_NUM_PARTE
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= ('" _
             & NumParte & "')"
             
    LV_Cod_Campo = E_NUM_CAMPO_TBL_INSTRUMENT.CAMPO_NUM_SERIE
    lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
             & BD_Get_Field_Name(Tabla_Instrumentos, LV_Cod_Campo) _
             & ")= ('" _
             & NumSerie & "')"
             
    lv_Sql = lv_Sql & ");"
        
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
            
            BD_Get_Cod_Instrument = !Codigo
        Else
            
            BD_Get_Cod_Instrument = 0
            
        End If
        
    End With
    
End Function

Function BD_Get_Cod_Project() As Integer

Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim i                   As Integer
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = "Proyectos"
    
    If LV_Nombre_Tabla = "" Then
        
        Exit Function
    
    End If
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE (("
            
    With GV_Actual_Project
        
        lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
                 & BD_Get_Field_Name(LISTA_PROYECTOS, CAMPO_NOMBRE) _
                 & ")= ('" _
                 & .Project_Name & "')"
                 
        lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
                 & BD_Get_Field_Name(LISTA_PROYECTOS, CAMPO_DISPOSITIVO) _
                 & ")= ('" _
                 & .Dispositivo & "')"
                 
        lv_Sql = lv_Sql & " AND (" & LV_Nombre_Tabla & "." _
                 & BD_Get_Field_Name(LISTA_PROYECTOS, CAMPO_PJT_NUM_PARTE) _
                 & ")= ('" _
                 & .Num_Parte & "')"
                 
        lv_Sql = lv_Sql & " AND(" & LV_Nombre_Tabla & "." _
                 & BD_Get_Field_Name(LISTA_PROYECTOS, CAMPO_PJT_NUM_SERIE) _
                 & ")= ('" _
                 & .Num_Serie & "')"
                 
        lv_Sql = lv_Sql & ");"
                 
    End With
    
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
            
            BD_Get_Cod_Project = !Codigo
        Else
            
            BD_Get_Cod_Project = 0
            
        End If
        
    End With
        
End Function

Function BD_Get_Field_Name_Proyectos(LV_Index_Field As Integer) As String

Dim LV_Field            As String
    
    Select Case LV_Index_Field
    
    Case CAMPO_CODIGO
        LV_Field = "Codigo"
    Case CAMPO_NOMBRE
        LV_Field = "Nombre"
    Case CAMPO_UBICACION
        LV_Field = "Ubicacion"
    Case CAMPO_DISPOSITIVO
        LV_Field = "Disp_Prueba"
    Case CAMPO_PJT_NUM_PARTE
        LV_Field = "Num_Parte"
    Case CAMPO_PJT_NUM_SERIE
        LV_Field = "Num_Serie"
    Case CAMPO_FECHA
        LV_Field = "Fecha"
    Case CAMPO_RESULT_FILE
        LV_Field = "Result_File"
    Case CAMPO_OCOMPRA
        LV_Field = "OrdenCompra"
    Case CAMPO_ENCARGADO
        LV_Field = "Encargado"
    Case CAMPO_COD_LISTAINSTRU
        LV_Field = "Cod_ListaInstrumento"
    Case CAMPO_COD_LISTACOMPON
        LV_Field = "Cod_ListaComponente"
    Case CAMPO_COD_LISTAPARAM
        LV_Field = "Cod_ListaParamControl"

    End Select
    
    BD_Get_Field_Name_Proyectos = LV_Field
    
End Function

Function BD_Get_Field_Name(LV_Index_Tbl As Integer, LV_Index_Field As Integer) As String

Dim LV_Field            As String

    If LV_Index_Field Then
    
        Select Case LV_Index_Tbl
        
        Case LISTA_COMPONENTES
        Case LISTA_INSTRUMENTOS
        Case LISTA_PARAMCONTROL
        Case LISTA_PROYECTOS
            LV_Field = BD_Get_Field_Name_Proyectos(LV_Index_Field)
        Case TABLA_COMANDOS
        Case TABLA_COMANDOS_GPIB
        Case Tabla_Instrumentos
            LV_Field = BD_Get_Field_Name_Tbl_Instrument(LV_Index_Field)
        Case TABLA_PARAM_COMPONENTES
        Case TABLA_PRECISIONES
        Case TIPO_COMPONENTE
        Case Tipo_Comunicacion
        Case TIPO_DIMENSION
        Case TIPO_DISPOSITIVOS
        Case TIPO_FABRICANTES
        Case TIPO_FUNCIONES
        Case TIPO_INSTRUMENTOS
        Case TIPO_PARAMETROS
        
        End Select
    
    Else
    
        Select Case LV_Index_Field
        
        Case CAMPO_CODIGO
            LV_Field = "Codigo"
        End Select
        
    End If
    
    If LV_Field = "" Then
        MsgBox "No hay Campo Válido"
        'End
    End If
    
    BD_Get_Field_Name = LV_Field
    
    
    
End Function

Function BD_Get_Field_Name_Tbl_Instrument(LV_Index_Field As Integer) As String

Dim LV_Field            As String
    
    Select Case LV_Index_Field
    
    Case CAMPO_CODIGO_INSTRUMENT
        LV_Field = "Codigo"
    Case CAMPO_COD_COMPONENTE
        LV_Field = "Cod_Componente"
    Case CAMPO_COD_INSTRUMENTO
        LV_Field = "Cod_Instrumento"
    Case CAMPO_COD_COMUNICACION
        LV_Field = "Cod_Comunicacion"
    Case CAMPO_COD_DISPOSITIVO
        LV_Field = "Cod_Dispositivo"
    Case CAMPO_COD_FABRICANTE
        LV_Field = "Cod_Fabricante"
    Case CAMPO_COD_FUNCION
        LV_Field = "Cod_Funcion"
    Case CAMPO_MODELO
        LV_Field = "Modelo"
    Case CAMPO_NUM_PARTE
        LV_Field = "Numero_Parte"
    Case CAMPO_NUM_SERIE
        LV_Field = "Numero_Serie"
    Case CAMPO_COD_TBL_PARAM
        LV_Field = "Cod_Tabla_Param"
    Case CAMPO_COD_TBL_CMDS
        LV_Field = "Cod_Tabla_Comandos"

    End Select
    
    BD_Get_Field_Name_Tbl_Instrument = LV_Field
    
End Function

Function BD_Get_Rangos_Control(LV_Cod_Pjt As Integer, _
                               LV_Rangos() As Type_RangoControl) As Integer

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = "Lista_Param_Control"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla & ";"
        
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    i = 0
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
                        
            Do
            
                If LV_Cod_Pjt = 0 Or _
                    LV_Cod_Pjt = !Cod_Proyecto Then
                    
                    ReDim Preserve LV_Rangos(i)
                    
                    LV_Rangos(i).Cod_Dimension = !Cod_Dimension
                    LV_Rangos(i).Cod_Parametro = !Cod_Parametro
                    LV_Rangos(i).Etapa = !Etapa
                    LV_Rangos(i).Parametro = GV_Lbls_Parametros(LV_Rangos(i).Cod_Parametro - 1)
                    LV_Rangos(i).Paso = !Paso
                    LV_Rangos(i).Unidad = GV_Lbls_Dimensiones(LV_Rangos(i).Cod_Dimension - 1)
                    LV_Rangos(i).ValorMax = !ValorMax
                    LV_Rangos(i).ValorMin = !ValorMin
                    LV_Rangos(i).PRI = 1000
                    LV_Rangos(i).PW = 1
                    On Error Resume Next
                    LV_Rangos(i).PRI = !PRI
                    LV_Rangos(i).PW = !PW
                    LV_Rangos(i).AccionFinEtapa = !AccionFinEtapa
                    On Error GoTo 0
                    If !AplicarCurva Then
                        LV_Rangos(i).AplicarPV = 1
                        LV_Rangos(i).b50Ohms = !b50Ohms
                        LV_Rangos(i).VoltDiv = !Escala
                        LV_Rangos(i).CurvaPV = !CurvaPV
                    Else
                        LV_Rangos(i).AplicarPV = 0
                        LV_Rangos(i).b50Ohms = 0
                        LV_Rangos(i).VoltDiv = 0
                        LV_Rangos(i).CurvaPV = ""
                    End If
                    i = i + 1
                End If
                
                .MoveNext
            
            Loop Until .EOF Or .BOF
            
        End If
        
    End With
        
    BD_Get_Rangos_Control = i
    
End Function

Function BD_Index_Elemento(ByVal IndexTipo As Integer, lsData As String) As Integer

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = BD_Nombre_Tabla(IndexTipo)
    
    If LV_Nombre_Tabla = "" Then
        
        Exit Function
    
    End If
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE ((" & LV_Nombre_Tabla & ".Nombre)= ('" _
            & lsData & "'));"
        
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
            
            BD_Index_Elemento = !Codigo
        Else
            
            BD_Index_Elemento = -1
            
        End If
        
    End With
        
End Function

Function BD_Leer_Tipos(LV_Nombre_Tabla As String, ByRef LV_Lbls() As String)

Dim i           As Integer
Dim lv_Sql              As String
Dim LV_Query            As Recordset

    If LV_Nombre_Tabla = "" Then
        Exit Function
    End If
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla & ";"
    
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            .MoveFirst
            
            i = 0
            
            Do
            
                ReDim Preserve LV_Lbls(i)
                
                LV_Lbls(i) = !Nombre
                
                i = i + 1
                
                .MoveNext
                
            Loop Until (.EOF Or .BOF)
            
        End If
        
    End With
    
End Function

Function BD_Nombre_Tabla(Index As Integer) As String

    BD_Nombre_Tabla = ""
    
    Select Case Index
        
        Case Is = Enum_Orden_Tabla.Tipo_Disp
            
            BD_Nombre_Tabla = "Tipo_Dispositivos"
    
        Case Is = Enum_Orden_Tabla.Funcion
            
            BD_Nombre_Tabla = "Tipo_Funciones"
    
        Case Is = Enum_Orden_Tabla.Clase_Instru
            
            BD_Nombre_Tabla = "Tipo_Instrumentos"
    
        Case Is = Enum_Orden_Tabla.Fabricante
            
            BD_Nombre_Tabla = "Tipo_Fabricantes"
    
        Case Is = Enum_Orden_Tabla.Comunicacion
            
            BD_Nombre_Tabla = "Tipo_Comunicacion"
    
        Case Is = Enum_Orden_Tabla.Clase_Componen
            
            BD_Nombre_Tabla = "Tipo_Componentes"
    
    End Select
    
End Function

Sub BD_Get_Controles_Proyecto()

Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim i                   As Integer
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Pjt          As Integer

    LV_Nombre_Tabla = "Proyectos"
    LV_Cod_Pjt = GV_Actual_Project.Cod_Project
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE (("
            
    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
             & "Codigo" _
             & ")= (" _
             & LV_Cod_Pjt & ")"
             
    lv_Sql = lv_Sql & ");"
             
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        If Not .EOF And Not .BOF Then
            .MoveFirst
            GV_Actual_Project.Controles.AddressGPIB = 19
            GV_Actual_Project.Controles.Adquirir = 1
            GV_Actual_Project.Controles.AplicarCurvaVideoPot = 0
            GV_Actual_Project.Controles.ArchivoCompensaSalida = " "
            GV_Actual_Project.Controles.CapturarPot = 0
            GV_Actual_Project.Controles.ControlAnalizaEspec = 0
            GV_Actual_Project.Controles.ControlOscilos = 0
            GV_Actual_Project.Controles.ControlPowerMeter = 0
            GV_Actual_Project.Controles.EsperarEstabiliza = 1
            GV_Actual_Project.Controles.FileCurvaVideoPot = " "
            GV_Actual_Project.Controles.FileTablaParam = " "
            GV_Actual_Project.Controles.OperacionManual = 0
            GV_Actual_Project.Controles.TpoEspera = 10
            GV_Actual_Project.Controles.UsarTablaParam = 0
            On Error Resume Next
            GV_Actual_Project.Controles.AddressGPIB = !AddressGPIB
            GV_Actual_Project.Controles.Adquirir = !Adquirir
            GV_Actual_Project.Controles.AplicarCurvaVideoPot = !AplicarCurvaVideoPot
            GV_Actual_Project.Controles.ArchivoCompensaSalida = !ArchivoCompensaSalida
            GV_Actual_Project.Controles.CapturarPot = !CapturarPot
            GV_Actual_Project.Controles.ControlAnalizaEspec = !ControlAnalizaEspec
            GV_Actual_Project.Controles.ControlOscilos = !ControlOscilos
            GV_Actual_Project.Controles.ControlPowerMeter = !ControlPowerMeter
            GV_Actual_Project.Controles.EsperarEstabiliza = !EsperarEstabilizacion
            GV_Actual_Project.Controles.FileCurvaVideoPot = !FileCurvaVideoPot
            GV_Actual_Project.Controles.FileTablaParam = !FileTablaParam
            GV_Actual_Project.Controles.OperacionManual = !OperacionManual
            GV_Actual_Project.Controles.TpoEspera = !TpoEspera
            GV_Actual_Project.Controles.UsarTablaParam = !UsarTablaParam
            On Error GoTo 0
            
        End If
        
    End With

End Sub

Function BD_Open_Project(LV_Cod_Pjt, LV_Campos() As String)

Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim i                   As Integer
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = "Proyectos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE (("
            
    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
             & "Codigo" _
             & ")= (" _
             & LV_Cod_Pjt & ")"
             
    lv_Sql = lv_Sql & ");"
             
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            ReDim LV_Campos(9)
            
            .MoveFirst
            
            On Error Resume Next
            LV_Campos(0) = !Nombre
            LV_Campos(1) = !Ubicacion
            LV_Campos(2) = !Disp_Prueba
            LV_Campos(3) = !Num_Parte
            LV_Campos(4) = !Num_Serie
            LV_Campos(5) = !Fecha
            LV_Campos(6) = !Result_File
            LV_Campos(7) = !OrdenCompra
            LV_Campos(8) = !Encargado
            LV_Campos(9) = !CompensacionSetup
            On Error GoTo 0
            
        End If
        
    End With

End Function

Function BD_Read_Commands_Instrument(LV_Cod_Instru As Integer, _
                            LV_Cmds() As Type_Comando) As Integer

Dim i                   As Integer
Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim LV_Nombre_Tabla     As String

    If LV_Cod_Instru = 0 Then
        BD_Read_Commands_Instrument = 0
        Exit Function
    End If
    
    LV_Nombre_Tabla = "Tabla_Comandos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE "

    lv_Sql = lv_Sql & LV_Nombre_Tabla & "." _
             & "Cod_Instrumento" _
             & "= (" _
             & LV_Cod_Instru & ");"
             
    
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            i = 0
            
            .MoveFirst
            
            Do
                ReDim Preserve LV_Cmds(i)
                
                LV_Cmds(i).Cod_Comunicacion = !Cod_Comunicacion
                LV_Cmds(i).Cod_Funcion = !Cod_Funcion
                LV_Cmds(i).Cod_Instrumento = !Cod_Instrumento
                LV_Cmds(i).Cod_Parametro = !Cod_Parametro
                LV_Cmds(i).Comando = !Comando
                LV_Cmds(i).Valor = !Valor
                i = i + 1
                
                .MoveNext
            
            Loop Until .EOF Or .BOF
            
        End If
        
    End With
    
    BD_Read_Commands_Instrument = i
    
End Function

Function BD_Read_Instrument(LV_Cod_Instru As Integer, _
                            LV_Campos() As String, _
                            LV_Indexs() As Integer) As Boolean

Dim i               As Integer
Dim LV_Nombre_Tabla As String
Dim lv_Sql          As String
Dim LV_Query        As Recordset
'Dim LV_Campos()     As String
    
    If LV_Cod_Instru = 0 Then
        BD_Read_Instrument = False
        Exit Function
    End If
    
    LV_Nombre_Tabla = "Tabla_Instrumentos"
    
    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE " & LV_Nombre_Tabla & ".Codigo = (" _
            & LV_Cod_Instru & ") ;"
            

    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With LV_Query
        
        If Not .EOF And Not .BOF Then
        
            ReDim LV_Campos(4)
            ReDim LV_Indexs(5)
            
            .MoveFirst
            
            LV_Indexs(0) = !Cod_Dispositivo
            LV_Indexs(1) = !Cod_Funcion
            LV_Indexs(2) = !Cod_Componente
            LV_Indexs(3) = !Cod_Instrumento
            LV_Indexs(4) = !Cod_Fabricante
            LV_Indexs(5) = !Cod_Comunicacion
            
'            LV_Campos(0) = GV_Lbl(Enum_Orden_Tabla.Tipo_Disp).Etiqueta(!Cod_Dispositivo - 1)
'            LV_Campos(1) = GV_Lbl(Enum_Orden_Tabla.Funcion).Etiqueta(!Cod_Funcion - 1)
'            If !Cod_Componente Then
'                LV_Campos(2) = GV_Lbl(Enum_Orden_Tabla.Clase_Componen).Etiqueta(!Cod_Componente - 1)
'            Else
'                LV_Campos(2) = GV_Lbl(Enum_Orden_Tabla.Clase_Instru).Etiqueta(!Cod_Instrumento - 1)
'            End If
'            LV_Campos(3) = GV_Lbl(Enum_Orden_Tabla.Fabricante).Etiqueta(!Cod_Fabricante - 1)
            LV_Campos(0) = !Modelo
            LV_Campos(1) = !Numero_Parte
            LV_Campos(2) = !Numero_Serie
            If !Archivo_Carac <> Null Then
                LV_Campos(3) = !Archivo_Carac
            End If
            If !GPIB_Address <> Null Then
                LV_Campos(4) = !GPIB_Address
            End If
'            If !Cod_Comunicacion Then
'                LV_Campos(7) = GV_Lbl(Enum_Orden_Tabla.Comunicacion).Etiqueta(!Cod_Comunicacion - 1)
'            End If
        Else
            BD_Read_Instrument = False
            Exit Function
        End If
        
    End With

    BD_Read_Instrument = True
    
End Function

Function BD_Update_Controles_Proyecto()

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Pjt      As Integer
Dim LV_Fecha_Update     As String

    LV_Nombre_Tabla = "Proyectos"
    
    LV_Cod_Pjt = GV_Actual_Project.Cod_Project
    
    LV_Fecha_Update = format(Now(), "DD-MM-YYYY")

    With GV_Actual_Project.Controles
        lv_Sql = " UPDATE " & LV_Nombre_Tabla _
                & " SET " _
                & LV_Nombre_Tabla & ".AplicarCurvaVideoPot = (" _
                & .AplicarCurvaVideoPot & "), " _
                & LV_Nombre_Tabla & ".EsperarEstabilizacion = (" _
                & .EsperarEstabiliza & "), " _
                & LV_Nombre_Tabla & ".TpoEspera = (" _
                & .TpoEspera & "), " _
                & LV_Nombre_Tabla & ".Adquirir = (" _
                & .Adquirir & "), " _
                & LV_Nombre_Tabla & ".CapturarPot = (" _
                & .CapturarPot & "), "
        
        If .FileCurvaVideoPot = "" Then
            .FileCurvaVideoPot = " "
        End If
        
        lv_Sql = lv_Sql _
                & LV_Nombre_Tabla & ".AddressGPIB = (" _
                & .AddressGPIB & "), " '_
'                & LV_Nombre_Tabla & ".ControlOscilos = (" _
'                & .ControlOscilos & "), " _
'                & LV_Nombre_Tabla & ".ControlPowerMeter = (" _
'                & .ControlPowerMeter & "), " _
'                & LV_Nombre_Tabla & ".ControlAnalizaEspec = (" _
'                & .ControlAnalizaEspec & "), " _
'                & LV_Nombre_Tabla & ".FileCurvaVideoPot = ('" _
'                & .FileCurvaVideoPot & "'), "
        lv_Sql = lv_Sql _
                & LV_Nombre_Tabla & ".OperacionManual = (" _
                & .OperacionManual & "), " _
                & LV_Nombre_Tabla & ".ArchivoCompensaSalida = ('" _
                & .ArchivoCompensaSalida & "'), " _
                & LV_Nombre_Tabla & ".UsarTablaParam = (" _
                & .UsarTablaParam & "), " _
                & LV_Nombre_Tabla & ".FileTablaParam = ('" _
                & .FileTablaParam & "') "
        lv_Sql = lv_Sql _
                & " WHERE " & LV_Nombre_Tabla & ".Codigo = (" _
                & LV_Cod_Pjt & ") ;"
    End With

    GV_BD_INSTRUMENT.Execute (lv_Sql)

End Function

Function BD_Update_Info_Gral_Proyecto( _
                                        Nombre As String, _
                                        Ubicacion As String, _
                                        Disp_Prueba As String, _
                                        Num_Parte As String, _
                                        Num_Serie As String, _
                                        Fecha As String, _
                                        Result_File As String, _
                                        OCompra As String, _
                                        Encargado As String, _
                                        LV_CompSetup As String)



Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Pjt      As Integer
Dim LV_Fecha_Update     As String

    LV_Nombre_Tabla = "Proyectos"
    
    LV_Cod_Pjt = GV_Actual_Project.Cod_Project
    
    LV_Fecha_Update = format(Now(), "DD-MM-YYYY")

    lv_Sql = " UPDATE " & LV_Nombre_Tabla _
            & " SET " _
            & LV_Nombre_Tabla & ".Nombre = ('" _
            & Nombre & "'), " _
            & LV_Nombre_Tabla & ".Ubicacion = ('" _
            & Ubicacion & "'), " _
            & LV_Nombre_Tabla & ".Disp_Prueba = ('" _
            & Disp_Prueba & "'), " _
            & LV_Nombre_Tabla & ".Num_Parte = ('" _
            & Num_Parte & "'), " _
            & LV_Nombre_Tabla & ".Num_Serie = ('" _
            & Num_Serie & "'), "
    lv_Sql = lv_Sql _
            & LV_Nombre_Tabla & ".Fecha = ('" _
            & Fecha & "'), " _
            & LV_Nombre_Tabla & ".Result_File = ('" _
            & Result_File & "'), " _
            & LV_Nombre_Tabla & ".OrdenCompra = ('" _
            & OCompra & "'), " _
            & LV_Nombre_Tabla & ".Encargado = ('" _
            & Encargado & "'), " _
            & LV_Nombre_Tabla & ".CompensacionSetup = ('" _
            & LV_CompSetup & "'), " _
            & LV_Nombre_Tabla & ".UltimaModificacion = ('" _
            & LV_CompSetup & "') "
    lv_Sql = lv_Sql _
            & " WHERE " & LV_Nombre_Tabla & ".Codigo = (" _
            & LV_Cod_Pjt & ") ;"
            

    GV_BD_INSTRUMENT.Execute (lv_Sql)

    
End Function

Function BD_Update_Instrumento(LV_Codigo As Integer, _
            Cod_Compo As Integer, _
            Cod_Instrumento As Integer, _
            Cod_Comu As Integer, _
            Cod_Dispo As Integer, _
            Cod_Fabri As Integer, _
            Cod_Func As Integer, _
            Modelo As String, _
            NumParte As String, _
            NumSerie As String, _
            Cod_TabParam As Integer, _
            Cod_Tab_Command As Integer, _
            address As Integer _
            ) As Integer

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer
Dim LV_Nombre_Tabla     As String
Dim LV_Cod_Instrum      As Integer

    LV_Nombre_Tabla = "Tabla_Instrumentos"
    
    lv_Sql = " UPDATE " & LV_Nombre_Tabla _
            & " SET " _
            & LV_Nombre_Tabla & ".Cod_Componente = ('" _
            & Cod_Compo & "'), " _
            & LV_Nombre_Tabla & ".Cod_Instrumento = ('" _
            & Cod_Instrumento & "'), " _
            & LV_Nombre_Tabla & ".Cod_Comunicacion = ('" _
            & Cod_Comu & "'), " _
            & LV_Nombre_Tabla & ".Cod_Dispositivo = ('" _
            & Cod_Dispo & "'), " _
            & LV_Nombre_Tabla & ".Cod_Fabricante = ('" _
            & Cod_Fabri & "'), "
    lv_Sql = lv_Sql _
            & LV_Nombre_Tabla & ".Cod_Funcion = ('" _
            & Cod_Func & "'), " _
            & LV_Nombre_Tabla & ".Modelo = ('" _
            & Modelo & "'), " _
            & LV_Nombre_Tabla & ".Numero_Parte = ('" _
            & NumParte & "'), " _
            & LV_Nombre_Tabla & ".Numero_Serie = ('" _
            & NumSerie & "'), " _
            & LV_Nombre_Tabla & ".GPIB_Address = ('" _
            & address & "') " _
            & " WHERE " & LV_Nombre_Tabla & ".Codigo = (" _
            & LV_Codigo & ") ;"
            
    GV_BD_INSTRUMENT.Execute (lv_Sql)

End Function


Function BD_Update_Rango_Control(LV_Cod_Pjt As Integer, LV_Etapa As Integer)

Dim i                   As Integer
Dim lv_Sql              As String
Dim LV_Query            As Recordset
Dim LV_Nombre_Tabla     As String

    LV_Nombre_Tabla = "Lista_Param_Control"

    lv_Sql = " SELECT  " & LV_Nombre_Tabla _
            & ".* From " & LV_Nombre_Tabla _
            & " WHERE " & LV_Nombre_Tabla & ".Etapa = (" _
            & LV_Etapa & ") AND " _
            & LV_Nombre_Tabla & ".Cod_Proyecto = (" _
            & LV_Cod_Pjt & ") ;"
    
    Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
    
    With GV_Actual_Project.Rango(LV_Etapa)
        
        If Not LV_Query.EOF And Not LV_Query.BOF Then
        
            lv_Sql = " UPDATE " & LV_Nombre_Tabla _
                    & " SET " _
                    & LV_Nombre_Tabla & ".Cod_Parametro = (" _
                    & .Cod_Parametro & "), " _
                    & LV_Nombre_Tabla & ".Cod_Dimension = (" _
                    & .Cod_Dimension & "), " _
                    & LV_Nombre_Tabla & ".ValorMin = (" _
                    & .ValorMin & "), " _
                    & LV_Nombre_Tabla & ".ValorMax = (" _
                    & .ValorMax & "), " _
                    & LV_Nombre_Tabla & ".AplicarCurva = (" _
                    & .AplicarPV & "), " _
                    & LV_Nombre_Tabla & ".b50Ohms = (" _
                    & .b50Ohms & "), " _
                    & LV_Nombre_Tabla & ".Escala = (" _
                    & .VoltDiv & "), "
            lv_Sql = lv_Sql _
                    & LV_Nombre_Tabla & ".CurvaPV = ('" _
                    & .CurvaPV & "'), " _
                    & LV_Nombre_Tabla & ".PRI = ('" _
                    & .PRI & "'), " _
                    & LV_Nombre_Tabla & ".PW = ('" _
                    & .PW & "'), " _
                    & LV_Nombre_Tabla & ".Paso = (" _
                    & .Paso & ") " _
                    & " WHERE " & LV_Nombre_Tabla & ".Etapa = (" _
                    & LV_Etapa & ") AND " _
                    & LV_Nombre_Tabla & ".Cod_Proyecto = (" _
                    & LV_Cod_Pjt & ") ;"
            
        Else
            
            lv_Sql = " INSERT INTO  " & LV_Nombre_Tabla _
                    & " (Cod_Proyecto,Cod_Instrumento,Cod_Parametro" _
                    & ",Cod_Dimension,ValorMin,ValorMax" _
                    & ",PRI,PW" _
                    & ",Paso,AplicarCurva,b50Ohms,Escala,CurvaPV,Etapa) " _
                    & " VALUES ((" & LV_Cod_Pjt & ")" _
                    & ",(" & 0 & ")" _
                    & ",(" & .Cod_Parametro & ")" _
                    & ",(" & .Cod_Dimension & ")" _
                    & ",(" & .ValorMin & ")" _
                    & ",(" & .ValorMax & ")" _
                    & ",(" & .PRI & ")" _
                    & ",(" & .PW & ")" _
                    & ",(" & .Paso & ")" _
                    & ",(" & .AplicarPV & ")" _
                    & ",(" & .b50Ohms & ")" _
                    & ",(" & .VoltDiv & ")" _
                    & ",('" & .CurvaPV & "')" _
                    & ",(" & LV_Etapa & ") );"
                
            'lv_Sql = " INSERT INTO  " & LV_Nombre_Tabla _
                    & " (Cod_Proyecto,Cod_Instrumento,Cod_Parametro" _
                    & ",Cod_Dimension,ValorMin,ValorMax" _
                    & ",Paso,Etapa) " _
                    & " VALUES ((" & LV_Cod_Pjt & ")" _
                    & ",(" & 0 & ")" _
                    & ",(" & .Cod_Parametro & ")" _
                    & ",(" & .Cod_Dimension & ")" _
                    & ",(" & .ValorMin & ")" _
                    & ",(" & .ValorMax & ")" _
                    & ",(" & .Paso & ")" _
                    & ",(" & LV_Etapa & ") );"
                
            
        End If
        
        GV_BD_INSTRUMENT.Execute lv_Sql
    
    End With
    
End Function

