Attribute VB_Name = "modFuncionesVarias"
Option Explicit

Function CalcMeanValue(LV_Data() As String) As Double

Dim i           As Integer
Dim LV_Total    As Double

    LV_Total = 0
    For i = 0 To UBound(LV_Data)
        LV_Total = LV_Total + Val(LV_Data(i))
    Next
    
    CalcMeanValue = LV_Total / (UBound(LV_Data) + 1)
    
End Function

Function QuitarValores(LV_Str As String, LV_Sep As String, LV_Qty As Integer)

Dim i       As Integer
Dim k       As Integer

    For i = 1 To LV_Qty
        k = InStr(LV_Str, LV_Sep)
        If k Then
            LV_Str = Right$(LV_Str, Len(LV_Str) - k)
        Else
            Exit For
        End If
    Next
    
End Function

Sub Fill_LstVw_Ptos_Carac(LV_LstVw As ListView, LV_Ptos() As Type_Ptos_Charac)

Dim LV_Campos(3)        As String
Dim i                   As Integer

    If UBound(LV_Ptos) >= 0 Then
        For i = 0 To UBound(LV_Ptos)
            With LV_Ptos(i)
                LV_Campos(i) = .Pot_In
                LV_Campos(i) = .Freq
                LV_Campos(i) = .Pot_Out
                LV_Campos(i) = .Gain
                AddItemListView LV_LstVw, LV_Campos, True
            End With
        Next
    Else
        LV_LstVw.ListItems.Clear
    End If
    
End Sub

Function If_Index_Meet(Index As Integer, LV_List() As Integer) As Boolean

Dim i           As Integer

    For i = 0 To UBound(LV_List)
        If LV_List(i) = Index Then
            If_Index_Meet = True
            Exit Function
        End If
    Next
    
    If_Index_Meet = False
    
End Function

Sub Iniciar_Estructura_Proyecto_Vacia()

    With GV_Actual_Project
        
        .Cod_Project = 0
        .Dispositivo = ""
        .Encargado = ""
        .EtapasDeControl = 0
        .Fecha = ""
        .Flag_UpDate = False
        .Num_Parte = ""
        .Num_Serie = ""
        .Path_Project = ""
        .Project_Name = ""
        'Set .Rango = Nothing
        .Result_File = ""
        
    End With
    
End Sub

Sub Set_Column_Lista_Instrumentos(LstVw As ListView, LV_Column() As String)

    ReDim LV_Column(7)
    
    LV_Column(0) = "Dispositivo"
    LV_Column(1) = "Funcion"
    LV_Column(2) = "Nombre"
    LV_Column(3) = "Fabricante"
    LV_Column(4) = "Modelo"
    LV_Column(5) = "Numero de Parte"
    LV_Column(6) = "Numero de Serie"
    LV_Column(7) = "Comunicacion"
    
    AddColumListView LstVw, LV_Column
    
End Sub

Sub SeleccionarText(LV_TextBox As TextBox)

    With LV_TextBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Sub Update_Project(LV_Campos() As String)

    With GV_Actual_Project
        
        '.Cod_Project = LV_Campos()
        .Dispositivo = LV_Campos(2)
        .Encargado = LV_Campos(8)
        .EtapasDeControl = 0
        .Fecha = LV_Campos(5)
        .Flag_UpDate = False
        .Num_Parte = LV_Campos(3)
        .Num_Serie = LV_Campos(4)
        .Path_Project = LV_Campos(1)
        .Project_Name = LV_Campos(0)
        'Set .Rango = Nothing
        .Result_File = LV_Campos(6)
        .CompensacionSetup = LV_Campos(9)
        
    End With
    
End Sub

Function ConvTextToNumeric(ByVal LV_Txt As String, LV_Num As Double) As Boolean
' Devuelve Verdadero si el número es válido

Dim LV_Ch_Decimal       As String
Dim LV_Ch_Find          As String

    LV_Ch_Decimal = RegGetValue$(HKEY_CURRENT_USER, "Control Panel\International", "sDecimal")

    If LV_Ch_Decimal = "," Then
        LV_Ch_Find = "."
    Else
        LV_Ch_Find = ","
    End If
    
    LV_Txt = Replace(LV_Txt, LV_Ch_Find, LV_Ch_Decimal)
    
    If IsNumeric(LV_Txt) = True Then
        ConvTextToNumeric = True
        LV_Num = Val(LV_Txt)
    Else
        ConvTextToNumeric = False
        LV_Num = 0
    End If
    
End Function

Function Get_Decimal_From_Regional_Config() As String

    Get_Decimal_From_Regional_Config = RegGetValue$(HKEY_CURRENT_USER, "Control Panel\International", "sList")

End Function

Function GetValidDataFromStr(ByVal LV_Str As String) As String

Dim LV_Campos()         As String
Dim i                   As Integer

    LV_Str = Trim$(LV_Str)
    
    i = InStr(1, LV_Str, vbLf)
    
    If i Then
        LV_Str = Left$(LV_Str, i - 1)
    End If

    i = InStr(1, LV_Str, vbCr)
    
    If i Then
        LV_Str = Left$(LV_Str, i - 1)
    End If

    i = InStr(1, LV_Str, vbCrLf)
    
    If i Then
        LV_Str = Left$(LV_Str, i - 1)
    End If

    GetValidDataFromStr = LV_Str
    
End Function

Function Verify_Valid_Digit(ByVal LV_Ascii As Byte) As Boolean

    If LV_Ascii >= 48 And LV_Ascii <= 57 Then
        Verify_Valid_Digit = True
    ElseIf LV_Ascii >= 43 And LV_Ascii <= 46 Then
        Verify_Valid_Digit = True
    Else
        Verify_Valid_Digit = False
    End If
    
End Function

Function Obtener_Archivo_Salida() As String

    With GV_Actual_Project
        Obtener_Archivo_Salida = .Path_Project & "\" & .Result_File
    End With
    
End Function

Function Put_Cod_Instru_To_Comandos(LV_Cod_Instru As Integer, LV_Cmds() As Type_Comando)

Dim i           As Integer

    'If IsEmpty(LV_Cmds) = False Then
        For i = 0 To UBound(LV_Cmds)
            With LV_Cmds(i)
                .Cod_Instrumento = LV_Cod_Instru
            End With
        Next
    'End If
    
End Function

Function Retroceder_Path(ByVal lsPath As String) As String

Dim i           As Integer
Dim lsTemp()    As String
    
    lsTemp = Split(lsPath, "\")
    
    ReDim Preserve lsTemp(UBound(lsTemp) - 1)
    
    Retroceder_Path = Join(lsTemp, "\")

End Function

Function IsArrayDoubleVacio(LV_Array() As Double) As Boolean

Dim i       As Long

    On Error GoTo Array_Empty
    
    i = UBound(LV_Array)
    
    IsArrayDoubleVacio = False
    
    Exit Function
    
Array_Empty:

    IsArrayDoubleVacio = True
    On Error GoTo 0

End Function

Function ExtraerNumeric(ByVal LV_Value As String) As Double

Dim i       As Integer
Dim LV_Ch   As Integer

    Do
        LV_Ch = Asc(Left$(LV_Value, 1))
        If LV_Ch >= 65 And LV_Ch <= 90 Then
            LV_Value = Right$(LV_Value, Len(LV_Value) - 1)
        ElseIf LV_Ch >= 97 And LV_Ch <= 122 Then
            LV_Value = Right$(LV_Value, Len(LV_Value) - 1)
        Else
            Exit Do
        End If
    Loop Until LV_Value = ""
    
End Function

Function VerificarExiste(LV_File As String) As Boolean

Dim filesys                 ', newfolder, newfolderpath

    Set filesys = CreateObject("Scripting.FileSystemObject")
    
    If filesys.FileExists(LV_File) Then
        VerificarExiste = True
    Else
        VerificarExiste = False
    End If
    
End Function

Function Verificar_Archivo(LV_File As String) As Boolean

        
End Function


Function Convertir_Video_en_Pot(LV_Video As Double, LV_Frec As Long) As Double

Dim LV_Index_F      As Integer
Dim LV_Index_P      As Integer
Dim LV_Frec_Ok      As Boolean
Dim LV_Inc_P        As Double

    LV_Frec_Ok = False
    For LV_Index_F = 0 To UBound(GV_Lista_Frec)
        If GV_Lista_Frec(LV_Index_F) = LV_Frec Then
            LV_Frec_Ok = True
            Exit For
        End If
    Next
    
    If LV_Frec_Ok = True Then
        For LV_Index_P = 1 To UBound(GV_Lista_Pot)
            If GV_Tabla_Vid_Pot(LV_Index_F).Filas(LV_Index_P) > LV_Video Then
                Convertir_Video_en_Pot = GV_Lista_Pot(LV_Index_P - 1)
                LV_Inc_P = LV_Video - GV_Tabla_Vid_Pot(LV_Index_F).Filas(LV_Index_P - 1)
                LV_Inc_P = LV_Inc_P / (GV_Tabla_Vid_Pot(LV_Index_F).Filas(LV_Index_P) - GV_Tabla_Vid_Pot(LV_Index_F).Filas(LV_Index_P - 1))
                LV_Inc_P = LV_Inc_P * (GV_Lista_Pot(LV_Index_P) - GV_Lista_Pot(LV_Index_P - 1))
                Convertir_Video_en_Pot = Convertir_Video_en_Pot + LV_Inc_P
                Exit For
            End If
        Next
    Else
        Convertir_Video_en_Pot = LV_Video
    End If
    
    
    
    

End Function
