Attribute VB_Name = "modGPIBInstrumentos"
Option Explicit

Global Const GPIB0 = 0

Const NO_SECONDARY_ADDR = 0         ' Secondary address of device
'Const timeout = T10s                ' Timeout value = 10 seconds
Const EOTMODE = 1                   ' Enable the END message
Const EOSMODE = 0                   ' Disable the EOS mode

Enum TIPO_INSTRUMENTO

    MEDICION = 0
    GENERACION

End Enum

Public Enum Type_Communica
    COMM_RS232 = 1
    COMM_GPIB
    COMM_USB
End Enum

Type Type_Instrumento

    address         As Long
    Cmd_Config      As String
    Cmd_Set_Var()   As String
    Cmd_Set_Param() As Integer
    Cmd_Consult()   As String
    Cmd_End         As String
    Cmd_Consu_Param()   As Integer
    Dev_Handle      As Long
    Tipo            As TIPO_INSTRUMENTO
    Comunicacion    As Type_Communica
    TpoEspera       As Long
    Name            As String

End Type

Type Type_Anali_Sp
    instrID         As Long
    CenterFreq      As Double
    SPAN            As Double
    RefLvl          As Double
End Type

Global GV_Instrumentos()     As Type_Instrumento
Global GV_Analizador_Sp()   As Type_Anali_Sp

Global GV_Address_List()    As Integer
Global GV_Result_List()     As Integer
Global Dev                  As Integer
Private PV_Volt_Div         As Double

Function CerrarCommInstrumento()

End Function

Sub Gen_Lists()

Dim i           As Integer

    ReDim GV_Address_List(UBound(GV_Instrumentos) + 1)
    ReDim GV_Result_List(UBound(GV_Instrumentos) + 1)
    'ReDim GV_Freq_List(UBound(GV_Instrumentos))
    
    For i = 0 To UBound(GV_Instrumentos)
        If GV_Instrumentos(i).Comunicacion = COMM_GPIB Then
            GV_Address_List(i) = GV_Instrumentos(i).address
        Else
            GV_Address_List(i) = 0
        End If
    Next
    
    'GV_Address_List(i) = NOADDR
    
End Sub

Function GetGPIBData(ByVal LV_Handle As Long, _
                     ByVal LV_Command As String, _
                     ByRef LV_Data As Double) As Boolean


Dim err_status      As Long
    
    'err_status = hp837xx_cmdReal(LV_Handle, LV_Command, LV_Data)
    
    If err_status = 0 Then
        GetGPIBData = True
    Else
        GetGPIBData = False
    End If
    
End Function

Function Get_Data_From_Instr(LV_Index As Integer) As String

Dim Str_Cmd         As String
Dim LV_Reading      As String
Dim LV_Data()       As String
Dim LV_Offset       As Double
'Dim LV_Volt_Div     As Double
Dim i               As Integer
Dim LV_Cmds()       As String

    Str_Cmd = GV_Instrumentos(LV_Index).Cmd_Consult(0)
    
    If GV_Instrumentos(LV_Index).TpoEspera Then
        'Sleep GV_Instrumentos(LV_Index).TpoEspera
    End If
    
    'Call Send(GPIB0, GV_Result_List(LV_Index), Str_Cmd, NLend)
    LV_Cmds = Split(Str_Cmd, ";")
    
    For i = 0 To UBound(LV_Cmds)
        'Call Send(GPIB0, GV_Result_List(LV_Index), Str_Cmd, NLend)
        'Call Send(GPIB0, GV_Result_List(LV_Index), LV_Cmds(i) & vbCrLf, NLend)
        'Call Send(GPIB0, GV_Result_List(LV_Index), LV_Cmds(i), NLend)
        Str_Cmd = Trim$(LV_Cmds(i))
    Next
    
'    If (ibsta And EERR) Then
'        Error_GPIB
''        MsgBox "Error sending '*IDN?'. "
''        End
'    End If
    
    If Str_Cmd = "CURVE?" Then
        LV_Reading = Space$(740 * 4)
        'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
        LV_Reading = Right$(LV_Reading, Len(LV_Reading) - 5)
        QuitarValores LV_Reading, ",", 520
        LV_Data = Split(LV_Reading, ",")
        If UBound(LV_Data) > 240 Then
        ReDim Preserve LV_Data(240)
        End If
        Get_Data_From_Instr = (GV_Volt_Div * (CalcMeanValue(LV_Data) - GV_Offset * 25))
        Exit Function
    ElseIf Str_Cmd = "AVG?" Then
        LV_Reading = Space$(&H32)
        'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
        LV_Reading = Right$(LV_Reading, Len(LV_Reading) - 4)
        Get_Data_From_Instr = (GV_Volt_Div * (Val(LV_Reading) - GV_Offset * 25))
        Exit Function
    ElseIf Str_Cmd = "MFL?" Then
        LV_Reading = Space$(&H32)
        'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
        LV_Reading = Right$(LV_Reading, Len(LV_Reading) - 4)
        LV_Data = Split(LV_Reading, ",")
        If UBound(LV_Data) Then
            'LV_Reading = GetValidDataFromStr(LV_Reading)
            Get_Data_From_Instr = GetValidDataFromStr(LV_Data(1))
        End If
        Exit Function
    Else
        LV_Reading = Space$(&H32)
        
        'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
        
    End If
    
    'End If
    
'    If ibcntl > 0 Then
'        LV_Reading = Left$(LV_Reading, ibcntl - 1)
'
'        LV_Reading = GetValidDataFromStr(LV_Reading)
'    Else
'    End If
    'LV_Reading = ""
    Get_Data_From_Instr = LV_Reading
    
End Function

Private Sub GpibErr(Msg$)
    Msg$ = Msg$ '+ AddIbsta() + AddIberr() + AddIbcnt() + Chr$(13) + Chr$(13) + "I'm quitting!"
    MsgBox Msg$, vbOKOnly + vbExclamation, "Error"
    
End Sub

Function IniciarCommInstrumento() As Long

Dim err_status          As Long
Dim SessionID           As Long
Dim LV_Command          As String
Dim ErrMsg              As String
Dim LV_i                As Integer
    
    
    ' Inicializar
'    Call SendIFC(GPIB0)
'    If (ibsta And EERR) Then
'        Error_GPIB
''         MsgBox "Error sending IFC."
''         End
'    End If

    
    Gen_Lists
    
'    Call FindLstn(GPIB0, GV_Address_List, GV_Result_List, UBound(GV_Address_List))
'    If (ibsta And EERR) Then
'        Error_GPIB
'
'         'MsgBox "Error finding all listeners."
'    End If

'    GV_Result_List(ibcntl) = NOADDR
    
'    Call DevClearList(GPIB0, GV_Result_List)
'    If (ibsta And EERR) Then
'        MsgBox "Error in clearing the devices. "
'    End If

    ' 3000
    
    For LV_i = 0 To UBound(GV_Instrumentos)
        
        With GV_Instrumentos(LV_i)
            If .Comunicacion = COMM_GPIB Then
                If .Name = "Analizador de Espectro" Then
                    Iniciar_Analizador
                End If
                    If GV_Address_List(LV_i) <> GV_Result_List(LV_i) Then
                        GV_Result_List(LV_i) = GV_Address_List(LV_i)
'                        Dev = ildev(0, GV_Address_List(LV_i), _
                                    0, 3, 1, 0)
'                        If (ibsta And EERR) Then
'                            GpibErr ("Error opening device.")
'                        End If
        
'                        ilclr Dev
'                        If (ibsta And EERR) Then
'                            GpibErr ("Error clearing device.")
'                        End If
        
                        If .Cmd_Config <> "" Then
'                            ilwrt Dev, .Cmd_Config, Len(.Cmd_Config)
'                            If (ibsta And EERR) Then
'                                GpibErr ("Error Sending Data to " & .Name)
'                                'GpibErr ("Error writing to device.")
'                            End If
                        End If
                    Else
                        If .Cmd_Config <> "" Then
'                            Call Send(GPIB0, GV_Result_List(LV_i), .Cmd_Config, NLend)
'                            If (ibsta And EERR) Then
'                                Error_GPIB
'    '                            MsgBox "Error sending '*IDN?'. "
'    '                            End
'                            End If
                        End If
                    End If
    '            Else
    '                IniciarCommPort
    '                If .Cmd_Config <> "" Then
    '                    SendRS232 .Cmd_Config
    '                End If
                'End If
            End If
        End With
    Next



'        .Dev_Handle = ildev(Index _
'                    , .Address _
'                    , NO_SECONDARY_ADDR _
'                    , TIMEOUT _
'                    , EOTMODE _
'                    , EOSMODE)
'
'        If (ibsta And EERR) Then
'            ErrMsg = "Unable to open device" & Chr(13) & "ibsta = &H" _
'                      & Hex(ibsta) & Chr(13) & "iberr = " & iberr
'            MsgBox ErrMsg, vbCritical, "Error"
'            End
'        End If
'
'        If .Cmd_Config <> "" Then
'            ilwrt .Dev_Handle, .Cmd_Config, Len(.Cmd_Config)
'
'            If (ibsta And EERR) Then
'
'                ilonl .Dev_Handle, 0
'
'            End If
'        End If
'
'    End With
'
'    Select Case Cod_Instrumento
'        Case Is = 0
'            ' Agilent
'            LV_Command = "GPIB0::19::INSTR"
'        Case Is = 1
'            ' Power Meter
'            LV_Command = "GPIB0::14::INSTR"
'        Case Is = 2
'            ' Tektronix
'            LV_Command = "GPIB0::2::INSTR"
'        Case Is = 3
'
'    End Select
    
    
'    err_status = hp837xx_init(LV_Command, True, True, SessionID)
'
'    If err_status = 1073676413 Then
'        'Error
'    Else
'        'Ok
'    End If
'
'    IniciarCommInstrumento = SessionID
    
End Function

Function SendCmd(LV_Cmds As String, Index As Integer)

'    Call Send(GPIB0, GV_Result_List(Index), LV_Cmd & vbCrLf, NLend)
    fMainForm.SendRS232 LV_Cmds & vbCrLf
'    If (ibsta And EERR) Then
'        Error_GPIB
''        MsgBox "Error sending '*IDN?'. "
''        End
'    End If

End Function

' LV_Freq in [Hz]
Function SendRFPowerOff(ByVal Index As Integer)

Dim LV_Cmds             As String

    LV_Cmds = "OUTP OFF"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(Index), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf
    
End Function

Function SendIntModulacionON()

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULM:SOUR INT; STAT ON"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf
    
End Function

Function SendIntFMModulacionON()

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:FM:SOUR INT; STAT ON"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf
    
End Function

Function SendFmModulacionOFF()

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:FM:STAT OFF"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendFmModulacionON()

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:FM:STAT ON"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendModulacionON()

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULM:STAT ON"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf
    
End Function

Function SendModulacionOFF()

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULM:STAT OFF"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf
    
End Function

Function SendRFPowerOn(ByVal Index As Integer)

Dim LV_Cmds             As String

    LV_Cmds = "OUTPUT ON"
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(Index), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf
    
End Function

Function SendALCState(ByVal Index As Integer, LV_ALC_ON As Boolean)

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:POW:ALC "
    
    If LV_ALC_ON = True Then
        LV_Cmds = LV_Cmds & "ON"
    Else
        LV_Cmds = LV_Cmds & "OFF"
    End If
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(Index), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendPRI(ByVal Index As Integer, LV_PRI As Double)

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULS:PER "
    
    LV_Cmds = LV_Cmds & Trim$(LV_PRI) & "us"
    LV_Cmds = Replace(LV_Cmds, ",", ".")
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(Index), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendPW(ByVal Index As Integer, LV_PW As Double)

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULS:WIDT "
    
    LV_Cmds = LV_Cmds & Trim$(LV_PW) & "us"
    LV_Cmds = Replace(LV_Cmds, ",", ".")
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(Index), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendPulseDelay(LV_PulseDelay As Double)

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULS:DEL "
    
    LV_Cmds = LV_Cmds & Trim$(LV_PulseDelay) & "us"
    LV_Cmds = Replace(LV_Cmds, ",", ".")
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendTriggerMode(LV_Mode As String)

Dim LV_Cmds             As String

    LV_Cmds = "SOUR:PULM:SOUR INT"
    fMainForm.SendRS232 LV_Cmds & vbCrLf

    LV_Cmds = "TRIG:PULS:SOUR "
    
    LV_Cmds = LV_Cmds & Trim$(LV_Mode)
    
    On Error Resume Next
'    Call Send(GPIB0, GV_Result_List(0), LV_Cmds & vbCrLf, NLend)
    On Error GoTo 0
    fMainForm.SendRS232 LV_Cmds & vbCrLf

End Function

Function SendFrec(ByVal Index As Integer, _
                    ByVal LV_Freq As Double, _
                    Optional LV_Str As String) As Boolean

Dim Str_Cmd         As String
Dim LV_Str_Val      As String
Dim LV_Cmds()       As String
Dim i               As Integer

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
        'Call Send(GPIB0, GV_Result_List(Index), Str_Cmd, NLend)
        fMainForm.SendRS232 Str_Cmd & vbCrLf
        'LV_Cmds(i) = "FREQ 1000MHz"
        'ParseCommand LV_Cmds(i)
        
'        Call Send(GPIB0, GV_Result_List(Index), LV_Cmds(i) & vbCrLf, NLend)
        'fMainForm.SendRS232 Str_Cmd & vbCrLf
        'Call Send(GPIB0, GV_Result_List(Index), LV_Cmds(i), NLend)
    Next
    
'    If (ibsta And EERR) Then
'        Error_GPIB
''        MsgBox "Error sending '*IDN?'. "
''        End
'    End If
    


'Dim LV_Comando      As String
'Dim err_status      As Long
'
'    LV_Comando = "FREQ:CW"
'
'    err_status = hp837xx_cmdReal(LV_Handle, LV_Comando, LV_Freq)
'
'    If err_status = 0 Then
'        SendFrec = True
'    Else
'        SendFrec = False
'    End If
'
End Function

'Public Function SendGPIBCommand(ByVal LV_Handle As Long, _
'                                ByVal LV_Command As String, _
'                                ByVal LV_Data As Double) As Boolean
'
'Dim LV_Comando      As String
'Dim err_status      As Long
'
'    err_status = hp837xx_cmdReal(LV_Handle, LV_Comando, LV_Data)
'
'    If err_status = 0 Then
'        SendGPIBCommand = True
'    Else
'        SendGPIBCommand = False
'    End If
'
'End Function
Function Replace_Value(LV_Cmd As String, LV_Value As String, LV_Id As String) As String

Dim i       As Integer

    i = InStr(LV_Cmd, LV_Id)
    
    If i Then
        Replace_Value = Left$(LV_Cmd, i - 1) & LV_Value & Right$(LV_Cmd, Len(LV_Cmd) - i + 1 - Len(LV_Id))
    Else
        Replace_Value = LV_Cmd & LV_Value
    End If

End Function

Function SendPow(ByVal Index As Integer, _
                    ByVal LV_Pow As Double) As Boolean
                    
Dim Str_Cmd         As String
Dim LV_Str_Val      As String

    LV_Str_Val = LV_Pow
    LV_Str_Val = Replace(LV_Str_Val, ",", ".")
    'Str_Cmd = GV_Instrumentos(Index).Cmd_Set_Var(1) & " " & LV_Str_Val
    Str_Cmd = Replace_Value(GV_Instrumentos(Index).Cmd_Set_Var(1), LV_Str_Val, "%P")
    
    'Str_Cmd = "POW -12dBm"
    
'    Call Send(GPIB0, GV_Result_List(Index), Str_Cmd, NLend)
    fMainForm.SendRS232 Str_Cmd & vbCrLf
    
    
'    If (ibsta And EERR) Then
'        Error_GPIB
''        MsgBox "Error sending '*IDN?'. "
''        End
'    End If
    

End Function


Function Verify_Frec_State(ByVal LV_Index As Integer, _
                            ByRef LV_Freq As Double) As Boolean

Dim Str_Cmd         As String
Dim LV_Reading      As String
Dim LV_Cmds()       As String
Dim i               As Integer

    Str_Cmd = GV_Instrumentos(LV_Index).Cmd_Consult(0)
    
'    Call Send(GPIB0, GV_Result_List(LV_Index), Str_Cmd, NLend)
    fMainForm.SendRS232 Str_Cmd & vbCrLf

'    If (ibsta And EERR) Then
'        Error_GPIB
''        MsgBox "Error sending '*IDN?'. "
''        End
'    End If
    
    
    LV_Reading = Space$(&H32)
    'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
'    If (ibsta And EERR) Then
'         MsgBox "Error in receiving response to '*IDN?'. "
'    End If
    
'    LV_Reading = Left$(LV_Reading, ibcntl - 1)
    LV_Freq = Val(LV_Reading)
    
'Dim LV_Comando      As String
'Dim err_status      As Long
'
'    LV_Comando = "FREQ:CW?"
'
'    err_status = hp837xx_cmdReal64_Q(LV_Handle, "FREQ:CW?", LV_Freq)
'
'    If err_status = 0 Then
'        Verify_Frec_State = True
'    Else
'        Verify_Frec_State = False
'    End If
    
End Function

Function Verify_Pow_State(ByVal LV_Index As Integer, _
                            ByRef LV_Pow As Double) As Boolean

Dim Str_Cmd         As String
Dim LV_Reading      As String

    Str_Cmd = GV_Instrumentos(LV_Index).Cmd_Consult(1)
    
'    Call Send(GPIB0, GV_Result_List(LV_Index), Str_Cmd, NLend)
    fMainForm.SendRS232 Str_Cmd & vbCrLf
    
'    If (ibsta And EERR) Then
'        Error_GPIB
''        MsgBox "Error sending '*IDN?'. "
''        End
'    End If
    
    
    LV_Reading = Space$(&H32)
    'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
'    If (ibsta And EERR) Then
'         MsgBox "Error in receiving response to '*IDN?'. "
'    End If
    
'    LV_Reading = Left$(LV_Reading, ibcntl - 1)
    LV_Pow = Val(LV_Reading)
    
    
End Function

Function Error_GPIB()

    MsgBox "Error sending '*IDN?'. "
    End

End Function

Function Capturar_Pot_Gen_RS232(LV_Index As Integer) As Double

Dim Str_Cmd         As String
Dim LV_Reading      As String

    Str_Cmd = GV_Instrumentos(LV_Index).Cmd_Consult(1)
    
    fMainForm.SendRS232 Str_Cmd

    LV_Reading = Space$(&H32)
    
    Do
        DoEvents
    Loop Until GV_Data_Instrument_Ok = True
    
    If GV_Data_Instrument <> "" Then
        LV_Reading = GV_Data_Instrument
        Capturar_Pot_Gen_RS232 = Val(LV_Reading)
    Else
        Capturar_Pot_Gen_RS232 = "0"
    End If
    
End Function

Function Capturar_Pot_Gen(LV_Index As Integer)

Dim Str_Cmd         As String
Dim LV_Reading      As String

    Str_Cmd = GV_Instrumentos(LV_Index).Cmd_Consult(1)
    
'    Call Send(GPIB0, GV_Result_List(LV_Index), Str_Cmd, NLend)
    fMainForm.SendRS232 Str_Cmd & vbCrLf
    
'    If (ibsta And EERR) Then
'        Error_GPIB
'    End If
        
    LV_Reading = Space$(&H32)
    'Call Receive(GPIB0, GV_Result_List(LV_Index), LV_Reading, STOPend)
'    If (ibsta And EERR) Then
'         MsgBox "Error in receiving response to '*IDN?'. "
'    End If
    
'    LV_Reading = Left$(LV_Reading, ibcntl - 1)
    Capturar_Pot_Gen = Val(LV_Reading)
    
End Function

Function Iniciar_Analizador()

Dim instrID         As Long
Dim resetDevice     As Integer

    With GV_Analizador_Sp(0)
'        AVULIS_init "GPIB::8::INSTR", VI_ON, resetDevice, .instrID
        
        'AVULIS_Set_Span .instrID, 0, .Span
        
        'AVULIS_Set_Freq .instrID, 0, .CenterFreq
        
        'AVULIS_Set_PeakSearchOptions
        
    End With
    
End Function

Function ParseCommand(LV_Cmd As String)

Dim i           As Integer

    i = InStr(LV_Cmd, "DELAY")
    
    If i Then
'        Sleep Mid$(LV_Cmd, i + 5)
    End If
    
End Function

Function ComandoOsciloscopio(b_50Ohms As Integer, _
                            b_Inverter As Integer, _
                            ByVal LV_Volt_Div As Integer)

Dim LV_50Ohms               As String
Dim LV_Inverter             As String

    GV_Volt_Div = LV_Volt_Div / 25
    GV_Offset = -4
    
    LV_Volt_Div = LV_Volt_Div / 1000#
    LV_50Ohms = "OFF"
    LV_Inverter = "OFF"
    If b_50Ohms = 1 Then
        LV_50Ohms = "ON"
    End If
    If b_Inverter = 1 Then
        LV_Inverter = "ON"
    End If
    
    ComandoOsciloscopio = "DAT:SOU CH1" _
                & ";:DAT:ENC ASCII" _
                & ";:DAT:WID 1" _
                & ";:DAT:STAR 1" _
                & ";:DAT:STOP 50" _
                & ";CH1 VOLts:" & LV_Volt_Div _
                & ";CH1 POSition:-4" _
                & ";CH1 FIFty:" & LV_50Ohms _
                & ";CH1 INVERT:" & LV_Inverter _
                & ";:HOR:MAIN:SCALE 5e-4"

End Function
