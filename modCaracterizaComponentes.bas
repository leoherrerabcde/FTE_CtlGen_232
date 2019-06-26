Attribute VB_Name = "modCaracterizaComponentes"
Option Explicit


Function Open_Characteriz_File(LV_File As String, LV_Ptos() As Type_Ptos_Charac) As Boolean

Dim i               As Integer
Dim h               As Integer
Dim LV_Data         As String
Dim LV_Campos()     As String
Dim LV_Separa       As String

    If LV_File = "" Then
        Open_Characteriz_File = False
        Exit Function
    End If
    
    Open_Characteriz_File = False
    
    h = FreeFile
    
    Open LV_File For Input As h
    
    For i = 1 To 5
        If EOF(h) = False Then
            Line Input #h, LV_Data
        Else
            Open_Characteriz_File = False
            Exit Function
        End If
    Next
    
    LV_Separa = Get_Decimal_From_Regional_Config
    
    i = 0
    Do
        If EOF(h) = False Then
            
            ReDim Preserve LV_Ptos(i)
            Line Input #h, LV_Data
            LV_Campos() = Split(LV_Data, LV_Separa)
            
            If UBound(LV_Campos) >= 2 Then
                With LV_Ptos(i)
                    .Freq = LV_Campos(1)
                    .Pot_In = LV_Campos(0)
                    .Pot_Out = LV_Campos(2)
                    .Gain = .Pot_Out - .Pot_In
                End With
            Else
                Exit Do
            End If
            i = i + 1
        Else
        
            Open_Characteriz_File = True
            Exit Do
            
        End If
    Loop
    
    Close #h
    h = 0
    
    
End Function
