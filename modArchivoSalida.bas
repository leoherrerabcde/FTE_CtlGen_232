Attribute VB_Name = "modArchivoSalida"
Option Explicit

Global GV_Archivo_Salida        As String

Sub Crear_Archivo_Salida(ByRef LV_Hdl As Integer, _
                        ByVal LV_File_Name As String)

Dim LV_Text         As String
Dim i               As Integer

    LV_Hdl = FreeFile
    
    Open LV_File_Name For Output As LV_Hdl
    
    Print #LV_Hdl, GV_Actual_Project.Dispositivo
    LV_Text = "Num Parte;" & GV_Actual_Project.Num_Parte
    LV_Text = LV_Text & ";Num Serie;" & GV_Actual_Project.Num_Serie
    Print #LV_Hdl, LV_Text
    Print #LV_Hdl, GV_Actual_Project.OCompra
    Print #LV_Hdl, "Hora Comienzo:" & GV_Ch_Decimal & Format(Time(), "hh:mm:ss")
    
    For i = 0 To UBound(GV_Data_Captur.NombreCampo)
        If i = 0 Then
            LV_Text = GV_Data_Captur.NombreCampo(i)
        Else
            LV_Text = LV_Text & GV_Ch_Decimal & GV_Data_Captur.NombreCampo(i)
        End If
    Next
    
    Print #LV_Hdl, LV_Text
    
End Sub

Sub Guardar_Data(ByVal LV_Hdl As Integer, _
                  LV_Data() As Double)
                  
Dim LV_Text     As String
Dim i           As Integer

    
    For i = 0 To UBound(LV_Data)
        If i = 0 Then
            LV_Text = LV_Data(i)
        Else
            LV_Text = LV_Text & GV_Ch_Decimal & LV_Data(i)
        End If
    Next
    
    Print #LV_Hdl, LV_Text
    
End Sub
                

