Attribute VB_Name = "modInstrumentos"
Option Explicit



Enum Enum_Orden_Tabla

    Tipo_Disp = 0
    Funcion
    Clase_Instru
    Fabricante
    Modelo
    PartNumber
    SerialNumber
    Comunicacion
    Clase_Componen
    
End Enum


Type TipoInfoDispositivo

    Componente      As String
    Comunicación    As String
    Dispositivo     As String
    Fabricante      As String
    Funcion         As String
    Modelo          As String
    Num_Parte       As String
    Num_Serie       As String
    
End Type


Global Info_Dispositivo         As TipoInfoDispositivo


Public GV_BD_INSTRUMENT             As DAO.Database
'Public PV_BD_MISION                 As DAO.Database
'Public PV_BD_PASSWORD               As DAO.Database
'Public PV_BD_REG_MISION             As DAO.Database

Sub Abrir_BD_Instrumentos()

Dim lv_DB_File      As String

    lv_DB_File = Retroceder_Path(App.Path) & "\Bd\" & "CyI.mdb"
    
    Set GV_BD_INSTRUMENT = OpenDatabase(lv_DB_File, False, False, "")
    
    
End Sub


Sub Consulta_Info_Dispositivo(ByVal Cod_Instrumento As Integer)

End Sub

Sub DB_New_Instrument()

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer


End Sub

Function DB_Sql_Select_Tipo(Index As Integer) As String

    DB_Sql_Select_Tipo = ""
    
    Select Case Index
        
        Case Is = Enum_Orden_Tabla.Tipo_Disp
            
            DB_Sql_Select_Tipo = "SELECT Tipo_Dispositivos.* From Tipo_Dispositivos;"
    
        Case Is = Enum_Orden_Tabla.Funcion
            
            DB_Sql_Select_Tipo = "SELECT Tipo_Funciones.* From Tipo_Funciones;"
    
        Case Is = Enum_Orden_Tabla.Clase_Instru
            
            DB_Sql_Select_Tipo = "SELECT Tipo_Instrumentos.* From Tipo_Instrumentos;"
    
        Case Is = Enum_Orden_Tabla.Fabricante
            
            DB_Sql_Select_Tipo = "SELECT Tipo_Fabricantes.* From Tipo_Fabricantes;"
    
        Case Is = Enum_Orden_Tabla.Comunicacion
            
            DB_Sql_Select_Tipo = "SELECT Tipo_Comunicacion.* From Tipo_Comunicacion;"
    
        Case Is = Enum_Orden_Tabla.Clase_Componen
            
            DB_Sql_Select_Tipo = "SELECT Tipo_Componentes.* From Tipo_Componentes;"
    
    End Select
    
End Function

Sub DB_Leer_Tipos(ByRef lLista() As String, Index As Integer)

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer

    lv_Sql = DB_Sql_Select_Tipo(Index)
            '& _
            '"WHERE ((Tipo_Dispositivos.Codigo)= ('" & Me.ComboNombreEmisor & "'));"
    
    If lv_Sql <> "" Then
        Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
        
        With LV_Query
            
            If Not .EOF And Not .BOF Then
            
                .MoveFirst
                
                i = 0
                
                Do
                
                    ReDim Preserve lLista(i)
                
                    lLista(i) = !Nombre
                
                    i = i + 1
                    
                    .MoveNext
                    
                Loop Until .EOF Or .BOF
                
            End If
            
        End With
        
    End If
    
End Sub

Sub Leer_Lista_Dispositivos(ByRef lLista() As String)

Dim lv_Sql          As String
Dim LV_Query        As Recordset
Dim i               As Integer

        lv_Sql = "SELECT Tipo_Dispositivos.* From Tipo_Dispositivos;"
                '& _
                '"WHERE ((Tipo_Dispositivos.Codigo)= ('" & Me.ComboNombreEmisor & "'));"
        
        Set LV_Query = GV_BD_INSTRUMENT.OpenRecordset(lv_Sql)
        With LV_Query
            If Not .EOF And Not .BOF Then
            
                ReDim lLista(.RecordCount)
                
                .MoveFirst
                
                i = 0
                
                Do
                
                    lLista(i) = !Nombre
                
                    i = i + 1
                    
                    .MoveNext
                    
                Loop Until .EOF Or .BOF
                
            End If
            
        End With
    
End Sub
