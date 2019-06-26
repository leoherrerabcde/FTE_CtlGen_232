VERSION 5.00
Begin VB.Form frmDialogNewProject 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Proyecto"
   ClientHeight    =   4215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6495
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameNewProject 
      Caption         =   "Informaci�n General"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3600
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   255
         Left            =   4920
         TabIndex        =   11
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscarPath 
         Caption         =   "Command1"
         Height          =   195
         Left            =   5760
         TabIndex        =   10
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   8
         Left            =   2160
         TabIndex        =   9
         Text            =   "txtProyecto"
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   8
         Text            =   "txtProyecto"
         Top             =   2880
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   7
         Text            =   "txtProyecto"
         Top             =   2520
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   6
         Text            =   "txtProyecto"
         Top             =   2160
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   5
         Text            =   "txtProyecto"
         Top             =   1800
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   4
         Text            =   "txtProyecto"
         Top             =   1440
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   3
         Text            =   "txtProyecto"
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   2
         Text            =   "txtProyecto"
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtProyecto 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   1
         Text            =   "txtProyecto"
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Encargado de la Prueba :"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   21
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Orden de Compra"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Archivo de Resultados :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Fecha de Creaci�n :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   18
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "N�m Serie :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "N�m Parte :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Dispositivo Bajo Prueba :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Ubicaci�n :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblProyecto 
         Caption         =   "Nombre Proyecto :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
End
Attribute VB_Name = "frmDialogNewProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Function Guardar_Proyecto() As Boolean

Dim Nombre As String
Dim Ubicacion As String
Dim Disp_Prueba As String
Dim Num_Parte As String
Dim Num_Serie As String
Dim Fecha As String
Dim Result_File As String
Dim OCompra As String
Dim Encargado As String

    With Me
    
        Nombre = .txtProyecto(0).Text
        Ubicacion = .txtProyecto(1).Text
        Disp_Prueba = .txtProyecto(2).Text
        Num_Parte = .txtProyecto(3).Text
        Num_Serie = .txtProyecto(4).Text
        Fecha = .txtProyecto(5).Text
        Result_File = .txtProyecto(6).Text
        OCompra = .txtProyecto(7).Text
        Encargado = .txtProyecto(8).Text
    
    End With
    
    If GV_Actual_Project.Cod_Project = 0 Then
        
        GV_Actual_Project.Cod_Project = BD_Get_Cod_Project()
        
        If GV_Actual_Project.Cod_Project Then
        
            If MsgBox("El proyecto ya existe. �Desea Sobreescribirlo?", vbYesNo) = vbNo Then
                
                GV_Actual_Project.Cod_Project = 0
                
                Guardar_Proyecto = False
                
                Exit Function
                
            End If
            
            BD_Update_Info_Gral_Proyecto Nombre, _
                        Ubicacion, _
                        Disp_Prueba, _
                        Num_Parte, _
                        Num_Serie, _
                        Fecha, _
                        Result_File, _
                        OCompra, _
                        Encargado, _
                        ""
            
        Else
        
            BD_Agregar_Proyecto Nombre, _
                    Ubicacion, _
                    Disp_Prueba, _
                    Num_Parte, _
                    Num_Serie, _
                    Fecha, _
                    Result_File, _
                    OCompra, _
                    Encargado, _
                    ""
                    
            GV_Actual_Project.Cod_Project = BD_Get_Cod_Project()
            
        End If
        
    Else
    
        BD_Update_Info_Gral_Proyecto Nombre, _
                    Ubicacion, _
                    Disp_Prueba, _
                    Num_Parte, _
                    Num_Serie, _
                    Fecha, _
                    Result_File, _
                    OCompra, _
                    Encargado, _
                    ""
        
    End If
    
    Guardar_Proyecto = True
    
End Function

Function Verificar_Campos_OK() As Boolean

Dim i       As Integer

    ' Verificar
    With Me
        For i = 0 To .txtProyecto.UBound
            If .txtProyecto(i).Text = "" Then
                MsgBox "El campo " & .lblProyecto(i) & " no ha sido ingresado correctamente", vbExclamation
                Verificar_Campos_OK = False
                Exit Function
            End If
        Next
    End With
    
    Verificar_Campos_OK = True

End Function

Private Sub cmdAceptar_Click()

Dim i           As Integer

    If Verificar_Campos_OK = True Then
    
        If Guardar_Proyecto = False Then
            Exit Sub
        End If
        
        Me.cmdAceptar.Enabled = False
        
        Unload Me
        
        fMainForm.LoadFormProject
        
    End If
    
End Sub

Private Sub cmdBuscarPath_Click()

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    lFlags = BIF_RETURNONLYFSDIRS
    lPath = ""
    
    sDir = BrowseForFolder(Me.hWnd, "Seleccionar Directorio", lPath, lFlags)

    If Err = 0 Then
        Me.txtProyecto(1).Text = sDir
    Else
        'MsgBox "Se ha cancelado la operaci�n, el error devuelto es:" & vbCrLf & _
               "Source: " & Err.Source & vbCrLf & "Description: " & Err.Description
        Err = 0
    End If

    ' Pero si es conveniente poner de nuevo el valor a cero

'    With dlgCommonDialog
'        .DialogTitle = "Abrir Proyecto"
'        .CancelError = False
'        .FileName = Me.txtProyecto(0).Text
'        'ToDo: set the flags and attributes of the common dialog control
'        '.Filter = "All Files (*.*)|*.*"
'        .Filter = "*.stp"
'        .ShowOpen
'        If Len(.FileName) = 0 Then
'            Exit Sub
'        End If
'        sFile = .FileName
'        Me.txtProyecto(1).Text = sFile
'    End With
    
End Sub

Private Sub cmdCancelar_Click()

    Unload Me
    
    Iniciar_Estructura_Proyecto_Vacia
    
    'MdiMain.Enabled = True
    
End Sub

Private Sub Form_Load()

Dim i           As Integer

    Iniciar_Estructura_Proyecto_Vacia
    
    With Me
        For i = 0 To .txtProyecto.UBound
            .txtProyecto(i).Text = ""
        Next
        .txtProyecto(5).Text = format(Now(), "DD-MM-YYYY")
        
        
'        .txtProyecto(0).Text = "Prueba"
'        .txtProyecto(1).Text = App.Path
'        .txtProyecto(2).Text = "DLVA"
'        .txtProyecto(3).Text = "ER"
'        .txtProyecto(4).Text = "070"
'        .txtProyecto(7).Text = "a"
'        .txtProyecto(8).Text = "b"
        
    End With
    
    
End Sub

Private Sub Form_LostFocus()

    Me.Show
    
End Sub

Private Sub lblProyecto_Click(Index As Integer)

    With Me
    
        .txtProyecto(Index).SetFocus
        
    End With
    
End Sub

Private Sub txtProyecto_Change(Index As Integer)

Dim lsFecha()           As String

    With Me
        
        If Index = 0 Or Index = 2 Or Index = 5 Or Index = 4 Then
            
            lsFecha = Split(.txtProyecto(5).Text, "-")
            
            If .txtProyecto(0) <> "" _
                And .txtProyecto(2) <> "" _
                And .txtProyecto(4) <> "" _
                Then
                '.txtProyecto(6).Text = .txtProyecto(0).Text & "-" _
                                        & .txtProyecto(2) _
                                        & "-" & .txtProyecto(4) _
                                        & "-" & lsFecha(2) _
                                        & "-" & lsFecha(1) _
                                        & "-" & lsFecha(0) _
                                        & ".csv"
                .txtProyecto(6).Text = .txtProyecto(0).Text & " NP" _
                                        & .txtProyecto(2) _
                                        & " NS" & .txtProyecto(4) _
                                        & ".csv"
                GV_Archivo_Salida = .txtProyecto(6).Text
            End If
        End If
        
        Select Case Index
        
        Case 0
            
            .Caption = "Proyecto " & .txtProyecto(Index).Text
            Call PostMessage(GV_hWnd_Mdi, NV_PROJECT_CHANGE_MSG, Project_Name, 0&)
            
            GV_Actual_Project.Project_Name = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True

        Case 1
            GV_Actual_Project.Path_Project = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        Case 2
            GV_Actual_Project.Dispositivo = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        Case 3
            GV_Actual_Project.Num_Parte = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        Case 4
            GV_Actual_Project.Num_Serie = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        Case 5
            GV_Actual_Project.Fecha = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        Case 6
            GV_Actual_Project.Result_File = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        Case 7
            GV_Actual_Project.OCompra = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
            
        Case 8
            GV_Actual_Project.Encargado = .txtProyecto(Index).Text
            GV_Actual_Project.Flag_UpDate = True
        
        End Select
    
        .cmdAceptar.Enabled = True
        
    End With
    
End Sub

Private Sub txtProyecto_GotFocus(Index As Integer)

    With Me.txtProyecto(Index)
    
        .SelStart = 0
        .SelLength = Len(.Text)
        
    End With
    
End Sub


