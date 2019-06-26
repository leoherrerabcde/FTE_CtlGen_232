VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbrirProyecto 
   BorderStyle     =   0  'None
   Caption         =   "Listado de Proyectos"
   ClientHeight    =   5220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameAbrirProyecto 
      Caption         =   "Frame1"
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton cmdAbrir 
         Caption         =   "&Abrir"
         Height          =   195
         Left            =   8160
         TabIndex        =   2
         Top             =   3960
         Width           =   855
      End
      Begin MSComctlLib.ListView LstVwProyectos 
         Height          =   2775
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   4895
         View            =   3
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
End
Attribute VB_Name = "frmAbrirProyecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PV_Cabecera()       As String

Sub Get_Proyectos()

    BD_Fill_LstVw_Projects Me.LstVwProyectos
    
End Sub

Private Sub cmdAbrir_Click()

    If Me.LstVwProyectos.SelectedItem Is Nothing Then
        Exit Sub
    End If
    GV_Actual_Project.Cod_Project = Me.LstVwProyectos.SelectedItem.Tag
    
    Unload Me
    
    fMainForm.LoadFormProject
    
    
End Sub

Private Sub cmdCancelar_Click()
    
    Unload Me
    
    Iniciar_Estructura_Proyecto_Vacia
    
End Sub

Private Sub Form_Load()

Dim i                   As Integer

    ReDim PV_Cabecera(7)
    
    PV_Cabecera(0) = "Nombre"
    PV_Cabecera(1) = "Ubicación"
    PV_Cabecera(2) = "Dispositivo"
    PV_Cabecera(3) = "Num Parte"
    PV_Cabecera(4) = "Num Serie"
    PV_Cabecera(5) = "Fecha"
    PV_Cabecera(6) = "Encargado"
    PV_Cabecera(7) = "Resultados"
    
    AddColumListView Me.LstVwProyectos, PV_Cabecera
    
    Me.Get_Proyectos
    
End Sub

Private Sub Form_Resize()

    With Me
        .frameAbrirProyecto.Width = .Width - 2 * .frameAbrirProyecto.Left
        .frameAbrirProyecto.Height = .Height - 2 * .frameAbrirProyecto.Top
        
        .cmdAbrir.Left = .frameAbrirProyecto.Width _
                        - .cmdAbrir.Width _
                        - 2 * .cmdCancelar.Left
        .cmdAbrir.Top = .frameAbrirProyecto.Height - 2 * .frameAbrirProyecto.Top
        .cmdCancelar.Top = .cmdAbrir.Top
        
        .LstVwProyectos.Width = .frameAbrirProyecto.Width _
                                - 2 * .LstVwProyectos.Left
        .LstVwProyectos.Height = .cmdAbrir.Top - 2 * .LstVwProyectos.Top
        
    End With
    
End Sub

Private Sub LstVwProyectos_DblClick()

    Me.cmdAbrir.value = True
    
End Sub

Private Sub LstVwProyectos_KeyPress(KeyAscii As Integer)

    If (Me.LstVwProyectos.SelectedItem Is Nothing) = False Then
        If KeyAscii = 13 Then
            Me.cmdAbrir.value = True
        End If
    End If
    
End Sub
