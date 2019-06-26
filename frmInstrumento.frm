VERSION 5.00
Begin VB.Form frmInstrumento 
   Caption         =   "Información Instrumento"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7830
   LinkTopic       =   "Form3"
   ScaleHeight     =   6225
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Comandos de Comunicación"
      Height          =   2655
      Left            =   4920
      TabIndex        =   10
      Top             =   240
      Width           =   2775
   End
   Begin VB.Frame FrameDescParam 
      Caption         =   "Descripción de Parámetros"
      Height          =   2415
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   4575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información Dispositivo"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   7
         Left            =   2160
         TabIndex        =   18
         Text            =   "txtInfoDisp"
         Top             =   2880
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   6
         Left            =   2160
         TabIndex        =   17
         Text            =   "txtInfoDisp"
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   5
         Left            =   2160
         TabIndex        =   16
         Text            =   "txtInfoDisp"
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   4
         Left            =   2160
         TabIndex        =   15
         Text            =   "txtInfoDisp"
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   3
         Left            =   2160
         TabIndex        =   14
         Text            =   "txtInfoDisp"
         Top             =   1440
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   2
         Left            =   2160
         TabIndex        =   13
         Text            =   "txtInfoDisp"
         Top             =   1080
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Text            =   "txtInfoDisp"
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox txtInfoDisp 
         Height          =   285
         Index           =   0
         Left            =   2160
         TabIndex        =   11
         Text            =   "txtInfoDisp"
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "Serial Number :"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Comunicación :"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Part Number"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Modelo :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Instrumento / Componente :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fabricante :"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Funcion :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Dispositivo :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmInstrumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public PV_CodigoInstrumento        As Integer

Sub Refrescar_Datos()

    'Consulta_Info_Dispositivo PV_CodigoInstrumento
    
End Sub

Private Sub Form_Load()

    PV_CodigoInstrumento = 1
    
    Abrir_BD_Instrumentos
    
    Refrescar_Datos
    
End Sub
