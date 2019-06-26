VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmResultados 
   BorderStyle     =   0  'None
   Caption         =   "Resultados"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7170
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   720
      TabIndex        =   1
      Top             =   360
      Width           =   5175
   End
   Begin RichTextLib.RichTextBox RichTxtResultados 
      Height          =   2655
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   4683
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmResultados.frx":0000
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
