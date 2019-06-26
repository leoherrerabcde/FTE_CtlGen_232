VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form formTextDisplay 
   Caption         =   "Form1"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConfigComm 
      Caption         =   "RS232"
      Height          =   255
      Left            =   8880
      TabIndex        =   1
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtComCapture 
      Height          =   4935
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "formTextDisplay.frx":0000
      Top             =   120
      Width           =   9495
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   0
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
End
Attribute VB_Name = "formTextDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PV_CommPort             As Boolean

Sub ShowData(msg As String)

    With Me.txtComCapture
        .SelStart = Len(msg)
        .SelText = msg
    End With
    
End Sub

Function Capturar_Pot_Gen_RS232(LV_Index As Integer) As Double

Dim Str_Cmd         As String
Dim LV_Reading      As String

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

Function IniciarCommPort() As Boolean


End Function

Private Sub cmdConfigComm_Click()

    fMainForm.LoadFormRs232Props
    
End Sub

Private Sub Form_Load()

    With Me
        .txtComCapture.Text = ""
    End With
    
    'fMainForm.LoadFormRs232Props
    
    PV_CommPort = fMainForm.IniciarCommPort
    
    'If PV_CommPort = True Then
    '    Me.Capturar_Pot_Gen_RS232 0
    'End If
    'If .Cmd_Config <> "" Then
        'fMainForm.SendRS232 .Cmd_Config
    'End If

End Sub

Private Sub Form_Terminate()

    End
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End
    
End Sub

