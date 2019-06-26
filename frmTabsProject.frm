VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTabsProject 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   4890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7695
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTabProject 
      Height          =   4095
      Left            =   480
      TabIndex        =   0
      Top             =   240
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmTabsProject.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "pictureDock(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Dispositivos"
      TabPicture(1)   =   "frmTabsProject.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pictureDock(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Setup"
      TabPicture(2)   =   "frmTabsProject.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "pictureDock(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Rangos"
      TabPicture(3)   =   "frmTabsProject.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "pictureDock(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Salida"
      TabPicture(4)   =   "frmTabsProject.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "pictureDock(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Prueba"
      TabPicture(5)   =   "frmTabsProject.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "pictureDock(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Resultados"
      TabPicture(6)   =   "frmTabsProject.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "pictureDock(6)"
      Tab(6).ControlCount=   1
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   6
         Left            =   -74640
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   7
         Top             =   600
         Width           =   4215
      End
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   5
         Left            =   -74760
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   6
         Top             =   600
         Width           =   4215
      End
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   4
         Left            =   -74760
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   5
         Top             =   660
         Width           =   4215
      End
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   3
         Left            =   -74880
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   4
         Top             =   780
         Width           =   4215
      End
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   3
         Top             =   780
         Width           =   4215
      End
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   1
         Left            =   -74760
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   2
         Top             =   660
         Width           =   4215
      End
      Begin VB.PictureBox pictureDock 
         Height          =   2415
         Index           =   0
         Left            =   120
         ScaleHeight     =   2355
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   660
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmTabsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_GotFocus()


    With Me
        If .WindowState <> vbMaximized Then
            .WindowState = vbMaximized
        End If
        
    End With
    
End Sub

Private Sub Form_Load()

Dim i           As Integer

    With Me
        For i = 0 To .pictureDock.UBound
            '.pictureDock(i).BorderStyle = 0
        Next
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim i           As Integer

    With Me

        Unload frmNewProject
        Unload frmDispositivos
        Unload frmRangoControl
        Unload frmPrueba
        Unload frmDispositivos
        Unload frmConfigSalida
        Unload frmResultados
        
        '.Show
        
        '.SSTabProject.TabIndex = 0
        
    End With
    
End Sub

Private Sub Form_Resize()

Dim i           As Integer

    With Me
        .SSTabProject.Left = 0
        .SSTabProject.Top = 0
        .SSTabProject.Width = .ScaleWidth
        .SSTabProject.Height = .ScaleHeight
        
        i = 0
        'For i = .pictureDock.UBound To 0 Step -1
        
            .pictureDock(i).Top = .SSTabProject.TabHeight * Int((.SSTabProject.Tabs + .SSTabProject.TabsPerRow - 1) / .SSTabProject.TabsPerRow)
            .pictureDock(i).Left = 0
            .pictureDock(i).Height = .SSTabProject.Height - .pictureDock(i).Top
            .pictureDock(i).Width = .SSTabProject.Width
            .pictureDock(i).Tag = "OK"
            
        'Next
        
    End With
    
End Sub

Private Sub pictureDock_Resize(Index As Integer)


    Select Case Index
        Case Is = 0
            dockForm frmNewProject.hWnd, Me.pictureDock(Index), True
            
        Case Is = 1
            dockForm frmDispositivos.hWnd, Me.pictureDock(Index), True
        Case Is = 2
        
        Case Is = 3
            dockForm frmRangoControl.hWnd, Me.pictureDock(Index), True
        
        Case Is = 4
            dockForm frmConfigSalida.hWnd, Me.pictureDock(Index), True
        
        Case Is = 5
            dockForm frmPrueba.hWnd, Me.pictureDock(Index), True
        
        Case Is = 6
            dockForm frmResultados.hWnd, Me.pictureDock(Index), True
            
    End Select
    
End Sub

Private Sub SSTabProject_Click(PreviousTab As Integer)

Dim i       As Integer

    With Me
        
        i = .SSTabProject.Tab
        
        If .pictureDock(i).Tag <> "OK" Then
            .pictureDock(i).Top = .SSTabProject.TabHeight * Int((.SSTabProject.Tabs + .SSTabProject.TabsPerRow - 1) / .SSTabProject.TabsPerRow)
            .pictureDock(i).Left = 0
            .pictureDock(i).Height = .SSTabProject.Height - .pictureDock(i).Top
            .pictureDock(i).Width = .SSTabProject.Width
        Else
            .pictureDock(i).Refresh
        End If
    End With
    
End Sub

