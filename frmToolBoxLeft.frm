VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmToolBoxLeft 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Left"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTabProject 
      Height          =   1335
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmToolBoxLeft.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmToolBoxLeft.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frmToolBoxLeft.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmToolBoxLeft.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmToolBoxLeft.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).ControlCount=   0
   End
   Begin MSComctlLib.TreeView TreeViewProjectNav 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   2990
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "frmToolBoxLeft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With Me
    
        If X > .ScaleWidth - 25 Then
        
            X = X
            
            
            
        End If
        
    End With
 
 End Sub

Private Sub Form_Resize()

    With Me
        
        .TreeViewProjectNav.Top = 0
        .TreeViewProjectNav.Left = 0
        
        .TreeViewProjectNav.Height = .ScaleHeight / 2
        .TreeViewProjectNav.Width = .ScaleWidth
        
        .SSTabProject.Top = .TreeViewProjectNav.Height
        .SSTabProject.Left = 0
        
        .SSTabProject.Height = .ScaleHeight - .TreeViewProjectNav.Height
        .SSTabProject.Width = .ScaleWidth
        
    End With
    
End Sub
