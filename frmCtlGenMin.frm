VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmCtlGenMin 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Prueba"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6990
   Icon            =   "frmCtlGenMin.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRS232 
      BackColor       =   &H00404040&
      Caption         =   "chkRS232"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3960
      TabIndex        =   66
      Top             =   3600
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   840
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm 
      Left            =   0
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RThreshold      =   1
      SThreshold      =   1
   End
   Begin VB.Frame framePrueba 
      Caption         =   "Estado de Control"
      Height          =   3975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      Begin VB.Frame FrameModPulsos 
         BackColor       =   &H00404040&
         Caption         =   "Modulación de Pulsos"
         ForeColor       =   &H0000FF00&
         Height          =   2055
         Left            =   0
         TabIndex        =   48
         Top             =   1920
         Width           =   6975
         Begin VB.CommandButton cmdConfigComm 
            Caption         =   "RS232"
            Height          =   255
            Left            =   5280
            TabIndex        =   77
            Top             =   1680
            Width           =   615
         End
         Begin VB.Timer tmrNotifyIcon 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   6600
            Top             =   0
         End
         Begin VB.CommandButton cmdIncMod 
            BackColor       =   &H00404040&
            Caption         =   "+"
            Height          =   135
            Index           =   2
            Left            =   2400
            TabIndex        =   75
            Top             =   1080
            Width           =   255
         End
         Begin VB.ComboBox cboStep 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   2760
            TabIndex        =   74
            Text            =   "Combo1"
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton cmdDecMod 
            BackColor       =   &H00404040&
            Caption         =   "-"
            Height          =   135
            Index           =   2
            Left            =   2400
            TabIndex        =   73
            Top             =   1200
            Width           =   255
         End
         Begin VB.ComboBox cboStep 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   2760
            TabIndex        =   72
            Text            =   "Combo1"
            Top             =   600
            Width           =   735
         End
         Begin VB.CommandButton cmdDecMod 
            BackColor       =   &H00404040&
            Caption         =   "-"
            Height          =   135
            Index           =   1
            Left            =   2400
            TabIndex        =   71
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton cmdIncMod 
            BackColor       =   &H00404040&
            Caption         =   "+"
            Height          =   135
            Index           =   1
            Left            =   2400
            TabIndex        =   70
            Top             =   600
            Width           =   255
         End
         Begin VB.ComboBox cboStep 
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   2760
            TabIndex        =   69
            Text            =   "Combo1"
            Top             =   240
            Width           =   735
         End
         Begin VB.CommandButton cmdDecMod 
            BackColor       =   &H00404040&
            Caption         =   "-"
            Height          =   135
            Index           =   0
            Left            =   2400
            TabIndex        =   68
            Top             =   360
            Width           =   255
         End
         Begin VB.CommandButton cmdIncMod 
            BackColor       =   &H00404040&
            Caption         =   "+"
            Height          =   135
            Index           =   0
            Left            =   2400
            TabIndex        =   67
            Top             =   240
            Width           =   255
         End
         Begin VB.CheckBox chkOnTop 
            BackColor       =   &H00404040&
            Caption         =   "On Top"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   6000
            TabIndex        =   65
            Top             =   1680
            Width           =   855
         End
         Begin VB.CheckBox chkAplicarCurvaSetup 
            BackColor       =   &H00404040&
            Caption         =   "Aplicar Curva Setup"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   1815
         End
         Begin VB.TextBox txtFileCurvaSetup 
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   120
            TabIndex        =   60
            Text            =   "txtFileCurvaSetup"
            Top             =   1680
            Width           =   2895
         End
         Begin VB.CommandButton cmdSelCompSetup 
            Caption         =   "Command1"
            Height          =   195
            Left            =   1920
            TabIndex        =   59
            Top             =   1440
            Width           =   255
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Control"
            ForeColor       =   &H0000FF00&
            Height          =   1215
            Left            =   3600
            TabIndex        =   56
            Top             =   120
            Width           =   1215
            Begin VB.CheckBox chkFMState 
               BackColor       =   &H00404040&
               Caption         =   "FM Mod ON"
               ForeColor       =   &H0000FF00&
               Height          =   195
               Left            =   120
               TabIndex        =   78
               Top             =   720
               Width           =   1695
            End
            Begin VB.CheckBox chkALCOn 
               BackColor       =   &H00404040&
               Caption         =   "ALC ON"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   120
               TabIndex        =   58
               Top             =   240
               Width           =   975
            End
            Begin VB.CheckBox chkModulacionOn 
               BackColor       =   &H00404040&
               Caption         =   "MOD ON"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Left            =   120
               TabIndex        =   57
               Top             =   480
               Width           =   975
            End
         End
         Begin VB.TextBox txtDelay 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Left            =   960
            TabIndex        =   55
            Text            =   "txtDelay"
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtPW 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Left            =   960
            TabIndex        =   54
            Text            =   "txtPW"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtPRI 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Left            =   960
            TabIndex        =   53
            Text            =   "txtPRI"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BackColor       =   &H00404040&
            Caption         =   "Trigger"
            ForeColor       =   &H0000FF00&
            Height          =   1215
            Left            =   4920
            TabIndex        =   49
            Top             =   120
            Width           =   1335
            Begin VB.OptionButton OptionTrigger 
               BackColor       =   &H00404040&
               Caption         =   "Auto"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   360
               Width           =   735
            End
            Begin VB.OptionButton OptionTrigger 
               BackColor       =   &H00404040&
               Caption         =   "Ext Trigger"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   51
               Top             =   600
               Width           =   1095
            End
            Begin VB.OptionButton OptionTrigger 
               BackColor       =   &H00404040&
               Caption         =   "Ext Gated"
               ForeColor       =   &H0000FF00&
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   50
               Top             =   840
               Width           =   1095
            End
         End
         Begin VB.Label Label9 
            BackColor       =   &H00404040&
            Caption         =   "PW[us]:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00404040&
            Caption         =   "PRI[us]:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   63
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label8 
            BackColor       =   &H00404040&
            Caption         =   "Delay [us]:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   62
            Top             =   960
            Width           =   855
         End
      End
      Begin VB.Frame frameReadOut 
         BackColor       =   &H00404040&
         Caption         =   "Panel Principal"
         ForeColor       =   &H0000FF00&
         Height          =   1935
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   6975
         Begin VB.CommandButton cmdClose 
            Caption         =   "X"
            Height          =   255
            Left            =   6690
            TabIndex        =   76
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton cmdBajarPot 
            BackColor       =   &H00404040&
            Caption         =   "-5 PERDIZ"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   4440
            TabIndex        =   43
            Top             =   1560
            Width           =   735
         End
         Begin VB.CommandButton cmdSubirPot 
            BackColor       =   &H00404040&
            Caption         =   "+5 PERDIZ"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   3600
            TabIndex        =   42
            Top             =   1560
            Width           =   735
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "DEM PERDIZ"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   2640
            TabIndex        =   41
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtStepPriPot 
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   4320
            TabIndex        =   40
            Text            =   "txtStepPriPot"
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox chkIncPRI 
            BackColor       =   &H00404040&
            Caption         =   "Variar PRI"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   2640
            TabIndex        =   39
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdDecFreq 
            BackColor       =   &H00404040&
            Caption         =   "-100"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   1560
            TabIndex        =   38
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "+100"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   1560
            TabIndex        =   37
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdDecFreq 
            BackColor       =   &H00404040&
            Caption         =   "-250"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1080
            TabIndex        =   36
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "+250"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   4
            Left            =   1080
            TabIndex        =   35
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox txtPrueba 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   9.75
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Index           =   3
            Left            =   5160
            TabIndex        =   34
            Text            =   "txtPrueba"
            Top             =   300
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtUnidad 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   9.75
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Index           =   2
            Left            =   6240
            TabIndex        =   33
            Text            =   "txtPrueba"
            Top             =   300
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "+10"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2040
            TabIndex        =   32
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdDecFreq 
            BackColor       =   &H00404040&
            Caption         =   "-10"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   2040
            TabIndex        =   31
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdSetFrec 
            Caption         =   "18GHz"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   1080
            TabIndex        =   30
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton cmdSetFrec 
            Caption         =   "8GHz"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   29
            Top             =   1560
            Width           =   495
         End
         Begin VB.CommandButton cmdSetFrec 
            Caption         =   "2.5G"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox txtUnidad 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Index           =   1
            Left            =   3960
            TabIndex        =   27
            Text            =   "txtPrueba"
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtUnidad 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Index           =   0
            Left            =   1440
            TabIndex        =   26
            Text            =   "txtPrueba"
            Top             =   300
            Width           =   1095
         End
         Begin VB.TextBox txtPrueba 
            Height          =   285
            Index           =   2
            Left            =   5880
            TabIndex        =   25
            Text            =   "txtPrueba"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtPrueba 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,000"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Text            =   "txtPrueba"
            Top             =   300
            Width           =   1335
         End
         Begin VB.TextBox txtPrueba 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   13322
               SubFormatType   =   1
            EndProperty
            BeginProperty Font 
               Name            =   "Roman"
               Size            =   12
               Charset         =   255
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   390
            Index           =   1
            Left            =   2640
            TabIndex        =   23
            Text            =   "txtPrueba"
            Top             =   300
            Width           =   1335
         End
         Begin VB.CommandButton cmdBajarPot 
            BackColor       =   &H00404040&
            Caption         =   "-.1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   22
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdSubirPot 
            BackColor       =   &H00404040&
            Caption         =   "+.1"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   21
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdBajarPot 
            BackColor       =   &H00404040&
            Caption         =   "-5"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   20
            Top             =   1320
            Width           =   375
         End
         Begin VB.CommandButton cmdSubirPot 
            BackColor       =   &H00404040&
            Caption         =   "+5"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   3000
            TabIndex        =   19
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdSubirPot 
            BackColor       =   &H00404040&
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2640
            TabIndex        =   18
            Top             =   1080
            Width           =   375
         End
         Begin VB.CommandButton cmdBajarPot 
            BackColor       =   &H00404040&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2640
            MaskColor       =   &H00404040&
            TabIndex        =   17
            Top             =   1320
            Width           =   375
         End
         Begin VB.TextBox txtStepFreq 
            BackColor       =   &H00404040&
            ForeColor       =   &H0000FF00&
            Height          =   285
            Left            =   600
            TabIndex        =   16
            Text            =   "txtStepFreq"
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton cmdDecFreq 
            BackColor       =   &H00404040&
            Caption         =   "-"
            Height          =   195
            Index           =   0
            Left            =   1920
            TabIndex        =   15
            Top             =   840
            Width           =   255
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "+"
            Height          =   135
            Index           =   0
            Left            =   1920
            TabIndex        =   14
            Top             =   720
            Width           =   255
         End
         Begin VB.CommandButton cmdDecFreq 
            BackColor       =   &H00404040&
            Caption         =   "-1G"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "+1G"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton cmdDecFreq 
            BackColor       =   &H00404040&
            Caption         =   "-500"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   11
            Top             =   1320
            Width           =   495
         End
         Begin VB.CommandButton cmdIncFreq 
            BackColor       =   &H00404040&
            Caption         =   "+500"
            BeginProperty Font 
               Name            =   "MS Serif"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   600
            TabIndex        =   10
            Top             =   1080
            Width           =   495
         End
         Begin VB.CheckBox chkRFOn 
            BackColor       =   &H00404040&
            Caption         =   "RF ON"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   5280
            TabIndex        =   9
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label11 
            BackColor       =   &H00404040&
            Caption         =   "Step:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   3840
            TabIndex        =   47
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00404040&
            Caption         =   "Valor Corregido"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   5280
            TabIndex        =   46
            Top             =   720
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label7 
            BackColor       =   &H00404040&
            Caption         =   "[KHz]"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   1440
            TabIndex        =   45
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label6 
            BackColor       =   &H00404040&
            Caption         =   "Step:"
            ForeColor       =   &H0000FF00&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   495
         End
      End
      Begin VB.TextBox txtAgregarPot 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "txtAgregarPot"
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox chkIncPot 
         Caption         =   "Agregar"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtCronometro 
         Height          =   285
         Left            =   6360
         TabIndex        =   5
         Text            =   "txtCronometro"
         Top             =   600
         Width           =   1095
      End
      Begin MSComctlLib.ListView LstVwRangoControl 
         Height          =   1095
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   1931
         View            =   3
         MultiSelect     =   -1  'True
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
      Begin VB.Label lblPrueba 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado :"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   3
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblPrueba 
         Alignment       =   1  'Right Justify
         Caption         =   "Potencia Actual :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblPrueba 
         Alignment       =   1  'Right Justify
         Caption         =   "Frecuencia Actual :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.Timer tmrPrueba 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmCtlGenMin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PV_Ini_Tmr          As Long
Dim PV_End_Tmr          As Long

Dim PV_TckCnt_E1        As Long

Dim PV_Ptos_Prueba      As Long
Dim PV_Ptos_Now         As Long
Dim PV_Tpo_Ini          As Single
Dim PV_Estado           As Integer
Dim PV_Frec_Min         As Long
Dim PV_Frec_Max         As Long
Dim PV_Frec_Paso        As Long
Dim PV_Pot_Min          As Double
Dim PV_Pot_Max          As Double
Dim PV_Pot_Paso         As Double
Dim PV_Factor_Pot       As Double
Dim PV_Etapa            As Integer
Dim PV_EmpezarEtapa     As Boolean
Dim PV_Frec_Now         As Long
Dim PV_Pot_Now          As Double
Dim PV_Index_Tabla      As Integer

Dim PV_Flag_RF_On       As Boolean

Dim PV_Estabiliza_Counter       As Integer

Dim PV_Handle_Instruments()    As Long

Dim PV_Address_List()       As Long
Dim PC_Config_Commands()    As String


Dim PV_Flag_Freq            As Boolean
Dim PV_Flag_Pow             As Boolean

Dim PV_File_Hdl             As Integer

Dim PV_Column_Header()           As String

Dim PV_Compensa()           As Type_Correccion
Dim PV_Correccion()         As Type_Correccion
Dim PV_TablaParam()         As Type_ParamControl

Dim PV_Index_Correc         As Integer
Dim PV_CommPort             As Boolean

Private Declare Sub ReleaseCapture Lib "user32" ()

Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
'
'
'
'
'
' Nueva funcion para iniciarlizar los temas de XP
'Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagINITCOMMONCONTROLSEX) As Boolean
' Creamos una instancia de la estructura
'Public Ini As tagINITCOMMONCONTROLSEX
' Estructura del notify icon (Version 5 o posterior)
Private Type NOTIFYICONDATA ' declaracion del tipo de datos para notificar el Notify
        cbSize As Long
        hWnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 128   'Desde aqui el nuevo notify
        dwState As Long
        dwStateMask As Long
        szInfo As String * 256
        uTimeout As Long        ' Este es compartido con (uVersion as Long)
        szInfoTitle As String * 64
        dwInfoFlags As Long
        ' guidItem As GUID (solo para la version 6)
End Type
' Para la version de el Notify icon (por defecto en XP ya esta inicializado)
Private Const NOTIFYICON_VERSION = 3

'constantes relacionas con el raton
Private Const WM_RBUTTONUP = &H205
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_MOUSEMOVE = &H200
Private Const WM_USER = &H400
' Constantes relacionadas con el Ballon tool tip
Private Const NIN_BALLOONSHOW = (WM_USER + 2)
Private Const NIN_BALLOONHIDE = (WM_USER + 3)
Private Const NIN_BALLOONTIMEOUT = (WM_USER + 4)
' Esta es la que me gusta !!!!
Private Const NIN_BALLOONUSERCLICK = (WM_USER + 5)
'constantes de lo que queremos que muestre el Notify
Private Const NIF_ICON = &H2 ' queremos que muestre un Notify
Private Const NIF_MESSAGE = &H1 ' queremos que nos envie un mensaje
Private Const NIF_TIP = &H4 ' queremos que muestre un texto al posicionarnos encima
' Para la version 5
Private Const NIF_STATE = &H8 ' Devuelve el estado
Private Const NIF_INFO = &H10 ' Muestra un ballon en el notify icon

' Aqui las constantes para los Notifys de los ballons tips
' No muestra nada
Private Const NIIF_NONE = &H0
' Muestra un Notify de Informacion
Private Const NIIF_INFO = &H1
' Muestra un Notify de Precaucion
Private Const NIIF_WARNING = &H2
' Muestra un Notify de Error
Private Const NIIF_ERROR = &H3

'constantes para añadir, borrar o modificar el Notify
Private Const NIM_ADD = &H0             ' añadirlo a la barra de tareas
Private Const NIM_DELETE = &H2  ' borrarlo de la barra de tareas
Private Const NIM_MODIFY = &H1  ' modificarlo
' Para la version 5
Private Const NIM_SETFOCUS = &H3                ' Da el foco a la barra de tareas
Private Const NIM_SETVERSION = &H4      ' Asigna la version del Notify icon

' declaracion de la funcion
Private Declare Function Shell_NotifyIcon Lib "shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Boolean

' Creamos una instancia del notify
Private Notify As NOTIFYICONDATA
'
'
'
'

Sub Activar_Display_Valor_Corregido(LV_Flag As Boolean)

    With Me
        If LV_Flag = True Then
            .txtPrueba(3).Visible = True
            .txtUnidad(2).Visible = True
            .txtPrueba(3).Text = ""
            '.txtUnidad(2).Text = ""
        Else
            .txtPrueba(3).Visible = False
            .txtUnidad(2).Visible = False
        End If
    End With
    
End Sub

Sub Enviar_Pot_From_TextBox()

Dim LV_Pow_Cor      As Double
Dim LV_F            As Double

    If PV_Flag_Freq = True Then
        PV_Flag_Freq = False
        If IsNumeric(Me.txtPrueba(1).Text) = False Then
            Exit Sub
        End If
        LV_Pow_Cor = Me.txtPrueba(1).Text
        If Me.chkAplicarCurvaSetup.value Then
            PV_Frec_Now = Me.txtPrueba(0).Text
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, LV_Pow_Cor)
            If LV_Pow_Cor > 15 Then
                Exit Sub
            End If
            Me.txtPrueba(3).Text = Int(10 * LV_Pow_Cor) / 10
        End If
        
        If Me.chkRS232.value = 0 Then
            SendPow 0, LV_Pow_Cor   'PV_Pot_Now - PV_Correccion(PV_Index_Correc)
        Else
            fMainForm.SendRS232 CommandPow(0, LV_Pow_Cor)
        End If
    End If
    
End Sub

Sub Enviar_Frec_From_TextBox()

Dim LV_F_I          As Double
Dim LV_F            As Double

    If PV_Flag_Freq = True Then
        PV_Flag_Freq = False
        LV_F = Me.txtPrueba(0).Text
'        If Me.chkRS232.value = 0 Then
'            SendFrec 0, LV_F
'        Else
            fMainForm.SendRS232 CommandFrec(0, LV_F)
'        End If
        
'        If Me.chkRS232.value = 0 Then
'            Verify_Frec_State 0, LV_F_I
'        Else
'        End If
        
        If LV_F_I <> LV_F Then
            Me.txtPrueba(2).Text = "Falla Envío/Respuesta Frec"
        End If
    End If
    

End Sub

Function Calc_Factor_Pot(ByVal LV_Paso As Double) As Double

Dim i           As Integer
Dim LV_Factor   As Double
Dim lv

    LV_Factor = 1
    
    i = 0
    Do
        If LV_Paso = Int(LV_Paso * LV_Factor) / LV_Factor Then
            
            Calc_Factor_Pot = LV_Factor
        
            Exit Do
            
        End If
        
        i = i + 1
        
        LV_Factor = LV_Factor * 10
        
    Loop Until i = 2
    
    Calc_Factor_Pot = LV_Factor
    
End Function

Function Calcular_Compensa(LV_Frec As Long, LV_Pot As Double) As Double

Dim i       As Integer

    Calcular_Compensa = LV_Pot

    For i = 0 To UBound(PV_Compensa)
        With PV_Compensa(i)
            If .Freq = LV_Frec Then
                Calcular_Compensa = LV_Pot - .Correccion
                Exit Function
            End If
        End With
    Next

End Function

Function Calcular_Correccion(LV_Frec As Long, LV_Pot As Double) As Double

Dim i       As Integer
Dim j       As Integer

    Calcular_Correccion = LV_Pot

    For i = 0 To UBound(PV_Correccion)
        With PV_Correccion(i)
            If .Freq = LV_Frec Then
                Calcular_Correccion = LV_Pot - .Correccion
                Exit Function
            ElseIf .Freq <= LV_Frec Then
                j = i
            End If
        End With
    Next
    
    With PV_Correccion(j)
        Calcular_Correccion = LV_Pot - .Correccion
    End With
    
End Function

Function Calcular_Correccion_Inv(LV_Frec As Long, LV_Pot As Double) As Double

Dim i       As Integer

    Calcular_Correccion_Inv = LV_Pot

    For i = 0 To UBound(PV_Correccion)
        With PV_Correccion(i)
            If .Freq = LV_Frec Then
                Calcular_Correccion_Inv = LV_Pot + .Correccion
                Exit Function
            End If
        End With
    Next
    
End Function

'Sub Cargar_Curva_Video_Pot(Optional LV_File As String)
'
'Dim h               As Integer
'Dim i               As Integer
'Dim LV_Line         As String
'Dim LV_Campos()     As String
''Dim LV_File         As String
'Dim j               As Integer
'
'    ReDim GV_Lista_Frec(99)
'    ReDim GV_Tabla_Vid_Pot(99)
'    i = 1
'    h = FreeFile
'
'    If LV_File = "" Then
'        LV_File = Me.txtCurvaVideoPot.Text
'    Else
'        Me.txtCurvaVideoPot.Text = LV_File
'    End If
'
'    If VerificarExiste(LV_File) = True Then
'
'        Open LV_File For Input As h
'
'        For i = 1 To 2
'            If EOF(h) = True Then
'                Exit Sub
'            End If
'            Line Input #h, LV_Line
'        Next
'
'        If InStr(LV_Line, ";") Then
'            ' Leer Potencias
'            LV_Campos = Split(LV_Line, ";")
'            ReDim GV_Lista_Pot(UBound(LV_Campos) - 1)
'            For i = 1 To UBound(LV_Campos)
'                GV_Lista_Pot(i - 1) = LV_Campos(i)
'            Next
'
'            ' Leer Frecuencias
'            i = 0
'            Do
'                If EOF(h) = False Then
'                    Line Input #h, LV_Line
'                    If LV_Line <> "" Then
'                        LV_Campos = Split(LV_Line, ";")
'                        If UBound(GV_Lista_Frec) < i Then
'                            ReDim Preserve GV_Lista_Frec(99 + i)
'                            ReDim Preserve GV_Tabla_Vid_Pot(99 + i)
'                        End If
'                        GV_Lista_Frec(i) = LV_Campos(0)
'                        ReDim GV_Tabla_Vid_Pot(i).Filas(UBound(GV_Lista_Pot))
'                        For j = 0 To UBound(GV_Lista_Pot)
'                            GV_Tabla_Vid_Pot(i).Filas(j) = LV_Campos(j + 1)
'                        Next
'                        i = i + 1
'                    Else
'                        Exit Do
'                    End If
'                Else
'                    Exit Do
'                End If
'            Loop
'        End If
'
'        Close h
'
'    End If
'
'    ReDim Preserve GV_Tabla_Vid_Pot(i - 1)
'
'    ReDim Preserve GV_Lista_Frec(i - 1)
'
'End Sub
'
'Sub Cargar_Compensa()
'
'Dim h               As Integer
'Dim i               As Integer
'Dim LV_Line         As String
'Dim LV_Campos()     As String
'Dim LV_File         As String
'
'    h = FreeFile
'    'PV_Index_Correc = 0
'
'    ReDim PV_Compensa(0)
'
'    LV_File = Me.txtArchCompSal
'
'    If VerificarExiste(LV_File) = True Then
'
'        Open LV_File For Input As h
'
'        For i = 1 To 5
'            If EOF(h) = True Then
'                Exit Sub
'            End If
'            Line Input #h, LV_Line
'        Next
'
'        i = 0
'        Do
'            If EOF(h) = False Then
'                Line Input #h, LV_Line
'                If LV_Line <> "" Then
'                    LV_Campos = Split(LV_Line, ";")
'                    ReDim Preserve PV_Compensa(i)
'                    PV_Compensa(i).Correccion = LV_Campos(3)
'                    PV_Compensa(i).Freq = LV_Campos(0)
'                    i = i + 1
'                Else
'                    Exit Do
'                End If
'            Else
'                Exit Do
'            End If
'        Loop
'
'        Close h
'
'    End If
'
'    Exit Sub
'
'End Sub
'
'Sub CargarTablaParam()
'
'Dim h               As Integer
'Dim i               As Integer
'Dim LV_Line         As String
'Dim LV_Campos()     As String
'Dim LV_File         As String
'Dim j               As Integer
'
'    ReDim PV_TablaParam(99)
'    i = 1
'    h = FreeFile
'
'    LV_File = Me.txtTablaParam.Text
'
'    If VerificarExiste(LV_File) = True Then
'
'        Open LV_File For Input As h
'
'        For i = 1 To 5
'            If EOF(h) = True Then
'                Exit Sub
'            End If
'            Line Input #h, LV_Line
'        Next
'
'        ' Leer Frecuencias
'        i = 0
'        Do
'            If EOF(h) = False Then
'                Line Input #h, LV_Line
'                If LV_Line <> "" Then
'                    LV_Campos = Split(LV_Line, ";")
'                    If UBound(PV_TablaParam) < i Then
'                        ReDim Preserve PV_TablaParam(99 + i)
'                    End If
'                    PV_TablaParam(i).Frec = LV_Campos(0)
'                    PV_TablaParam(i).Pot_Gen = LV_Campos(1)
'                    i = i + 1
'                Else
'                    Exit Do
'                End If
'            Else
'                Exit Do
'            End If
'        Loop
'
'        Close h
'
'    End If
'
'    ReDim Preserve PV_TablaParam(i - 1)
'    PV_Index_Tabla = 0
'
'End Sub
'
Sub Cargar_Correccion()

Dim h               As Integer
Dim i               As Integer
Dim LV_Line         As String
Dim LV_Campos()     As String

    h = FreeFile
    PV_Index_Correc = 0
    
    ReDim PV_Correccion(0)
    
    'Set PV_Correccion = vbNull

    On Error GoTo NoHayArchivo
    
    Open Me.txtFileCurvaSetup.Text For Input As h
    
    On Error GoTo 0
    
    For i = 1 To 5
        If EOF(h) = True Then
            Exit Sub
        End If
        Line Input #h, LV_Line
    Next
    
    i = 0
    Do
        If EOF(h) = False Then
            Line Input #h, LV_Line
            If LV_Line <> "" Then
                LV_Campos = Split(LV_Line, ";")
                ReDim Preserve PV_Correccion(i)
                PV_Correccion(i).Correccion = Val(LV_Campos(3))
                PV_Correccion(i).Freq = LV_Campos(0)
                i = i + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    Close h
    
    Exit Sub
    
NoHayArchivo:

    MsgBox "No se ha encontrado archivo de Corrección.", vbInformation
    
    On Error GoTo 0
    

End Sub

Sub Cargar_Valores_Etapa(LV_Etapa As Integer)

    PV_EmpezarEtapa = True
    
    With GV_Actual_Project.Rango(LV_Etapa)
    
        PV_Frec_Min = .ValorMin
        PV_Frec_Max = .ValorMax
        PV_Frec_Paso = .Paso
    
    End With
    
    With GV_Actual_Project.Rango(LV_Etapa + 1)
        
        PV_Pot_Min = .ValorMin
        PV_Pot_Max = .ValorMax
        PV_Pot_Paso = .Paso
        
        PV_Factor_Pot = Calc_Factor_Pot(PV_Pot_Paso)
        
    End With
    
End Sub

Sub Cargar_Valores()

    PV_Etapa = 0
    
'    If GV_Actual_Project.EtapasDeControl = 0 Then
'        Me.cmdComenzar.Enabled = False
'        Exit Sub
'    Else
'        Me.cmdComenzar.Enabled = True
'    End If
    
    Me.Cargar_Valores_Etapa PV_Etapa
    
    With Me
        
        PV_Frec_Now = PV_Frec_Min
        PV_Pot_Now = PV_Pot_Min
        
'        If Me.chkUsarTablaParam.value Then
'            CargarValoresFromTablaParam
'        End If
        
        PV_Flag_Freq = True
        PV_Flag_Pow = True
        
        .txtPrueba(0).Text = PV_Frec_Min
        .txtPrueba(1).Text = PV_Pot_Min
        .txtPrueba(2).Text = ""
        
    End With
    
End Sub

Sub CargarValoresFromTablaParam()

    If PV_Index_Tabla <= UBound(PV_TablaParam) Then
        PV_Frec_Now = PV_TablaParam(PV_Index_Tabla).Frec
        PV_Pot_Now = PV_TablaParam(PV_Index_Tabla).Pot_Gen
        PV_Index_Tabla = PV_Index_Tabla + 1
    End If
    
End Sub

Sub Repetir_Prueba()

    With Me
        .tmrPrueba.Enabled = False
'        .cmdComenzar.Caption = "&Reanudar"
'        .cmdComenzar.Refresh
        PV_Index_Tabla = 0
        PV_EmpezarEtapa = True
        PV_Etapa = 0
        PV_Estado = 0
        PV_Ptos_Now = 0
        PV_Tpo_Ini = Timer
        .Cargar_Valores_Etapa PV_Etapa
        Cargar_Valores
'        If .chkCortarRFalTerminar.value = 1 Then
'            .chkRFOn.value = 0
'            SendRFPowerOff 0
'        End If
    End With
    
End Sub

Sub Cancelar_Prueba()

    With Me
'        If .chkCortarRFalTerminar.value = 1 Then
'            .chkRFOn.value = 0
'            SendRFPowerOff 0
'        End If

        .chkRS232.Enabled = True
        .tmrPrueba.Enabled = False
'        .cmdCancelar.Enabled = False
'        .cmdComenzar.Caption = "&Comenzar"
        .Close_Devices
        .Cerrar_Archivo_Salida
        .txtPrueba(2).Text = ""
    End With

End Sub

Function Capturar_Valores() As Boolean

Dim LV_Data         As String
Dim LV_Value        As Double
Dim i               As Integer

    
    
    Capturar_Valores = True

For i = 1 To UBound(GV_Instrumentos)
    LV_Data = Get_Data_From_Instr(i)
    
    If LV_Data <> "" Then
        If IsNumeric(LV_Data) = False Then
            LV_Value = ExtraerNumeric(LV_Data)
        End If
        LV_Value = Val(LV_Data)
        ' Obtiene Compensación de Salidas
        LV_Value = Calcular_Compensa(PV_Frec_Now, LV_Value)
        GV_Data_Captur.Data(i + 1) = LV_Value
'        If Me.chkCurvaVideoPot.value Then
'            LV_Value = Convertir_Video_en_Pot(LV_Value, PV_Frec_Now)
'        End If
        If UBound(GV_Data_Captur.Data) > i + 1 Then
            GV_Data_Captur.Data(2 + i) = LV_Value
        End If
    End If
    Capturar_Valores = True
Next

End Function

Sub Cerrar_Archivo_Salida()

    If PV_File_Hdl <> 0 Then
    
        Print #PV_File_Hdl, vbCrLf
        Print #PV_File_Hdl, "Hora Término:" & GV_Ch_Decimal & Format(Time(), "hh:mm:ss") & vbCrLf
    
        Close #PV_File_Hdl
        
        PV_File_Hdl = 0
        
    End If
    
End Sub

Sub Close_Devices()

Dim i           As Integer

    If Me.chkRS232.value = 0 Then
'        If GPIBglobalsRegistered = 1 Then
'             For i = 0 To UBound(GV_Instrumentos)
'                With GV_Instrumentos(i)
'                    If .Cmd_End <> "" Then
'                        SendCmd .Cmd_Config, i
'                    End If
'                End With
'            Next
'
'           Call DevClearList(GPIB0, GV_Result_List)
'
'            ilonl GPIB0, 0
'
'            GPIBglobalsRegistered = 0
'
'        End If
    Else
        fMainForm.CloseMsComm
    End If

End Sub

Sub Crear_Estructuras()

Dim i           As Integer

    ReDim GV_Data_Captur.Data(3)
    
    ReDim GV_Data_Captur.NombreCampo(3)
    
    With GV_Data_Captur
    
        For i = 0 To 3
        
                    
            .NombreCampo(i) = ""
            
        Next
        
        .NombreCampo(0) = "Frecuencia Gen"
        .NombreCampo(1) = "Potencia Gen"
        .NombreCampo(2) = "Voltaje"
        .NombreCampo(3) = "Potencia"
        
    End With
    
End Sub

Function CommandFrec(ByVal Index As Integer, _
                    ByVal LV_Freq As Double, _
                    Optional LV_Str As String) As String

Dim Str_Cmd         As String
Dim LV_Str_Val      As String
Dim LV_Cmds()       As String
Dim i               As Integer

    CommandFrec = ""
    
    Str_Cmd = GV_Instrumentos(Index).Cmd_Set_Var(0) '& " " & LV_Freq
    
    If Str_Cmd = "" Then
        Exit Function
    End If
    
    LV_Str_Val = LV_Freq
    Str_Cmd = Replace_Value(Str_Cmd, LV_Str_Val, "%F")
    
    If LV_Str <> "" Then
        Str_Cmd = Str_Cmd & LV_Str
    End If
    
    LV_Cmds = Split(Str_Cmd, ";")
    For i = 0 To UBound(LV_Cmds)
        ParseCommand LV_Cmds(i)
        
        CommandFrec = LV_Cmds(i) & vbCrLf
        CommandFrec = Replace(CommandFrec, ",", ".")
    Next
    
End Function

Sub Enviar_Frecuencia()

Dim LV_F        As Double
Dim LV_F_I      As Double
Dim i           As Integer

    If PV_Flag_Freq = True Then
            
        For i = 0 To UBound(GV_Instrumentos)
            With GV_Instrumentos(i)
                LV_F = PV_Frec_Now * 1000000#
                
                'SendFrec Index, LV_F
                'If i Then
                If Me.chkRS232.value = 0 Then
                    SendFrec i, PV_Frec_Now
                Else
                    fMainForm.SendRS232 CommandFrec(i, PV_Frec_Now)
                End If
                GV_Data_Captur.Data(0) = PV_Frec_Now
                
                If Me.chkRS232.value = 0 Then
                    'Verify_Frec_State 0, LV_F_I
                Else
                End If
                
                If LV_F_I <> LV_F Then
                    
                    'Me.cmdCancelar.value = True
                
                End If
                
                PV_Flag_Freq = False
                    
                'SendFrec 2, PV_Frec_Now, "MZ"
            End With
        Next
    End If
    
End Sub

Function CommandPow(ByVal Index As Integer, _
                    ByVal LV_Pow As Double) As String
                    
Dim Str_Cmd         As String
Dim LV_Str_Val      As String

    LV_Str_Val = LV_Pow
    LV_Str_Val = Replace(LV_Str_Val, ",", ".")
    
    Str_Cmd = Replace_Value(GV_Instrumentos(Index).Cmd_Set_Var(1), LV_Str_Val, "%P")
    
    CommandPow = Str_Cmd & vbCrLf
    
End Function

Sub Enviar_Potencia()

Dim LV_Pow          As Double
Dim LV_Pow_Cor      As Double

    If PV_Flag_Pow = True Then
        
        If Me.chkIncPot.value And IsNumeric(Me.txtAgregarPot.Text) = True Then
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, PV_Pot_Now + Me.txtAgregarPot.Text)
        Else
            LV_Pow_Cor = Calcular_Correccion(PV_Frec_Now, PV_Pot_Now)
        End If
                
        If Me.chkRS232.value = 0 Then
            SendPow 0, LV_Pow_Cor   'PV_Pot_Now - PV_Correccion(PV_Index_Correc)
        Else
            fMainForm.SendRS232 CommandPow(0, LV_Pow_Cor)
        End If
        
        GV_Data_Captur.Data(1) = PV_Pot_Now
        
        If Me.chkRS232.value = 0 Then
            'Verify_Pow_State 0, LV_Pow
        End If
        
        If LV_Pow <> PV_Pot_Now Then
            
            'Me.cmdCancelar.value = True
        
        End If
        
        PV_Flag_Pow = False
        
    End If
    
End Sub

Sub Enviar_Frecuencia_Potencia()

    Enviar_Frecuencia
    
    Enviar_Potencia
    
    If PV_Flag_RF_On = True Then
        PV_Flag_RF_On = False
        SendRFPowerOn 0
        Me.chkRFOn.value = 1
    End If
    
End Sub
    
Function GetFileCorreccion() As String

    'GetFileCorreccion = GV_Actual_Project.Path_Project & "\" & "Correccion Setup.csv"
    GetFileCorreccion = GV_Actual_Project.CompensacionSetup
    
End Function

Sub Guardar_Valores()

Dim LV_Data()         As String
Dim i               As Integer

    If PV_File_Hdl = 0 Then
    
        Crear_Archivo_Salida PV_File_Hdl, GV_Archivo_Salida
        
        'Iniciar_List_View
        
    End If
        
    'GV_Data_Captur.Data(3) = GV_Data_Captur.Data(2)
    
    Guardar_Data PV_File_Hdl, GV_Data_Captur.Data
        
    ReDim LV_Data(UBound(GV_Data_Captur.Data))
    
    For i = 0 To UBound(LV_Data)
        LV_Data(i) = GV_Data_Captur.Data(i)
    Next i
    
    'AddItemListView Me.LstVwVisualTest, LV_Data
    
End Sub

Sub Inicializar_Comandos_Instrumentos()

Dim LV_Cod_Instru()     As Integer
Dim LV_Qty              As Integer
Dim i                   As Integer
Dim LV_Volt_Div         As String
Dim LV_50Ohms           As String
Dim LV_Inverter         As String

    'i = 1
    'ReDim GV_Instrumentos(i)
    
    i = 0
    ReDim GV_Instrumentos(i)
    
    LV_Qty = 0      'BD_Get_Cod_Instruments(LV_Cod_Instru, GV_Actual_Project.Cod_Project, 1)
    
    If LV_Qty Then
        ReDim GV_Instrumentos(LV_Qty)
    End If
    
    'BD_Fill_Comandos_Instru GV_Instrumentos(i), GV_Actual_Project.Cod_Project
    
    With GV_Instrumentos(i)
    
        .Comunicacion = Type_Communica.COMM_GPIB
        .address = 19
'        .address = Me.txtAddressGen.Text
        '.Address = 28
        
        '.Cmd_Config = "UNIT:POW:DBM"
        '.Cmd_Config = "OUTPUT ON"
        .Cmd_Config = ""
        
        ReDim .Cmd_Consult(1)
        ReDim .Cmd_Set_Var(1)
        
        '.Cmd_Set_Var(0) = "FREQ:CW"
        '.Cmd_Set_Var(0) = "FREQ %F MHz"
        .Cmd_Set_Var(0) = "FREQ %FMHz"
        .Cmd_Consult(0) = "FREQ:CW?"
        
        '.Cmd_Set_Var(1) = "POW:LEV"
        '.Cmd_Set_Var(1) = "POW %P dBm"
        .Cmd_Set_Var(1) = "POW %PdBm"
        .Cmd_Consult(1) = "POW:LEV?"
        
        .Cmd_End = ":OUTput1:PON OFF"
    End With
    
'    If Me.chkMedirOsciloscopio.value Then
'        i = i + 1
'        ReDim Preserve GV_Instrumentos(i)
'        '----------------------
'        ' Osciloscopio
'        With GV_Instrumentos(i)
'
'            .Name = "Osciloscopio"
'            .Comunicacion = COMM_GPIB
'            .address = 2
'
'            LV_Volt_Div = Trim$(Me.cboOscVoltDiv.Text / 1000#)
'            LV_50Ohms = "OFF"
'            LV_Inverter = "OFF"
'            If Me.chk50Ohm.value Then
'                LV_50Ohms = "ON"
'            End If
'            If Me.chkInvertir.value Then
'                LV_Inverter = "ON"
'            End If
'
'            .Cmd_Config = "DAT:SOU CH1" _
'                        & ";:DAT:ENC ASCII" _
'                        & ";:DAT:WID 1" _
'                        & ";:DAT:STAR 1" _
'                        & ";:DAT:STOP 1" _
'                        & ";CH1 VOLts:" & LV_Volt_Div _
'                        & ";CH1 POSition:-4" _
'                        & ";CH1 FIFty:" & LV_50Ohms _
'                        & ";CH1 INVERT:" & LV_Inverter _
'                        & ";:HOR:MAIN:SCALE 5e-4"
'
'            '.Cmd_Config = .Cmd_Config _
'                        & ";CH2 VOLts:" & LV_Volt_Div _
'                        & ";CH2 POSition:-4" _
'                        & ";CH2 FIFty:" & LV_50Ohms _
'                        & ";CH2 INVERT:" & LV_Inverter _
'                        & ";:HOR:MAIN:SCALE 5e-4"
'
'
'            'GV_Volt_Div = 500 / 25
'            GV_Volt_Div = Me.cboOscVoltDiv.Text / 25
'            GV_Offset = -4
'
'            ReDim .Cmd_Consult(0)
'            ReDim .Cmd_Set_Var(0)
'            '.Cmd_Set_Var(0) = "DAT:SOU CH2"
'            'ReDim .Cmd_Set_Var(1)
'
'            If Me.optOscNivel(0).value Then
'                .Cmd_Consult(0) = "AVG?"
'            Else
'                .Cmd_Consult(0) = "CURVE?"
'            End If
'
'        End With
'    End If
    
    'Exit Sub
    'i = 2
'    If Me.chkMedirAnalizador.value Then
'        i = i + 1
'        ReDim Preserve GV_Instrumentos(i)
'        'i = 1
'        '----------------------
'        ' Analizador de Espectro
'        With GV_Instrumentos(i)
'
'            .Comunicacion = COMM_GPIB
'            .address = 8
'            .Name = "Analizador de Espectro"
'            .Cmd_Config = "RL1" _
'                        & ";RE " & Trim$(Me.txtRefLvl) & "DB" _
'                        & ";AT " & Trim$(Me.txtAtt) & "DB" _
'                        & ";SP " & Trim$(Me.txtSpan) & "MZ" _
'                        & ";PS"
'
'            '.Cmd_Config = "RL1" _
'                        & ";RE 0DBM" _
'                        & ";AT 0DB" _
'                        & ";SP 20MZ" _
'                        & ";PS"
'
'
'            '.Cmd_Config = "CF 1900MZ" & vbCrLf
'            '.Cmd_Config = "CH1 FIFty:ON"
'
'            ReDim .Cmd_Consult(0)
'            ReDim .Cmd_Set_Var(0)
'
'            '.Cmd_Set_Var(0) = "SP 20MZ ; CF %FMZ ; PS "
'            .Cmd_Set_Var(0) = "CF %FMZ ; PS "
'            If Me.cboPeakGraph.ListIndex = 0 Then
'                .Cmd_Consult(0) = "PS; MFL?"
'            Else
'                .Cmd_Consult(0) = "PS; PLOT"
'            End If
'            .TpoEspera = 0
'
'            ReDim GV_Analizador_Sp(0)
'
'            With GV_Analizador_Sp(0)
'                .RefLvl = Me.txtRefLvl
'                .CenterFreq = Me.txtCenterFreq
'                .SPAN = Me.txtSpan
'
'            End With
'
'        End With
'        '----------------------
'        'Exit Sub
'        'i = i
'    End If
'
'    'If i <= UBound(GV_Instrumentos) Then
'    If Me.chkMedirPowerMeter.value Then
'        i = i + 1
'        ReDim Preserve GV_Instrumentos(i)
'        ' Power Meter
'        With GV_Instrumentos(i)
'
'            .Comunicacion = COMM_GPIB
'            .address = 13
'
'            .Cmd_Config = "LG;OC1"
'    '
'            ReDim .Cmd_Consult(0)
'    '        ReDim .Cmd_Set_Var(1)
'    '
'    '        .Cmd_Set_Var(0) = "FREQ:CW"
'            '.Cmd_Consult(0) = "MEAS1?"
'            .Cmd_Consult(0) = "LG"
'    '
'    '        .Cmd_Set_Var(1) = "POW:LEV"
'    '        .Cmd_Consult(1) = "POW:LEV?"
'            ReDim .Cmd_Set_Var(0)
'
'            .Cmd_Set_Var(0) = "FREQ %F MHz"
'        .TpoEspera = 0
'
'        End With
'    End If

End Sub

Function Iniciar_Instrumentos_RS232() As Boolean

Dim LV_i            As Integer

    For LV_i = 0 To UBound(GV_Instrumentos)
        
        With GV_Instrumentos(LV_i)
            'If .Comunicacion = COMM_RS232 Then
                'IniciarCommPort
                PV_CommPort = fMainForm.IniciarCommPort
                If .Cmd_Config <> "" Then
                    'SendRS232 .Cmd_Config
                    fMainForm.SendRS232 .Cmd_Config
                End If
            'End If
        End With
        
    Next

End Function

Function Inicializar_Comm_GPIB() As Boolean

Dim i           As Integer

    'ReDim PV_Handle_Instruments(2)
    
    Inicializar_Comandos_Instrumentos
    
    If Me.chkRS232.value = 0 Then
        IniciarCommInstrumento
    Else
        Iniciar_Instrumentos_RS232
    End If
    'For i = 0 To 2
    
        'PV_Handle_Instruments(i) =
        
     '   IniciarCommInstrumento (i)
        
'        If PV_Handle_Instruments(i) = 0 Then
'
'            Inicializar_Comm_GPIB = False
'
'            'Exit Function
'
'        End If
        
    'Next
    
    Inicializar_Comm_GPIB = True

End Function

Function IniciarCommPort() As Boolean

    With Me
        If .MSComm.PortOpen = False Then
            .MSComm.CommPort = 2
            .MSComm.PortOpen = True
            IniciarCommPort = True
        Else
            IniciarCommPort = True
        End If
    End With

End Function

Sub Running_in_Background()

    With Me
        .Hide
        SetNotifyIcon
        'Load frmCtlGenMin
        'frmCtlGenMin.Show vbModal
    End With
    
End Sub

Sub Delete_Notify_Icon()

    Shell_NotifyIcon NIM_DELETE, Notify

End Sub

Function SendRS232(LV_Cmd As String)

    With Me.MSComm
        If .PortOpen = True Then
            .Output = LV_Cmd
        End If
    End With
    
End Function

Sub SetNotifyIcon()        '(LV_Title As String, LV_Msg As String)

    Notify.cbSize = Len(Notify)
    'Notify.dwInfoFlags = NIIF_INFO      ' NIIF_WARNING
    Notify.szTip = "COM Desconectada." & Chr$(0)       ' tool tip clasico
    Notify.uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_INFO Or NIF_TIP  'NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    Notify.szInfoTitle = App.Title + Chr$(0)
    Notify.szInfo = "COM desconectada." + vbCr '+ "Doble Click para" + Chr$(0)
    Notify.hIcon = Me.Icon
    Notify.hWnd = Me.hWnd
    Notify.uCallbackMessage = WM_MOUSEMOVE
    Notify.uID = 1& ' un identificador del Notify
    ' llamamos a NIM_MODIFY para mostrar de nuevo el ballon
    Shell_NotifyIcon NIM_ADD, Notify

End Sub

Function LeerRS2323() As String

    With Me.MSComm
        If .PortOpen = True Then
            LeerRS2323 = .Input
        End If
    End With
    
End Function

'Sub Iniciar_List_View()
'
'    With Me
'        With .LstVwVisualTest
'            .ColumnHeaders.Clear
'            .ListItems.Clear
'        End With
'
'        AddColumListView .LstVwVisualTest, GV_Data_Captur.NombreCampo
'
'    End With
'
'End Sub
        
Sub LoadControles()

    BD_Get_Controles_Proyecto
    
    With Me
        '.txtAddressGen = GV_Actual_Project.Controles.AddressGPIB
'        .chkAdquirir.value = GV_Actual_Project.Controles.Adquirir
'        .chkCurvaVideoPot = GV_Actual_Project.Controles.AplicarCurvaVideoPot
'        .txtArchCompSal = GV_Actual_Project.Controles.ArchivoCompensaSalida
'        .chkCapPotGen = GV_Actual_Project.Controles.CapturarPot
'        .chkMedirAnalizador = GV_Actual_Project.Controles.ControlAnalizaEspec
'        .chkMedirOsciloscopio = GV_Actual_Project.Controles.ControlOscilos
'        .chkMedirPowerMeter = GV_Actual_Project.Controles.ControlPowerMeter
'        .chkEsperaEstabi = GV_Actual_Project.Controles.EsperarEstabiliza
'        .txtCurvaVideoPot = GV_Actual_Project.Controles.FileCurvaVideoPot
'        .txtTablaParam = GV_Actual_Project.Controles.FileTablaParam
'        .chkManual = GV_Actual_Project.Controles.OperacionManual
'        .txtTpoEspera = GV_Actual_Project.Controles.TpoEspera
'        .chkUsarTablaParam = GV_Actual_Project.Controles.UsarTablaParam
    End With
    
End Sub

Sub LoadRangos()

Dim LV_Etapas           As Integer
Dim LV_Campos()         As Integer
Dim i                   As Integer

    Me.LstVwRangoControl.ListItems.Clear
    
    LV_Etapas = BD_Get_Rangos_Control(GV_Actual_Project.Cod_Project, GV_Actual_Project.Rango)

    GV_Actual_Project.EtapasDeControl = LV_Etapas
    
    Me.UpDate_LstVw_Rangos
    
End Sub

Sub ProcesaEventoIniEtapa()

Dim LV_Cmd          As String
Dim LV_Inv          As Integer
Dim i               As Integer

    Me.txtPRI.Text = GV_Actual_Project.Rango(PV_Etapa).PRI
    Me.txtPW.Text = GV_Actual_Project.Rango(PV_Etapa).PW
    
'    If Me.chkMedirOsciloscopio.value Then
'        If GV_Actual_Project.Rango(PV_Etapa).AplicarPV Then
''            Me.chkCurvaVideoPot.value = 1
'            LV_Inv = 0
'            If Me.chkInvertir.value Then
'                LV_Inv = 1
'            End If
'            With GV_Actual_Project.Rango(PV_Etapa)
'                LV_Cmd = ComandoOsciloscopio(.b50Ohms, LV_Inv, .VoltDiv)
'                '.CurvaPV
'                'Cargar_Curva_Video_Pot .CurvaPV
'            End With
'            For i = 0 To UBound(GV_Instrumentos)
'                If GV_Instrumentos(i).Name = "Osciloscopio" Then
'                    'Call Send(GPIB0, GV_Result_List(i), LV_Cmd, NLend)
'                    Exit For
'                End If
'            Next
'        End If
'    End If
    
End Sub

Sub ProcesaEventoFinEtapa()

    With GV_Actual_Project.Rango(PV_Etapa)
        If .AccionFinEtapa = "1" Then
            Me.chkRFOn.value = 1
            SendRFPowerOff 0
            Me.tmrPrueba.Enabled = False
'            Me.cmdComenzar.Caption = "&Reanudar"
'            Me.cmdComenzar.Refresh
        End If
    End With
    
End Sub


Sub Refresh_Column_Header()

    ReDim PV_Column_Header(7)
    
    PV_Column_Header(0) = "Etapa"
    PV_Column_Header(1) = "Parámetro"
    PV_Column_Header(2) = "Valor Mín"
    PV_Column_Header(3) = "Valor Máx"
    PV_Column_Header(4) = "Paso"
    PV_Column_Header(5) = "Unidad"
    PV_Column_Header(6) = "Puntos"
    PV_Column_Header(7) = "Duración"
    
    
    AddColumListView Me.LstVwRangoControl, PV_Column_Header
    
End Sub

Sub Refresh_Estado()

    With Me.txtPrueba(2)
        Select Case PV_Estado
        
            Case Is = 0
                .Text = "Enviando Potencia y Frecuencia..."
                
            Case Is = 1
                .Text = "Esperando Estabilización de la Medición..."
                
            Case Is = 2
                .Text = "Adquiriendo Medidas..."
                
            Case Is = 3
                .Text = "Guardando Medidas..."
                
            Case Is = 4
                .Text = "Incrementando Potencia y Frecuencia..."
                
        End Select
        .Refresh
    End With
    
End Sub
        

Sub Refresh_Valores()

    With Me
        
        .txtPrueba(0).Text = PV_Frec_Now
        .txtPrueba(1).Text = PV_Pot_Now
        
    End With
    
End Sub

Sub UpDateControlesProyecto()
    
    BD_Update_Controles_Proyecto
    
End Sub

Sub UpDate_LstVw_Rangos()

Dim LV_Etapas           As Integer
Dim LV_Campos()         As String
Dim i                   As Integer
Dim LV_Ptos             As Double
Dim LV_Ptos_2           As Double

    PV_Ptos_Prueba = 0
    
    Me.LstVwRangoControl.ListItems.Clear
    
    LV_Etapas = GV_Actual_Project.EtapasDeControl
    
    If LV_Etapas Then
    
        ReDim LV_Campos(UBound(PV_Column_Header))
        
        For i = 1 To LV_Etapas / 2
        
            With GV_Actual_Project.Rango(2 * (i - 1))
                LV_Campos(0) = i
                LV_Campos(1) = .Parametro
                LV_Campos(2) = .ValorMin
                LV_Campos(3) = .ValorMax
                LV_Campos(4) = .Paso
                LV_Campos(5) = .Unidad
                LV_Campos(6) = ""
                If .Paso <> 0 Then
                    LV_Ptos = (.ValorMax - .ValorMin) / .Paso + 1
                Else
                    LV_Ptos = 1
                End If
            End With
            
            AddItemListView Me.LstVwRangoControl, LV_Campos
            
            With GV_Actual_Project.Rango(2 * i - 1)
                LV_Campos(0) = ""
                LV_Campos(1) = .Parametro
                LV_Campos(2) = .ValorMin
                LV_Campos(3) = .ValorMax
                LV_Campos(4) = .Paso
                LV_Campos(5) = .Unidad
                If .Paso <> 0 Then
                    LV_Ptos_2 = (.ValorMax - .ValorMin) / .Paso + 1
                Else
                    LV_Ptos_2 = 1
                End If
                LV_Campos(6) = LV_Ptos_2 * LV_Ptos
                PV_Ptos_Prueba = PV_Ptos_Prueba + LV_Ptos_2 * LV_Ptos
            End With
            
            
            AddItemListView Me.LstVwRangoControl, LV_Campos
            
        Next
        
    End If
    
End Sub




Private Sub chkALCOn_Click()

    With Me.chkALCOn
        If .value Then
            SendALCState 0, True
        Else
            SendALCState 0, False
        End If
    End With
    
End Sub

Private Sub chkAplicarCurvaSetup_Click()

    With Me.chkAplicarCurvaSetup
        SaveSetting App.Title, "Properties", .Name & ".Value", .value
        If .value Then
            Cargar_Correccion
            Activar_Display_Valor_Corregido True
        Else
            Activar_Display_Valor_Corregido False
        End If
    End With
    
End Sub






Private Sub chkFMState_Click()

    With Me.chkFMState
        SaveSetting App.Title, "Properties", .Name & ".value", .value
        If .value Then
            SendFmModulacionON
            SendIntFMModulacionON
        Else
            SendFmModulacionOFF
        End If
    End With

End Sub

Private Sub chkIncPRI_Click()

    With Me.chkIncPRI
        SaveSetting App.Title, "Properties", .Name & ".value", .value
    End With
    
End Sub




Private Sub chkModulacionOn_Click()

    With Me.chkModulacionOn
        SaveSetting App.Title, "Properties", .Name & ".value", .value
        If .value Then
            SendModulacionON
            SendIntModulacionON
        Else
            SendModulacionOFF
        End If
    End With

End Sub



Private Sub chkOnTop_Click()

    With Me.chkOnTop
        SaveSetting App.Title, "Properties", .Name & ".Value", .value
        If .value Then
            Form_On_Top Me.hWnd, True
        Else
            Form_On_Top Me.hWnd, False
        End If
    End With

End Sub

Private Sub chkRFOn_Click()

    With Me.chkRFOn
        If .value Then
            PV_Ini_Tmr = GetTickCount
            SendRFPowerOn 0
        Else
            PV_End_Tmr = GetTickCount
            Me.txtCronometro.Text = PV_End_Tmr - PV_Ini_Tmr
            SendRFPowerOff 0
            
        End If
    End With
    
End Sub

Private Sub chkRS232_Click()

    With Me.chkRS232
        If GetSetting(App.Title, "GPIB Config", "Comunicacion RS 232", .value) <> .value Then
            If .value Then
                'Load frmProps
                fMainForm.LoadFormRs232Props
            End If
        End If
        SaveSetting App.Title, "GPIB Config", "Comunicacion RS 232", .value
    End With

End Sub



Private Sub cmdBajarPot_Click(Index As Integer)

Dim LV_Pow_Cor      As Double
Dim LV_Inc          As Double
Dim LV_PRI          As Double

    Select Case Index
        Case Is = 0
            LV_Inc = 1
        Case Is = 1
            LV_Inc = 5
        Case Is = 2
            LV_Inc = 0.1
        Case Is = 3
            LV_Inc = 5
    End Select
    
    With Me
        'If .cmdComenzar.Caption <> "&Comenzar" Then
        LV_Pow_Cor = .txtPrueba(1).Text
        LV_Pow_Cor = LV_Pow_Cor - LV_Inc
        .txtPrueba(1).Text = Int(10 * LV_Pow_Cor) / 10
        .Enviar_Pot_From_TextBox
        If Index = 3 Then
            If .chkIncPRI.value Then
                LV_PRI = .txtPRI.Text
                LV_PRI = LV_PRI - LV_Inc * .txtStepPriPot.Text
                If LV_PRI >= 30 Then
                    .txtPRI.Text = LV_PRI
                End If
            End If
            .cmdSetFrec(0).value = True
        End If
    End With

End Sub

Private Sub cmdClose_Click()

    'Unload Me
    fMainForm.CloseMsComm
    'Me.Hide
    'SetNotifyIcon
    'fMainForm.Running_in_Background
    Me.Running_in_Background
    
End Sub

Private Sub cmdConfigComm_Click()

    fMainForm.LoadFormRs232Props
    
End Sub

Private Sub cmdDecFreq_Click(Index As Integer)

Dim LV_F        As Double
Dim LV_F_I      As Double
Dim LV_Dec      As Double
Dim i           As Integer
    
        Select Case Index
            Case Is = 0
                If IsNumeric(Me.txtStepFreq.Text) = True Then
                    LV_Dec = Me.txtStepFreq.Text / 1000#
                End If
            Case Is = 1
                LV_Dec = 1000
            Case Is = 2
                LV_Dec = 500
            Case Is = 3
                LV_Dec = 10
            Case Is = 4
                LV_Dec = 250
            Case Is = 5
                LV_Dec = 100
        End Select
        
        LV_F = Me.txtPrueba(0).Text
        LV_F = LV_F - LV_Dec
        Me.txtPrueba(0).Text = LV_F
        
        'LV_F = LV_F * 1000000#
        Enviar_Frec_From_TextBox

End Sub


Private Sub cmdDecMod_Click(Index As Integer)

Dim LV_Val          As Double
Dim LV_Inc          As Double
Dim LV_Txt          As TextBox
Dim lv_Min          As Double

    With Me
        LV_Inc = .cboStep(Index).Text
        Select Case Index
            Case Is = 0
                Set LV_Txt = .txtPRI
                lv_Min = 0.4
            Case Is = 1
                Set LV_Txt = .txtPW
                lv_Min = 0.01
            Case Is = 2
                Set LV_Txt = .txtDelay
                lv_Min = 0
        End Select
        LV_Val = Val(LV_Txt.Text)
        LV_Val = LV_Val - LV_Inc
        LV_Val = Int(LV_Val * 100) / 100
        If lv_Min > LV_Val Then
            Exit Sub
        End If
        LV_Txt.Text = LV_Val
        LV_Txt.Text = Replace(LV_Txt.Text, ",", ".")
    End With

End Sub

Private Sub cmdIncFreq_Click(Index As Integer)

Dim LV_F        As Double
Dim LV_F_I      As Double
Dim LV_Inc      As Double
Dim i           As Integer
    
        Select Case Index
            Case Is = 0
                If IsNumeric(Me.txtStepFreq.Text) = True Then
                    LV_Inc = Me.txtStepFreq.Text / 1000#
                End If
            Case Is = 1
                LV_Inc = 1000
            Case Is = 2
                LV_Inc = 500
            Case Is = 3
                LV_Inc = 10
            Case Is = 4
                LV_Inc = 250
            Case Is = 5
                LV_Inc = 100
            Case Is = 6
                If Me.txtPrueba(0).Text = 2500 Then
                    If Me.chkRFOn.value = 0 Then
                        Me.chkRFOn.value = 1
                        Exit Sub
                    End If
                    LV_Inc = 500
                Else
                    LV_Inc = 1000
                End If
        End Select
        
        LV_F = Me.txtPrueba(0).Text
        LV_F = LV_F + LV_Inc
        If LV_F > 18000 Then
            Beep
            If Index = 6 Then
                If Me.chkRFOn.value = 1 Then
                    Me.chkRFOn.value = 0
                    Exit Sub
                End If
            End If
            Exit Sub
        End If
        Me.txtPrueba(0).Text = LV_F
        
        Enviar_Frec_From_TextBox
        
End Sub


Private Sub cmdIncMod_Click(Index As Integer)

Dim LV_Val          As Double
Dim LV_Inc          As Double
Dim LV_Txt          As TextBox
    With Me
        LV_Inc = .cboStep(Index).Text
        Select Case Index
            Case Is = 0
                Set LV_Txt = .txtPRI
            Case Is = 1
                Set LV_Txt = .txtPW
            Case Is = 2
                Set LV_Txt = .txtDelay
        End Select
        LV_Val = Val(LV_Txt.Text)
        LV_Val = LV_Val + LV_Inc
        LV_Txt.Text = LV_Val
        LV_Txt.Text = Replace(LV_Txt.Text, ",", ".")
    End With
    
End Sub

Private Sub cmdSelCompSetup_Click()

Dim sDir        As String
Dim lFlags      As Long
Dim lPath       As String
Dim sFile       As String

    lPath = GetSetting(App.Title, "Properties", "Path Project", App.Path)
    With Me.CommonDialog
        .Filter = "*.csv"
        .InitDir = lPath
        .CancelError = True
        .DialogTitle = "Archivo de Curva Video Potencia"
        On Error Resume Next
        .ShowOpen
        If Err <> cdlCancel Then
            SaveSetting App.Title, "Properties", "Path Project", .InitDir
            Me.txtFileCurvaSetup.Text = .FileName
            Cargar_Correccion
        End If
        On Error GoTo 0
    End With

End Sub





Private Sub cmdSetFrec_Click(Index As Integer)

    Select Case Index
        Case Is = 0
            Me.txtPrueba(0).Text = 2500
        Case Is = 1
            Me.txtPrueba(0).Text = 8000
        Case Is = 2
            Me.txtPrueba(0).Text = 18000
    End Select
    
End Sub

Private Sub cmdSubirPot_Click(Index As Integer)

Dim LV_Pow_Cor      As Double
Dim LV_Inc          As Double
Dim LV_PRI          As Double

    Select Case Index
        Case Is = 0
            LV_Inc = 1
        Case Is = 1
            LV_Inc = 5
        Case Is = 2
            LV_Inc = 0.1
        Case Is = 3
            LV_Inc = 5
    End Select
    
    With Me
        LV_Pow_Cor = 10 * .txtPrueba(1).Text
        LV_Pow_Cor = LV_Pow_Cor + 10 * LV_Inc
        .txtPrueba(1).Text = Int(LV_Pow_Cor) / 10
        .Enviar_Pot_From_TextBox
        If Index = 3 Then
            If .chkIncPRI.value Then
                LV_PRI = .txtPRI.Text
                LV_PRI = LV_PRI + LV_Inc * .txtStepPriPot.Text
                If LV_PRI >= 30 Then
                    .txtPRI.Text = LV_PRI
                End If
            End If
            .cmdSetFrec(0).value = True
        End If
    End With

End Sub


Private Sub Command1_Click()

    
End Sub

Private Sub Form_Load()

Dim i           As Integer
Dim j           As Integer

Dim LV_Default  As Double

    With Me
        
        For i = .cboStep.LBound To .cboStep.UBound
            With .cboStep(i)
                .Clear
                'LV_Default = 0.02
                For j = 0 To 6
                    If j > 3 And i Then
                        Exit For
                    End If
                    Select Case j
                        Case Is = 0
                            LV_Default = 0.02
                        Case Is = 1
                            LV_Default = 0.1
                        Case Is = 2
                            LV_Default = 1
                        Case Is = 3
                            LV_Default = 10
                        Case Is = 4
                            LV_Default = 100
                        Case Is = 5
                            LV_Default = 1000
                        Case Is = 6
                            LV_Default = 10000
'                        Case j = 7
'                        LV_Default = 0.1
'                        Case j = 8
'                        LV_Default = 0.1
                    End Select
                    .AddItem LV_Default
                Next
                .ListIndex = 0
            End With
        Next
        
        .Left = GetSetting(Me.Name, "Settings", "MainLeft", 1000)
        .Top = GetSetting(Me.Name, "Settings", "MainTop", 1000)
        PV_Estado = 0
        .txtStepFreq.Text = "250"
        .txtUnidad(0).Text = "[MHz]"
        .txtUnidad(1).Text = "[dBm]"
        .txtUnidad(2).Text = "[dBm]"
'        With .txtAddressGen
'            .Text = GetSetting(App.Title, "GPIB Config", "Address Generator", 28)
'        End With
        Inicializar_Comandos_Instrumentos
        
'        .chkRS232.value = GetSetting(App.Title, "GPIB Config", "Comunicacion RS 232", 0)
        .txtAgregarPot.Text = GetSetting(App.Title, "Properties", .txtAgregarPot.Name, 0)
        
        With .txtFileCurvaSetup
            .Text = GetSetting(App.Title, "Properties", .Name & ".Text", "")
        End With
        With .chkAplicarCurvaSetup
            .value = GetSetting(App.Title, "Properties", .Name & ".Value", .value)
        End With
        
        With .txtPRI
            .Text = GetSetting(App.Title, "Generador Config", "PRI", 1000)
        End With
        
        With .txtPW
            .Text = GetSetting(App.Title, "Generador Config", "PW", 1)
        End With
        For i = 0 To .txtPrueba.UBound
            With Me.txtPrueba(i)
                Select Case i
                    Case Is = 0
                        LV_Default = 9000
                    Case Is = 1
                        LV_Default = -45
                End Select
                .Text = GetSetting(App.Title, "Properties", .Name & Trim(i) & ".Text", LV_Default)
            End With
        Next
        
        With .chkModulacionOn
            .value = GetSetting(App.Title, "Properties", .Name & ".value", 1)
        End With
        With .txtDelay
            .Text = GetSetting(App.Title, "Properties", .Name & ".Text", 0)
        End With
        With .OptionTrigger(0)
            i = GetSetting(App.Title, "Properties", .Name, 0)
        End With
        .OptionTrigger(i).value = True
        Iniciar_Instrumentos_RS232
'        .chkGeneradorEn.value = 1
        With .txtStepPriPot
            .Text = GetSetting(App.Title, "Properties", .Name & ".Text", 10)
        End With
        With Me.chkIncPRI
            .value = GetSetting(App.Title, "Properties", .Name & ".value", .value)
        End With
        With Me.chkOnTop
            .value = GetSetting(App.Title, "Properties", .Name & ".Value", .value)
        End With
    End With
        
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim lngReturnValue As Long
Static rec As Boolean
Dim Msg As Long
    
    If Button = 1 And Me.WindowState <> vbNormal Then
       Call ReleaseCapture
       lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                    HTCAPTION, 0&)
    Else
        Msg = X / Screen.TwipsPerPixelX
        If rec = False Then
            rec = True
            Select Case Msg
                Case WM_LBUTTONDBLCLK:                      ' doble click con el boton izquierdo del raton
                    fMainForm.IniciarCommPort
                    Me.Show                                                                 ' mostramos el formulario
                    Me.tmrNotifyIcon.Enabled = True
                Case WM_RBUTTONUP:
                    fMainForm.Levantar_Menu_Cerrar
                    '.PopupMenu mnuCerrar                       ' click con el boton secundario, mostramos el menu correspondiente
                Case NIN_BALLOONUSERCLICK:  'Click al ballon Tool Tip
                    'MsgBox "hizo click al ballon", vbExclamation, "Mensaje"
                    'Me.Show
            End Select
            rec = False
        End If
    End If

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Close_Devices
    
    Me.tmrPrueba.Enabled = False
    
End Sub

Private Sub Form_Resize()

Dim i           As Integer
Dim lHeight     As Long

    With Me
        
'        If .WindowState = vbMinimized Or .Width = 0 Then
'            Exit Sub
'        End If
'
'        For i = 0 To .framePrueba.UBound
'            .framePrueba(i).Width = .Width - 2 * .framePrueba(i).Left
'        Next
'
'        i = 2
'        '.framePrueba(i).Top = .Height - .framePrueba(i).Height
'        .cmdComenzar.Left = .Width - .cmdCancelar.Left - .cmdComenzar.Width
'        '.cmdModificar(i).Left = (.Width - .cmdModificar(i).Width) / 2
'
'        i = 1
'        '.framePrueba(i).Top = .framePrueba(i - 1).Height
'        '.framePrueba(i).Height = .framePrueba(i + 1).Top - .framePrueba(i - 1).Height
'
'        .LstVwVisualTest.Width = .framePrueba(i).Width - 2 * .LstVwVisualTest.Left
'        .LstVwVisualTest.Height = .framePrueba(i).Height _
'                                     - 2 * .LstVwVisualTest.Top
'
'        .LstVwRangoControl.Width = .framePrueba(0).Width - 2 * .LstVwRangoControl.Left
'        'lHeight = (.Height - .framePrueba(i).Height) / .framePrueba.UBound
'
'        .frameControlGral.Width = .framePrueba(i).Width - 2 * .frameControlGral.Left
'        .frameParamGPIB.Width = .frameControlGral.Width
'
'        '.frameReadOut.Left = 0
'        '.frameReadOut.Top = 0
'        .frameReadOut.Width = .framePrueba(0).Width
'        .frameReadOut.Height = .framePrueba(0).Height
'
'        '.FrameModPulsos.Left = 0
'        '.FrameModPulsos.Top = 0
'        .FrameModPulsos.Width = .framePrueba(1).Width
'        .FrameModPulsos.Height = .framePrueba(1).Height
    End With
    
End Sub

Private Sub Text1_Change()

End Sub


Private Sub Form_Unload(Cancel As Integer)

    If Me.WindowState <> vbMinimized Then
        SaveSetting Me.Name, "Settings", "MainLeft", Me.Left
        SaveSetting Me.Name, "Settings", "MainTop", Me.Top
    End If

    Me.Delete_Notify_Icon
    
End Sub

Private Sub FrameModPulsos_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngReturnValue As Long

    If Button = 1 Then
       Call ReleaseCapture
       lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                    HTCAPTION, 0&)
    End If

End Sub

Private Sub frameReadOut_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim lngReturnValue As Long

    If Button = 1 Then
       Call ReleaseCapture
       lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, _
                                    HTCAPTION, 0&)
    End If

End Sub

Private Sub OptionTrigger_Click(Index As Integer)

    With Me.OptionTrigger(Index)
        SaveSetting App.Title, "Properties", .Name, Index
    End With
    
    Select Case Index
        Case Is = 0
            SendTriggerMode "AUTO"
        Case Is = 1
            SendTriggerMode "EXT"
        Case Is = 2
            SendTriggerMode "EXT_G"
    End Select
    
End Sub

Private Sub tmrNotifyIcon_Timer()

    Shell_NotifyIcon NIM_DELETE, Notify
    Me.tmrNotifyIcon.Enabled = False

End Sub

Private Sub tmrPrueba_Timer()

Dim LV_Enable           As Boolean

    With Me
    
        If PV_EmpezarEtapa = True Then
            PV_EmpezarEtapa = False
            ProcesaEventoIniEtapa
        End If
        
        Refresh_Estado
        
        Select Case PV_Estado
        
            Case Is = 0
                Enviar_Frecuencia_Potencia
                
                PV_Ptos_Now = PV_Ptos_Now + 1
                
                
                PV_Estado = PV_Estado + 1
                
'                If Me.chkManual.value = 1 Then
'                    Me.cmdComenzar.value = True
'                End If
                
                PV_TckCnt_E1 = GetTickCount
                LV_Enable = False
                
            Case Is = 1
                'Estabilizando Medidas
                PV_Estabiliza_Counter = PV_Estabiliza_Counter - 1
                'If PV_Estabiliza_Counter <= 0 Or Me.chkEsperaEstabi.value = 0 Then
                '    PV_Estado = PV_Estado + 1
                'End If
                PV_Estado = PV_Estado + 1
                
            Case Is = 2
            
                If LV_Enable = True Then
                    LV_Enable = False
                End If
                
            Case Is = 3
                Guardar_Valores
            
                PV_Estado = PV_Estado + 1
            
            Case Is = 4
                
            
                Refresh_Valores
                
                fMainForm.Refresca_Contador_Tpo PV_Tpo_Ini, PV_Ptos_Prueba, PV_Ptos_Now
                
                PV_Estado = 0
        End Select
    End With
    
End Sub


Private Sub txtAgregarPot_Change()

    With Me.txtAgregarPot
        If IsNumeric(.Text) = True Then
            SaveSetting App.Title, "Properties", .Name, .Text
            'GV_Actual_Project.Controles.TpoEspera = .Text
            'UpDateControlesProyecto
        End If
    End With

End Sub






Private Sub txtFrecGen_Change()

End Sub

Private Sub txtDelay_Change()

    With Me.txtDelay
        If IsNumeric(.Text) = True Then
            SaveSetting App.Title, "Properties", .Name & ".Text", .Text
            SendPulseDelay Val(.Text)
        End If
    End With
    
End Sub

Private Sub txtFileCurvaSetup_Change()

    With Me.txtFileCurvaSetup
        SaveSetting App.Title, "Properties", .Name & ".Text", .Text
    End With
    
End Sub

Private Sub txtPRI_Change()

'    If GPIBglobalsRegistered = 0 Then
'        'Exit Sub
'    End If
    
    With Me.txtPRI
        If IsNumeric(.Text) = True Then
            If Val(.Text) > 0 Then
                SendPRI 0, Val(.Text)
                SaveSetting App.Title, "Generador Config", "PRI", .Text
            End If
        End If
    End With
    
End Sub

Private Sub txtPrueba_Change(Index As Integer)

    With Me.txtPrueba(Index)
        If IsNumeric(.Text) = True Then
            SaveSetting App.Title, "Properties", .Name & Trim(Index) & ".Text", .Text
        Else
            Exit Sub
        End If
    End With
    
    Select Case Index
        Case Is = 0
            PV_Flag_Freq = True
            Enviar_Frec_From_TextBox
            If Me.chkAplicarCurvaSetup.value Then
                PV_Flag_Freq = True
                Enviar_Pot_From_TextBox
            End If
        Case Is = 1
            PV_Flag_Freq = True
            Enviar_Pot_From_TextBox
        
    End Select
End Sub

Private Sub txtPW_Change()

'    If GPIBglobalsRegistered = 0 Then
'        'Exit Sub
'    End If
    
    With Me.txtPW
        If IsNumeric(.Text) = True Then
            If Val(.Text) > 0 Then
                If Val(.Text) <= 0.22 Then
                    Me.chkALCOn.value = 0
                    SendALCState 0, False
                Else
                    Me.chkALCOn.value = 1
                    SendALCState 0, True
                End If
                SendPW 0, Val(.Text)
                SaveSetting App.Title, "Generador Config", "PW", .Text
            End If
        End If
    End With

End Sub



Private Sub txtStepPriPot_Change()

    With Me.txtStepPriPot
        SaveSetting App.Title, "Properties", .Name & ".Text", .Text
    End With

End Sub





Private Sub txtUnidad_KeyPress(Index As Integer, KeyAscii As Integer)

    KeyAscii = 0
    
End Sub
