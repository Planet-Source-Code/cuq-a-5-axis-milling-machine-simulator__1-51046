VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Machine_SIMUL_FRM 
   Caption         =   "Machine Simulation"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   14475
   FillColor       =   &H00FF0000&
   Icon            =   "Machine_simul.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   965
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   10680
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Support_BTN 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   120
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   953
      TabIndex        =   2
      Top             =   3480
      Width           =   14295
      Begin VB.Frame FrameLumiere 
         Caption         =   "Light"
         Height          =   975
         Left            =   0
         TabIndex        =   92
         Top             =   4920
         Width           =   8415
         Begin VB.TextBox Light 
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
            Left            =   7560
            TabIndex        =   98
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox Light 
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
            Left            =   7560
            TabIndex        =   97
            Text            =   "0"
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox Light 
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
            Left            =   7560
            TabIndex        =   96
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.HScrollBar HScroll 
            Height          =   135
            Index           =   0
            Left            =   4680
            Max             =   4000
            Min             =   -4000
            TabIndex        =   95
            Top             =   240
            Width           =   2775
         End
         Begin VB.HScrollBar HScroll 
            Height          =   135
            Index           =   1
            Left            =   4680
            Max             =   4000
            Min             =   -4000
            TabIndex        =   94
            Top             =   480
            Width           =   2775
         End
         Begin VB.HScrollBar HScroll 
            Height          =   135
            Index           =   2
            Left            =   4680
            Max             =   4000
            Min             =   -4000
            TabIndex        =   93
            Top             =   720
            Width           =   2775
         End
         Begin MSComctlLib.Slider sldAlpha 
            Height          =   255
            Left            =   360
            TabIndex        =   99
            Top             =   360
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Max             =   100
            TickFrequency   =   15
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Transparency"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   102
            Top             =   150
            Width           =   1335
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Full"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   101
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.Frame FrameOrigine 
         Caption         =   "Program file Origine"
         ForeColor       =   &H00000000&
         Height          =   2175
         Left            =   3720
         TabIndex        =   65
         Top             =   2760
         Width           =   2175
         Begin VB.CommandButton CommandFixOrigine 
            Caption         =   "Fix Point"
            Height          =   255
            Left            =   120
            Picture         =   "Machine_simul.frx":030A
            TabIndex        =   75
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox AxeOrigine 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   69
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox AxeOrigine 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   68
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox AxeOrigine 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   67
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox AxeOrigine 
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   66
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label LabelOrigineProgramme 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   73
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LabelOrigineProgramme 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   72
            Top             =   600
            Width           =   255
         End
         Begin VB.Label LabelOrigineProgramme 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   71
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LabelOrigineProgramme 
            Caption         =   "ED"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   70
            Top             =   1320
            Width           =   255
         End
      End
      Begin VB.Frame FrameControl 
         Caption         =   "Manual Control"
         ForeColor       =   &H00000000&
         Height          =   4815
         Left            =   1800
         TabIndex        =   40
         Top             =   120
         Width           =   1935
         Begin VB.PictureBox PictureAxeMachine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   1
            Left            =   120
            Picture         =   "Machine_simul.frx":1138
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   52
            Top             =   600
            Width           =   330
         End
         Begin VB.TextBox AxeMachine 
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   51
            Text            =   "0"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox AxeMachine 
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   50
            Text            =   "0"
            Top             =   960
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeMachine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   2
            Left            =   120
            Picture         =   "Machine_simul.frx":162A
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   49
            Top             =   1320
            Width           =   330
         End
         Begin VB.PictureBox PictureAxeMachine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   3
            Left            =   120
            Picture         =   "Machine_simul.frx":1B1C
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   48
            Top             =   2040
            Width           =   330
         End
         Begin VB.TextBox AxeMachine 
            Height          =   285
            Index           =   3
            Left            =   840
            TabIndex        =   47
            Text            =   "0"
            Top             =   1680
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeMachine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   4
            Left            =   120
            Picture         =   "Machine_simul.frx":200E
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   46
            Top             =   2760
            Width           =   330
         End
         Begin VB.TextBox AxeMachine 
            Height          =   285
            Index           =   4
            Left            =   840
            TabIndex        =   45
            Text            =   "0"
            Top             =   2400
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeMachine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   5
            Left            =   120
            Picture         =   "Machine_simul.frx":2500
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   44
            Top             =   3480
            Width           =   330
         End
         Begin VB.TextBox AxeMachine 
            Height          =   285
            Index           =   5
            Left            =   840
            TabIndex        =   43
            Text            =   "0"
            Top             =   3120
            Width           =   855
         End
         Begin VB.PictureBox PictureAxeMachine 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   330
            Index           =   6
            Left            =   120
            Picture         =   "Machine_simul.frx":29F2
            ScaleHeight     =   260.87
            ScaleMode       =   0  'User
            ScaleWidth      =   260.87
            TabIndex        =   42
            Top             =   4200
            Width           =   330
         End
         Begin VB.TextBox AxeMachine 
            Height          =   285
            Index           =   6
            Left            =   840
            TabIndex        =   41
            Text            =   "0"
            Top             =   3840
            Width           =   855
         End
         Begin MSComctlLib.Slider SliderAxeMachine 
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   53
            Top             =   600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeMachine 
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   54
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeMachine 
            Height          =   255
            Index           =   3
            Left            =   720
            TabIndex        =   55
            Top             =   2040
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeMachine 
            Height          =   255
            Index           =   4
            Left            =   720
            TabIndex        =   56
            Top             =   2760
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeMachine 
            Height          =   255
            Index           =   5
            Left            =   720
            TabIndex        =   57
            Top             =   3480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin MSComctlLib.Slider SliderAxeMachine 
            Height          =   255
            Index           =   6
            Left            =   720
            TabIndex        =   58
            Top             =   4200
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   -360
            Max             =   360
            TickFrequency   =   50
         End
         Begin VB.Label LabelAxeMachine 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LabelAxeMachine 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   63
            Top             =   960
            Width           =   615
         End
         Begin VB.Label LabelAxeMachine 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label LabelAxeMachine 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   61
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label LabelAxeMachine 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   60
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label LabelAxeMachine 
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   59
            Top             =   3840
            Width           =   495
         End
      End
      Begin VB.Frame FrameXYZABC 
         Caption         =   "Absolute Position"
         ForeColor       =   &H00000000&
         Height          =   2655
         Left            =   3720
         TabIndex        =   18
         Top             =   120
         Width           =   2175
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   30
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   29
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   28
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   4
            Left            =   360
            TabIndex        =   27
            Text            =   "0"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   5
            Left            =   360
            TabIndex        =   26
            Text            =   "0"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   6
            Left            =   960
            TabIndex        =   25
            Text            =   "0"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   7
            Left            =   960
            TabIndex        =   24
            Text            =   "0"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   9
            Left            =   1560
            TabIndex        =   23
            Text            =   "0"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   10
            Left            =   1560
            TabIndex        =   22
            Text            =   "0"
            Top             =   1920
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   11
            Left            =   1560
            TabIndex        =   21
            Text            =   "0"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   3
            Left            =   360
            TabIndex        =   20
            Text            =   "0"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox AxeAbsolu 
            Height          =   285
            Index           =   8
            Left            =   960
            TabIndex        =   19
            Text            =   "0"
            Top             =   2280
            Width           =   495
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   39
            Top             =   240
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   38
            Top             =   600
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   37
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "I"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   1560
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "J"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   1920
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "K"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   34
            Top             =   2280
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Vx"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   33
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Vy"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1080
            TabIndex        =   32
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LabelAxeAbsolu 
            Caption         =   "Vz"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   1680
            TabIndex        =   31
            Top             =   1320
            Width           =   255
         End
      End
      Begin VB.Frame FrameMagasin 
         Caption         =   "Tool List Control"
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   6000
         TabIndex        =   15
         Top             =   3360
         Width           =   2415
         Begin VB.CommandButton CommandUnloadTool 
            Caption         =   "Unload Tool"
            Height          =   255
            Left            =   240
            TabIndex        =   91
            Top             =   1200
            Width           =   1935
         End
         Begin VB.CommandButton CommandLoadTool 
            Caption         =   "Load Tool"
            Height          =   255
            Left            =   240
            TabIndex        =   90
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox Magasin 
            Alignment       =   2  'Center
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1036
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   16
            Text            =   "1"
            Top             =   600
            Width           =   975
         End
         Begin MSComctlLib.Slider SliderMagasin 
            Height          =   255
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   393216
            Min             =   1
            Max             =   8
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "M06 T"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   89
            Top             =   600
            Width           =   615
         End
      End
      Begin VB.PictureBox MessageLOG 
         AutoRedraw      =   -1  'True
         Height          =   5775
         Left            =   8520
         ScaleHeight     =   5715
         ScaleWidth      =   5355
         TabIndex        =   14
         Top             =   120
         Width           =   5415
      End
      Begin VB.Frame FrameDestination 
         Caption         =   "Goto Point"
         ForeColor       =   &H00000000&
         Height          =   3255
         Left            =   6000
         TabIndex        =   10
         Top             =   120
         Width           =   2415
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   6
            Left            =   720
            TabIndex        =   108
            Text            =   "0"
            Top             =   2040
            Width           =   1215
         End
         Begin VB.PictureBox PictureVoyant 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   470
            Index           =   2
            Left            =   2040
            Picture         =   "Machine_simul.frx":2EE4
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   87
            Top             =   2520
            Width           =   470
         End
         Begin VB.PictureBox PictureVoyant 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   470
            Index           =   1
            Left            =   1920
            Picture         =   "Machine_simul.frx":3AC6
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   86
            Top             =   2520
            Width           =   470
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   84
            Text            =   "0"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   82
            Text            =   "0"
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   80
            Text            =   "0"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   78
            Text            =   "0"
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox AxeDestination 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   76
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.PictureBox PictureVoyant 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   470
            Index           =   0
            Left            =   1680
            Picture         =   "Machine_simul.frx":46A8
            ScaleHeight     =   465
            ScaleWidth      =   465
            TabIndex        =   74
            Top             =   2520
            Width           =   470
         End
         Begin VB.CommandButton CommandGoto 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   240
            Picture         =   "Machine_simul.frx":528A
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   2400
            Width           =   1215
         End
         Begin VB.CommandButton CommandRedo 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   840
            Picture         =   "Machine_simul.frx":60B8
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2800
            Width           =   615
         End
         Begin VB.CommandButton CommandUndo 
            BackColor       =   &H00E0E0E0&
            Height          =   375
            Left            =   240
            Picture         =   "Machine_simul.frx":6CB2
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2800
            Width           =   615
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   109
            Top             =   2040
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   85
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   83
            Top             =   1320
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "Z"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   81
            Top             =   960
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "Y"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   79
            Top             =   600
            Width           =   255
         End
         Begin VB.Label LabelAxeDestination 
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   77
            Top             =   240
            Width           =   255
         End
      End
      Begin VB.Frame FrameAffichage 
         Caption         =   "Display"
         Height          =   4815
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton VueMachine 
            Caption         =   "View-Y"
            Height          =   195
            Index           =   3
            Left            =   840
            TabIndex        =   113
            Top             =   3000
            Width           =   855
         End
         Begin VB.OptionButton VueMachine 
            Caption         =   "View-X"
            Height          =   195
            Index           =   2
            Left            =   840
            TabIndex        =   112
            Top             =   2760
            Width           =   855
         End
         Begin VB.OptionButton VueMachine 
            Caption         =   "ViewY"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   111
            Top             =   3000
            Width           =   855
         End
         Begin VB.OptionButton VueMachine 
            Caption         =   "ViewX"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   110
            Top             =   2760
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CheckBox CheckBox 
            Caption         =   "Box controle"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.TextBox Zgrille 
            Height          =   285
            Left            =   960
            TabIndex        =   106
            Text            =   "0"
            Top             =   3840
            Width           =   615
         End
         Begin VB.CheckBox VueFix 
            Caption         =   "View tool holder X"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   105
            Top             =   2040
            Width           =   1575
         End
         Begin VB.CheckBox OptionRTCP 
            Caption         =   "RTCP"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   3480
            Width           =   1215
         End
         Begin VB.CheckBox OptionTracerParcours 
            Caption         =   "Tool Path"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1680
            Width           =   1215
         End
         Begin VB.CheckBox CHK 
            Caption         =   "RasterZ"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   3840
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.OptionButton OPT 
            Caption         =   "Shade"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OPT 
            Caption         =   "Wire"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton OPT 
            Caption         =   "Depth Wire"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox VueFix 
            Caption         =   "Machine View"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   2400
            Width           =   1575
         End
         Begin MSComctlLib.Slider SliderIncrement 
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   4440
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            LargeChange     =   10
            Min             =   1
            Max             =   50
            SelStart        =   10
            TickFrequency   =   10
            Value           =   10
         End
         Begin VB.Line Line1 
            Index           =   2
            X1              =   120
            X2              =   1680
            Y1              =   1560
            Y2              =   1560
         End
         Begin VB.Line Line1 
            Index           =   1
            X1              =   120
            X2              =   1680
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   120
            X2              =   1680
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Label Label2 
            Caption         =   "Speed Simulation"
            Height          =   255
            Left            =   120
            TabIndex        =   104
            Top             =   4200
            Width           =   1455
         End
      End
   End
   Begin RichTextLib.RichTextBox Programme 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5953
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   3
      TextRTF         =   $"Machine_simul.frx":7834
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   3435
      Left            =   4920
      ScaleHeight     =   227
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   364
      TabIndex        =   0
      Top             =   120
      Width           =   5490
   End
   Begin MSComctlLib.ImageList ImageListTOOL 
      Left            =   10680
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Machine_simul.frx":78B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Machine_simul.frx":7EA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Machine_simul.frx":8AB1
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Machine_simul.frx":96C7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeViewTOOL 
      Height          =   4815
      Left            =   11280
      TabIndex        =   114
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   8493
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageListTOOL"
      Appearance      =   1
   End
   Begin VB.Menu mnuFichier 
      Caption         =   "&File"
      Begin VB.Menu mnuOuvrirProjet 
         Caption         =   "Open Project"
      End
      Begin VB.Menu mnuSauverProjet 
         Caption         =   "Save Project"
      End
      Begin VB.Menu Separator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOuvrir 
         Caption         =   "&Open ISO File"
      End
      Begin VB.Menu mnuSauver 
         Caption         =   "&Save ISO File"
      End
      Begin VB.Menu mnuNouveau 
         Caption         =   "&New ISO File"
      End
      Begin VB.Menu Separateur1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChargerMachine 
         Caption         =   "Load  &Machine"
      End
      Begin VB.Menu mnuCharger 
         Caption         =   "&Load Part"
      End
      Begin VB.Menu mnuDeCharger 
         Caption         =   "Unload Part"
      End
      Begin VB.Menu separateur2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuiter 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuSimulation 
      Caption         =   "&Simulation"
      Begin VB.Menu mnuExecute 
         Caption         =   "S&imulate file"
      End
      Begin VB.Menu SeparatorS0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAffichageProgramme 
         Caption         =   "Display &File"
      End
      Begin VB.Menu mnuRenum 
         Caption         =   "Renum ISO FIle"
      End
   End
   Begin VB.Menu mnuOutil 
      Caption         =   "Tool"
      Begin VB.Menu mnuListeOutil 
         Caption         =   "&Tool List"
      End
      Begin VB.Menu mnuChargerOutil 
         Caption         =   "Load Tool"
      End
      Begin VB.Menu mnuDeChargerOutil 
         Caption         =   "Unload Tool"
      End
   End
   Begin VB.Menu mnuVue 
      Caption         =   "View"
      Begin VB.Menu mnuVueDessus 
         Caption         =   "Top View"
      End
      Begin VB.Menu mnuVueAuxilliaire 
         Caption         =   "Auxiliarry View"
      End
      Begin VB.Menu mnuVueBase 
         Caption         =   "Base View"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "O&ptions"
      Begin VB.Menu mnuInfoVue 
         Caption         =   "Information &View"
      End
      Begin VB.Menu mnuUpdateInfoVue 
         Caption         =   "&Update Information View"
      End
      Begin VB.Menu mnuUpdateInfoVueAuxiliaire 
         Caption         =   "Update Info Auxilliray View"
      End
      Begin VB.Menu separatorO1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color Part"
      End
      Begin VB.Menu separatorO2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRAZPO 
         Caption         =   "&Reset ToolLpath"
      End
      Begin VB.Menu mnuRazLog 
         Caption         =   "&Reset Log"
      End
   End
   Begin VB.Menu mnuApropos 
      Caption         =   "About"
      Begin VB.Menu mnuVersion 
         Caption         =   "Release"
      End
   End
End
Attribute VB_Name = "Machine_SIMUL_FRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Variable pour indiquer le chargement Tool
Dim Chargement As Boolean


Dim mnIndex As Integer   ' Index du noueud pour le treeview
Dim mbIndrag As Boolean  ' Flag pour Drag Drop operation
Dim moDragNode As Object ' Item qui est glissé




Private Sub Axemachine_Change(Index As Integer)
    If Val(AxeMachine(Index)) <> 0 Then
        AxeMachine(Index) = Round(Val(AxeMachine(Index)), 4)
    End If
End Sub

Private Sub AxeOrigine_Change(Index As Integer)
    OrigineProg.X = Val(AxeOrigine(1))
    OrigineProg.Y = Val(AxeOrigine(2))
    OrigineProg.Z = Val(AxeOrigine(3))
    OrigineProg.ED = Val(AxeOrigine(4))

    ' affectation piece
    Piece.Origine.X = OrigineProg.X
    Piece.Origine.Y = OrigineProg.Y
    Piece.Origine.Z = OrigineProg.Z
    
    Piece.Valeur_axe = OrigineProg.ED
    Call Pic_Paint ' mise a jour affichage
End Sub

Private Sub Bouton1_Click()
PreVisu.Show
End Sub



'Initialisation de la machine
Sub Init_Machine(Repertoire As String, Fichier_Def_machine As String)
Dim Fichier_Stl As String
Dim I, J As Integer
Dim Pt As Point3
Dim Box0 As Box3
Dim Elem_tempo As Element3D
Dim Obb_Box_tempo As OBB_box
Dim Dec1 As DecOrigine
Dim Mere As Integer
Dim Element_Test As Integer
' Init si plusieurs chargement machine
Machine.Pivot = Pt
Machine.DecalageMachine = Pt
Machine.Tool_current = 0


' Init des sliders pour les axes machines
For I = 3 To 6
    PictureAxeMachine(I).Picture = LoadResPicture(99, vbResBitmap)
    SliderAxeMachine(I).Visible = False
    AxeMachine(I).Visible = False
    LabelAxeMachine(I).Visible = False
    PictureAxeMachine(I).Visible = False
Next

'Chargement caractéristique machine
Call Charger_Machine(Fichier_Def_machine, Machine)


'Recuperation de la definition geométrique via fichier STL ascii
For I = 0 To UBound(Machine.Element)
    Call Reinit_Element(Machine.Element(I))
    Call ChargeFichierSTL(App.Path + "\Machine_def\" + Repertoire + "\" + Machine.Element(I).fichier, Machine.Element(I).STL_def)
    'Machine.Element(I).Color = RVB(QBColor(I + 1))
Next



ReDim ToolC(Machine.MagasinPo.MaxiAxe)
ReDim POTool(Machine.MagasinPo.MaxiAxe)

'Magasin Tool
Call Reinit_Element(Machine.MagasinPo)
        
If Machine.PositionMagasin Then
    For I = 1 To Machine.MagasinPo.MaxiAxe
        'ToolC(I).Name = "T" & I
        'POTool(I).Name = "PO" & I
        ToolC(I).Origine.X = 280 * Sin((360 / Machine.MagasinPo.MaxiAxe) * (I - 1) * DEGTORAD)
        ToolC(I).Origine.Y = 280 * Cos((360 / Machine.MagasinPo.MaxiAxe) * (I - 1) * DEGTORAD)
        ToolC(I).Origine.Z = 0
        POTool(I).Origine = ToolC(I).Origine
        ToolC(I).Vecteur.X = 0
        ToolC(I).Vecteur.Y = 0
        ToolC(I).Vecteur.Z = 1
        POTool(I).Vecteur = ToolC(I).Vecteur
        'POTool(I).NB_point = 9
        'ReDim POTool(I).Coord(POTool(I).NB_point)
    Next I
    

        Call ChargeFichierSTL(App.Path + "\Machine_def\" + Repertoire + "\" + Machine.MagasinPo.fichier, Machine.MagasinPo.STL_def)
        'Machine.MagasinPo.Color = GetRVB(QBColor(4))
        Machine.MagasinPo.Valeur_axe = 1
    
End If


    ' traitement des Obb pour l'element 5
    Element_Test = Machine.Element_Collision
    
    ReDim Machine.Element(Element_Test).Obb_Col(1)
    Machine.Element(Element_Test).Obb_Col(0).Pointeur_Mere = -1
    Machine.Element(Element_Test).Obb_Col(0).Box_Descr = Def_Box(Machine.Element(Element_Test).STL_def)
    'decalage according config machine
    Machine.Element(Element_Test).Box_Col = Machine.Element(Element_Test).Obb_Col(0).Box_Descr
    
    Call MessageLOG_CLS 'Réinit fentre LOG


    ' Stock pour la durée des modif les infos
    ' Elem_tempo = Machine.Element(5)
    Obb_Box_tempo = Machine.Element(Element_Test).Obb_Col(0)
    Call Creer_Obb_Box(Obb_Box_tempo, Machine.Element(Element_Test), False, 0)
    'réintègre les infos à leur place
    'Machine.Element(5) = Elem_tempo
    Machine.Element(Element_Test).Obb_Col(0) = Obb_Box_tempo
    'Debug.Print UBound(Machine.Element(Element_Test).Obb_Col)
    'Deuxieme niveau
    For I = 1 To 8
        Obb_Box_tempo = Machine.Element(Element_Test).Obb_Col(I)
        Mere = I
        Call Creer_Obb_Box(Obb_Box_tempo, Machine.Element(Element_Test), True, Mere)
        Machine.Element(Element_Test).Obb_Col(I) = Obb_Box_tempo
        'Debug.Print UBound(Machine.Element(Element_Test).Obb_Col)
    Next I


 
 
'Tool & Porte Tool according machine
 Call Init_PO(Fichier_Def_machine, POTool)
 Call init_Tool(Fichier_Def_machine, ToolC)


' Init des sliders pour les axes machines
For I = 1 To UBound(Machine.Element) - 1
    PictureAxeMachine(I).Picture = LoadResPicture(Machine.Element(I).Type_axe, vbResBitmap)
    SliderAxeMachine(I).Max = Machine.Element(I).MaxiAxe
    SliderAxeMachine(I).Min = Machine.Element(I).MiniAxe

    SliderAxeMachine(I).TickFrequency = Abs(SliderAxeMachine(I).Max - SliderAxeMachine(I).Min) / 10
    SliderAxeMachine(I).Visible = True
    AxeMachine(I).Visible = True
    LabelAxeMachine(I).Visible = True
    PictureAxeMachine(I).Visible = True
    LabelAxeMachine(I).Caption = Machine.Element(I).name
Next

' position de l'origine au point pivot
' ( Element5 = Point pivot )
For I = 1 To Machine.NB_axe
    Machine.DecalageMachine.X = Machine.DecalageMachine.X + Machine.Element(I).Origine.X
    Machine.DecalageMachine.Y = Machine.DecalageMachine.Y + Machine.Element(I).Origine.Y
    Machine.DecalageMachine.Z = Machine.DecalageMachine.Z + Machine.Element(I).Origine.Z
Next

' position de point pivot a nez de broche
For I = (Machine.NB_axe + 1) To UBound(Machine.Element)
    Machine.Pivot.X = Machine.Pivot.X + Machine.Element(I).Origine.X
    Machine.Pivot.Y = Machine.Pivot.Y + Machine.Element(I).Origine.Y
    Machine.Pivot.Z = Machine.Pivot.Z + Machine.Element(I).Origine.Z
Next

 PositionPivot = Machine.Pivot
' Affiche le plan
Call Def_Plan(Val(Zgrille))




' Chargement effectuée
Chargement = True
End Sub


Private Sub Bouton2_KeyPress(KeyAscii As Integer)
'Stop la simulation
If KeyAscii = 27 Then
    stop_exec = True
End If
End Sub


Private Sub CheckBox_Click()

If Chargement Then
 Call Collision
End If
 
 Call Pic_Paint
End Sub

Private Sub CHK_Click()
 Call Pic_Paint
End Sub



'Execute un mouvement entre deux position du machine
' Overtravel peremet de supprimer le test des dépassements d'axe
' Cas du chargement Tool par exemple
Function Execute_mouvement(Optional OverTravel As Boolean = False) As Boolean
Dim Increment_max As Integer
Dim Increment_controle As Integer
Dim J As Integer
Dim I As Integer
Dim Ret As Boolean

Execute_mouvement = False

   

'Voyant Orange allumée
Call Initialise_Voyant

' Controle des limites sur les axes
If OverTravel = False Then
For J = 1 To Machine.NB_axe + 1
    If Position_accordinge.Join(J) < Machine.Element(J).MiniAxe Or Position_accordinge.Join(J) > Machine.Element(J).MaxiAxe Then
        Call Affiche_Message(" Attention Axe " & J & " Hors Course")
        Exit Function
    End If
Next
End If


Increment_max = 1

'Cherche l'incrément sur les axes "Lineaire" = ceux qui provoquent les plus grands déplacements
For J = 1 To Machine.NB_axe
     If Machine.Element(J).Type_axe = 2 Then
        Increment_controle = Abs((Position_accordinge.Join(J) - Position_precedente.Join(J))) / SliderIncrement.Value
        If Increment_controle > Increment_max Then
           Increment_max = Increment_controle
        End If
     End If
Next
' Increment_max est toujours egal a un dans ce
' Cas il y a un mouvement uniquement sur des axes circulaires
If Increment_max = 1 Then
    For J = 1 To Machine.NB_axe
            Increment_controle = Abs((Position_accordinge.Join(J) - Position_precedente.Join(J))) / SliderIncrement.Value
            If Increment_controle > Increment_max Then
               Increment_max = Increment_controle
            End If
    Next
End If


For I = 0 To Increment_max - 1
    
    
    For J = 1 To Machine.NB_axe
        Machine.Element(J).Valeur_axe = Machine.Element(J).Valeur_axe + ((Position_accordinge.Join(J) - Position_precedente.Join(J))) / Increment_max
    Next
    
    
    
    
    
    'Permet d'annuler la simulation
    DoEvents
    
    Call GetPoint
    Call Pic_Paint
        
    ' calcul la collision
    If CheckBox Then
        If Collision Then
            Call Affiche_Message(" Collision")
            Call Pic_Paint
            GoTo Fin
        End If
    End If
    
Next

' Position exacte demandée
For J = 1 To Machine.NB_axe
    'Machine.Element(J).Valeur_axe = Position_accordinge.Join(J)
Next

    Call GetPoint
    Call Pic_Paint
        
    ' calcul la collision
    If CheckBox Then
        If Collision Then
            Call Affiche_Message(" Collision")
            Call Pic_Paint
            GoTo Fin
        End If
    End If

' Affiche la coordonnées obtenues
Call Affiche_coord

' Position exacte affichée
For J = 1 To Machine.NB_axe
  AxeMachine(J) = Machine.Element(J).Valeur_axe
Next

 'Voyant vert allumé
Call Allume_Voyant(1)

Execute_mouvement = True

Fin:

End Function

Sub Affiche_Message(Texte As String, Optional voyant As Integer = 2)
Dim J As Integer
        For J = 1 To 2
            PictureVoyant(J).Visible = False
        Next

        PictureVoyant(voyant).Visible = True
        MessageLOG.Cls
        MessageLOG.CurrentY = 0
        MessageLOG.Print Texte
End Sub

Private Sub CommandFixOrigine_Click()
Dim I As Integer

For I = 1 To 3

AxeOrigine(I) = AxeAbsolu(I - 1)

Next

End Sub

Private Sub CommandGoto_Click()
Dim retour As Boolean
Dim J As Integer
Dim xyzac_lu As Interpolation


If Not Chargement Then
    Call Affiche_Message("Vous devez charger le machine !!!!!")
    Exit Sub
End If

'Sauvegarde position actuelle
For J = 1 To Machine.NB_axe + 1
    Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
    
    
'Recuperation coordonees
xyzac_lu.Coord.X = Val(AxeDestination(1))
xyzac_lu.Coord.Y = Val(AxeDestination(2))
xyzac_lu.Coord.Z = Val(AxeDestination(3))
xyzac_lu.Pos.A = Val(AxeDestination(4))
xyzac_lu.Pos.C = Val(AxeDestination(5))
xyzac_lu.Pos.B = Val(AxeDestination(6))

Debug.Print UVW_Actuelle.W
' Execute deplacement
retour = Execute_Deplacement(xyzac_lu, UVW_Actuelle)

End Sub

Private Sub CommandLoadTool_Click()
  Call Charge_Tool(Val(Magasin))
  Pic_Paint
End Sub
Private Sub CommandUnloadTool_Click()
  Call Charge_Tool(0)
  Pic_Paint
End Sub

Private Sub CommandRedo_Click()
Dim J As Integer

If Not Chargement Then
Call MessageLOG_Print("You must Load a Machine !!!!!")
Exit Sub
End If

Position_accordinge = Position_precedente
For J = 1 To 6
Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
Call Execute_mouvement
    Call GetPoint
    Call Affiche_coord
    Call Sauvegarde_Position
    Position_Actuelle.Coord = Pt0
End Sub

Private Sub CommandUndo_Click()
Dim J As Integer

If Not Chargement Then
    Call MessageLOG_Print("You must load a machine !!!!!")
Exit Sub
End If

Position_accordinge = Position_precedente
For J = 1 To 6
Position_precedente.Join(J) = Machine.Element(J).Valeur_axe
Next
Call Execute_mouvement
    Call GetPoint
    Call Affiche_coord
    Call Sauvegarde_Position
    Position_Actuelle.Coord = Pt0
    
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    stop_exec = True
End If

If KeyCode = vbKeyDown Then PosX = PosX + 1
If KeyCode = vbKeyUp Then PosX = PosX - 1
If KeyCode = vbKeyRight Then PosY = PosY + 1
If KeyCode = vbKeyLeft Then PosY = PosY - 1


Call Pic_Paint

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    stop_exec = True
End If
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim SysInfo As SYSTEM_INFO
Dim strMHZ As String
Dim rc As Long

' Init pour Calcul de convertion degree radian
    RADTODEG = 180 / (4 * Atn(1))
    DEGTORAD = (4 * Atn(1)) / 180
    PI = (4 * Atn(1))
    
    ' init des voyants
    For I = 1 To 2
        PictureVoyant(I).Left = PictureVoyant(0).Left
        PictureVoyant(I).Top = PictureVoyant(0).Top
        PictureVoyant(I).Visible = False
    Next
    
    xm_base = 30
    ym_base = 0
    zm_base = 45
    
    Zoom_base = 0.001
    
    PosX_base = 0
    PosY_base = 0
    ' type d'affichage
    Render = 1
    
    ' Donne la vitesse en Mhz du processeur (WindowsXP).
    rc = GetKeyValue(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\CentralProcessor\0", "~MHz", strMHZ)

    '---------- Infos System ----------
    'SysInfo.dwProcessorType
    '   PROCESSOR_INTEL_386 = 386
    '   PROCESSOR_INTEL_486 = 486
    '   PROCESSOR_INTEL_PENTIUM = 586
    Call GetSystemInfo(SysInfo)

    ' pc pas assez performant
    If SysInfo.dwProcessorType <= 586 And Val(strMHZ) < 2096 Then
        SliderIncrement.Value = SliderIncrement.Max
    End If
    
    'Init des types d'axes
    For I = 1 To 6
        PictureAxeMachine(I).Picture = LoadResPicture(3, vbResBitmap)
        CheckBox.Value = False
    Next


    For I = 4 To 6
        SliderAxeMachine(I).Visible = False
        AxeMachine(I).Visible = False
        LabelAxeMachine(I).Visible = False
        PictureAxeMachine(I).Visible = False
    Next

   ' init treeview Tool
   Set moDragNode = Nothing
   TreeViewTOOL.Nodes.Add , tvwLast, "LIB", "Tool Lib", 1, 1
   TreeViewTOOL.Visible = True
   Programme.Visible = False
   
    'Init openGL lumiere
    ObjectAlpha = 1
    Machine_SIMUL_FRM.HScroll(0) = 650
    Machine_SIMUL_FRM.HScroll(1) = 3300
    Machine_SIMUL_FRM.HScroll(2) = -1000
    
    Vecteur_vue.X = 1


    'Previsu et chargement Machine
    PreVisu.Show

    ' Initialisation du controle picturebox en opengl
    'LoadGL Pic
    'Call Pic_DblClick
    'Call Pic_Paint


End Sub


Private Sub Form_Resize()
Dim W, H As Integer

'Rien si minimisé
If Me.WindowState = 1 Then
    Exit Sub
End If

    W = Me.ScaleWidth
    H = Me.ScaleHeight
    
    'Image
    Pic.Width = W - 30 - Programme.Width
    
    
    With Support_BTN
        .Width = W - 15
        .Top = H - .ScaleHeight - 5
        Pic.Height = H - .ScaleHeight - 20
    End With
    
    Programme.Height = Pic.Height
    ' dimenssion du treeview
    TreeViewTOOL.Height = Programme.Height
    TreeViewTOOL.Top = Programme.Top
    TreeViewTOOL.Left = Programme.Left
    TreeViewTOOL.Width = Programme.Width
    
    MessageLOG.Width = W - 590
    
    
    LoadGL Pic
    Call Pic_Paint
End Sub
'Fin
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

' Reglage position lumière
Private Sub HScroll_Change(Index As Integer)
 Light(Index).Text = HScroll(Index)
 Call Lampe
 Call Pic_Paint
End Sub

Private Sub mnuAffichageProgramme_Click()
        'rend invisible la biliotheque Tool
        ' et visible le controle programme
        TreeViewTOOL.Visible = False
        Programme.Visible = True
End Sub

Private Sub mnuCharger_Click()
Dim PtOrigine As Point3
Dim ValAxe As Double

' Ouvrir un Fichier STL
        ' Filtres.
        CMDialog1.Filter = LoadResString(102)
        ' Filtre par défaut
        CMDialog1.FilterIndex = 1
        ' Affiche le dialogue Ouvrir
        CMDialog1.ShowOpen
        FichierStl = CMDialog1.FileName
    
    'Si pas de fichier indiquée
    If Len(FichierStl) = 0 Then
        Exit Sub
    End If

    PtOrigine.X = Val(AxeOrigine(1).Text)
    PtOrigine.Y = Val(AxeOrigine(2).Text)
    PtOrigine.Z = Val(AxeOrigine(3).Text)
    ValAxe = Val(AxeOrigine(4).Text)
    
    
    'Piece
     Call Chargement_piece(FichierStl, ValAxe, PtOrigine)
    
    
    ' refresh
    Call Pic_Paint

End Sub

Private Sub mnuChargerMachine_Click()
    PreVisu.Show
End Sub

Private Sub mnuChargerTool_Click()
  Call Charge_Tool(Val(Magasin))
  Pic_Paint
End Sub

Private Sub mnuChargerOutil_Click()
On Error GoTo Fin

  Call Charge_Tool(TreeViewTOOL.SelectedItem.Index - 1)
  Pic_Paint

Exit Sub

  
Fin:
  Call Charge_Tool(Val(Magasin))
  Pic_Paint
End Sub

Private Sub mnuColor_Click()


        CMDialog1.ShowColor
        Piece.Color = GetRVB(CMDialog1.Color)
        
        Call Pic_Paint ' mise a jour couleur visu

End Sub



Private Sub mnuDeCharger_Click()
    Dim PO As Point3
    
    MessageLOG_CLS 'Réinit fentre LOG
    Call Reinit_Element(Piece)

    Piece.Type_axe = 0
    Piece.name = ""
    Piece.Valeur_axe = 0
    Piece.Origine = PO

    ReDim Piece.Obb_Col(0)
    Piece.Obb_Col(0).Pointeur_Mere = -1  ' Pointeur sur la boite mere
    Piece.Obb_Col(0).Nb_filles = 0         ' Nb_filles
    ReDim Piece.Obb_Col(0).Pointeur_filles(0)    ' Pointeur sur les boite filles (8)
    Piece.Obb_Col(0).Box_Descr.Centre = PO
    Piece.Obb_Col(0).Box_Descr.Longueurs(0) = 0 ' Longueur de la boite
    Piece.Obb_Col(0).Box_Descr.Longueurs(1) = 0
    Piece.Obb_Col(0).Box_Descr.Longueurs(2) = 0
   
    Piece.Obb_Col(0).Nb_facettes = 0      ' namebre de facettes dans la boite
    ReDim Piece.Obb_Col(0).Maillage_Box(0)    ' facettes élements de cette boite
   
    Piece_charger = False
   
   
    Call Pic_Paint
    
End Sub

Private Sub mnuDeChargerTool_Click()
  Call Charge_Tool(0)
  Pic_Paint
End Sub

Private Sub mnuDeChargerOutil_Click()
  Call Charge_Tool(0)
  Pic_Paint
End Sub

Private Sub mnuExecute_Click()
Dim J As Integer
Dim T1
Dim T2
    MessageLOG.Cls
    MessageLOG.CurrentY = 0
    
    If Not Chargement Then
        Call MessageLOG_Print("You must load a machine !!!!!")
        Exit Sub
    End If
     
    If FichierIso = "" Then
     Call MessageLOG_Print("First you have to load an Iso file !!!!!")
     Exit Sub
    End If
     
    ' pour optimisation comparaison de temp
    MessageLOG_CLS
    T1 = Time
    MessageLOG_Print ("Start times " & T1)
    Call Simul_Fichier(Programme)
    T2 = Time
    MessageLOG_Print ("End times " & T2)
    T1 = T2 - T1
    MessageLOG_Print ("Times " & Format(T1, "hh:mm:ss"))
End Sub

' Affiche des infos sur les paramètres de vue actuelle

Private Sub mnuInfoVue_Click()

    MessageLOG_CLS
    Call MessageLOG_Print("INFO VIEW")
    Call MessageLOG_Print("========")
    Call MessageLOG_Print("Zoom : " & Zoom)
    Call MessageLOG_Print("xm : " & xm)
    Call MessageLOG_Print("ym : " & ym)
    Call MessageLOG_Print("zm : " & zm)
    Call MessageLOG_Print("PosX : " & PosX)
    Call MessageLOG_Print("PosY : " & PosY)


End Sub

Private Sub mnuListeTool_Click()
        'rend visible la biliotheque Tool
        ' et invisible le controle programme
        TreeViewTOOL.Visible = True
        Programme.Visible = False
End Sub


Private Sub mnuListeOutil_Click()
        'rend invisible la biliotheque Tool
        ' et visible le controle programme
        TreeViewTOOL.Visible = True
        Programme.Visible = False
End Sub

' Efface le contenu du controle ritch text box
Private Sub mnuNouveau_Click()
'rend visible la biliotheque Tool
' et invisible le controle programme
 TreeViewTOOL.Visible = False
 Programme.Visible = True
 Programme.TextRTF = ""
 CMDialog1.FileName = "*.nc"
End Sub


Private Sub mnuOuvrirProjet_Click()
Dim fichier As String

        CMDialog1.FileName = "" 'FichierIso
        ' Filtres.
        CMDialog1.Filter = LoadResString(103)
        ' Filtre par défaut
        CMDialog1.FilterIndex = 1
      
        ' Affichage du dialogue Enregistrer
        CMDialog1.ShowOpen
        Debug.Print CMDialog1.CancelError
        fichier = CMDialog1.FileName
        '
        If fichier <> "" Then
            'Sauvegarde au format Texte
            Call Charger_Projet(fichier)
        End If
End Sub

Private Sub mnuRazLog_Click()

    MessageLOG_CLS

End Sub

Private Sub mnuRAZPO_Click()
    ReDim Parcours(0)
    Call Pic_Paint
End Sub

Private Sub mnuRenum_Click()
        'rend invisible la biliotheque Tool
        ' et visible le controle programme
        TreeViewTOOL.Visible = False
        Programme.Visible = True
        
        Call Renumerotation(Programme)
End Sub

Private Sub mnuSauver_Click()
Dim fichier As String

        CMDialog1.FileName = "" 'FichierIso
        ' Filtres.
        CMDialog1.Filter = LoadResString(101)
        ' Filtre par défaut
        CMDialog1.FilterIndex = 1
      
        ' Affichage du dialogue Enregistrer
        CMDialog1.ShowSave
        Debug.Print CMDialog1.CancelError
        fichier = CMDialog1.FileName
        '
        If fichier <> "" Then
            FichierIso = fichier
            'Sauvegarde au format Texte
            Programme.SaveFile FichierIso, rtfText
        End If
End Sub

Private Sub mnuSauverProjet_Click()
Dim fichier As String

        CMDialog1.FileName = "" 'FichierIso
        ' Filtres.
        CMDialog1.Filter = LoadResString(103)
        ' Filtre par défaut
        CMDialog1.FilterIndex = 1
      
        ' Affichage du dialogue Enregistrer
        CMDialog1.ShowSave
        Debug.Print CMDialog1.CancelError
        fichier = CMDialog1.FileName
        '
        If fichier <> "" Then
            'Sauvegarde au format Texte
            Call Sauver_Projet(fichier)
        End If
End Sub

' Permet de mettre a jour le fichier dat pour visualisation
Private Sub mnuUpdateInfoVue_Click()
Dim Ret As Integer
Dim Chaine As String

    Call mnuInfoVue_Click
    
    ' Affichage depart
    Chaine = xm
    Ret = mfncWriteIni("Machine", "Xm_base", Chaine, Fichier_Machine)
    Chaine = ym
    Ret = mfncWriteIni("Machine", "Ym_base", Chaine, Fichier_Machine)
    Chaine = zm
    Ret = mfncWriteIni("Machine", "Zm_base", Chaine, Fichier_Machine)
    Chaine = Zoom
    Ret = mfncWriteIni("Machine", "Zoom_base", Chaine, Fichier_Machine)
    Chaine = PosX
    Ret = mfncWriteIni("Machine", "PosX_base", Chaine, Fichier_Machine)
    Chaine = PosY
    Ret = mfncWriteIni("Machine", "PosY_base", Chaine, Fichier_Machine)


End Sub

' Ouvrir un Fichier ISO
Private Sub mnuOuvrir_Click()
Dim fichier As String
        'rend invisible la biliotheque Tool
        ' et visible le controle programme
        TreeViewTOOL.Visible = False
        Programme.Visible = True
        
        
        ' Filtres.
        CMDialog1.Filter = LoadResString(101)
        ' Filtre par défaut
        CMDialog1.FilterIndex = 1
        ' Affiche le dialogue Ouvrir
        CMDialog1.ShowOpen
        fichier = CMDialog1.FileName
        If fichier <> "" Then
            FichierIso = fichier
            Programme.LoadFile FichierIso
        End If
End Sub

Private Sub mnuQuiter_Click()
      Unload Me
End Sub


Private Sub mnuUpdateInfoVueAuxiliaire_Click()
Dim Ret As Integer
Dim Chaine As String

    Call mnuInfoVue_Click
    
    ' Affichage depart
    Chaine = xm
    Ret = mfncWriteIni("Machine", "Xm_Auxiliaire", Chaine, Fichier_Machine)
    Chaine = ym
    Ret = mfncWriteIni("Machine", "Ym_Auxiliaire", Chaine, Fichier_Machine)
    Chaine = zm
    Ret = mfncWriteIni("Machine", "Zm_Auxiliaire", Chaine, Fichier_Machine)
    Chaine = Zoom
    Ret = mfncWriteIni("Machine", "Zoom_Auxiliaire", Chaine, Fichier_Machine)
    Chaine = PosX
    Ret = mfncWriteIni("Machine", "PosX_Auxiliaire", Chaine, Fichier_Machine)
    Chaine = PosY
    Ret = mfncWriteIni("Machine", "PosY_Auxiliaire", Chaine, Fichier_Machine)
End Sub

Private Sub mnuVersion_Click()
 frmAbout.Show
End Sub

Private Sub mnuVueAuxilliaire_Click()
    xm = Val(mfncGetFromIni("Machine", "Xm_Auxiliaire", Fichier_Machine))
    ym = Val(mfncGetFromIni("Machine", "Ym_Auxiliaire", Fichier_Machine))
    zm = Val(mfncGetFromIni("Machine", "Zm_Auxiliaire", Fichier_Machine))
    Zoom = Val(mfncGetFromIni("Machine", "Zoom_Auxiliaire", Fichier_Machine))
    PosX = Val(mfncGetFromIni("Machine", "PosX_Auxiliaire", Fichier_Machine))
    PosY = Val(mfncGetFromIni("Machine", "PosY_Auxiliaire", Fichier_Machine))
    Call Pic_Paint
End Sub

Private Sub mnuVueBase_Click()
      ' Affichage depart
    xm = Val(mfncGetFromIni("Machine", "Xm_base", Fichier_Machine))
    ym = Val(mfncGetFromIni("Machine", "Ym_base", Fichier_Machine))
    zm = Val(mfncGetFromIni("Machine", "Zm_base", Fichier_Machine))
    Zoom = Val(mfncGetFromIni("Machine", "Zoom_base", Fichier_Machine))
    PosX = Val(mfncGetFromIni("Machine", "PosX_base", Fichier_Machine))
    PosY = Val(mfncGetFromIni("Machine", "PosY_base", Fichier_Machine))
    Call Pic_Paint
End Sub

Private Sub mnuVueDessus_Click()
    xm = 90
    ym = 0
    zm = 0
    Zoom = Val(mfncGetFromIni("Machine", "Zoom_base", Fichier_Machine))
    'PosX = 0
    'PosY = 0
    Call Pic_Paint
End Sub

' Juste un Choix du type d'affichage
Private Sub OPT_Click(Index As Integer)
     'trait caché
    If Abs(CInt(OPT(0).Value)) = 1 Then
        Render = 1
    End If
    'Fil de fer
    If Abs(CInt(OPT(1).Value)) = 1 Then
        Render = 2
    End If
    'ombrée
    If Abs(CInt(OPT(2).Value)) = 1 Then
        Render = 3
    End If
    
    'Remettre a jour l'affichage
     Call Pic_Paint
    
End Sub
'Affiche le parcours au Point Pivot
Private Sub OptionTracerParcours_click()
    'Remettre a jour l'affichage
    Call Pic_Paint
End Sub

'RAZ de la visu de la pièce
Private Sub Pic_DblClick()
    xm = xm_base '30
    ym = ym_base '0
    zm = zm_base '45
    
    Zoom = Zoom_base '0.001
    
    PosX = PosX_base '0
    PosY = PosY_base '0
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim diffx As Single
Dim diffy As Single
Dim facteur_diff
Dim facteur_zoom

facteur_diff = 20
facteur_zoom = 0.0001

'Translation
If Button = 1 And Shift = 0 Then
' Recherche la direction principale
  diffx = Abs(SVGposX - X)
  diffy = Abs(SVGposY - Y)
  
  If diffx > diffy Then
       If X > SVGposX Then
       PosY = PosY + facteur_diff
       Else
       PosY = PosY - facteur_diff
      End If
  Else
       If Y > SVGposY Then
       PosX = PosX - facteur_diff
      Else
       PosX = PosX + facteur_diff
      End If
  End If
End If

'  Zoom global
If Button = 2 And Shift = 0 Then
 If Y < SVGposY Then
 Zoom = Zoom - facteur_zoom
 Else
 Zoom = Zoom + facteur_zoom
 End If
 
 If Zoom < 0 Then Zoom = 0.0001
 'Zoom = Y / 10000
End If

'  Zoom fenetre
If Button = 3 And Shift = 0 Then

End If

'rotation
If Button = 1 And Shift = 1 Then
' Recherche la direction principale
  diffx = Abs(SVGposX - X)
  diffy = Abs(SVGposY - Y)
  
  If diffx > diffy Then
       If X > SVGposX Then
       ym = ym - X * 0.005
       Else
       ym = ym + X * 0.005
       End If
  Else
       If Y > SVGposY Then
       xm = xm - Y * 0.005
       Else
       xm = xm + Y * 0.005
       End If
  End If
End If

If Button = 2 And Shift = 1 Then
       If X > SVGposX Then
       zm = zm - X * 0.005
       Else
       zm = zm + X * 0.005
       End If
End If

If Button > 0 Then
SVGposX = X
SVGposY = Y
End If


Call Pic_Paint

End Sub

' Mise à jour de l'affichage
Public Sub Pic_Paint()
    Call DessineMachine(Pic, CBool(CHK.Value), Render, OptionTracerParcours, CheckBox)
End Sub

Private Sub PictureAxemachine_dblClick(Index As Integer)
Dim Cancel As Boolean
    AxeMachine(Index) = 0
    Call Axemachine_Validate(Index, Cancel)
End Sub

Private Sub Magasin_Change()

If Val(Magasin) = 0 Then
Magasin = 1
End If

If Val(Magasin) > 8 Then
Magasin = 8
End If

If Chargement Then
    Call Rotation_Magasin_Tool(Val(Magasin))
End If

End Sub



Private Sub sldAlpha_Click()
    ObjectAlpha = (100 - sldAlpha.Value) / 100 'Modif du  alpha level
    Call Pic_Paint

End Sub

Private Sub SliderAxemachine_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Cancel As Boolean
AxeMachine(Index) = SliderAxeMachine(Index).Value
Call Axemachine_Validate(Index, Cancel)
End Sub

Private Sub Axemachine_Validate(Index As Integer, Cancel As Boolean)
Dim I As Integer

   MessageLOG.Print "Validate " & Index
    ' force un namebre
   AxeMachine(Index) = Val(AxeMachine(Index))

If Chargement Then

    For I = 1 To Machine.NB_axe
       Position_precedente.Join(I) = Machine.Element(I).Valeur_axe
    Next
    
    
    If Val(AxeMachine(Index)) < SliderAxeMachine(Index).Min Then
        AxeMachine(Index) = SliderAxeMachine(Index).Min
        Call MessageLOG_Print("Axis " & Index & " < limite " & SliderAxeMachine(Index).Min)
    End If
    
    If Val(AxeMachine(Index)) > SliderAxeMachine(Index).Max Then
        AxeMachine(Index) = SliderAxeMachine(Index).Max
        Call MessageLOG_Print("Axis " & Index & " > limite " & SliderAxeMachine(Index).Max)
    End If
    
    
    
    For I = 1 To Machine.NB_axe '+ 1
        Position_accordinge.Join(I) = AxeMachine(I)
    Next
    
    
    
    Call Execute_mouvement

    

    'Machine.Element(Index).Valeur_axe = Val(Axemachine(Index))
    'Recuperation de position via OpenGl fonction
    Call GetPoint
    Call Affiche_coord
    Call Pic_Paint
    Call Sauvegarde_Position
    Position_Actuelle.Coord = Pt0
    
End If

End Sub



Private Sub SliderMagasin_Click()
Magasin = SliderMagasin.Value
End Sub



Private Sub TreeViewTOOL_DblClick()
  Call Charge_Tool(TreeViewTOOL.SelectedItem.Index - 1)
  Pic_Paint
End Sub

Private Sub TreeViewTOOL_NodeClick(ByVal Node As MSComctlLib.Node)
    Magasin.Text = Node.Index - 1
End Sub

Private Sub VueFix_Click(Index As Integer)
Select Case Index

Case 0 ' vue X
VueFix(1).Value = 0
Case 1 'vue Y
VueFix(0).Value = 0
End Select

Call Pic_Paint
End Sub

'Definition vecteur de vuE machine
Private Sub VueMachine_Click(Index As Integer)
Select Case Index
    Case 0 ' vue X
        Vecteur_vue.X = 1
        Vecteur_vue.Y = 0
        Vecteur_vue.Z = 0
    
    Case 1 ' vue Y
        Vecteur_vue.X = 0
        Vecteur_vue.Y = 1
        Vecteur_vue.Z = 0
    
    Case 2 ' vue -X
        Vecteur_vue.X = -1
        Vecteur_vue.Y = 0
        Vecteur_vue.Z = 0
    
    Case 3 ' vue -Y
        Vecteur_vue.X = 0
        Vecteur_vue.Y = -1
        Vecteur_vue.Z = 0
End Select

Call Pic_Paint
End Sub

Private Sub Zgrille_Change()
    Call Def_Plan(Val(Zgrille))
End Sub
