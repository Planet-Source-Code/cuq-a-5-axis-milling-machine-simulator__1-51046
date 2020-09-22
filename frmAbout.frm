VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About the application"
   ClientHeight    =   4485
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7005
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3097.725
   ScaleMode       =   0  'User
   ScaleWidth      =   6567.34
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   600
      Left            =   360
      Picture         =   "frmAbout.frx":030A
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label Warranty 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmAbout.frx":07DC
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   5055
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.503
      X2              =   6465.15
      Y1              =   497.294
      Y2              =   497.294
   End
   Begin VB.Label lblTitle 
      Caption         =   "Titre de l'application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3285
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   6338.584
      Y1              =   497.294
      Y2              =   497.294
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   2532
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = LoadResString(501) & " " & App.Title
    lblVersion.Caption = LoadResString(502) & " " & App.Major & "." & App.Minor & "." & App.Revision & " " '& LoadResString(100)
    lblTitle.Caption = App.Title
    'image1.picture = LoadResPicture(103, vbResBitmap)
End Sub







