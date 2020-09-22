VERSION 5.00
Begin VB.Form PreVisu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine PreLoad"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
   FillColor       =   &H00808080&
   ForeColor       =   &H00808080&
   Icon            =   "PreVisu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   4800
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2220
      Left            =   120
      TabIndex        =   1
      Top             =   3840
      Width           =   4575
   End
   Begin VB.PictureBox PicPrevisu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   305
      TabIndex        =   0
      Top             =   120
      Width           =   4605
   End
End
Attribute VB_Name = "PreVisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Dim Machine_tempo As Machine3D
' position fenetre toujours dessus
Private Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Load()
    'Me.Show
    'PreVisu.SetFocus
     Call MessageLOG_CLS


    SetWindowPos PreVisu.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Me.Top = Machine_SIMUL_FRM.Top + 150
     Me.Left = Machine_SIMUL_FRM.Left + Machine_SIMUL_FRM.Width / 3
       

    ' permet d'initialiser la grille
    xm = 45
ym = 0
zm = 45

Zoom = 0.0018

PosX = 0
PosY = 0

' Initialisation du controle picturebox en opengl
LoadGL Me.PicPrevisu
    
Call PrevisuMachine(Me.PicPrevisu, Machine_tempo)
     
Fichier_Ini = App.Path + "\Machine_simul.ini"

Call Init_Liste_Machine(Fichier_Ini)
'init sur premier Machine par defaut
If Fichier_Machine = "" Then
    Indice_Machine = "Machine1"
    Fichier_Machine = App.Path + "\Machine_def\" + mfncGetFromIni(Indice_Machine, "Repertoire", Fichier_Ini) + "\" + mfncGetFromIni(Indice_Machine, "Fichier", Fichier_Ini)
End If

End Sub

Sub Init_Liste_Machine(fichier As String)
Dim nb_Machine As Integer
Dim i As Integer
Dim name_Item As String

    nb_Machine = Val(mfncGetFromIni("MACHINE_SIMUL", "NB_Machine", fichier))

List1.Clear

For i = 1 To nb_Machine
    name_Item = "Machine" & i
    List1.AddItem (mfncGetFromIni(name_Item, "name", fichier))
Next i


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Chargement_machine(mfncGetFromIni(Indice_Machine, "Repertoire", Fichier_Ini), Fichier_Machine)
End Sub
' lance la previsualisation de la machine
Private Sub List1_Click()
Dim name_Item As String
Dim i As Integer

i = List1.ListIndex + 1
name_Item = "Machine" & i

Indice_Machine = name_Item
Fichier_Machine = App.Path + "\Machine_def\" + mfncGetFromIni(name_Item, "Repertoire", Fichier_Ini) + "\" + mfncGetFromIni(name_Item, "Fichier", Fichier_Ini)
' Affichage depart
xm_base = Val(mfncGetFromIni("Machine", "Xm_base", Fichier_Machine))
ym_base = Val(mfncGetFromIni("Machine", "Ym_base", Fichier_Machine))
zm_base = Val(mfncGetFromIni("Machine", "Zm_base", Fichier_Machine))
    
Zoom_base = Val(mfncGetFromIni("Machine", "Zoom_base", Fichier_Machine))
    
PosX_base = Val(mfncGetFromIni("Machine", "PosX_base", Fichier_Machine))
PosY_base = Val(mfncGetFromIni("Machine", "PosY_base", Fichier_Machine))

xm = xm_base
ym = ym_base
zm = zm_base

Zoom = Zoom_base

PosX = PosX_base
PosY = PosY_base

   
Call Init_Previsu_Machine(mfncGetFromIni(Indice_Machine, "Repertoire", Fichier_Ini), Fichier_Machine, Machine_tempo)
Call PrevisuMachine(PicPrevisu, Machine_tempo)
    
End Sub

Private Sub PicPrevisu_Paint()
        LoadGL PicPrevisu
        Call PrevisuMachine(PicPrevisu, Machine_tempo)
End Sub
