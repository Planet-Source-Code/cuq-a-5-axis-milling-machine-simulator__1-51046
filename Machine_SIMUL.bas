Attribute VB_Name = "Machine_SIMUL"
Option Explicit

' Stock du name du fichier ISO
Public FichierIso As String
' Stock du name du fichier PIECE STL
Public FichierStl As String

Sub Affiche_Message(Texte As String, Optional voyant As Integer = 2)
Dim J As Integer
        For J = 1 To 2
            Machine_SIMUL_FRM.PictureVoyant(J).Visible = False
        Next

        Machine_SIMUL_FRM.PictureVoyant(voyant).Visible = True
        Machine_SIMUL_FRM.MessageLOG.Cls
        Machine_SIMUL_FRM.MessageLOG.CurrentY = 0
        Machine_SIMUL_FRM.MessageLOG.Print Texte
End Sub

' Efface et réinitialise la fenêtre de sortie des infos
Sub MessageLOG_CLS()

    Machine_SIMUL_FRM.MessageLOG.Cls
    Machine_SIMUL_FRM.MessageLOG.CurrentY = 0
    
End Sub
' Ecrit dans la fenetre les infos et active un voyant
Sub MessageLOG_Print(Chaine As String, Optional voyant As Integer)

    If Machine_SIMUL_FRM.MessageLOG.CurrentY >= Machine_SIMUL_FRM.MessageLOG.ScaleHeight Then
        Call MessageLOG_CLS
    End If
    
    Machine_SIMUL_FRM.MessageLOG.Print Chaine

    
    ' Allumage d'un voyant
    If voyant Then
       Machine_SIMUL_FRM.PictureVoyant(voyant).Visible = True
    End If

End Sub

'Initialise les voyants
' 0 = Orange
' 1 = Vert
' 2 = Rouge
Sub Initialise_Voyant()
Dim J As Integer
    
    For J = 1 To 2
    Machine_SIMUL_FRM.PictureVoyant(J).Visible = False
    Next
    
End Sub

Sub Allume_Voyant(voyant As Integer)

    Machine_SIMUL_FRM.PictureVoyant(voyant).Visible = True

End Sub

'Execute un mouvement entre deux position XYZ AC demandé
' Overtravel peremet de supprimer le test des dépassements d'axe
' Cas du chargement Tool par exemple
Function Execute_Deplacement(XYZ_ABC As Interpolation, Uvw As AxePositionne, Optional OverTravel As Boolean = False) As Boolean
Dim Increment_max As Integer
Dim Increment_controle As Integer
Dim J As Integer
Dim i As Integer
Dim Ret As Boolean

Dim Xyzabc_Memoire As Interpolation
Dim Uvw_Memoire As AxePositionne

Dim xyzac_lu As Interpolation
Dim Uvw_lu As AxePositionne
Dim xyzac_Int As Interpolation
Dim Uvw_Int As AxePositionne

Dim Position_int As position
Dim Position_prec As position
Dim Position_suite As position


Execute_Deplacement = False
'Voyant Orange allumée
Call Initialise_Voyant

'Recalcul avec origine programme
' Decalage fixé dans la fenetre FRM_simul
'Xyzabc_Memoire = TransCoord(Position_Actuelle, OrigineProg)

Xyzabc_Memoire = Position_Actuelle
Uvw_Memoire = UVW_Actuelle
xyzac_lu = XYZ_ABC
Uvw_lu = Uvw

'Calcul selon la position demandée les valeurs des axes machines
'Sauvegarde position actuelle
For J = 1 To Machine.NB_axe + 1
    Position_prec.Join(J) = Machine.Element(J).Valeur_axe
Next
    Position_precedente = Position_prec
' Une fois la position a atteindre calculée
' Calcul Des axes machines
' et affectation de Position_accordinge
Position_suite = Position_Machine(xyzac_lu, Uvw_lu)



' Controle des limites sur les axes ( chargement Tool fait sans controle de limite )
If OverTravel = False Then
For J = 1 To Machine.NB_axe + 1
    If Position_suite.Join(J) < Machine.Element(J).MiniAxe Or Position_suite.Join(J) > Machine.Element(J).MaxiAxe Then
        Call Affiche_Message(" Attention Axis " & J & " Out of range")
        Call MessageLOG_Print(" Axis " & Machine.Element(J).name & " Limited between " & Machine.Element(J).MiniAxe & " and " & Machine.Element(J).MaxiAxe)
        Call MessageLOG_Print(" Requested Value " & Format(Position_suite.Join(J), "# ###.###"))
        Exit Function
    End If
Next
End If


Increment_max = 1

'Cherche l'incrément sur les axes "Lineaire" = ceux qui provoquent les plus grands déplacements
For J = 1 To Machine.NB_axe
     If Machine.Element(J).Type_axe = 2 Then
        Increment_controle = Abs((Position_suite.Join(J) - Position_prec.Join(J))) / Machine_SIMUL_FRM.SliderIncrement.Value
        If Increment_controle > Increment_max Then
           Increment_max = Increment_controle
        End If
     End If
Next
' Increment_max est toujours egal a un dans ce
' Cas il y a un mouvement uniquement sur des axes circulaires
If Increment_max = 1 Then
    For J = 1 To Machine.NB_axe
            Increment_controle = Abs((Position_suite.Join(J) - Position_prec.Join(J))) / Machine_SIMUL_FRM.SliderIncrement.Value
            If Increment_controle > Increment_max Then
               Increment_max = Increment_controle
            End If
    Next
End If


For i = 1 To Increment_max '- 1
    xyzac_Int.Coord.X = Xyzabc_Memoire.Coord.X + i * ((XYZ_ABC.Coord.X - Xyzabc_Memoire.Coord.X) / Increment_max)
    xyzac_Int.Coord.Y = Xyzabc_Memoire.Coord.Y + i * ((XYZ_ABC.Coord.Y - Xyzabc_Memoire.Coord.Y) / Increment_max)
    xyzac_Int.Coord.Z = Xyzabc_Memoire.Coord.Z + i * ((XYZ_ABC.Coord.Z - Xyzabc_Memoire.Coord.Z) / Increment_max)
    xyzac_Int.Pos.A = Xyzabc_Memoire.Pos.A + i * ((XYZ_ABC.Pos.A - Xyzabc_Memoire.Pos.A) / Increment_max)
    xyzac_Int.Pos.B = Xyzabc_Memoire.Pos.B + i * ((XYZ_ABC.Pos.B - Xyzabc_Memoire.Pos.B) / Increment_max)
    xyzac_Int.Pos.C = Xyzabc_Memoire.Pos.C + i * ((XYZ_ABC.Pos.C - Xyzabc_Memoire.Pos.C) / Increment_max)
    
    Uvw_Int.U = Uvw_Memoire.U + i * ((Uvw.U - Uvw_Memoire.U) / Increment_max)
    Uvw_Int.V = Uvw_Memoire.V + i * ((Uvw.V - Uvw_Memoire.V) / Increment_max)
    Uvw_Int.W = Uvw_Memoire.W + i * ((Uvw.W - Uvw_Memoire.W) / Increment_max)

    ' Une fois la position a atteindre calculée
    ' Calcul Des axes machines
    ' et affectation de Position_accordinge
    Position_int = Position_Machine(xyzac_Int, Uvw_Int)
    

    For J = 1 To Machine.NB_axe
        Machine.Element(J).Valeur_axe = Position_int.Join(J)
    Next
    

    Call Machine_SIMUL_FRM.Pic_Paint

    
' calcul la collision
    If Machine_SIMUL_FRM.CheckBox Then
        If Collision Then
            ' Affiche la coordonnées obtenues
            Call Affiche_coord
            Call Affiche_Message(" Collision")
            Call Machine_SIMUL_FRM.Pic_Paint
            Position_accordinge = Position_int
            GoTo Fin
        End If
    End If
    
Next

    'Call GetPoint
    ' Sauvegared position
    'Call Sauvegarde_Position
    'Position_Actuelle = xyzac_Int
    'Permet d'annuler la simulation
    'DoEvents
    

    ' Position exacte demandée
    For J = 1 To Machine.NB_axe
         Machine.Element(J).Valeur_axe = Position_suite.Join(J)
    Next
    
    Position_accordinge = Position_suite


    For J = 1 To Machine.NB_axe
        Machine.Element(J).Valeur_axe = Position_int.Join(J)
    Next
    
    Position_accordinge = Position_int
    
    ' calcul la collision
    If Machine_SIMUL_FRM.CheckBox Then
        If Collision Then
            Call Affiche_Message(" Collision")
            Call Machine_SIMUL_FRM.Pic_Paint
            GoTo Fin
        End If
    End If






 'Voyant vert allumé
Call Allume_Voyant(1)



Execute_Deplacement = True

Fin:

Call GetPoint
Call Machine_SIMUL_FRM.Pic_Paint

' Sauvegared position
Call Sauvegarde_Position
Position_Actuelle = XYZ_ABC
'Permet d'annuler la simulation
DoEvents
    
' Affiche la coordonnées obtenues
Call Affiche_coord
    
' Affiche la Position obtenues
For J = 1 To Machine.NB_axe
        Machine_SIMUL_FRM.AxeMachine(J).Text = Machine.Element(J).Valeur_axe
Next
    
If Machine_SIMUL_FRM.OptionRTCP.Value Then
    XYZ_ABC = xyzac_lu
Else
    XYZ_ABC.Coord = Pt0
End If

End Function

Sub Sauvegarde_Position()
Dim i As Integer
Dim Element As Element3D

For i = 1 To UBound(Machine.Element)
    Element = Machine.Element(i)
    ' Sauvegarde position
    Select Case Element.name
        Case "U"
              UVW_Actuelle.U = Element.Valeur_axe
        Case "V"
              UVW_Actuelle.V = Element.Valeur_axe
        Case "W"
              UVW_Actuelle.W = Element.Valeur_axe
        Case "A"
              Position_Actuelle.Pos.A = Element.Valeur_axe
        Case "B"
              Position_Actuelle.Pos.B = Element.Valeur_axe
        Case "C"
              Position_Actuelle.Pos.C = Element.Valeur_axe
    End Select
Next i


End Sub

' Sauvegarde un projet
' Machine
' Piece
' Fichier Iso
' Reglages
' Position Machine
Sub Sauver_Projet(fichier As String)
Dim Ret As Integer
Dim Chaine As String
Dim i As Integer


    'Machine
    Ret = mfncWriteIni("Machine", "Fichier_Machine", Fichier_Machine, fichier)
    
    Chaine = Machine.Tool_current
    Ret = mfncWriteIni("Machine", "Tool_current", Chaine, fichier)
    
    Chaine = ""
    For i = 1 To Machine.NB_axe
    Chaine = Chaine & Machine.Element(i).Valeur_axe & ","
    Next i
    Ret = mfncWriteIni("Machine", "Position_Actuelle", Mid(Chaine, 1, Len(Chaine) - 1), fichier)

    
    ' Affichage depart
    Chaine = xm
    Ret = mfncWriteIni("Machine_Simu", "Xm", Chaine, fichier)
    Chaine = ym
    Ret = mfncWriteIni("Machine_Simu", "Ym", Chaine, fichier)
    Chaine = zm
    Ret = mfncWriteIni("Machine_Simu", "Zm", Chaine, fichier)
    Chaine = Zoom
    Ret = mfncWriteIni("Machine_Simu", "Zoom", Chaine, fichier)
    Chaine = PosX
    Ret = mfncWriteIni("Machine_Simu", "PosX", Chaine, fichier)
    Chaine = PosY
    Ret = mfncWriteIni("Machine_Simu", "PosY", Chaine, fichier)
    
    Chaine = Machine_SIMUL_FRM.OptionTracerParcours.Value
    Ret = mfncWriteIni("Machine_Simu", "TracerParcours", Chaine, fichier)
    
    Chaine = Machine_SIMUL_FRM.SliderIncrement.Value
    Ret = mfncWriteIni("Machine_Simu", "Increment", Chaine, fichier)
    
    'Lumiere
    Chaine = Machine_SIMUL_FRM.sldAlpha.Value
    Ret = mfncWriteIni("Machine_Simu", "Transparence", Chaine, fichier)
    
    For i = 0 To 2
        Chaine = "Light" & i
        Ret = mfncWriteIni("Machine_Simu", Chaine, Machine_SIMUL_FRM.Light(i).Text, fichier)
    Next i

    
    Chaine = Machine_SIMUL_FRM.CHK.Value
    Ret = mfncWriteIni("Machine_Simu", "Grille", Chaine, fichier)
    Ret = mfncWriteIni("Machine_Simu", "ZGrille", Machine_SIMUL_FRM.Zgrille.Text, fichier)
    
    Chaine = Machine_SIMUL_FRM.OptionRTCP.Value
    Ret = mfncWriteIni("Machine_Simu", "RTCP", Chaine, fichier)
      
    Chaine = Machine_SIMUL_FRM.CheckBox.Value
    Ret = mfncWriteIni("Machine_Simu", "Collision", Chaine, fichier)
    
    
    For i = 1 To 4
        Chaine = "AxeOrigine" & i
        Ret = mfncWriteIni("Machine_Simu", Chaine, Machine_SIMUL_FRM.AxeOrigine(i).Text, fichier)
    Next i
    
    ' Piece
    
    Ret = mfncWriteIni("Piece", "Fichier_Piece", FichierStl, fichier)
    Chaine = Piece.Color.Rouge
    Ret = mfncWriteIni("Piece", "Couleur_Piece_Rouge", Chaine, fichier)
    Chaine = Piece.Color.Vert
    Ret = mfncWriteIni("Piece", "Couleur_Piece_Vert", Chaine, fichier)
    Chaine = Piece.Color.Bleu
    Ret = mfncWriteIni("Piece", "Couleur_Piece_Bleu", Chaine, fichier)
    
    'Fichier Iso
    Ret = mfncWriteIni("Iso", "Fichier_Iso", FichierIso, fichier)
    
End Sub

' Charger un projet
' Machine
' Piece
' Fichier Iso
' Reglages
' Position Machine
Sub Charger_Projet(fichier As String)
Dim Chaine As String
Dim i As Integer
Dim PtOrigine As Point3
Dim ValAxe As Double
Dim TableauPosition

'desactif la controle de collision pour effectuer les chragement machine et autre
Machine_SIMUL_FRM.CheckBox.Value = False
    
    Machine_SIMUL_FRM.CHK.Value = Val(mfncGetFromIni("Machine_Simu", "Grille", fichier))
    Machine_SIMUL_FRM.Zgrille.Text = mfncGetFromIni("Machine_Simu", "ZGrille", fichier)
    
    Machine_SIMUL_FRM.OptionRTCP.Value = Val(mfncGetFromIni("Machine_Simu", "RTCP", fichier))

    
    Machine_SIMUL_FRM.OptionTracerParcours.Value = Val(mfncGetFromIni("Machine_Simu", "TracerParcours", fichier))
    
    Machine_SIMUL_FRM.SliderIncrement.Value = Val(mfncGetFromIni("Machine_Simu", "Increment", fichier))
    
    'Lumiere
    Machine_SIMUL_FRM.sldAlpha.Value = Val(mfncGetFromIni("Machine_Simu", "Transparence", fichier))
    
    For i = 0 To 2
        Chaine = "Light" & i
        Machine_SIMUL_FRM.Light(i).Text = mfncGetFromIni("Machine_Simu", Chaine, fichier)
        Machine_SIMUL_FRM.HScroll(i).Value = Val(mfncGetFromIni("Machine_Simu", Chaine, fichier))
    Next i
    
    Call Lampe
    
    'origine
    For i = 1 To 4
        Chaine = "AxeOrigine" & i
        Machine_SIMUL_FRM.AxeOrigine(i).Text = mfncGetFromIni("Machine_Simu", Chaine, fichier)
    Next i
    
   'Origine programme
    PtOrigine.X = Val(Machine_SIMUL_FRM.AxeOrigine(1).Text)
    PtOrigine.Y = Val(Machine_SIMUL_FRM.AxeOrigine(2).Text)
    PtOrigine.Z = Val(Machine_SIMUL_FRM.AxeOrigine(3).Text)
    ValAxe = Val(Machine_SIMUL_FRM.AxeOrigine(4).Text)
    
    OrigineProg.X = PtOrigine.X
    OrigineProg.Y = PtOrigine.Y
    OrigineProg.Z = PtOrigine.Z
    OrigineProg.ED = ValAxe
    
    'machine
    Fichier_Machine = mfncGetFromIni("Machine", "Fichier_Machine", fichier)
    'Debug.Print TokRightLeft(Fichier_Machine, "\")
    Call Chargement_machine(TokRightRight(TokRightLeft(Fichier_Machine, "\"), "\"), Fichier_Machine)
    
    Call Charge_Tool(Val(mfncGetFromIni("Machine", "Tool_current", fichier)))
    'Machine.Tool_current = Val(mfncGetFromIni("Machine", "Tool_current", fichier))
       
    
    
    'Debug.Print mfncGetFromIni("Machine", "Position_Actuelle", fichier)
    TableauPosition = Split(mfncGetFromIni("Machine", "Position_Actuelle", fichier), ",")
    
    For i = 1 To UBound(TableauPosition)
     'Machine.Element(I).Valeur_axe = Val(TableauPosition(I))
     Position_accordinge.Join(i) = Val(TableauPosition(i))
    Next i
   
     
   Call Machine_SIMUL_FRM.Execute_mouvement
   
    xm = Val(mfncGetFromIni("Machine_Simu", "Xm", fichier))
    ym = Val(mfncGetFromIni("Machine_Simu", "Ym", fichier))
    zm = Val(mfncGetFromIni("Machine_Simu", "Zm", fichier))
    Zoom = Val(mfncGetFromIni("Machine_Simu", "Zoom", fichier))
    PosX = Val(mfncGetFromIni("Machine_Simu", "PosX", fichier))
    PosY = Val(mfncGetFromIni("Machine_Simu", "PosY", fichier))
     
     
   ' active ou non le controle de collision après le chargement Tool
   Machine_SIMUL_FRM.CheckBox.Value = Val(mfncGetFromIni("Machine_Simu", "Collision", fichier))
   
   

    
    'Piece
        ' Piece
        Chaine = mfncGetFromIni("Piece", "Fichier_Piece", fichier)
        If Chaine <> "" Then
            FichierStl = Chaine
            Call Chargement_piece(FichierStl, ValAxe, PtOrigine)
            Piece.Color.Rouge = Val(mfncGetFromIni("Piece", "Couleur_Piece_Rouge", fichier))
            Piece.Color.Vert = Val(mfncGetFromIni("Piece", "Couleur_Piece_Vert", fichier))
            Piece.Color.Bleu = Val(mfncGetFromIni("Piece", "Couleur_Piece_Bleu", fichier))
        End If
        
    'Fichier ISO
        Chaine = mfncGetFromIni("Iso", "Fichier_Iso", fichier)
        If Chaine <> "" Then
        'rend invisible la biliotheque Tool
        ' et visible le controle programme
        Machine_SIMUL_FRM.TreeViewTOOL.Visible = False
        Machine_SIMUL_FRM.Programme.Visible = True
        
            FichierIso = Chaine
            Machine_SIMUL_FRM.Programme.LoadFile FichierIso
        End If
        
    
    Call Machine_SIMUL_FRM.Pic_Paint
    
    
End Sub

Sub Chargement_piece(fichier As String, Valeur_axe As Double, OriginePt As Point3)
Dim Box0 As Box3
Dim Elem_tempo As Element3D
Dim Obb_Box_tempo As OBB_box
Dim i As Integer
Dim Mere As Integer

    MessageLOG_CLS 'Réinit fentre LOG

    
    Call Reinit_Element(Piece)
    Piece.name = ChargeFichierSTL(fichier, Piece.STL_def)
    Piece.Type_axe = 5
    Piece.Origine = OriginePt
    
    Piece.Valeur_axe = Valeur_axe
    Piece.Vecteur = Machine.Element(0).Vecteur ' La piece a le même vecteur que l'element O
    ' element 0 de la machine = plateau
    ' test si le veceteur a une orientation
    If Longueur(Piece.Vecteur) = 0 Then
        Piece.Vecteur.Z = 1
    End If
    
    Box0 = Def_Box(Piece.STL_def)
    ReDim Piece.Obb_Col(1)
    Piece.Obb_Col(0).Pointeur_Mere = -1
    Piece.Obb_Col(0).Box_Descr = Box0
    Piece.Box_Col = Box0
    ' Stock pour la durée des modif les infos
    Obb_Box_tempo = Piece.Obb_Col(0)
    Call Creer_Obb_Box(Obb_Box_tempo, Piece, False, 0)
    'réintègre les infos à leur place
    Piece.Obb_Col(0) = Obb_Box_tempo
    'Deuxieme niveau
    For i = 1 To 8
        Obb_Box_tempo = Piece.Obb_Col(i)
        Mere = i
        Call Creer_Obb_Box(Obb_Box_tempo, Piece, True, Mere)
        Piece.Obb_Col(i) = Obb_Box_tempo
    Next i
    
    Piece_charger = True
End Sub

Sub Chargement_machine(Repertoire As String, Fichier_Def_machine As String)
Dim J As Integer

'reinit les labels
    For J = 1 To 6
        Machine_SIMUL_FRM.AxeDestination(J).Text = "0"
        Machine_SIMUL_FRM.AxeMachine(J).Text = "0"
        Position_accordinge.Join(J) = 0
        Position_precedente.Join(J) = 0
    Next
    
    
' rend invisible les labels axes avant chargement
    For J = 1 To 6
        Machine_SIMUL_FRM.AxeDestination(J).Visible = False
        Machine_SIMUL_FRM.LabelAxeDestination(J).Visible = False
    Next
    
    
   'Charge le Machine
    Call MessageLOG_CLS

    LoadGL Machine_SIMUL_FRM.Pic


    Call Machine_SIMUL_FRM.Init_Machine(Repertoire, Fichier_Def_machine)
    
    Call DessineMachine(Machine_SIMUL_FRM.Pic, False, 1, False, True)
    Call GetPoint
    
    Call Affiche_coord
    Call Sauvegarde_Position
    Position_Actuelle.Coord = Pt0
    
    
 Call Initialise_Voyant
    
' rend visible les labels axes selon machine

For J = 1 To Machine.NB_axe
    Select Case Machine.Element(J).name
    
    Case "X"
        Machine_SIMUL_FRM.AxeDestination(1).Visible = True
        Machine_SIMUL_FRM.LabelAxeDestination(1).Visible = True
        
    Case "Y"
        Machine_SIMUL_FRM.AxeDestination(2).Visible = True
        Machine_SIMUL_FRM.LabelAxeDestination(2).Visible = True
        
    Case "Z"
        Machine_SIMUL_FRM.AxeDestination(3).Visible = True
        Machine_SIMUL_FRM.LabelAxeDestination(3).Visible = True
    
    Case "A"
        Machine_SIMUL_FRM.AxeDestination(4).Visible = True
        Machine_SIMUL_FRM.LabelAxeDestination(4).Visible = True
       
    Case "C"
        Machine_SIMUL_FRM.AxeDestination(5).Visible = True
        Machine_SIMUL_FRM.LabelAxeDestination(5).Visible = True
        
    Case "B"
        Machine_SIMUL_FRM.AxeDestination(6).Visible = True
        Machine_SIMUL_FRM.LabelAxeDestination(6).Visible = True
        
    Case Else
    End Select
Next J

'
ReDim Parcours(0)
End Sub
