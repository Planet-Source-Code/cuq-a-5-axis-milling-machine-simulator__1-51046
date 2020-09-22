Attribute VB_Name = "Import_STL"
'---------------------------------------------------------------------------
' Position Angulaire Tetes Machines 5 axes
'---------------------------------------------------------------------------
Public Type PosAngulaire
    A As Double
    C As Double
    B As Double
End Type

'-----------------------------------------
' Coordonnees programmées Machines 5 axes
'-----------------------------------------
Public Type Interpolation
    Coord As Point3
    Pos As PosAngulaire
End Type

'--------------------------------------------
' Position Axes Optionnelles Machines 5 axes
'--------------------------------------------
Public Type AxePositionne
    U As Double
    V As Double
    W As Double
End Type


'----------------------------------------------
' Definition d'un maillage constitué de facette
'----------------------------------------------
Public Type Maillage
    NmbVertex As Integer
    NmbNormal As Integer
    Normal() As Point3
    Vertex() As Point3
End Type

'--------------------------------------------------------
' Definition d'un Element en 3D issu d'une relecture STL
'--------------------------------------------------------
Public Type Element3D
    name As String
    fichier As String
    STL_def As Maillage
    Origine As Point3
    Type_axe As Integer
    Vecteur As Point3
    Valeur_axe As Double
    Color As CouleurRVB
    MiniAxe As Long           ' Limite inferieur sur l'axe
    MaxiAxe As Long           ' Limite supérieur sur  l' axe
    Matrix(1 To 16) As Double ' Matrice de transformation de l'element
    Obb_Col() As OBB_box      ' Décomposition de l'élément pour diminuer les temps de calcul de la collision
    Box_Col As Box3           ' Boite en collision utilisée pour affichage graphique
End Type


'-----------------------------
' Definition Tool de fraisage
'-----------------------------
Public Type Tool
    ' 1 = Ball
    ' 2 = FlatTool
    ' 3 = Drill
    Type As Integer
    name As String           ' name de l'Tool
    Remarque As String
    Origine As Point3
    Vecteur As Point3
    Diameter As Single
    CornerRadius As Single
    LgCone As Single
    DiameterCorp As Single
    LG_Coupe As Single
    LG As Single
    box As Box3
End Type

'----------------------------
' Definition Porte Tool
'--------------------------
Public Type PO_Tool
       NB_point As Integer
       name As String
       Remarque As String
       Dec_Z As Double ' Hauteur PO de base a Point0
       Origine As Point3
       Vecteur As Point3
       Coord() As Point2
       box As Box3
End Type

'Sauvegarde des positions axes machine
' il peut Y en avoir 6 pour une fraiseuse 5 axes conventionnel
' X Y Z
' A ou B   C
' et axe fourreau W
Public Type position
    Join(7) As Double
End Type

'----------------------------
' Decalage Origine en 3D
'---------------------------
Public Type Machine3D
    name As String
    Type As Integer   'type 1=A/B C 2 A/B
    Element_Fixe As Integer   'Element fixe de la machine
    Element_Collision As Integer ' element testé en collision
    NB_axe As Integer ' Nb axe sur la machine
    Remarque As String
    DecalageMachine As Point3  'Position de Origine a position Point Pivot
    Pivot As Point3            'Position de Point Pivot à Nez de broche
    Tool_current As Integer   ' Tool chargé sur la machine
    LG_Tool_current As Double   ' Longeur Tool chargé sur la machine
    Element() As Element3D
    MagasinPo As Element3D
    PositionMagasin As Integer ' position de l'elemet magasin dans la chaine de creation de la machine
End Type

' Position Pivot varie en fonction de l'Tool chargé
Public PositionPivot As Point3

'----------------------------
' Decalage Origine en 3D
'---------------------------
Public Type DecOrigine
    X As Double
    Y As Double
    Z As Double
    ED As Double ' Rotation autour de l'axe piece
End Type

' Origine de programmation utilisée
Public OrigineProg As DecOrigine

' Les Tools
Public ToolC() As Tool
' Les portes Tools
Public POTool() As PO_Tool

' La fraiseuse
Public Machine As Machine3D
Public Parcours() As Point3

'une piece
Public Piece As Element3D

Public Position_Actuelle As Interpolation
Public UVW_Actuelle As AxePositionne


Public Position_precedente As position
' Global Position_currente As position
Public Position_accordinge As position
Public Piece_charger As Boolean
' Affiche les coordonnées
Sub Affiche_coord()
    Call Sub_Affiche_coord(Pt0, 0)
    Call Sub_Affiche_coord(Vx, 3)
    Call Sub_Affiche_coord(Vy, 6)
    Call Sub_Affiche_coord(Vz, 9)
End Sub

' Affiche les coordonnees d'un point dans le controle machine_SIMUL_FRM.AxeAbsolu(Index)
Sub Sub_Affiche_coord(P1 As Point3, Index As Integer)
    Machine_SIMUL_FRM.AxeAbsolu(Index) = Round(P1.X, 3)
    Machine_SIMUL_FRM.AxeAbsolu(Index + 1) = Round(P1.Y, 3)
    Machine_SIMUL_FRM.AxeAbsolu(Index + 2) = Round(P1.Z, 3)
End Sub

'Recuperation du fichier STL
Function ChargeFichierSTL(fichier As String, Mesh As Maillage, Optional Log As Boolean = True, Optional Facteur As Integer = 1) As String
Dim Chaine As String
Dim Donnéeslues As String
Dim Poi As Point3

If Log Then Call MessageLOG_Print("LOADING " & fichier)


On Error GoTo Fin

Open fichier For Input As #1   ' Ouvre le fichier en lecture
    Do While Not EOF(1) ' Cherche la fin du fichier.
  
    Line Input #1, Donnéeslues  ' Lit une ligne de données.

    ' récupération name du solid
    If InStr(1, Donnéeslues, "solid") Then
        Chaine = Replace(Donnéeslues, Chr(9), Chr(32)) 'Remplace les tabulations par des espaces
        Chaine = LTrim(Chaine) ' Suprime les espaces de gauche.
        Chaine = mReplaceCharacter(Chr(13), "", Chaine) 'Supprime Chr(13)
        Chaine = mReplaceCharacter(Chr(10), "", Chaine) 'Supprime Chr(10)
        Chaine = RTrim(Chaine) ' Suprime les espaces de droite.
        ChargeFichierSTL = LTrim(TokLeftRight(Chaine, " "))
    End If
    
    If InStr(1, Donnéeslues, "facet normal") Then
        Mesh.NmbNormal = Mesh.NmbNormal + 1
         ReDim Preserve Mesh.Normal(Mesh.NmbNormal)
        Call Decodage(Donnéeslues, Poi)
        Mesh.Normal(Mesh.NmbNormal - 1) = Poi
        Line Input #1, Donnéeslues  ' Lit une ligne de données.
    End If
    
    If InStr(1, Donnéeslues, "outer loop") Then
    Mesh.NmbVertex = Mesh.NmbVertex + 3
    ReDim Preserve Mesh.Vertex(Mesh.NmbVertex)
        Line Input #1, Donnéeslues  ' 1Er vertex
        Call Decodage(Donnéeslues, Poi)
        Mesh.Vertex(Mesh.NmbVertex - 3) = Poi
        Line Input #1, Donnéeslues  ' 2Eme vertex
        Call Decodage(Donnéeslues, Poi)
        Mesh.Vertex(Mesh.NmbVertex - 2) = Poi
        Line Input #1, Donnéeslues  ' 3Eme vertex
        Call Decodage(Donnéeslues, Poi)
        Mesh.Vertex(Mesh.NmbVertex - 1) = Poi
        ' Recalcul de la normal par moi
        ' Juste pour un test
        ' Call NormVec(Mesh.Vertex(Mesh.NmbVertex - 3), Mesh.Vertex(Mesh.NmbVertex - 2), Mesh.Vertex(Mesh.NmbVertex - 1), Mesh.Normal(Mesh.NmbNormal - 1))
         
    End If
    
  Loop ' fin boucle traitement fichier
Close #1

If Log Then Call MessageLOG_Print("Mesh of " & Mesh.NmbNormal & " Elements")

Exit Function


Fin:
    MsgBox Err.Description, 16, "Error #" & Err.Number

    
End Function

Public Function Decodage(ByVal Ligne As String, Poi As Point3) As Boolean
Dim Chaine As String
Dim Pt As Point3

Decodage = False

    Chaine = Replace(Ligne, Chr(9), Chr(32)) 'Remplace les tabulations par des espaces
    'chaine = LTrim(chaine) ' Suprime les espaces de gauche.
    'chaine = mReplaceCharacter(Chr(13), "", chaine) 'Supprime Chr(13)
    'chaine = mReplaceCharacter(Chr(10), "", chaine) 'Supprime Chr(10)
    'chaine = RTrim(chaine) ' Suprime les espaces de droite.
    
    If InStr(1, Ligne, "normal") Then
       Chaine = LTrim((TokRightRight(Chaine, "l")))

       Poi.X = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
       
       Poi.Y = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
    
       Poi.Z = Val(Chaine)
       Decodage = True
    End If
    
    If InStr(1, Ligne, "vertex") Then
       Chaine = LTrim((TokRightRight(Chaine, "x")))

       Poi.X = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
       
       Poi.Y = Val(TokLeftLeft(Chaine, " "))
       Chaine = LTrim(TokLeftRight(Chaine, " "))
    
       Poi.Z = Val(Chaine)
       Decodage = True
    End If
  
FinSub:

 
End Function



'def Box control
' initialisation de la boite mère
Function Def_Box(Mesh As Maillage) As Box3
Dim Pt_min As Point3
Dim Pt_max As Point3
Dim i As Long

Call MessageLOG_Print("INIT BOX " & Mesh.NmbVertex & " Vertex")
Pt_min.X = INFINITY
Pt_min.Y = INFINITY
Pt_min.Z = INFINITY
Pt_max.X = -INFINITY
Pt_max.Y = -INFINITY
Pt_max.Z = -INFINITY

' reinit axe
For i = 0 To 2
    Def_Box.Axes(i).X = 0
    Def_Box.Axes(i).Y = 0
    Def_Box.Axes(i).Z = 0
Next i

' boite alligné aux axes XYZ
Def_Box.Axes(0).X = 1
Def_Box.Axes(1).Y = 1
Def_Box.Axes(2).Z = 1


For i = 0 To Mesh.NmbVertex
    If Mesh.Vertex(i).X > Pt_max.X Then
      Pt_max.X = Mesh.Vertex(i).X
    End If
    
    If Mesh.Vertex(i).Y > Pt_max.Y Then
      Pt_max.Y = Mesh.Vertex(i).Y
    End If

    If Mesh.Vertex(i).Z > Pt_max.Z Then
      Pt_max.Z = Mesh.Vertex(i).Z
    End If

    If Mesh.Vertex(i).X < Pt_min.X Then
      Pt_min.X = Mesh.Vertex(i).X
    End If
    
    If Mesh.Vertex(i).Y < Pt_min.Y Then
      Pt_min.Y = Mesh.Vertex(i).Y
    End If
    
    If Mesh.Vertex(i).Z < Pt_min.Z Then
      Pt_min.Z = Mesh.Vertex(i).Z
    End If
Next i


Def_Box.Longueurs(0) = Abs((Pt_max.X - Pt_min.X)) / 2
Def_Box.Longueurs(1) = Abs((Pt_max.Y - Pt_min.Y)) / 2
Def_Box.Longueurs(2) = Abs((Pt_max.Z - Pt_min.Z)) / 2

' Centre
' ATTENTION LE CENTRE DOIT ETRE REMIS DANS LE CONTEXTE MACHINE
Def_Box.Centre = PointMillieu(Pt_max, Pt_min)

End Function

Public Sub Creer_Obb_Box(Box_Mere As OBB_box, Element As Element3D, Dernier_Niveau As Boolean, Mere As Integer)

Dim i As Integer
Dim J As Integer
Dim K As Integer
Dim Compteur_Triangle As Integer
Dim i0 As Integer
Dim i1 As Integer
Dim i2 As Integer
Dim Current_pointeur As Single
Dim PtC As Point3
Dim PtT1 As Point3
Dim PtT2 As Point3
Dim PtT3 As Point3
Dim TrT As Triangle3

'Svg namebre d'element
Current_pointeur = UBound(Element.Obb_Col)
' La mere a 8 filles
ReDim Box_Mere.Pointeur_filles(7)
Box_Mere.Nb_filles = 8
' L'element a donc huit box de controle en  plus
ReDim Preserve Element.Obb_Col(Current_pointeur + 8)

' mise a jour des pointeurs mere
For i = 0 To 7
    Box_Mere.Pointeur_filles(i) = Current_pointeur + i
Next i

i = 0
' creation des box filles
PtC = Box_Mere.Box_Descr.Centre
For i0 = -1 To 1 Step 2
    For i1 = -1 To 1 Step 2
        For i2 = -1 To 1 Step 2
            
            For J = 0 To 2
                Element.Obb_Col(Current_pointeur + i).Box_Descr.Axes(J) = Box_Mere.Box_Descr.Axes(J)
                Element.Obb_Col(Current_pointeur + i).Box_Descr.Longueurs(J) = Box_Mere.Box_Descr.Longueurs(J) / 2
            Next J
            
            PtT1 = VecAdd(PtC, Box_Mere.Box_Descr.Axes(0), i0 * Box_Mere.Box_Descr.Longueurs(0) / 2)
            PtT2 = VecAdd(PtT1, Box_Mere.Box_Descr.Axes(1), i1 * Box_Mere.Box_Descr.Longueurs(1) / 2)
            PtT3 = VecAdd(PtT2, Box_Mere.Box_Descr.Axes(2), i2 * Box_Mere.Box_Descr.Longueurs(2) / 2)
            Element.Obb_Col(Current_pointeur + i).Box_Descr.Centre = PtT3
            Element.Obb_Col(Current_pointeur + i).Pointeur_Mere = Mere
            i = i + 1 'creation de l'indice
        Next i2
    Next i1
Next i0

If Dernier_Niveau Then
    Call MessageLOG_Print("Sharing elements")
    'Repartition des facettes
    
    For i = 0 To 7
        Compteur_Triangle = 0
        For J = 0 To Element.STL_def.NmbVertex - 3 Step 3
            'creation du triangle
            For K = 0 To 2
            TrT.S(K) = Element.STL_def.Vertex(J + K)
            Next K
            'test du triangle
            col = TestIntersectionTriangleBox(TrT, Element.Obb_Col(Current_pointeur + i).Box_Descr)
            If col Then
                Compteur_Triangle = Compteur_Triangle + 1
                Element.Obb_Col(Current_pointeur + i).Nb_facettes = Compteur_Triangle
                ReDim Preserve Element.Obb_Col(Current_pointeur + i).Maillage_Box(Compteur_Triangle)
                Element.Obb_Col(Current_pointeur + i).Maillage_Box(Compteur_Triangle - 1) = TrT
            End If
            
        Next J
        Call MessageLOG_Print("SubBoxs  " & i & " with" & Compteur_Triangle & " elements")
    Next i
End If

End Sub
