Attribute VB_Name = "ControleCollision"
Option Explicit


'calcul la collision
Function Collision() As Boolean
Dim I As Integer
Dim J As Integer
Dim Mx2(2, 2) As Double
Dim MxP(2, 2) As Double
Dim TRB As Triangle3
Dim coll As Boolean
Dim PT_dec0 As Point3
'Dim PT_dec1 As Point3
Dim PT_dec0P As Point3
'Dim PT_dec1P As Point3

Dim Element_Machine As Element3D
Dim Element_Piece As Element3D



Collision = False

Call GetPointColl

' Description local pour calcul
Element_Machine = Machine.Element(Machine.Element_Collision)

Mx2(0, 0) = Element_Machine.Matrix(1)
Mx2(1, 0) = Element_Machine.Matrix(2)
Mx2(2, 0) = Element_Machine.Matrix(3)

Mx2(0, 1) = Element_Machine.Matrix(5)
Mx2(1, 1) = Element_Machine.Matrix(6)
Mx2(2, 1) = Element_Machine.Matrix(7)

Mx2(0, 2) = Element_Machine.Matrix(9)
Mx2(1, 2) = Element_Machine.Matrix(10)
Mx2(2, 2) = Element_Machine.Matrix(11)

PT_dec0.X = Element_Machine.Matrix(13)
PT_dec0.Y = Element_Machine.Matrix(14)
PT_dec0.Z = Element_Machine.Matrix(15)

    'Debug.Print Element_Machine.Obb_Col(0).Box_Descr.Centre.Z
    'PT_dec1 = VecAdd(PT_dec0, Element_Machine.Obb_Col(0).Box_Descr.Centre)
    'Debug.Print PT_dec1.Z
    'Element_Machine.Obb_Col(0).Box_Descr.Centre = PT_dec1
    Element_Machine.Obb_Col(0).Box_Descr.Centre = Trans_Matrix(Mx2, PT_dec0, Element_Machine.Obb_Col(0).Box_Descr.Centre)
    
    Debug.Print Element_Machine.Obb_Col(0).Box_Descr.Centre.X
    Debug.Print Element_Machine.Obb_Col(0).Box_Descr.Centre.Y
    Debug.Print Element_Machine.Obb_Col(0).Box_Descr.Centre.Z
    
    For I = 0 To 2
        Element_Machine.Obb_Col(0).Box_Descr.Axes(I).X = Mx2(0, I)
        Element_Machine.Obb_Col(0).Box_Descr.Axes(I).Y = Mx2(1, I)
        Element_Machine.Obb_Col(0).Box_Descr.Axes(I).Z = Mx2(2, I)
    Next I
    
    Machine.Element(Machine.Element_Collision).Box_Col = Element_Machine.Obb_Col(0).Box_Descr

    
  
' collision avec sol ( grille)
       coll = TestIntersectionTriangleBox(Ta1, Element_Machine.Obb_Col(0).Box_Descr)
       If coll And Machine_SIMUL_FRM.CHK.Value Then
        Collision = True
        Machine.Element(Machine.Element_Collision).Box_Col = Element_Machine.Obb_Col(0).Box_Descr
        Exit Function
       End If
       coll = TestIntersectionTriangleBox(Ta2, Element_Machine.Obb_Col(0).Box_Descr)
       If coll And Machine_SIMUL_FRM.CHK.Value Then
        Collision = True
        Machine.Element(Machine.Element_Collision).Box_Col = Element_Machine.Obb_Col(0).Box_Descr
        Exit Function
       End If

' collision avec piece
       If Piece_charger Then
       
       'Mise a jour boite de collision piece
       Element_Piece = Piece
       
MxP(0, 0) = Element_Piece.Matrix(1)
MxP(1, 0) = Element_Piece.Matrix(2)
MxP(2, 0) = Element_Piece.Matrix(3)

MxP(0, 1) = Element_Piece.Matrix(5)
MxP(1, 1) = Element_Piece.Matrix(6)
MxP(2, 1) = Element_Piece.Matrix(7)

MxP(0, 2) = Element_Piece.Matrix(9)
MxP(1, 2) = Element_Piece.Matrix(10)
MxP(2, 2) = Element_Piece.Matrix(11)

PT_dec0P.X = Element_Piece.Matrix(13)
PT_dec0P.Y = Element_Piece.Matrix(14)
PT_dec0P.Z = Element_Piece.Matrix(15)

    'Debug.Print Element_Piece.Obb_Col(0).Box_Descr.Centre.Z
    'PT_dec1P = VecAdd(PT_dec0P, Element_Piece.Obb_Col(0).Box_Descr.Centre)
    'Debug.Print PT_dec1P.Z
    Element_Piece.Obb_Col(0).Box_Descr.Centre = Trans_Matrix(MxP, PT_dec0P, Piece.Obb_Col(0).Box_Descr.Centre)
    'Debug.Print Element_Piece.Obb_Col(0).Box_Descr.Centre.Z
    
    For I = 0 To 2
        Element_Piece.Obb_Col(0).Box_Descr.Axes(I).X = MxP(0, I)
        Element_Piece.Obb_Col(0).Box_Descr.Axes(I).Y = MxP(1, I)
        Element_Piece.Obb_Col(0).Box_Descr.Axes(I).Z = MxP(2, I)
    Next I
    
    Piece.Box_Col = Element_Piece.Obb_Col(0).Box_Descr
    'intersection avec la boite mère
    
            coll = TestIntersectionBoite(Element_Piece.Obb_Col(0).Box_Descr, Element_Machine.Obb_Col(0).Box_Descr)
            If coll Then
                coll = Collision_boite_recursif(Element_Piece, Element_Machine, 1, 1, MxP, PT_dec0P, Mx2, PT_dec0)
                'Mise a jour pour visualisation des infos elements
                Machine.Element(Machine.Element_Collision).Box_Col = Element_Machine.Box_Col
                Piece.Box_Col = Element_Piece.Box_Col
                
                Collision = coll
                Call Machine_SIMUL_FRM.Pic_Paint
            End If
       End If
End Function

Function Collision_boite_recursif(Elemt1 As Element3D, Elemt2 As Element3D, Niveau1 As Integer, Niveau2 As Integer, Mx1() As Double, Dec1 As Point3, Mx2() As Double, Dec2 As Point3) As Boolean
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim L As Integer
Dim coll As Boolean
Dim PT_dec0 As Point3
'Dim PT_dec1 As Point3
'Dim PT_dec2 As Point3
'Dim Elemt3 As Element3D
'Dim rien As Point3
Dim P1 As Integer
Dim P2 As Integer

'PT_dec0.Z = 1


    ' Init des autres boite si collision avec boite 1
    For I = Niveau2 To Niveau2 + 7
    
        'Debug.Print Machine.Element(Machine.Element_Collision).Obb_Col(I).Box_Descr.Centre.Y
        'PT_dec1 = VecAdd(Dec2, Machine.Element(Machine.Element_Collision).Obb_Col(I).Box_Descr.Centre)
        Debug.Print Dec2.Y
        Elemt2.Obb_Col(I).Box_Descr.Centre = Trans_Matrix(Mx2, Dec2, Elemt2.Obb_Col(I).Box_Descr.Centre)
    
        'Debug.Print Elemt2.Obb_Col(I).Box_Descr.Centre.Y

        For J = 0 To 2
            'Elemt2.Obb_Col(I).Box_Descr.Axes(J) = Elemt2.Obb_Col(0).Box_Descr.Axes(J)
            Elemt2.Obb_Col(I).Box_Descr.Axes(J).X = Mx2(0, J)
            Elemt2.Obb_Col(I).Box_Descr.Axes(J).Y = Mx2(1, J)
            Elemt2.Obb_Col(I).Box_Descr.Axes(J).Z = Mx2(2, J)
        Next J
    Next I
                    
    
    
    For I = Niveau1 To Niveau1 + 7
    
    

        Elemt1.Obb_Col(I).Box_Descr.Centre = Trans_Matrix(Mx1, Dec1, Elemt1.Obb_Col(I).Box_Descr.Centre)
        For J = 0 To 2
            'Elemt2.Obb_Col(I).Box_Descr.Axes(J) = Elemt2.Obb_Col(0).Box_Descr.Axes(J)
            Elemt1.Obb_Col(I).Box_Descr.Axes(J).X = Mx1(0, J)
            Elemt1.Obb_Col(I).Box_Descr.Axes(J).Y = Mx1(1, J)
            Elemt1.Obb_Col(I).Box_Descr.Axes(J).Z = Mx1(2, J)
        Next J
    
    
        For J = Niveau2 To Niveau2 + 7
                    coll = TestIntersectionBoite(Elemt1.Obb_Col(I).Box_Descr, Elemt2.Obb_Col(J).Box_Descr)
                    
                    If coll Then
                        Elemt1.Box_Col = Elemt1.Obb_Col(I).Box_Descr
                        Elemt2.Box_Col = Elemt2.Obb_Col(J).Box_Descr
                        
                        If Elemt1.Obb_Col(I).Nb_filles > 0 And Elemt2.Obb_Col(J).Nb_filles > 0 Then
                            'For K = 0 To 7
                               ' For L = 0 To 7
                                    P1 = Elemt1.Obb_Col(I).Pointeur_filles(0)
                                    P2 = Elemt2.Obb_Col(J).Pointeur_filles(0)
                                    'Debug.Print P1 & " & " & P2
                                        coll = Collision_boite_recursif(Elemt1, Elemt2, P1, P2, Mx1, Dec1, Mx2, Dec2)
                                        If coll Then
                                          Collision_boite_recursif = True
                                         Exit Function
                                        End If
                               ' Next L
                            'Next K
                        Else
                            ' test avec les facettes
                            ' si pas de facettes dans boite alors out
                            If Elemt1.Obb_Col(I).Nb_facettes > 0 And Elemt2.Obb_Col(J).Nb_facettes > 0 Then
                                coll = Collision_Boite_facettes(Piece.Obb_Col(I), Machine.Element(Machine.Element_Collision).Obb_Col(J), Mx1, Dec1, Mx2, Dec2)
                                If coll Then
                                    Collision_boite_recursif = True
                                    Exit Function
                                End If
                            End If
                        
                        End If
        
                    End If
        Next J
    Next I
Collision_boite_recursif = False
End Function

'Collision entre facettes a l'interieur d'une boite
Function Collision_Boite_facettes(obb_box1 As OBB_box, obb_box2 As OBB_box, Mx1() As Double, ptdec1 As Point3, Mx2() As Double, ptdec2 As Point3) As Boolean
Dim I As Integer
Dim J As Integer
Dim coll As Boolean
Dim Pt As Point3
Dim Maillage_tempo_A() As Triangle3
Dim Maillage_tempo_B() As Triangle3

'DoEvents

ReDim Maillage_tempo_A(UBound(obb_box1.Maillage_Box))
ReDim Maillage_tempo_B(UBound(obb_box2.Maillage_Box))

' reclacul pour les facettes de l'elements 1
For I = 0 To obb_box1.Nb_facettes - 1
    For J = 0 To 2
    'Debug.Print ptdec.Z
        Maillage_tempo_A(I).S(J) = Trans_Matrix(Mx1, ptdec1, obb_box1.Maillage_Box(I).S(J))
        'Maillage_tempo(I).S(J) = Pt ' VecAdd(Pt0, Pt)
    Next J
Next I

' reclacul pour les facettes de l'elements 2
For I = 0 To obb_box2.Nb_facettes - 1
    For J = 0 To 2
    'Debug.Print ptdec.Z
        Maillage_tempo_B(I).S(J) = Trans_Matrix(Mx2, ptdec2, obb_box2.Maillage_Box(I).S(J))
        'Maillage_tempo(I).S(J) = Pt ' VecAdd(Pt0, Pt)
    Next J
Next I


'DoEvents

For I = 0 To UBound(Maillage_tempo_A) - 1
    For J = 0 To UBound(Maillage_tempo_B) - 1
                coll = TestIntersectionTriangle(Maillage_tempo_A(I), Maillage_tempo_B(J))
                If coll Then
                    Collision_Boite_facettes = True
                    Exit Function
                End If
    Next J
Next I
Collision_Boite_facettes = False
End Function

