Attribute VB_Name = "CalculPosition"
Option Explicit

'Public Const Pi = 3.14159265358979
Public PI As Double
Public RADTODEG As Double
Public DEGTORAD As Double

Public Const INFINITY As Double = 9E+99


'CALCUL AVEC DEC ET RTCP
Function DEC_PLUS_RTCP(Interpo As Interpolation, Uvw As AxePositionne) As Interpolation
Dim SVG_position As position
Dim P0 As Point3
Dim J As Integer

 'Sauvegarde position actuelle
For J = 1 To Machine.NB_axe + 1
    SVG_position.Join(J) = Machine.Element(J).Valeur_axe
Next



If Machine_SIMUL_FRM.OptionRTCP = False Then
'-----------------------------------------------------------------
'  Calcul sans RTCP
'-----------------------------------------------------------------
    ' Affectation des axes rotatifs a 0 pour non RTCP
    For J = 1 To Machine.NB_axe
            Select Case Machine.Element(J).name
                    Case "A"
                        Machine.Element(J).Valeur_axe = 0
                    Case "B"
                        Machine.Element(J).Valeur_axe = 0
                    Case "C"
                        Machine.Element(J).Valeur_axe = 0
                    Case Else
            End Select
    Next J
Else
'-----------------------------------------------------------------
'  Calcul avec RTCP
'-----------------------------------------------------------------
    ' Affectation des axes rotatifs
    ' Les axes rotatifs doivent prendre obligatoirement la valeur de la consigne
    For J = 1 To Machine.NB_axe
            Select Case Machine.Element(J).name
                    Case "A"
                        Machine.Element(J).Valeur_axe = Interpo.Pos.A
                    Case "B"
                        Machine.Element(J).Valeur_axe = Interpo.Pos.B
                    Case "C"
                        Machine.Element(J).Valeur_axe = Interpo.Pos.C
                    Case Else
            End Select
    Next J
End If
' Affectation des axes optionnelles
For J = 1 To Machine.NB_axe
            Select Case Machine.Element(J).name
                    Case "U"
                        Machine.Element(J).Valeur_axe = Uvw.U
                    Case "V"
                        Machine.Element(J).Valeur_axe = Uvw.V
                    Case "W"
                        Machine.Element(J).Valeur_axe = Uvw.W
                    Case Else
            End Select
Next J
    
P0 = Interpo.Coord
'calcul position avec valeur angulaire des axes rotatifs
Call CalculPos(P0)


'Restaure position actuelle
For J = 1 To Machine.NB_axe + 1
    Machine.Element(J).Valeur_axe = SVG_position.Join(J)
Next

DEC_PLUS_RTCP.Coord = P0
DEC_PLUS_RTCP.Pos = Interpo.Pos
End Function

Function Calcul_Position(Inter As Interpolation, Uvw As AxePositionne) As position
Dim J As Integer
Dim position_calcule As position
    'Position_accordinge = Pos
    ' Affectation des axes linéaires
    ' according le calcul
    For J = 1 To Machine.NB_axe
            Select Case Machine.Element(J).name
                    Case "X"
                       position_calcule.Join(J) = Machine.Element(J).Valeur_axe - Inter.Coord.X
                    Case "Y"
                       position_calcule.Join(J) = Machine.Element(J).Valeur_axe - Inter.Coord.Y
                    Case "Z"
                       position_calcule.Join(J) = Machine.Element(J).Valeur_axe - Inter.Coord.Z
                       
                   ' Affectation des axes rotatifs
                   ' Les axes rotatifs doivent prendre obligatoirement la valeur de la consigne

                    Case "A"
                        position_calcule.Join(J) = Inter.Pos.A
                    Case "B"
                        position_calcule.Join(J) = Inter.Pos.B
                    Case "C"
                        position_calcule.Join(J) = Inter.Pos.C
                   
                   ' Affectation des axes optionnels
                   ' Les axes optionnels doivent prendre obligatoirement la valeur de la consigne
                    Case "U"
                        position_calcule.Join(J) = Uvw.U
                    Case "V"
                        position_calcule.Join(J) = Uvw.V
                    Case "W"
                        position_calcule.Join(J) = Uvw.W
                    Case Else
                    
            End Select
    Next J
    
Calcul_Position = position_calcule
End Function

Function Position_Machine(xyzac As Interpolation, u_v_w As AxePositionne) As position
Dim Inter As Interpolation
Dim J As Integer
Dim SVG_position As position
Dim retour As Boolean

'Recalcul avec origine programme
' Decalage fixé dans la fenetre FRM_simul
xyzac = TransCoord(xyzac, OrigineProg)
   
'Calcul Selon Decalage et RTCP (ou correction Lg Tool)
'Attention dans cette fonction les valeurs axes machines sont modifiées
Inter = DEC_PLUS_RTCP(xyzac, u_v_w)

' Une fois la position a atteindre calculée
' Calcul Des axes machines
' et affectation de Position_accordinge
Position_Machine = Calcul_Position(Inter, u_v_w)

End Function


Function Trans_Matrix(ByRef Mx2() As Double, ByRef P0 As Point3, ByRef p1 As Point3) As Point3
'Calcul according la  matrice de transformation
               Trans_Matrix.X = Mx2(0, 0) * p1.X + Mx2(0, 1) * p1.Y + Mx2(0, 2) * p1.Z + P0.X
               Trans_Matrix.Y = Mx2(1, 0) * p1.X + Mx2(1, 1) * p1.Y + Mx2(1, 2) * p1.Z + P0.Y
               Trans_Matrix.Z = Mx2(2, 0) * p1.X + Mx2(2, 1) * p1.Y + Mx2(2, 2) * p1.Z + P0.Z
End Function


Function Trans_Matrix_1_16(ByRef Matrix() As Double, ByRef p1 As Point3) As Point3
'Calcul according la  matrice de transformation

               Trans_Matrix_1_16.X = Matrix(1) * p1.X + Matrix(5) * p1.Y + Matrix(9) * p1.Z + Matrix(13)
               Trans_Matrix_1_16.Y = Matrix(2) * p1.X + Matrix(6) * p1.Y + Matrix(10) * p1.Z + Matrix(14)
               Trans_Matrix_1_16.Z = Matrix(3) * p1.X + Matrix(7) * p1.Y + Matrix(11) * p1.Z + Matrix(15)
End Function
