Attribute VB_Name = "Fonction"
Public Type couleur
  valeur As Long
End Type

'--------------
' 3 couleurs
'-------------
Public Type CouleurRVB
  Rouge As Long
  Vert As Long
  Bleu As Long
End Type


'---------------------------------------------------------------------------
' Point en 3D
'---------------------------------------------------------------------------
Public Type Point3
    X As Double
    Y As Double
    Z As Double
End Type

'---------------------------------------------------------------------------
' Point en 2D
'---------------------------------------------------------------------------
Public Type Point2
    X As Double
    Y As Double
End Type

'---------------------------------------------------------------------------
' Triangle en 3D
'---------------------------------------------------------------------------
Public Type Triangle3
   S(2) As Point3       'Points sommets (3) du triangle
End Type

'---------------------------------------------------------------------------
'Boite en 3D
'---------------------------------------------------------------------------
Public Type Box3
   Centre As Point3         ' Centre de la boite
   Axes(2) As Point3        ' Vecteur orientation de la boite
   Longueurs(2) As Double   ' Longueur de la boite
End Type

'---------------------------------------------------------------------------
'Subdivision de Boite en 3D pour control de collision
'---------------------------------------------------------------------------
Public Type OBB_box
   Pointeur_Mere As Single        ' Pointeur sur la boite mere
   Nb_filles As Integer           ' Nb_filles
   Pointeur_filles() As Single    ' Pointeur sur les boite filles (8)
   Box_Descr As Box3              ' Description de la boite
   Nb_facettes As Integer         ' namebre de facettes dans la boite
   Maillage_Box() As Triangle3    ' facettes élements de cette boite
End Type

'Angle entre 2 Vecteurs formés de 3 point ( attention ne gère pas de valeu + ou -)
Function Angle3Pt(ByRef Po1 As Point3, ByRef Po2 As Point3, ByRef Po3 As Point3) As Double
Dim VEC1 As Point3
Dim VEC2 As Point3

VEC1 = VecSub(Po1, Po2)
VEC2 = VecSub(Po2, Po3)
If Distance(Po3, Po2) = 0 Or Distance(Po1, Po2) = 0 Then
    Angle3Pt = 0
Else
    If Distance(Po1, Po3) = 0 Then
        Angle3Pt = 180
    Else
        Angle3Pt = RADTODEG * (ACOS((VEC1.X * VEC2.X + VEC1.Y * VEC2.Y + VEC1.Z * VEC2.Z) / (Distance(Po3, Po2) * Distance(Po1, Po2))))
    End If
End If
                    
End Function


Function Atan2(ByVal X, ByVal Y)
    'On Error Resume Next
    '0 a PI
    If Y = 0 Then
        If X = 0 Then
            Atan2 = 0
        ElseIf X > 0 Then
            Atan2 = 0 'PI / 2
        Else
            Atan2 = PI '-PI / 2
        End If
    ElseIf X = 0 Then
        If Y > 0 Then
            Atan2 = PI / 2
        Else
            Atan2 = -PI / 2
        End If
    ElseIf X > 0 Then
        Atan2 = Atn(Y / X)
    Else
        Atan2 = (PI - Atn(Abs(Y) / Abs(X))) * Sgn(Y)
    End If
End Function




'****************************************************************
' Name: A 'strtok' function for VB
' Description:I wrote four functions to tokenize strings. He
'     re they are...
'The functions work like this TokLeftLeft finds the leftmost token and
'then returns the left part of the string (empty if not there). You
'can figure out the rest. Note that if the token is more than 1 character
'then the function will always return "".
'****************************************************************
Public Function TokLeftLeft(ByRef Source As String, ByRef token As String) As String

       Dim i As Integer
       TokLeft = Source

              For i = 1 To Len(Source)

                            If Mid(Source, i, 1) = token Then
                                   TokLeftLeft = Left(Source, i - 1)
                                   Exit Function
                            End If

              Next i

End Function
Public Function TokLeftRight(ByRef Source As String, ByRef token As String) As String

       Dim i As Integer
       TokRight = Source

              For i = 1 To Len(Source)

                            If Mid(Source, i, 1) = token Then
                                   TokLeftRight = Right(Source, Len(Source) - i)
                                   Exit Function
                            End If

              Next i

End Function
Public Function TokRightLeft(ByRef Source As String, ByRef token As String) As String

       Dim i As Integer
       TokRightLeft = ""

              For i = Len(Source) To 1 Step -1

                            If Mid(Source, i, 1) = token Then
                                   TokRightLeft = Left(Source, i - 1)
                                   Exit Function
                            End If

              Next i

End Function
Public Function TokRightRight(ByRef Source As String, ByRef token As String) As String

       Dim i As Integer
       TokRightRight = ""

              For i = Len(Source) To 1 Step -1

                            If Mid(Source, i, 1) = token Then
                                   TokRightRight = Right(Source, Len(Source) - i)
                                   Exit Function
                            End If

              Next i

End Function



'****************************************************************
' Name: mReplaceCharacter
' Description:Replaces all instances of substring A with sub
'     string B in a string
' By: Ian Ippolito
'
' Inputs:strString==string to do replacing on
'strOrigChar==orig substring
'strReplaceChar==substring to replace orig substring

' Returns:strString after replacing all instances of strOrigChar with strReplaceChar
' Assumes:None
' Side Effects:None
'
'Code provided by Planet Source Code(tm) 'as is', without
'     warranties as to performance, fitness, merchantability,
'     and any other warranty (whether expressed or implied).
'****************************************************************
Function mReplaceCharacter(ByRef strOrigChar, ByRef strReplaceChar, ByVal strString)

       '     '**********************************
       '     'changes all strOrigChar
       '     ' to
       '     ' in strString
       '     '**********************************
       Dim strResult
       strResult = ""
       '     'traverse string
       Dim intIndex
        For intIndex = 1 To Len(strString)

              If (Mid(strString, intIndex, Len(strOrigChar)) = strOrigChar) Then
                                        '*************
                                        'match found
                                        '*************
                                        'MsgBox "found in" + strString
                                        strResult = strResult + strReplaceChar
                                        intIndex = intIndex + Len(strOrigChar) - 1
                                Else
                                        '*************
                                        'no match
                                        '*************
                     strResult = strResult + Mid(strString, intIndex, 1)
              End If

Next

mReplaceCharacter = strResult
End Function
Function Longueur(ByRef P1 As Point3) As Double
  Longueur = Sqr((P1.X ^ 2) + (P1.Y ^ 2) + (P1.Z ^ 2))
End Function

Function Distance(ByRef P1 As Point3, ByRef P2 As Point3) As Double
    Distance = Sqr((P2.X - P1.X) ^ 2 + (P2.Y - P1.Y) ^ 2 + (P2.Z - P1.Z) ^ 2)
End Function
'Addition de vecteur
Function VecAdd(ByRef P1 As Point3, ByRef P2 As Point3, Optional f As Double = 1) As Point3
 VecAdd.X = P1.X + f * P2.X
 VecAdd.Y = P1.Y + f * P2.Y
 VecAdd.Z = P1.Z + f * P2.Z
End Function

'Produit Scalaire
Function Dot(ByRef p As Point3, ByRef q As Point3) As Double
    Dot = p.X * q.X + p.Y * q.Y + p.Z * q.Z
End Function
Function SubVect(ByRef P1 As Point3, ByRef P2 As Point3, ByRef f As Double) As Point3
 SubVect.X = P1.X - P2.X * f
 SubVect.Y = P1.Y - P2.Y * f
 SubVect.Z = P1.Z - P2.Z * f
End Function

'Produit vectoriel
Function VecProd(ByRef P1 As Point3, ByRef P2 As Point3) As Point3
Dim P4 As Point3

 P4.X = (P1.Y * P2.Z) - (P1.Z * P2.Y)
 P4.Y = (P1.Z * P2.X) - (P1.X * P2.Z)
 P4.Z = (P1.X * P2.Y) - (P1.Y * P2.X)
 VecProd = P4
 
End Function


' Soustraction de vecteur
Function VecSub(ByRef P1 As Point3, ByRef P2 As Point3) As Point3
 VecSub.X = P1.X - P2.X
 VecSub.Y = P1.Y - P2.Y
 VecSub.Z = P1.Z - P2.Z
End Function
'récupération du vecteur normal de 3 points
Function NormVec(ByRef P1 As Point3, ByRef P2 As Point3, ByRef P3 As Point3) As Point3
 NormVec = VecteurUnitaire(VecProd(VecSub(P1, P2), VecSub(P3, P2)))
End Function

' transforme un vecteur en vecteur unitaire
Function VecteurUnitaire(ByRef P1 As Point3) As Point3
Dim Norm As Double
Norm = Sqr((P1.X) ^ 2 + (P1.Y) ^ 2 + (P1.Z) ^ 2)
If Norm = 0 Then
    Exit Function
End If
    
    VecteurUnitaire.X = P1.X / Norm
    VecteurUnitaire.Y = P1.Y / Norm
    VecteurUnitaire.Z = P1.Z / Norm
End Function
' Coordonées du point Milieu
Function PointMillieu(ByRef P1 As Point3, ByRef P2 As Point3) As Point3
 PointMillieu.X = 0.5 * (P1.X + P2.X)
 PointMillieu.Y = 0.5 * (P1.Y + P2.Y)
 PointMillieu.Z = 0.5 * (P1.Z + P2.Z)
End Function
'****************************************************************
' Name: Round
'
' Inputs:DP is the decimal place to round to (0 to 14) e.g
' Round (3.56376, 3) will give the result 3.564
' Round (3.56376, 1) will give the result 3.6
' Round (3.56376, 0) will give the result 4
' Round (3.56376, 2) will give the result 3.56
' Round (1.4999, 3) will give the result 1.5
' Round (1.4899, 2) will give the result 1.49
' Returns:None
' Assumes:None
' Side Effects:None
'
'****************************************************************
Function Round(x1 As Double, DP As Integer) As Double
    Round = Int((x1 * 10 ^ DP) + 0.5) / 10 ^ DP
End Function

Function ACOS(Ang)
    Select Case Ang
        Case 1
            ACOS = 0 '0
        Case -1
            ACOS = 4 * Atn(1) 'PI
        Case Else
            ACOS = 2 * Atn(1) - Atn(Ang / Sqr(1 - Ang * Ang))
    End Select
End Function

Function ASIN(Ang)
    Select Case Ang
        Case 1
            ASIN = 2 * Atn(1)
        Case -1
            ASIN = -2 * Atn(1)
        Case Else
            ASIN = Atn(Ang / Sqr(1 - Ang * Ang))
    End Select
End Function

'Angle entre 2 Vecteurs formés concurrent en 0
Function AngleVect(ByRef Po1 As Point3, ByRef Po2 As Point3, ByRef Normal As Point3) As Double
Dim VEC1 As Point3
Dim VEC2 As Point3
Dim Po3 As Point3
Dim Signe As Double

VEC1 = Po1
VEC2 = Po2
If Longueur(VEC1) = 0 Or Longueur(VEC2) = 0 Then
    AngleVect = 0
Else
    If Distance(Po1, Po2) = 0 Then
        AngleVect = 0
    Else

        Po3 = VecProd(VEC1, VEC2)
        Signe = Sgn(Dot(Po3, Normal))
        
        'Debug.Print " Dot/Normal         | " & Signe
        
        'Debug.Print " Po3         | " & Format(Po3.x, "#,###0.0000") & " | " & Format(Po3.y, "#,###0.0000") & " | " & Format(Po3.Z, "#,###0.0000") & " | "
        AngleVect = Signe * RADTODEG * (ACOS((VEC1.X * VEC2.X + VEC1.Y * VEC2.Y + VEC1.Z * VEC2.Z) / (Longueur(VEC1) * Longueur(VEC2))))
    End If
End If
                    
End Function




'Autre Fonction avec le même resultat
Public Function GetRVB(Couleur_long As Long) As CouleurRVB
Dim TempColor As CouleurRVB
    
    TempColor.Bleu = Int(Couleur_long / 65536)
    TempColor.Vert = Int((Couleur_long - (65536 * TempColor.Bleu)) / 256)
    TempColor.Rouge = Couleur_long - (65536 * TempColor.Bleu + 256 * TempColor.Vert)

GetRVB = TempColor
  
End Function
