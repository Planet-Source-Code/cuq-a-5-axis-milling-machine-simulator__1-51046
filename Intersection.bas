Attribute VB_Name = "Intersection"
'----------------------------------------------------------------------------
Sub ProjectionTriangle(rkD As Point3, akV As Triangle3, rfMin As Double, rfMax As Double)
Dim afDot(2) As Double


    afDot(0) = Dot(rkD, akV.S(0))
    afDot(1) = Dot(rkD, akV.S(1))
    afDot(2) = Dot(rkD, akV.S(2))


    rfMin = afDot(0)
    rfMax = rfMin

    If (afDot(1) < rfMin) Then
        rfMin = afDot(1)
    Else
        If (afDot(1) > rfMax) Then
            rfMax = afDot(1)
        End If
    End If

    If (afDot(2) < rfMin) Then
        rfMin = afDot(2)
    Else
        If (afDot(2) > rfMax) Then
            rfMax = afDot(2)
        End If
    End If

End Sub
'----------------------------------------------------------------------------
Sub ProjectionBoite(rkD As Point3, rbBox As Box3, rfMin As Double, rfMax As Double)
Dim fDdC As Double
Dim fR As Double
Dim cpBox As Box3

    cpBox = rbBox
    fDdC = Dot(rkD, cpBox.Centre)
    fR = cpBox.Longueurs(0) * Abs(Dot(rkD, cpBox.Axes(0))) + cpBox.Longueurs(1) * Abs(Dot(rkD, cpBox.Axes(1))) + cpBox.Longueurs(2) * Abs(Dot(rkD, cpBox.Axes(2)))
    rfMin = fDdC - fR
    rfMax = fDdC + fR
End Sub

'----------------------------------------------------------------------------
Function TestIntersectionTriangleBox(apkTri As Triangle3, rkBox As Box3) As Boolean
    Dim fMin0 As Double
    Dim fMax0 As Double
    Dim fMin1 As Double
    Dim fMax1 As Double
    Dim fDdC As Double
    
    Dim kD As Point3
    Dim akE(2) As Point3
    Dim i0 As Integer
    Dim i1 As Integer
    
    Dim cpBox As Box3

     cpBox = rkBox
     
    ' test la direction de la normale du triangle
    akE(0) = VecSub(apkTri.S(1), apkTri.S(0)) '
    akE(1) = VecSub(apkTri.S(2), apkTri.S(0))  '
    'Creation de la normale
    kD = VecProd(akE(0), akE(1)) '
    fMin0 = Dot(kD, apkTri.S(0)) '
    fMax0 = fMin0 '
    Call ProjectionBoite(kD, cpBox, fMin1, fMax1) '
    If (fMax1 < fMin0 Or fMax0 < fMin1) Then
        TestIntersectionTriangleBox = False
        Exit Function
    End If

    ' test la direction des faces de la boite
    For i = 0 To 2
        kD = cpBox.Axes(i) '
        Call ProjectionTriangle(kD, apkTri, fMin0, fMax0)
        fDdC = Dot(kD, cpBox.Centre) '
        fMin1 = fDdC - cpBox.Longueurs(i) '
        fMax1 = fDdC + cpBox.Longueurs(i) '
        If (fMax1 < fMin0 Or fMax0 < fMin1) Then
                TestIntersectionTriangleBox = False
                Exit Function
        End If
    Next i

    ' test la direction du triangle-boite limites du produit vectoriel
    akE(2) = VecSub(akE(1), akE(0))  '
    For i0 = 0 To 2
        For i1 = 0 To 2
            kD = VecProd(akE(i0), cpBox.Axes(i1)) '
            Call ProjectionTriangle(kD, apkTri, fMin0, fMax0)
            Call ProjectionBoite(kD, cpBox, fMin1, fMax1)
            If (fMax1 < fMin0 Or fMax0 < fMin1) Then
                TestIntersectionTriangleBox = False
                Exit Function
            End If
        Next i1
    Next i0
    TestIntersectionTriangleBox = True
    
End Function

'----------------------------------------------------------------------------
Function TestIntersectionBoite(Box0 As Box3, Box1 As Box3) As Boolean
    Const fCutoff As Double = 0.9999999  'f '
    Dim ExistePaireParallele As Boolean
    Dim i As Integer
    Dim akA(2) As Point3
    Dim akB(2) As Point3
    Dim afEA(2) As Double
    Dim afEB(2) As Double
    Dim kD As Point3
    
    Dim aafC(2, 2) As Double    ' matrice C = A^T B, c_{ij} = Dot(A_i,B_j)
    Dim aafAbsC(2, 2) As Double   ' |C(i,j)|
    Dim afAD(2) As Double         ' Dot(A_i,D)
    Dim fR0 As Double  ' Rayon interval et distance entre centres
    Dim fR1 As Double  ' Rayon interval et distance entre centres
    Dim fR As Double  ' Rayon interval et distance entre centres
    Dim fR01 As Double ' = R0 + R1
    
    
    ' Init variables de calcul
    For i = 0 To 2
        ' transforme en vecteur unitaires
        akA(i) = VecteurUnitaire(Box0.Axes(i))
        akB(i) = VecteurUnitaire(Box1.Axes(i))
        afEA(i) = Box0.Longueurs(i)
        afEB(i) = Box1.Longueurs(i)
    Next i

    ' Calcul le vecteurs entre les deux centres des boites D = C1-C0
    kD = VecSub(Box1.Centre, Box0.Centre)

    ' Axes C0+t*A0
    For i = 0 To 2
        aafC(0, i) = Dot(akA(0), akB(i)) '
        aafAbsC(0, i) = Abs(aafC(0, i)) '
        If (aafAbsC(0, i) > fCutoff) Then
            ExistePaireParallele = True
        End If
    Next i
    
    afAD(0) = Dot(akA(0), kD) '
    fR = Abs(afAD(0)) '
    fR1 = afEB(0) * aafAbsC(0, 0) + afEB(1) * aafAbsC(0, 1) + afEB(2) * aafAbsC(0, 2) '
    fR01 = afEA(0) + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False
        Exit Function
    End If

    ' Axes C0+t*A1
    For i = 0 To 2
        aafC(1, i) = Dot(akA(1), akB(i)) '
        aafAbsC(1, i) = Abs(aafC(1, i)) '
        If (aafAbsC(1, i) > fCutoff) Then
            ExistePaireParallele = True '
        End If
    Next i
    
    afAD(1) = Dot(akA(1), kD) '
    fR = Abs(afAD(1)) '
    fR1 = afEB(0) * aafAbsC(1, 0) + afEB(1) * aafAbsC(1, 1) + afEB(2) * aafAbsC(1, 2) '
    fR01 = afEA(1) + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If

    ' Axes C0+t*A2
    For i = 0 To 2
        aafC(2, i) = Dot(akA(2), akB(i)) '
        aafAbsC(2, i) = Abs(aafC(2, i)) '
        If (aafAbsC(2, i) > fCutoff) Then
            ExistePaireParallele = True '
        End If
    Next i
    
    afAD(2) = Dot(akA(2), kD) '
    fR = Abs(afAD(2)) '
    fR1 = afEB(0) * aafAbsC(2, 0) + afEB(1) * aafAbsC(2, 1) + afEB(2) * aafAbsC(2, 2) '
    fR01 = afEA(2) + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If

    ' Axes C0+t*B0
    fR = Abs(Dot(akB(0), kD)) '
    fR0 = afEA(0) * aafAbsC(0, 0) + afEA(1) * aafAbsC(1, 0) + afEA(2) * aafAbsC(2, 0) '
    fR01 = fR0 + afEB(0) '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If

    ' Axes C0+t*B1
    fR = Abs(Dot(akB(1), kD)) '
    fR0 = afEA(0) * aafAbsC(0, 1) + afEA(1) * aafAbsC(1, 1) + afEA(2) * aafAbsC(2, 1) '
    fR01 = fR0 + afEB(1) '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*B2
    fR = Abs(Dot(akB(2), kD)) '
    fR0 = afEA(0) * aafAbsC(0, 2) + afEA(1) * aafAbsC(1, 2) + afEA(2) * aafAbsC(2, 2) '
    fR01 = fR0 + afEB(2) '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Au moins une paire  d'axes est parralèle
    If (ExistePaireParallele) Then
        TestIntersectionBoite = True '
        Exit Function
    End If
    
    ' Axes C0+t*A0xB0
    fR = Abs(afAD(2) * aafC(1, 0) - afAD(1) * aafC(2, 0)) '
    fR0 = afEA(1) * aafAbsC(2, 0) + afEA(2) * aafAbsC(1, 0) '
    fR1 = afEB(1) * aafAbsC(0, 2) + afEB(2) * aafAbsC(0, 1) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A0xB1
    fR = Abs(afAD(2) * aafC(1, 1) - afAD(1) * aafC(2, 1)) '
    fR0 = afEA(1) * aafAbsC(2, 1) + afEA(2) * aafAbsC(1, 1) '
    fR1 = afEB(0) * aafAbsC(0, 2) + afEB(2) * aafAbsC(0, 0) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A0xB2
    fR = Abs(afAD(2) * aafC(1, 2) - afAD(1) * aafC(2, 2)) '
    fR0 = afEA(1) * aafAbsC(2, 2) + afEA(2) * aafAbsC(1, 2) '
    fR1 = afEB(0) * aafAbsC(0, 1) + afEB(1) * aafAbsC(0, 0) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A1xB0
    fR = Abs(afAD(0) * aafC(2, 0) - afAD(2) * aafC(0, 0)) '
    fR0 = afEA(0) * aafAbsC(2, 0) + afEA(2) * aafAbsC(0, 0) '
    fR1 = afEB(1) * aafAbsC(1, 2) + afEB(2) * aafAbsC(1, 1) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A1xB1
    fR = Abs(afAD(0) * aafC(2, 1) - afAD(2) * aafC(0, 1)) '
    fR0 = afEA(0) * aafAbsC(2, 1) + afEA(2) * aafAbsC(0, 1) '
    fR1 = afEB(0) * aafAbsC(1, 2) + afEB(2) * aafAbsC(1, 0) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A1xB2
    fR = Abs(afAD(0) * aafC(2, 2) - afAD(2) * aafC(0, 2)) '
    fR0 = afEA(0) * aafAbsC(2, 2) + afEA(2) * aafAbsC(0, 2) '
    fR1 = afEB(0) * aafAbsC(1, 1) + afEB(1) * aafAbsC(1, 0) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A2xB0
    fR = Abs(afAD(1) * aafC(0, 0) - afAD(0) * aafC(1, 0)) '
    fR0 = afEA(0) * aafAbsC(1, 0) + afEA(1) * aafAbsC(0, 0) '
    fR1 = afEB(1) * aafAbsC(2, 2) + afEB(2) * aafAbsC(2, 1) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A2xB1
    fR = Abs(afAD(1) * aafC(0, 1) - afAD(0) * aafC(1, 1)) '
    fR0 = afEA(0) * aafAbsC(1, 1) + afEA(1) * aafAbsC(0, 1) '
    fR1 = afEB(0) * aafAbsC(2, 2) + afEB(2) * aafAbsC(2, 0) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    
    ' Axes C0+t*A2xB2
    fR = Abs(afAD(1) * aafC(0, 2) - afAD(0) * aafC(1, 2)) '
    fR0 = afEA(0) * aafAbsC(1, 2) + afEA(1) * aafAbsC(0, 2) '
    fR1 = afEB(0) * aafAbsC(2, 1) + afEB(1) * aafAbsC(2, 0) '
    fR01 = fR0 + fR1 '
    If (fR > fR01) Then
        TestIntersectionBoite = False '
        Exit Function
    End If
    

    TestIntersectionBoite = True
        
End Function
'----------------------------------------------------------------------------
'----------------------------------------------------------------------------
Function TestIntersectionTriangle(akU As Triangle3, akV As Triangle3) As Boolean
    Dim kDir As Point3
    Dim fUMin As Double
    Dim fUMax As Double
    Dim fVMin As Double
    Dim fVMax As Double
    Dim fMdV0 As Double
    Dim i0, i1 As Integer
    Dim akE As Triangle3
    Dim akF As Triangle3
    
    Dim kN As Point3
    Dim kM As Point3
    Dim kNxM As Point3
    
    Dim fNdU0 As Double
    Const fEpsilon As Double = 0.000001 ' sin(Angle(N,M)) < epsilon
   

       
    ' direction N
    akE.S(0) = VecSub(akU.S(1), akU.S(0))
    akE.S(1) = VecSub(akU.S(2), akU.S(1))
    akE.S(2) = VecSub(akU.S(0), akU.S(2))


    ' Normal du triangle 1
    kN = VecProd(akE.S(0), akE.S(1))
    
    fNdU0 = Dot(kN, akU.S(0))
    Call ProjectionTriangle(kN, akV, fVMin, fVMax)
    If fNdU0 < fVMin Or fNdU0 > fVMax Then
        TestIntersectionTriangle = False
        Exit Function
    End If
    
    ' direction M
    akF.S(0) = VecSub(akV.S(1), akV.S(0))
    akF.S(1) = VecSub(akV.S(2), akV.S(1))
    akF.S(2) = VecSub(akV.S(0), akV.S(2))
    
    
    ' Normal du triangle 1
    kM = VecProd(akF.S(0), akF.S(1))
    
     ' Normal issu de normal Triangle 1 et normal triangle 2
    kNxM = VecProd(kN, kM)
    

    If Dot(kNxM, kNxM) >= (fEpsilon * Dot(kN, kN) * Dot(kM, kM)) Then
        ' les triangles ne sont paralleles
        fMdV0 = Dot(kM, akV.S(0))
        
        Call ProjectionTriangle(kM, akU, fUMin, fUMax)
        If (fMdV0 < fUMin Or fMdV0 > fUMax) Then
                TestIntersectionTriangle = False
                Exit Function
        End If

        ' directions E(i0)xF(i1)
        For i1 = 0 To 2
            For i0 = 0 To 2
                kDir = VecProd(akE.S(i0), akF.S(i1))
                Call ProjectionTriangle(kDir, akU, fUMin, fUMax)
                Call ProjectionTriangle(kDir, akV, fVMin, fVMax)
                If (fUMax < fVMin Or fVMax < fUMin) Then
                    TestIntersectionTriangle = False
                    Exit Function
                End If
            Next i0
        Next i1
    Else  ' triangles sont paralleles (et, coplanair)
        ' directions NxE(i0)
        For i0 = 0 To 2
            kDir = VecProd(kN, akE.S(i0))
            Call ProjectionTriangle(kDir, akU, fUMin, fUMax)
            Call ProjectionTriangle(kDir, akV, fVMin, fVMax)
            If (fUMax < fVMin Or fVMax < fUMin) Then
                TestIntersectionTriangle = False
                Exit Function
            End If
        Next i0

        ' directions NxF(i1)
        For i1 = 0 To 2
            kDir = VecProd(kM, akF.S(i1))
            Call ProjectionTriangle(kDir, akU, fUMin, fUMax)
            Call ProjectionTriangle(kDir, akV, fVMin, fVMax)
            If (fUMax < fVMin Or fVMax < fUMin) Then
                TestIntersectionTriangle = False
                Exit Function
            End If
        Next i1
        
    End If

    TestIntersectionTriangle = True
End Function
'----------------------------------------------------------------------------

