Attribute VB_Name = "ISO"
Option Explicit

'Arret de simulation si appui sur ESC
Public stop_exec As Boolean
         
'Renumerotation du fichier ISO
Function Renumerotation(Fenetre As Object) As Boolean

'Retour chariot fin de ligne en MSDOS
Dim newline As String

Dim Start As Double
Dim Donnéeslues As String
Dim message_erreur As String
Dim couleur As Single
Dim nb_ligne As Long

stop_exec = False

' Initialisation
newline = Chr(13) + Chr(10)
Fenetre.SelStart = 0
Fenetre.SelLength = 0

nb_ligne = 0


Do

    nb_ligne = nb_ligne + 1
    Start = Fenetre.SelStart
    
    Fenetre.Span newline, True, True
    Fenetre.SelLength = Fenetre.SelLength + Len(newline)
    Debug.Print Fenetre.SelLength
    Donnéeslues = SupprimeNumeroISO(Fenetre.SelText)
    
    If Len(Donnéeslues) > 0 Then
        Donnéeslues = "N" & nb_ligne & Donnéeslues
    End If
    
    
    
    'Decode le texte
    couleur = 16744576
    'Change la couleur du texte
    Fenetre.SelColor = couleur
    'Réaffecte le nouveau code
    Fenetre.SelText = Donnéeslues
    
   ' permet de stopper l'operation
   DoEvents
   If stop_exec Then
    Exit Function
   End If

Loop Until Fenetre.SelStart = Start
    
Renumerotation = True
Exit Function

Erreur:
    Renumerotation = False
End Function
Function SupprimeNumeroISO(Chaine As String) As String
Dim Chainetraitee As String

Chainetraitee = LTrim(Chaine)

' test si le premier charactere de la chaine est le N
If Mid(Chainetraitee, 1, 1) = "N" Then

Encore:
Chainetraitee = Mid(Chainetraitee, 2, Len(Chainetraitee) - 1)
             Select Case Mid(Chainetraitee, 1, 1)
                Case "0" To "9"
                    GoTo Encore
                Case Else
                       Debug.Print "Chainetraitee ="; Chainetraitee
                
            End Select
End If


SupprimeNumeroISO = Chainetraitee
End Function

'Simulation du fichier ISO
Function Simul_Fichier(Fenetre As Object) As Boolean

'Retour chariot fin de ligne en MSDOS
Dim newline As String

Dim Start As Double
Dim CodeG As Integer
Dim CodeM As Integer
Dim CodeS As Integer
Dim CodeF As Integer
Dim CodeT As Integer
Dim Donnéeslues As String
Dim xyzac As Interpolation
Dim xyzac_lu As Interpolation
Dim Uvw As AxePositionne
Dim couleur As Single
Dim nb_ligne As Long
Dim SVG_position As Interpolation


ReDim Parcours(0)
stop_exec = False

' Initialisation
 ' init de la premiere interpolation
CodeG = 999

newline = Chr(13) + Chr(10)
Fenetre.SelStart = 0
Fenetre.SelLength = 0

nb_ligne = 0


Do

    nb_ligne = nb_ligne + 1
    
    Start = Fenetre.SelStart
    Fenetre.Span newline, True, True
    Fenetre.SelLength = Fenetre.SelLength + Len(newline)

    Donnéeslues = Nettoyage_texte_ISO(Fenetre.SelText)
    CodeM = 99
    'Decode le texte
    xyzac_lu = SVG_position
    couleur = Decodage(Donnéeslues, xyzac_lu, Uvw, CodeG, CodeM, CodeF, CodeS, CodeT)
    SVG_position = xyzac_lu
    'Change la couleur du texte
    Fenetre.SelColor = couleur
    
    Select Case CodeM
    
    'Code M06= Chargement Tool
        Case 6
            If Not Charge_Tool(CodeT) Then
                Call MessageLOG_Print("Stop simulation on error tool load")
                GoTo Erreur
            End If
            

     'Code M30= fin de programme et decharge l'Tool current
        Case 30
            Call Charge_Tool(0)
     End Select
        
        
    ' Deplacement
    ' Si G1 ou G0
    Select Case CodeG
    
    Case 0, 1
       'Calcul Selon Decalage et RTCP (ou correction Lg Tool)
      
            If Not Execute_Deplacement(xyzac_lu, Uvw) Then
               Call MessageLOG_Print("Error on  : " & Fenetre.SelText)
               Fenetre.SelColor = 255
               Fenetre.SelBold = True
               Exit Function
            End If

        Parcours(UBound(Parcours)) = xyzac_lu.Coord
        ReDim Preserve Parcours(UBound(Parcours) + 1)
     
     
     
     
     ' modification origine programme
     Case 55
        Machine_SIMUL_FRM.AxeOrigine(1) = xyzac_lu.Coord.X
        Machine_SIMUL_FRM.AxeOrigine(2) = xyzac_lu.Coord.Y
        Machine_SIMUL_FRM.AxeOrigine(3) = xyzac_lu.Coord.Z
        OrigineProg.X = xyzac_lu.Coord.X
        OrigineProg.Y = xyzac_lu.Coord.Y
        OrigineProg.Z = xyzac_lu.Coord.Z
        Piece.Origine = xyzac_lu.Coord
        
        CodeG = 999
    ' DesActivation RTCP
    Case 150
         Machine_SIMUL_FRM.OptionRTCP.Value = 0
        CodeG = 150
        
    ' Activation RTCP
     Case 151
        Machine_SIMUL_FRM.OptionRTCP.Value = 1
        CodeG = 151
        

    
    
    End Select
 

   If stop_exec Then
    Exit Function
   End If
   


Fenetre.SelStart = Fenetre.SelStart + Fenetre.SelLength



Loop Until Fenetre.SelStart = Start
    
Simul_Fichier = True
Exit Function



Erreur:
    Simul_Fichier = False
End Function

' Remplace G0 en G00 pour eviter les problèmes de confusion
Function mG00Code(strString)
       Dim strResult
       Dim intIndex As Integer
       
       strResult = strString
       
 intIndex = InStr(1, strString, "G0")
 '1, "G0", strString)
 If intIndex Then
 'Debug.Print Mid(strString, intIndex + 2, 1)
 Select Case Mid(strString, intIndex + 2, 1)
 
 Case "0" To "9"
   'Do nothing
   strResult = strString
   
 Case Else
      strResult = Mid(strString, 1, intIndex + 1) + "0" + Mid(strString, intIndex + 2, Len(strString) - intIndex)
 End Select

End If
mG00Code = strResult
End Function
'filtre les chaine pour éviter de retourner des valeurs nulles dans le cas de chaine
Public Sub Filtre_Code_iso(Chaine As String, valeur As Double)
             Select Case Mid(Chaine, 1, 1)
                Case "0" To "9"
                    valeur = Val(TokLeftLeft(Chaine, " "))
                Case "-", "+"
                    If Val(TokLeftLeft(Chaine, " ")) Then
                        valeur = Val(TokLeftLeft(Chaine, " "))
                    End If
                Case Else
                       Debug.Print "Commentaire $"; Chaine
                
            End Select
End Sub

       '     '**********************************
       '     'changes all strOrigChar
       '     ' to strReplaceChar
       '     ' in strString
       '     '**********************************
Function mSpaceCode(strString)
       Dim strResult
       strResult = ""
       '     'traverse string
       Dim intIndex
        For intIndex = 1 To Len(strString)
              Select Case Mid(strString, intIndex, 1)
              Case "0" To "9", "."
                                        '*************
                                        'match found
                                        '*************
                                        'MsgBox "found in" + strString
                                        ' Debug.Print Mid(strString, intIndex + 1, 1)
                                         Select Case Mid(strString, intIndex + 1, 1)
                                               Case "0" To "9", ".", "=", ",", ")", "(", " ", "[", "]"  'Is this character a number or decimal?

                                                 '*************
                                                 'no modification
                                                 '*************
                                                 strResult = strResult + Mid(strString, intIndex, 1)

                                                
                                               Case Else
                                                strResult = strResult + Mid(strString, intIndex, 1) + " " ' Add it to the string being built
                                                intIndex = intIndex '+ 1
                     
                                         End Select
                            

                Case Else
                    '*************
                    'no match
                    '*************
                     strResult = strResult + Mid(strString, intIndex, 1)
               End Select
      Next

mSpaceCode = strResult
End Function

' Nettoyage d'une ligne ISO
Function Nettoyage_texte_ISO(Chaine As String) As String
  
    'Debug.Print "Debut ->" & Len(Chaine) & " |" & Chaine & "|"
    Chaine = mReplaceCharacter(Chr(9), Chr(32), Chaine) 'Remplace les tabulations par des espaces
    
    Chaine = mSpaceCode(Chaine)  'Rajoute des espaces dans la chaine pour faciliter le traitement
    Chaine = mG00Code(Chaine)  ' Transforme les G0 en G00 pour éviter les problèmes de décodage
    'Debug.Print "Apres ->" & Len(Chaine) & " |" & Chaine & "|"
     
    Chaine = LTrim(Chaine) ' Suprime les espaces de gauche.
    Chaine = mReplaceCharacter(Chr(13), "", Chaine) 'Supprime Chr(13)
    Chaine = mReplaceCharacter(Chr(10), "", Chaine) 'Supprime Chr(10)

   
    'Debug.Print "Fin ->" & Len(Chaine) & " |" & Chaine & "|"
    
    Nettoyage_texte_ISO = Chaine
End Function

' DECODAGE d'une ligne de code ISO
Public Function Decodage(Ligne As String, Interpo As Interpolation, Axe_option As AxePositionne, CodeG As Integer, CodeM As Integer, CodeF As Integer, CodeS As Integer, CodeT As Integer) As Single
 Dim Chaine As String
 'Dim ChaineTraitee As String

'commentaires
        If InStr(Ligne, "(") <> 0 Or InStr(Ligne, ")") <> 0 Then
            Decodage = 4144959
            Exit Function
        End If
        
        
        If InStr(Ligne, "G") <> 0 Then
            Chaine = TokRightRight(Ligne, "G")
            CodeG = Val(TokLeftLeft(Chaine, " "))
        End If
        
        If InStr(Ligne, "M") <> 0 Then
            Chaine = TokRightRight(Ligne, "M")
            CodeM = Val(TokLeftLeft(Chaine, " "))
            
            CodeG = 9999
            
            Select Case CodeM
             'M06
             Case 6
                Decodage = 32768
             'M02
             Case 2
                Decodage = 33023
             'M30
             Case 30
                Decodage = 16711935
            End Select
        End If
      'Tool
        If InStr(Ligne, "T") <> 0 Then
            Chaine = TokRightRight(Ligne, "T")
            CodeT = Val(TokLeftLeft(Chaine, " "))
        End If
        'VITESSE
        If InStr(Ligne, "F") <> 0 Then
            Chaine = TokRightRight(Ligne, "F")
            CodeF = Val(TokLeftLeft(Chaine, " "))
        End If
        
    'coordonnées
    
        If InStr(Ligne, "X") <> 0 Then
            Chaine = TokRightRight(Ligne, "X")
            Interpo.Coord.X = Val(TokLeftLeft(Chaine, " "))
        End If
        
        If InStr(Ligne, "Y") <> 0 Then
            Chaine = TokRightRight(Ligne, "Y")
             Interpo.Coord.Y = Val(TokLeftLeft(Chaine, " "))
        End If
        
        If InStr(Ligne, "Z") <> 0 Then
            Chaine = TokRightRight(Ligne, "Z")
             Interpo.Coord.Z = Val(TokLeftLeft(Chaine, " "))
        End If
        
        ' C en ISO
        If InStr(Ligne, "C") <> 0 Then
            Chaine = TokRightRight(Ligne, "C")
             Interpo.Pos.C = Val(TokLeftLeft(Chaine, " "))
             'Debug.Print xyzac_stock.Pos.C
        End If
        
        ' B  en ISO
        If InStr(Ligne, "B") <> 0 Then
            Chaine = TokRightRight(Ligne, "B")
             Interpo.Pos.B = Val(TokLeftLeft(Chaine, " "))
        End If
  
        ' A  en ISO
        If InStr(Ligne, "A") <> 0 Then
            Chaine = TokRightRight(Ligne, "A")
             Interpo.Pos.A = Val(TokLeftLeft(Chaine, " "))
        End If
        
        ' Traitement des axes optionelles
        ' U en ISO
        If InStr(Ligne, "U") <> 0 Then
             Chaine = TokRightRight(Ligne, "U")
             Axe_option.U = Val(TokLeftLeft(Chaine, " "))
        End If
        
        ' V  en ISO
        If InStr(Ligne, "V") <> 0 Then
             Chaine = TokRightRight(Ligne, "V")
             Axe_option.V = Val(TokLeftLeft(Chaine, " "))
        End If
  
        ' W  en ISO
        If InStr(Ligne, "W") <> 0 Then
             Chaine = TokRightRight(Ligne, "W")
             Axe_option.W = Val(TokLeftLeft(Chaine, " "))
        End If
        
        
FinSub:
            Select Case CodeG
             'G1
             Case 1
                Decodage = 16711680
             'G0
             Case 0
                Decodage = 255
            'G55
             Case 55
                Decodage = 14653050
            'G150
            Case 150
                Decodage = 33023 'orange
            'G151
            Case 151
                Decodage = 33023 'orange
                
            End Select
            
 
End Function

Function TransCoord(ByRef Interpo As Interpolation, ByRef DEC As DecOrigine) As Interpolation
Dim p1 As Point3
Dim Mx2(2, 2) As Double

        If Piece_charger Then
            TransCoord.Coord = Trans_Matrix_1_16(Piece.Matrix, Interpo.Coord)
        Else
            
            p1.X = DEC.X
            p1.Y = DEC.Y
            p1.Z = DEC.Z
            
            
            Mx2(0, 0) = Machine.Element(0).Matrix(1)
            Mx2(1, 0) = Machine.Element(0).Matrix(2)
            Mx2(2, 0) = Machine.Element(0).Matrix(3)
            Mx2(0, 1) = Machine.Element(0).Matrix(5)
            Mx2(1, 1) = Machine.Element(0).Matrix(6)
            Mx2(2, 1) = Machine.Element(0).Matrix(7)
            
            Mx2(0, 2) = Machine.Element(0).Matrix(9)
            Mx2(1, 2) = Machine.Element(0).Matrix(10)
            Mx2(2, 2) = Machine.Element(0).Matrix(11)


            TransCoord.Coord = Trans_Matrix(Mx2, p1, Interpo.Coord)
        End If
        
     TransCoord.Pos = Interpo.Pos

End Function


