Attribute VB_Name = "Gestion_outil"
'
Sub Rotation_Magasin_Tool(Num_Tool As Integer)
Dim i As Integer
Dim inc As Double
Dim inc_visu As Double



    inc = Num_Tool - Machine.MagasinPo.Valeur_axe
    inc_visu = Abs(inc) * 5
    
    For i = 1 To inc_visu
        ' = Magasin
        Machine.MagasinPo.Valeur_axe = Machine.MagasinPo.Valeur_axe + (inc / inc_visu)
        Call Machine_SIMUL_FRM.Pic_Paint
    Next i
    
   Machine.MagasinPo.Valeur_axe = Num_Tool

End Sub

Function Charge_Tool(Num_Tool As Integer) As Boolean
Dim SVG_pos As position
Dim i As Integer
Dim MvtA As Integer
Dim MvtB As Integer
Dim MvtC As Integer
Dim MvtD As Integer
Dim Mvt As String




If Num_Tool = Machine.Tool_current Then
   'Rien a faire
    Call MessageLOG_CLS
    Call MessageLOG_Print("Itencical tool !!!!!", 2)
Else

If Num_Tool > UBound(ToolC) Then
   'Rien a faire
    Call MessageLOG_CLS
    Call MessageLOG_Print("Inexisting Tool on Current machine", 2)
    Charge_Tool = False
    Exit Function
End If

' aucun Tool dans le magasin donc ne charge pas l'Tool si Num_Tool=0 cas de déchargement
    If ToolC(Num_Tool).Type = 0 And Num_Tool Then
        Call MessageLOG_CLS
        Call MessageLOG_Print("No tool at this place !!!!!", 2)
        Exit Function
    End If
    
    'Sauvegarde position
    For i = 1 To Machine.NB_axe
        SVG_pos.Join(i) = Machine.Element(i).Valeur_axe
    Next i
     Position_precedente = SVG_pos
     
    ' reinitialise le namebre de mouvement dans la macro de chargement de l'Tool
    '   MvtA = Mouvement entre le point d'origine et la position avant le chargement Tool
    '          Position d'attente avant rotation magazin Tool
    '   MvtB = Mouvement lors du chargement Tool
    '   MvtC = Mouvement entre la position de chargement Tool et le retour a la position  d'attente
    '   MvtD = Mouvement pour retourner au point d'origine
    MvtA = Val(mfncGetFromIni("Macro_Tool", "MvtA", Fichier_Machine))
    MvtB = Val(mfncGetFromIni("Macro_Tool", "MvtB", Fichier_Machine))
    MvtC = Val(mfncGetFromIni("Macro_Tool", "MvtC", Fichier_Machine))
    MvtD = Val(mfncGetFromIni("Macro_Tool", "MvtD", Fichier_Machine))
   
   
    '
    '   MvtA = Mouvement entre le point d'origine et la position avant le chargement Tool
    '
    For i = 1 To MvtA
        Mvt = "MvtA" & i
        Call Decode_Mvt(mfncGetFromIni("Macro_Tool", Mvt, Fichier_Machine), Position_accordinge, SVG_pos)
        Call Machine_SIMUL_FRM.Execute_mouvement(True)
        Position_precedente = Position_accordinge
    Next i
       
    ' Cas avec un Tool deja Charger on doit le decharger
      If Machine.Tool_current > 0 Then
         Call Decharge_Tool(SVG_pos)
      End If
    
    ' si T00 alors Dechargement Tool
    If Num_Tool > 0 Then
       If Val(Machine_SIMUL_FRM.Magasin) <> Num_Tool Then
          Machine_SIMUL_FRM.Magasin = Num_Tool
        'Call Rotation_Magasin_Tool(Num_Tool)
       End If
       '
       '   MvtB = Mouvement lors du chargement Tool
       '
        For i = 1 To MvtB
            Mvt = "MvtB" & i
            Call Decode_Mvt(mfncGetFromIni("Macro_Tool", Mvt, Fichier_Machine), Position_accordinge, SVG_pos)
            Call Machine_SIMUL_FRM.Execute_mouvement(True)
            Position_precedente = Position_accordinge
        Next i
        '
        Call Machine_SIMUL_FRM.Execute_mouvement(True)
        ' Update info machine
        Machine.Tool_current = Num_Tool
        
        ' A Modifier pour machine AB
        ' La position pivot depend de la lg Tool+ orientation broche
        Debug.Print PositionPivot.Z
        Machine.LG_Tool_current = ToolC(Machine.Tool_current).LG
        PositionPivot = VecAdd(PositionPivot, Machine.Element(Machine.NB_axe + 1).Vecteur, -ToolC(Machine.Tool_current).LG)
        
        '
        '   MvtC = Mouvement entre la position de chargement Tool et le retour a la position  d'attente
        '
        For i = 1 To MvtC
            Mvt = "MvtC" & i
            Call Decode_Mvt(mfncGetFromIni("Macro_Tool", Mvt, Fichier_Machine), Position_accordinge, SVG_pos)
            Call Machine_SIMUL_FRM.Execute_mouvement(True)
            Position_precedente = Position_accordinge
        Next i
    End If
    
    '
    '   MvtD = Mouvement pour retourner au point d'origine
    '
    Call Machine_SIMUL_FRM.Execute_mouvement(True)
    For i = 1 To MvtD
           Mvt = "MvtD" & i
           Call Decode_Mvt(mfncGetFromIni("Macro_Tool", Mvt, Fichier_Machine), Position_accordinge, SVG_pos)
           Call Machine_SIMUL_FRM.Execute_mouvement(True)
           Position_precedente = Position_accordinge
    Next i
        
    
    End If
    
    Call GetPoint
    Call Affiche_coord
    Call Sauvegarde_Position
    Position_Actuelle.Coord = Pt0
    
    Charge_Tool = True
End Function
Sub Decharge_Tool(SVG As position)
Dim i As Integer
Dim MvtB As Integer
Dim MvtC As Integer
Dim Mvt As String
MvtB = Val(mfncGetFromIni("Macro_Tool", "MvtB", Fichier_Machine))
MvtC = Val(mfncGetFromIni("Macro_Tool", "MvtC", Fichier_Machine))


  If Machine.Tool_current > 0 And Val(Machine_SIMUL_FRM.Magasin) <> Machine.Tool_current Then
     Call Rotation_Magasin_Tool(Machine.Tool_current)
  End If
  
    For i = 1 To MvtB
        Mvt = "MvtB" & i
        Call Decode_Mvt(mfncGetFromIni("Macro_Tool", Mvt, Fichier_Machine), Position_accordinge, SVG)
        Call Machine_SIMUL_FRM.Execute_mouvement(True)
        Position_precedente = Position_accordinge
    Next i

Machine_SIMUL_FRM.Magasin = Machine.Tool_current
PositionPivot.Z = PositionPivot.Z + ToolC(Machine.Tool_current).LG
Machine.Tool_current = 0
  
    For i = 1 To MvtC
        Mvt = "MvtC" & i
        Call Decode_Mvt(mfncGetFromIni("Macro_Tool", Mvt, Fichier_Machine), Position_accordinge, SVG)
        Call Machine_SIMUL_FRM.Execute_mouvement(True)
        Position_precedente = Position_accordinge
    Next i
End Sub
