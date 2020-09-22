Attribute VB_Name = "ChargerMachine"
Option Explicit

Public Fichier_Machine As String ' fichier dat machine ( definition machine )
Public Indice_Machine As String  ' indice machine du fichier ini
Public Fichier_Ini As String ' Fichier machine_simul.ini ( declaration des machines)

Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

'INI File Functions
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

' Les Tools
'Public Tool() As Tool
' Les portes Tools
'Public POTool() As PO_Tool'' Definition de différent porte-Tool

Sub Init_PO(Fichier_Def As String, PTool() As PO_Tool)
Dim i As Integer
Dim J As Integer
Dim NB_Tool As Integer
Dim Section_Tool As String
Dim Indice_point As String
Dim TabPt

NB_Tool = Val(mfncGetFromIni("Machine", "Nb_Tool", Fichier_Def))

For i = 1 To NB_Tool
    Section_Tool = "Porte_Tool" & i
    PTool(i).name = Val(mfncGetFromIni(Section_Tool, "name", Fichier_Def))
    PTool(i).Remarque = Val(mfncGetFromIni(Section_Tool, "Remarque", Fichier_Def))
    
    PTool(i).NB_point = Val(mfncGetFromIni(Section_Tool, "NB_Point", Fichier_Def))
    PTool(i).Origine.Z = Val(mfncGetFromIni(Section_Tool, "OrigineZ", Fichier_Def))
    PTool(i).Dec_Z = Val(mfncGetFromIni(Section_Tool, "Dec_Z", Fichier_Def))
    ReDim PTool(i).Coord(PTool(i).NB_point)
    
    For J = 1 To PTool(i).NB_point
        Indice_point = "P" & J
        TabPt = Split(mfncGetFromIni(Section_Tool, Indice_point, Fichier_Def), ",")
        With PTool(i)
            .Coord(J).X = Val(TabPt(0))
            .Coord(J).Y = Val(TabPt(1))
        End With
    Next J
Next i


End Sub
'def plan control
Sub Def_Plan(Zgrille As Double)
    Dim P1 As Point3
    
    P1.X = -1000
    P1.Y = -1000
    P1.Z = Val(Zgrille)
    
    Ta1.S(0) = P1
    Ta1.S(1) = P1
    Ta1.S(2) = P1
    
    Ta1.S(1).X = -P1.X
    Ta1.S(2).Y = -P1.Y
    
    Ta2 = Ta1
    
    Ta2.S(0).X = -P1.X
    Ta2.S(0).Y = -P1.Y
End Sub

'Decode une instruction de mouvement pour une macro machine ( exemple chargement Tool)
'
' On retrouve les axes machines et pour chaques axes
'   # veut dire ne change pas la valeur de l'axe = Valeur de l'axe donnée par Pos
'   C reprend la valeur original avant la macro ( point d'usinage ) donnée par Pos_SVG
'   Autre Indique la nouvelle valeur de l'axe'
'
Sub Decode_Mvt(Chaine As String, Pos As position, Pos_SVG As position)
Dim TabSplt
Dim i As Integer

TabSplt = Split(Chaine, ",")

For i = 0 To UBound(TabSplt)
    Select Case TabSplt(i)
    
    Case "#"
    
    Case "C"
        Pos.Join(i + 1) = Pos_SVG.Join(i + 1)
    Case Else
        Pos.Join(i + 1) = Val(TabSplt(i))
    End Select
Next i

End Sub

Sub init_Tool(Fichier_Def As String, Outl() As Tool)
Dim i As Integer
Dim NB_Tool As Integer
Dim Section_Tool As String


NB_Tool = Val(mfncGetFromIni("Machine", "Nb_Tool", Fichier_Def))

With Machine_SIMUL_FRM
If .TreeViewTOOL.Nodes.Count > 1 Then
        .TreeViewTOOL.Nodes.Clear
        .TreeViewTOOL.Nodes.Add , tvwLast, "LIB", NB_Tool & " Tool(s)", 1, 1

    Else
        .TreeViewTOOL.Nodes(1).Text = NB_Tool & " Tool(s)"
        .TreeViewTOOL.Nodes(1).Expanded = True
End If
.TreeViewTOOL.Nodes(1).Expanded = True
.TreeViewTOOL.Nodes(1).Bold = True  'gras

' Type Tool
'   Type = 1  ' Ball
'   Type = 2  ' FlatTool
'   Type = 3  ' Drill
For i = 1 To NB_Tool
    Section_Tool = "Tool" & i
    Outl(i).Type = Val(mfncGetFromIni(Section_Tool, "Type", Fichier_Def))
    Outl(i).name = Val(mfncGetFromIni(Section_Tool, "name", Fichier_Def))
    Outl(i).Remarque = Val(mfncGetFromIni(Section_Tool, "Remarque", Fichier_Def))
    
    Outl(i).Diameter = Val(mfncGetFromIni(Section_Tool, "Diameter", Fichier_Def))
    Outl(i).CornerRadius = Val(mfncGetFromIni(Section_Tool, "CornerRadius", Fichier_Def))
    Outl(i).LG = Val(mfncGetFromIni(Section_Tool, "LG", Fichier_Def))
    Outl(i).LG_Coupe = Val(mfncGetFromIni(Section_Tool, "LG_Coupe", Fichier_Def))
    Outl(i).LgCone = Val(mfncGetFromIni(Section_Tool, "LgCone", Fichier_Def))   ' pour Drill
    Outl(i).DiameterCorp = Val(mfncGetFromIni(Section_Tool, "DiameterCorp", Fichier_Def))
        
        
    .TreeViewTOOL.Nodes.Add "LIB", tvwChild, Section_Tool, Generation_name_Tool(Outl(i), i), Outl(i).Type + 1, Outl(i).Type + 1
    
    
    ' case des emplacements vide dans le magasin
    If Outl(i).Type = 0 Then
        .TreeViewTOOL.Nodes(i + 1).Bold = True 'gras
        .TreeViewTOOL.Nodes(i + 1).BackColor = &HFFC0C0 'bleu cyan  ' &HFF& ' Rouge
    End If
    
Next i

End With

End Sub
'Generate de name according infos Tool
Function Generation_name_Tool(Outl_current As Tool, Num As Integer) As String
Dim name As String

name = "T" & Num & "  "

If Outl_current.name <> "0" Then
          If Outl_current.Type = 0 Then
            name = "Empty Place"
          Else
            name = name & Outl_current.name
          End If
Else
            ' Type Tool
            '   Type = 1  ' Ball
            '   Type = 2  ' FlatTool
            '   Type = 3  ' Drill
            Select Case Outl_current.Type
             Case 0
                name = "Empty Place"
             Case 1
                name = name & "D" & Outl_current.Diameter
            
             Case 2
                name = name & "D" & Outl_current.Diameter & "R" & Outl_current.CornerRadius
                
             Case 3
                name = name & "Drill " & Outl_current.Diameter
                
            Case Else
                
                
            End Select
End If

Generation_name_Tool = name
End Function
Public Sub Charger_Machine(Fichier_Def As String, Mach As Machine3D)
Dim NBElement As Integer
Dim i As Integer
Dim SectionElement As String
Dim Color As Long

'
NBElement = Val(mfncGetFromIni("Machine", "Element", Fichier_Def))
ReDim Mach.Element(NBElement)

' Affichage depart
xm_base = Val(mfncGetFromIni("Machine", "Xm_base", Fichier_Def))
ym_base = Val(mfncGetFromIni("Machine", "Ym_base", Fichier_Def))
zm_base = Val(mfncGetFromIni("Machine", "Zm_base", Fichier_Def))
    
Zoom_base = Val(mfncGetFromIni("Machine", "Zoom_base", Fichier_Def))
    
PosX_base = Val(mfncGetFromIni("Machine", "PosX_base", Fichier_Def))
PosY_base = Val(mfncGetFromIni("Machine", "PosY_base", Fichier_Def))

Color = Val(mfncGetFromIni("Machine", "Couleur_Piece", Fichier_Def)) ' recuperation couleur piece
If Color = 0 Then Color = 33023 ' init couleur orange pour piece
Piece.Color = GetRVB(Color)

xm = xm_base
ym = ym_base
zm = zm_base
Zoom = Zoom_base
PosX = PosX_base
PosY = PosY_base

'Grille base
GrilleX = Val(mfncGetFromIni("Machine", "GrilleX", Fichier_Def))
If GrilleX = 0 Then GrilleX = 50
GrilleY = Val(mfncGetFromIni("Machine", "GrilleY", Fichier_Def))
If GrilleY = 0 Then GrilleY = 50
PasGrille = Val(mfncGetFromIni("Machine", "PasGrille", Fichier_Def))
If PasGrille = 0 Then PasGrille = 20
    
    
Mach.name = mfncGetFromIni("Machine", "name", Fichier_Def)
Mach.Type = Val(mfncGetFromIni("Machine", "Type", Fichier_Def))
Mach.NB_axe = Val(mfncGetFromIni("Machine", "NB_axe", Fichier_Def))
Mach.Element_Fixe = Val(mfncGetFromIni("Machine", "Element_Fixe", Fichier_Def))
Mach.Element_Collision = Val(mfncGetFromIni("Machine", "Element_Collision", Fichier_Def))
' Reinit objet
'Type_axe = 0 => Translation
'Type_axe = 1 => rotation
For i = 0 To NBElement
    'name de la section
    SectionElement = "Element" & i
       
    Mach.Element(i).name = mfncGetFromIni(SectionElement, "name", Fichier_Def)
    Mach.Element(i).fichier = mfncGetFromIni(SectionElement, "Fichier", Fichier_Def)

    Mach.Element(i).Color = GetRVB(QBColor(Val(mfncGetFromIni(SectionElement, "Couleur", Fichier_Def))))

    Mach.Element(i).MiniAxe = Val(mfncGetFromIni(SectionElement, "Mini_axe", Fichier_Def))
    Mach.Element(i).MaxiAxe = Val(mfncGetFromIni(SectionElement, "Maxi_axe", Fichier_Def))
    
    Mach.Element(i).Type_axe = Val(mfncGetFromIni(SectionElement, "Type_axe", Fichier_Def))
    Mach.Element(i).Origine.X = Val(mfncGetFromIni(SectionElement, "Origine_X", Fichier_Def))
    Mach.Element(i).Origine.Y = Val(mfncGetFromIni(SectionElement, "Origine_Y", Fichier_Def))
    Mach.Element(i).Origine.Z = Val(mfncGetFromIni(SectionElement, "Origine_Z", Fichier_Def))
    Mach.Element(i).Vecteur.X = Val(mfncGetFromIni(SectionElement, "Vecteur_X", Fichier_Def))
    Mach.Element(i).Vecteur.Y = Val(mfncGetFromIni(SectionElement, "Vecteur_Y", Fichier_Def))
    Mach.Element(i).Vecteur.Z = Val(mfncGetFromIni(SectionElement, "Vecteur_Z", Fichier_Def))
Next i

'Magasin porte Tool
Mach.PositionMagasin = Val(mfncGetFromIni("Magasin", "PositionMagasin", Fichier_Def))
Mach.MagasinPo.MiniAxe = Val(mfncGetFromIni("Magasin", "Mini_axe", Fichier_Def))
Mach.MagasinPo.MaxiAxe = Val(mfncGetFromIni("Magasin", "Maxi_axe", Fichier_Def))
    
'Magasin Tool
If Mach.PositionMagasin Then
    Mach.MagasinPo.name = mfncGetFromIni("Magasin", "name", Fichier_Def)
    Mach.MagasinPo.fichier = mfncGetFromIni("Magasin", "Fichier", Fichier_Def)
    Mach.MagasinPo.Color = GetRVB(QBColor(Val(mfncGetFromIni("Magasin", "Couleur", Fichier_Def))))
    Mach.MagasinPo.Type_axe = Val(mfncGetFromIni("Magasin", "Type_axe", Fichier_Def))
    Mach.MagasinPo.Origine.X = Val(mfncGetFromIni("Magasin", "Origine_X", Fichier_Def))
    Mach.MagasinPo.Origine.Y = Val(mfncGetFromIni("Magasin", "Origine_Y", Fichier_Def))
    Mach.MagasinPo.Origine.Z = Val(mfncGetFromIni("Magasin", "Origine_Z", Fichier_Def))
    Mach.MagasinPo.Vecteur.X = Val(mfncGetFromIni("Magasin", "Vecteur_X", Fichier_Def))
    Mach.MagasinPo.Vecteur.Y = Val(mfncGetFromIni("Magasin", "Vecteur_Y", Fichier_Def))
    Mach.MagasinPo.Vecteur.Z = Val(mfncGetFromIni("Magasin", "Vecteur_Z", Fichier_Def))
End If

End Sub


'****************************************************************
' Name: .INI read/write routines
' Description:.INI read/write routines
'mfncGetFromIni-- Reads from an *.INI file strFileName(full path & file name)
'mfncWriteIni--Writes to an *.INI file called strFileName (full path & file name)
'****************************************************************
Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFileName As String) As String
       '*** DESCRIPTION:Reads from an *.INI file strFileName (fullpath & file name)
       '     '*** RETURNS:The string stored in [strSectionHeader], line  beginning strVariableName=
       '     '*** NOTE: Requires declaration of API call GetPrivateProfileString
       '     'Initialise variable
       Dim strReturn As String
       '     'Blank the return string
       strReturn = String(255, Chr(0))
       '     'Get requested information, trimming the returned string
       mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFileName))
End Function
Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFileName As String) As Integer
       '*** DESCRIPTION:Writes to an *.INI file called strFileName
       '     (full       path & file name)
       '*** RETURNS:Integer indicating failure (0) or success (other)       to write
       '     '*** NOTE: Requires declaration of API call     WritePrivateProfileString
       '     'Call the API
       mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFileName)
End Function
Function mfncDeleteIniKey(strSectionHeader As String, strVariableName As String, strFileName As String) As Integer
       '*** DESCRIPTION:Writes to an *.INI file called strFileName
       '     (full       path & file name)
       '*** RETURNS:Integer indicating failure (0) or success (other)       to write
       '     '*** NOTE: Requires declaration of API call     WritePrivateProfileString
       '     'Call the API
       mfncDeleteIniKey = WritePrivateProfileString(strSectionHeader, strVariableName, 0&, strFileName)
End Function
Function mfncDeleteIniSection(strSectionHeader As String, strFileName As String) As Integer
       '*** DESCRIPTION:Writes to an *.INI file called strFileName
       '     (full       path & file name)
       '*** RETURNS:Integer indicating failure (0) or success (other)       to write
       '     '*** NOTE: Requires declaration of API call     WritePrivateProfileString
       '     'Call the API
       mfncDeleteIniSection = WritePrivateProfileString(strSectionHeader, 0&, 0&, strFileName)
End Function
Public Function GetIniInt(strSectionHeader As String, Key As String, strFileName As String, Optional Default As Long) As Long
    GetIniInt = GetPrivateProfileInt(strSectionHeader, Key, Default, strFileName)
End Function

