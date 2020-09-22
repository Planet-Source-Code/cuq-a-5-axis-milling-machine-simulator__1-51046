Attribute VB_Name = "OpenGL"
Option Explicit

Dim HGLRC As Long

' Variables pour la manipulation 3D opengl
Global Yangle As GLfloat
Global Langle As Integer
Global PosX As GLfloat
Global PosY As GLfloat
Global SVGposX As GLfloat
Global SVGposY As GLfloat
Global Zoom As Single
Global Zoom_base As Single
Global PosX_base As GLfloat
Global PosY_base As GLfloat
Global xm As Single, ym As Single, zm As Single
Global xm_base As Single, ym_base As Single, zm_base As Single

' Calcul des orientations broche et point pivot avec les fonctions opengl
Global Pt0 As Point3
Global Vx As Point3
Global Vy As Point3
Global Vz As Point3

Public Vecteur_vue As Point3 ' vecteur de vue Opengl

Global ObjectAlpha As Single 'Tranparence de la machine
Global Ta1 As Triangle3
Global Ta2 As Triangle3

Global Render As Integer

Public GrilleY As Integer
Public GrilleX As Integer
Public PasGrille As Integer







'Reglage du format de pixel
Sub SetupPixelFormat(ByVal hdc As Long)
    Dim PFD As PIXELFORMATDESCRIPTOR
    Dim PixelFormat As Integer
    PFD.nSize = Len(PFD)
    PFD.nVersion = 1
    PFD.dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
    PFD.iPixelType = PFD_TYPE_RGBA
    PFD.cColorBits = 16
    PFD.cDepthBits = 16
    PFD.iLayerType = PFD_MAIN_PLANE
    PixelFormat = ChoosePixelFormat(hdc, PFD)
    If PixelFormat = 0 Then MsgBox ("Pixel format unknown")
    SetPixelFormat hdc, PixelFormat, PFD
End Sub

'lumiere
Sub Lampe()
    Dim aflLightAmbient(4) As GLfloat
    Dim aflLightDiffuse(4) As GLfloat
    Dim aflLightPosition(4) As GLfloat
    Dim aflLightSpecular(4) As GLfloat
    
   
   glDisable glcLighting 'temporarily disable lighting
   glPushMatrix 'push the matrix stack down by one
   
    glEnable glcLight0
    

      
    ' Ambient light settings
    aflLightAmbient(0) = 1
    aflLightAmbient(1) = 1
    aflLightAmbient(2) = 1
    aflLightAmbient(3) = 1
    ' Diffuse light settings
    aflLightDiffuse(0) = 1
    aflLightDiffuse(1) = 1
    aflLightDiffuse(2) = 1
    aflLightDiffuse(3) = 1
    ' Light position settings
    'aflLightPosition(0) = 3000
    'aflLightPosition(1) = -1000
    'aflLightPosition(2) = -1000
    
    aflLightPosition(0) = Machine_SIMUL_FRM.HScroll(0)
    aflLightPosition(1) = Machine_SIMUL_FRM.HScroll(1)
    aflLightPosition(2) = Machine_SIMUL_FRM.HScroll(2)
    aflLightPosition(3) = 1
      
    ' Light position Specular
    aflLightSpecular(0) = 0
    aflLightSpecular(1) = 0
    aflLightSpecular(2) = 0
    aflLightSpecular(3) = 1
    
    ' Set the light's ambient and diffuse levels and its position
    glLightfv ltLight0, lpmAmbient, aflLightAmbient(0)
    glLightfv ltLight0, lpmDiffuse, aflLightDiffuse(0)
    glLightfv ltLight0, lpmSpecular, aflLightSpecular(0)
    glLightfv ltLight0, lpmPosition, aflLightPosition(0)
    
    ' Enable light0
    'glEnable glcLight0
    
            glPopMatrix 'pop the matrix stack up by one
    glEnable glcLighting    're-enable lighting
    
End Sub
' initialisation OpenGL
Sub LoadGL(p As PictureBox)
Dim i As Integer

    SetupPixelFormat p.hdc
    HGLRC = wglCreateContext(p.hdc)
    wglMakeCurrent p.hdc, HGLRC
    
    ' ????
    glEnable glcColorMaterial 'Allow material parameters track the current color
    glColorMaterial faceFront, cmmAmbientAndDiffuse 'Enable color tracking
    glClearDepth 1 'Set the clear value for the depth buffer
    
    
    '' Tres important pour transparance
    glEnable glcBlend 'enable alpha blending
    glBlendFunc sfSrcAlpha, dfOneMinusSrcAlpha 'set alpha blending options
    'glEnable glcDepthTest   'Enable depth comparisons
    glEnable glcColorMaterial

    ' Smooth shading
    'glShadeModel smSmooth
    glShadeModel smFlat
    
    ' Set the clear colour gris
    glClearColor 0.5, 0.5, 0.5, 0 '
    ' Set the clear depth
    glClearDepth 1#
    
    ' Enable Z-buffer
    glEnable glcDepthTest
    
    ' Set test type
     glDepthFunc cfLEqual
     
     'important
     ' Best perspective correction
     glHint htPerspectiveCorrectionHint, hmNicest
    
    'Meilleur Ombrage
    'Sinon les couleurs sont trop vive trop de contraste
     glEnable glcNormalize
    
    ' Init des traits cachés
    glEnable glcPolygonOffsetFill
    glPolygonOffset 2, 1
    'glPolygonMode faceFrontAndBack, pgmFILL

    glMatrixMode mmProjection
    glLoadIdentity
    gluPerspective 10, p.ScaleWidth / p.ScaleHeight, 1, 1000
    
    glMatrixMode mmModelView
    glLoadIdentity
    
    ' definition des lumieres
    Call Lampe
End Sub


Sub DessineMachine(Pict As PictureBox, Grille As Boolean, Render_mode As Integer, Option_tracer As Boolean, Option_box As Boolean)
Dim K As Integer
Dim Pt1 As Point3
Dim MMatrix(1 To 16) As Double
Dim BoxO As Box3
Dim i As Integer

On Error Resume Next

'---DEBUT :
    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
   ' glMatrixMode mmModelView

'---PETITE INITALISATION :
    ' glViewport 0, 0, Pict.ScaleWidth, Pict.ScaleHeight
     
'---POINT DE VUE : Camera
If Machine_SIMUL_FRM.VueFix(0).Value = False Then
  If Machine_SIMUL_FRM.VueFix(1).Value = False Then
        gluLookAt 20, 0, 0, 0, 0, 0, 0, 0, 20
        
        glRotatef xm, 0, 1, 0  ' rotation en y
        glRotatef ym, 1, 0, 0  ' rotation en x
        glRotatef zm, 0, 0, 1  ' rotation en z
        glTranslatef PosY * Zoom, PosY * Zoom, PosX * Zoom
  Else
       ' Vue Broche X ou Y
        gluLookAt (Pt0.X) * Zoom + Vx.X * 10, (Pt0.Y) * Zoom + Vx.Y * 10, (Pt0.Z) * Zoom + Vx.Z * 10, Pt0.X * Zoom, Pt0.Y * Zoom, Pt0.Z * Zoom, 0, 0, 1
  End If
  
 Else
         ' Vue Machine X ou Y
        gluLookAt Pt0.X * Zoom + Vecteur_vue.X * 20, Pt0.Y * Zoom + Vecteur_vue.Y * 20, Pt0.Z * Zoom, Pt0.X * Zoom, Pt0.Y * Zoom, Pt0.Z * Zoom, 0, 0, 1
        glTranslatef 0, PosY * Zoom, PosX * Zoom
End If

'Zoom
glScalef Zoom, Zoom, Zoom

' Annulation des transfortmation pour machine avec element fixe different de 0
If Machine.Element_Fixe Then
    'For i = 0 To Machine.Element_Fixe Step -1
    For i = Machine.Element_Fixe To 0 Step -1
    Select Case Machine.Element(i).Type_axe
    
    
    ' Element Fixe (Tool, torche...)
    Case 99
        glTranslatef -Machine.Element(i).Origine.X, -Machine.Element(i).Origine.Y, -Machine.Element(i).Origine.Z
        
    'Rotation
    Case 1
        glRotatef -Machine.Element(i).Valeur_axe, Machine.Element(i).Vecteur.X, Machine.Element(i).Vecteur.Y, Machine.Element(i).Vecteur.Z   ' rotation autour axe
        glTranslatef -Machine.Element(i).Origine.X, -Machine.Element(i).Origine.Y, -Machine.Element(i).Origine.Z

    'Translation
    Case 2
        glTranslatef -Machine.Element(i).Origine.X, -Machine.Element(i).Origine.Y, -Machine.Element(i).Origine.Z
        glTranslatef -Machine.Element(i).Vecteur.X * Machine.Element(i).Valeur_axe, Machine.Element(i).Vecteur.Y * -Machine.Element(i).Valeur_axe, Machine.Element(i).Vecteur.Z * -Machine.Element(i).Valeur_axe
        
    'Piece
    Case 5
        glRotatef -Machine.Element(i).Valeur_axe, Machine.Element(i).Vecteur.X, Machine.Element(i).Vecteur.Y, Machine.Element(i).Vecteur.Z   ' rotation autour axe
        glTranslatef -Machine.Element(i).Origine.X, -Machine.Element(i).Origine.Y, -Machine.Element(i).Origine.Z

        
    ' Rotule ???
    Case Else
    
    
    End Select
    Next
End If


'---GRILLE :
        
If Grille Then
    'GrilleX = 50
    'GrilleY = 40
    'PasGrille = 20
    glTranslatef 0, 0, 0
    glBegin bmLines
    glColor3d 1, 1, 1
    glEnable GL_DEPTH_TEST
    'X
    For i = -GrilleX To GrilleX Step 10
        glVertex3f i * PasGrille, -GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glVertex3f i * PasGrille, GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
    Next
    'Y
    For i = -GrilleY To GrilleY Step 10
        glVertex3f -GrilleX * PasGrille, i * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glVertex3f GrilleX * PasGrille, i * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
    Next
    glEnd
    
    'grille avec transparence
    glColor4f 0.5, 0.5, 1, 0.1
    glBegin bmTriangles
        glNormal3f 0, 0, 1
        glVertex3f -GrilleX * PasGrille, -GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glVertex3f GrilleX * PasGrille, -GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glVertex3f -GrilleX * PasGrille, GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glNormal3f 0, 0, 1
        glVertex3f GrilleX * PasGrille, GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glVertex3f GrilleX * PasGrille, -GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
        glVertex3f -GrilleX * PasGrille, GrilleY * PasGrille, Val(Machine_SIMUL_FRM.Zgrille)
    glEnd
    
End If
 


'---AFFICHAGE

glPushMatrix

If Machine.NB_axe = 0 Then GoTo Fin
    

' Trace le parcours
If Option_tracer Then
            glColor3d 1, 0, 0  'en rouge
            glBegin bmLineStrip
                For K = 1 To UBound(Parcours) - 1
                    glVertex3f Parcours(K).X, Parcours(K).Y, Parcours(K).Z
                Next
            glEnd
End If


'Dessine les boites en collision
If Option_box And Piece_charger Then
    BoxO = Piece.Box_Col
    Call Affiche_box(BoxO)
End If

If Option_box Then
    BoxO = Machine.Element(Machine.Element_Collision).Box_Col
    Call Affiche_box(BoxO)
End If

'Affichage piece
Call Affiche_Piece(Piece, Render_mode)

'Affichage element
For i = 0 To Machine.PositionMagasin - 1
    'Debug.Print Machine.Element(I).Valeur_axe
    Call Affiche_Element(Machine.Element(i), Render_mode)
Next

'Affichage magasin
If Machine.PositionMagasin > 0 Then
    Call Affiche_Magasin(Machine.MagasinPo, Render_mode)
End If
        
'Affichage element
For i = Machine.PositionMagasin To UBound(Machine.Element)
    Call Affiche_Element(Machine.Element(i), Render_mode)
Next


' Decalage Dernier element
'glTranslatef Machine.Element(UBound(Machine.Element)).Origine.X, Machine.Element(UBound(Machine.Element)).Origine.Y, Machine.Element(UBound(Machine.Element)).Origine.Z

' Tools chargé
If Machine.Tool_current > 0 Then
       Pt1.Z = -POTool(Machine.Tool_current).Dec_Z
       Call Affiche_PoTool(POTool(Machine.Tool_current), Pt1, Machine.Element(Machine.NB_axe + 1).Vecteur)
       Pt1.Z = 0
       Call Affiche_Tool(ToolC(Machine.Tool_current), Pt1)
End If

Fin:

glPopMatrix

SwapBuffers Pict.hdc

'Reinitialise la matrice ..... brrrr Neo si tu m'entends
glLoadIdentity
End Sub
Sub AFFICHE_MATRIX()
Dim i As Integer

Dim MMatrix(1 To 16) As Double
'Recuperation de la matrice OpengL
glGetDoublev glgModelViewMatrix, MMatrix(1)

Debug.Print
Debug.Print "--- T MMatrix-------------------------"
For i = 1 To 16 Step 4
Debug.Print " | " & Format(MMatrix(i), "#,###0.0000") & " | " & Format(MMatrix(i + 1), "#,###0.0000") & " | " & Format(MMatrix(i + 2), "#,###0.0000") & " | " & Format(MMatrix(i + 3), "#,###0.0000") & " | "
Next
Debug.Print "------------------------------------"
End Sub
' calcul XYZ according les axes machines rotatif et Fixe pour RTCP
Sub CalculPos(Point_calc As Point3)
Dim i As Integer
Dim MMatrix(1 To 16) As Double
Dim Element As Element3D
Dim Mx2(2, 2) As Double
Dim VxInt As Point3
Dim VyInt As Point3
Dim VzInt As Point3
Dim PtInt As Point3

'On Error Resume Next


'Translation au point désiré
glTranslatef -Point_calc.X, -Point_calc.Y, -Point_calc.Z


For i = 0 To UBound(Machine.Element)
    Element = Machine.Element(i)
    Select Case Element.Type_axe
    
    ' Element Fixe (Banc...)
    Case 99
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    'Rotation
    Case 1
    
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z

        glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe

    
    'Translation
    Case 2
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glTranslatef Element.Vecteur.X * Element.Valeur_axe, Element.Vecteur.Y * Element.Valeur_axe, Element.Vecteur.Z * Element.Valeur_axe
    
    'Piece
    Case 5
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
        
    ' Rotule ???
    Case Else
    
    
    End Select
            ' Récupération de la matrice de transformation au niveau de l'element Fixe
            If i = Machine.Element_Fixe Then
            'Recuperation de la matrice OpengL
                glGetDoublev glgModelViewMatrix, MMatrix(1)
                
                'Point_calc.X = MMatrix(13)
                'Point_calc.Y = MMatrix(14)
                'Point_calc.Z = MMatrix(15)
                
                VxInt.X = MMatrix(1)
                VxInt.Y = MMatrix(2)
                VxInt.Z = MMatrix(3)
                
                VyInt.X = MMatrix(5)
                VyInt.Y = MMatrix(6)
                VyInt.Z = MMatrix(7)
                
                VzInt.X = MMatrix(9)
                VzInt.Y = MMatrix(10)
                VzInt.Z = MMatrix(11)
            End If
        Next

'decalage Tool
If Machine.Tool_current > 0 Then
   Element = Machine.Element(i - 1)
   glTranslatef Element.Vecteur.X * -Machine.LG_Tool_current, Element.Vecteur.Y * -Machine.LG_Tool_current, Element.Vecteur.Z * -Machine.LG_Tool_current
End If
    

'Recuperation de la matrice OpengL
glGetDoublev glgModelViewMatrix, MMatrix(1)

PtInt.X = MMatrix(13)
PtInt.Y = MMatrix(14)
PtInt.Z = MMatrix(15)

'Recalcul dans le repère ABS pour avoir les XYZ machines
Point_calc.X = PtInt.X * VxInt.X + PtInt.Y * VxInt.Y + PtInt.Z * VxInt.Z
Point_calc.Y = PtInt.X * VyInt.X + PtInt.Y * VyInt.Y + PtInt.Z * VyInt.Z
Point_calc.Z = PtInt.X * VzInt.X + PtInt.Y * VzInt.Y + PtInt.Z * VzInt.Z

' Reinit de la matrice  pour la suite
glLoadIdentity
End Sub
'
' Getpoint permet de récupérer la matrice de transformation de la machine
' Cette opération ne peut être faite dans la procédure dessine machine
' car dans ce cas vient aussi se supperposer les zooms et autres rotation
'
Sub GetPoint()
Dim i As Integer
Dim Limite As Integer
Dim MMatrix(1 To 16) As Double
Dim Element As Element3D
Dim Mx2(2, 2) As Double
Dim Pt1 As Point3

'On Error Resume Next

'piece
If Piece_charger Then
    glTranslatef Piece.Origine.X, Piece.Origine.Y, Piece.Origine.Z
    glRotatef Piece.Valeur_axe, Piece.Vecteur.X, Piece.Vecteur.Y, Piece.Vecteur.Z   ' rotation autour axe
    'Recuperation de la matrice OpengL
    glGetDoublev glgModelViewMatrix, Piece.Matrix(1)
    ' Annulation des transformation pour piece
    glRotatef -Piece.Valeur_axe, Piece.Vecteur.X, Piece.Vecteur.Y, Piece.Vecteur.Z   ' rotation autour axe
    glTranslatef -Piece.Origine.X, -Piece.Origine.Y, -Piece.Origine.Z
End If

        
Limite = UBound(Machine.Element)

For i = 0 To Limite
    Element = Machine.Element(i)
    
    Select Case Element.Type_axe
    
        ' Element Fixe (Banc...)
        Case 99
            glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
            
        'Rotation
        Case 1
            glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
            glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
    
        'Translation
        Case 2
            glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
            glTranslatef Element.Vecteur.X * Element.Valeur_axe, Element.Vecteur.Y * Element.Valeur_axe, Element.Vecteur.Z * Element.Valeur_axe
        
        ' Rotule ???
        Case Else
    End Select
    ' mise a jour des matrices elements
    'Recuperation de la matrice OpengL
    glGetDoublev glgModelViewMatrix, Machine.Element(i).Matrix(1)
Next


' Tools chargé
If Machine.Tool_current > 0 Then
      ' Pt1.Z = -POTool(Machine.Tool_current).Dec_Z
      Pt1.Z = -ToolC(Machine.Tool_current).LG
       ' Gestion des machines avec Tool différent de Z
        ' Ce code n'est pas definitif car il ne gère pas les axes autres que X ou Y
        If Machine.Element(Machine.NB_axe + 1).Vecteur.Z = 0 Then
            If Machine.Element(Machine.NB_axe + 1).Vecteur.X = 1 Then
                glRotatef -90, 0, 1, 0
            End If
            If Machine.Element(Machine.NB_axe + 1).Vecteur.Y = 1 Then
                glRotatef -90, 1, 0, 0
            End If
        End If
        
        
        glTranslatef Pt1.X, Pt1.Y, Pt1.Z 'position objet

End If

    
'Recuperation de la matrice OpengL
glGetDoublev glgModelViewMatrix, MMatrix(1)

'Debug.Print
'Debug.Print "--- T MMatrix-------------------------"
'For I = 1 To 16 Step 4
'Debug.Print " | " & Format(Machine.Element(Limite).Matrix(I), "#,###0.0000") & " | " & Format(Machine.Element(Limite).Matrix(I + 1), "#,###0.0000") & " | " & Format(Machine.Element(Limite).Matrix(I + 2), "#,###0.0000") & " | " & Format(Machine.Element(I - 1).Matrix(I + 3), "#,###0.0000") & " | "
'Next
'Debug.Print "------------------------------------"

'Pt0.X = Machine.Element(Limite).Matrix(13)
'Pt0.Y = Machine.Element(Limite).Matrix(14)
'Pt0.Z = Machine.Element(Limite).Matrix(15)
Pt0.X = MMatrix(13)
Pt0.Y = MMatrix(14)
Pt0.Z = MMatrix(15)


Vx.X = Machine.Element(Limite).Matrix(1)
Vx.Y = Machine.Element(Limite).Matrix(2)
Vx.Z = Machine.Element(Limite).Matrix(3)

Vy.X = Machine.Element(Limite).Matrix(5)
Vy.Y = Machine.Element(Limite).Matrix(6)
Vy.Z = Machine.Element(Limite).Matrix(7)

Vz.X = Machine.Element(Limite).Matrix(9)
Vz.Y = Machine.Element(Limite).Matrix(10)
Vz.Z = Machine.Element(Limite).Matrix(11)

' Reinit de la matrice  pour la suite
glLoadIdentity
End Sub
'
' Getpoint permet de récupérer la matrice de transformation de la machine
' Cette opération ne peut être faite dans la procédure dessine machine
' car dans ce cas vient aussi se supperposer les zooms et autres rotation
'
Sub GetPointColl()
Dim i As Integer
Dim Limite As Integer
Dim MMatrix(1 To 16) As Double
Dim Element As Element3D
Dim Mx2(2, 2) As Double
'On Error Resume Next

'piece la piece est l'element fixe
If Piece_charger Then
    glTranslatef Piece.Origine.X, Piece.Origine.Y, Piece.Origine.Z
    glRotatef Piece.Valeur_axe, Piece.Vecteur.X, Piece.Vecteur.Y, Piece.Vecteur.Z   ' rotation autour axe
    'Recuperation de la matrice OpengL
    glGetDoublev glgModelViewMatrix, Piece.Matrix(1)
    ' Annulation des transformation pour piece
    glRotatef -Piece.Valeur_axe, Piece.Vecteur.X, Piece.Vecteur.Y, Piece.Vecteur.Z   ' rotation autour axe
    glTranslatef -Piece.Origine.X, -Piece.Origine.Y, -Piece.Origine.Z
End If

        
Limite = UBound(Machine.Element)

For i = 0 To Limite
    Element = Machine.Element(i)
    Select Case Element.Type_axe
    
        ' Element Fixe (Banc...)
        Case 99
            glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
            
        'Rotation
        Case 1
            glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
            glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
    
        'Translation
        Case 2
            glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
            glTranslatef Element.Vecteur.X * Element.Valeur_axe, Element.Vecteur.Y * Element.Valeur_axe, Element.Vecteur.Z * Element.Valeur_axe
        
        ' Rotule ???
        Case Else
    End Select
    ' mise a jour des matrices elements
    'Recuperation de la matrice OpengL
     glGetDoublev glgModelViewMatrix, Machine.Element(i).Matrix(1)
Next

' Reinit de la matrice  pour la suite
glLoadIdentity
End Sub
'Dessine avec les fonctions OPENGL ou Tool de Fraisage
Sub Affiche_Tool(ByRef Outl As Tool, ByRef Pt0 As Point3)
Dim i As Integer
Dim XE1 As Double
Dim ZE1 As Double
Dim XE2 As Double
Dim ZE2 As Double
Dim J As Integer
Dim ZE As Double
Dim Obj As Long '

Obj = gluNewQuadric 'create new quadric object

gluQuadricNormals Obj, qnSmooth
glShadeModel smFlat


glColor4f 0.7, 0.35, 0, 1

' Decalage pour origine de creation Tool
ZE = Pt0.Z - Outl.LG
glTranslatef Pt0.X, Pt0.Y, ZE 'position objet

Select Case Outl.Type

    Case 1 ' Ball
        gluSphere Obj, Outl.Diameter / 2, 16, 16
    
    Case 2 ' FlatTool
        ' Fraise platte
        If Outl.CornerRadius = 0 Then
            gluDisk Obj, 0, Outl.Diameter / 2, 16, 1
        
        Else ' FlatTool
            glTranslatef 0, 0, -Outl.CornerRadius  'position objet
        
            'Approx Rayon par plusieurs cone
            For J = 10 To 90 Step 10
            ZE1 = Outl.CornerRadius * (1 - Cos((J - 10) * DEGTORAD))
            XE1 = Outl.CornerRadius * (1 - Sin((J - 10) * DEGTORAD))
            
            ZE2 = Outl.CornerRadius * (1 - Cos(J * DEGTORAD))
            XE2 = Outl.CornerRadius * (1 - Sin(J * DEGTORAD))
            
            gluCylinder Obj, ((Outl.Diameter / 2) - XE1), ((Outl.Diameter / 2) - XE2), ZE2 - ZE1, 16, 1
            glTranslatef 0, 0, ZE2 - ZE1
            Next
            
            gluCylinder Obj, ((Outl.Diameter / 2) - Outl.CornerRadius), Outl.Diameter / 2, Outl.CornerRadius, 16, 1
            'glTranslatef 0, 0, Outl.CornerRadius  'position objet
       End If
    
    Case 3 ' Drill
       glTranslatef 0, 0, -Outl.LgCone  'position objet
       gluCylinder Obj, 0, Outl.Diameter / 2, Outl.LgCone, 16, 1
       glTranslatef 0, 0, Outl.LgCone  'position objet
   
   Case Else ' autre Tool
   
End Select

'Dessin corps Tool commun au trois type
gluCylinder Obj, Outl.Diameter / 2, Outl.Diameter / 2, Outl.LG_Coupe, 16, 1
ZE = ZE + Outl.LG_Coupe
glTranslatef 0, 0, Outl.LG_Coupe  'position objet
gluDisk Obj, Outl.DiameterCorp / 2, Outl.Diameter / 2, 16, 1
gluCylinder Obj, Outl.DiameterCorp / 2, Outl.DiameterCorp / 2, (Outl.LG - Outl.LG_Coupe), 16, 1


glTranslatef -Pt0.X, -Pt0.Y, -ZE  'Annule transformation objet

End Sub

'Dessine avec les fonctions OPENGL le Porte Tool de Fraisage ( element de revolution)
Sub Affiche_PoTool(ByRef PoOutl As PO_Tool, ByRef Pt0 As Point3, ByRef Axe As Point3)
Dim i As Integer
Dim ZE As Double
Dim Obj As Long '

Obj = gluNewQuadric 'create new quadric object

'gluQuadricNormals Obj, qnSmooth
'glShadeModel smFlat

glColor4f 0, 0.3, 0.6, 1

' Gestion des machines avec Tool différent de Z
' Ce code n'est pas definitif car il ne gère pas les axes autres que X ou Y
If Axe.Z = 0 Then
    If Axe.X = 1 Then
        glRotatef -90, 0, 1, 0
    End If
    If Axe.Y = 1 Then
        glRotatef -90, 1, 0, 0
    End If
End If

ZE = Pt0.Z
glTranslatef Pt0.X, Pt0.Y, Pt0.Z 'position objet


' Premier disque
ZE = ZE + PoOutl.Coord(1).Y
glTranslatef 0, 0, PoOutl.Coord(1).Y 'position objet
gluDisk Obj, 0, PoOutl.Coord(1).X, 16, 1

For i = 2 To PoOutl.NB_point


    If PoOutl.Coord(i).Y = PoOutl.Coord(i - 1).Y Then
    ' Disque
      If PoOutl.Coord(i - 1).X < PoOutl.Coord(i).X Then
        gluDisk Obj, PoOutl.Coord(i - 1).X, PoOutl.Coord(i).X, 16, 1
      Else
        gluDisk Obj, PoOutl.Coord(i).X, PoOutl.Coord(i - 1).X, 16, 1
      End If
    Else
    ' Cylindre
       gluCylinder Obj, PoOutl.Coord(i - 1).X, PoOutl.Coord(i).X, (PoOutl.Coord(i).Y - PoOutl.Coord(i - 1).Y), 16, 1
    End If
    
    ZE = ZE + (PoOutl.Coord(i).Y - PoOutl.Coord(i - 1).Y)
    glTranslatef 0, 0, (PoOutl.Coord(i).Y - PoOutl.Coord(i - 1).Y) 'position objet

Next i
       
        
'dernier disque
'ZE = ZE + (PoOutl.Coord(PoOutl.NB_point).Y - PoOutl.Coord(PoOutl.NB_point - 1).Y)
'glTranslatef 0, 0, (PoOutl.Coord(PoOutl.NB_point).Y - PoOutl.Coord(PoOutl.NB_point - 1).Y) 'position objet
gluDisk Obj, 0, PoOutl.Coord(PoOutl.NB_point).X, 16, 1


glTranslatef -Pt0.X, -Pt0.Y, -ZE  'Annule transformation objet

End Sub

' affiche la boite utilisée pour le controle de collision
Sub Affiche_box(ByRef Element As Box3)
Dim Pt0 As Point3
Dim Pt1 As Point3
Dim Pt2 As Point3
Dim i0 As Integer
Dim i1 As Integer


glColor4f 1, 0, 0, 0.8



With Element

For i0 = -1 To 1 Step 2
    For i1 = -1 To 1 Step 2
        Pt0 = VecAdd(VecAdd(.Centre, .Axes(0), .Longueurs(0)), .Axes(1), i1 * .Longueurs(1))
        Pt1 = VecAdd(Pt0, .Axes(2), .Longueurs(2) * i0)
        Pt0 = VecAdd(VecAdd(.Centre, .Axes(0), -1 * .Longueurs(0)), .Axes(1), i1 * .Longueurs(1))
        Pt2 = VecAdd(Pt0, .Axes(2), .Longueurs(2) * i0)
        
        glBegin bmLines
                       glVertex3f Pt1.X, Pt1.Y, Pt1.Z
                       glVertex3f Pt2.X, Pt2.Y, Pt2.Z

        Pt0 = VecAdd(VecAdd(.Centre, .Axes(0), i0 * .Longueurs(0)), .Axes(1), i1 * .Longueurs(1))
        Pt1 = VecAdd(Pt0, .Axes(2), .Longueurs(2))
        Pt0 = VecAdd(VecAdd(.Centre, .Axes(0), i0 * .Longueurs(0)), .Axes(1), i1 * .Longueurs(1))
        Pt2 = VecAdd(Pt0, .Axes(2), -.Longueurs(2))

                       glVertex3f Pt1.X, Pt1.Y, Pt1.Z
                       glVertex3f Pt2.X, Pt2.Y, Pt2.Z

        
        Pt0 = VecAdd(VecAdd(.Centre, .Axes(0), i0 * .Longueurs(0)), .Axes(1), .Longueurs(1))
        Pt1 = VecAdd(Pt0, .Axes(2), i1 * .Longueurs(2))
        Pt0 = VecAdd(VecAdd(.Centre, .Axes(0), i0 * .Longueurs(0)), .Axes(1), -.Longueurs(1))
        Pt2 = VecAdd(Pt0, .Axes(2), i1 * .Longueurs(2))

                       glVertex3f Pt1.X, Pt1.Y, Pt1.Z
                       glVertex3f Pt2.X, Pt2.Y, Pt2.Z
        glEnd
        
    Next i1
Next i0

    
End With


End Sub

' Affiche un element de la machine
'
'
' Render_mode = 1  => Traits cachés
' Render_mode = 2  => Ombrée
' Render_mode = autre => Filaire
'
Sub Affiche_Element(ByRef Element As Element3D, ByRef Render_mode As Integer)
Dim P1 As DecOrigine
Dim mode As Integer
Dim J As Integer
Dim i As Integer



' Test si Triangles present dans Element
If Element.STL_def.NmbVertex = 0 Then Exit Sub
    Select Case Element.Type_axe
    
    ' Element Fixe (Banc...)
    Case 99
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        
    'Rotation
    Case 1
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe

    'Translation
    Case 2
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glTranslatef Element.Vecteur.X * Element.Valeur_axe, Element.Vecteur.Y * Element.Valeur_axe, Element.Vecteur.Z * Element.Valeur_axe
        
    'magasin Tool
    Case 4
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef (-1 + Element.Valeur_axe) * (360 / Element.MaxiAxe), Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z  ' rotation en z
     
    'Piece
    Case 5
        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
        
    ' Rotule ???
    Case Else
    
    
    End Select

    
' mode par defaut = ombrée
mode = bmTriangles

If Render_mode = 2 Then
     mode = bmLineLoop
End If


glColor4f Element.Color.Rouge / 255, Element.Color.Vert / 255, Element.Color.Bleu / 255, ObjectAlpha


With Element.STL_def
    For J = 0 To .NmbVertex - 3 Step 3
        glBegin mode
            glNormal3f .Normal(J / 3).X, .Normal(J / 3).Y, .Normal(J / 3).Z
            glVertex3f .Vertex(J).X, .Vertex(J).Y, .Vertex(J).Z
            glVertex3f .Vertex(J + 1).X, .Vertex(J + 1).Y, .Vertex(J + 1).Z
            glVertex3f .Vertex(J + 2).X, .Vertex(J + 2).Y, .Vertex(J + 2).Z
        glEnd
    Next
End With

    

'trait cachés
If Render_mode = 1 Then
    With Element.STL_def
        For J = 0 To .NmbVertex - 3 Step 3
    
            glColor3f 0, 0, 0
            glBegin bmLines
                glVertex3f .Vertex(J).X, .Vertex(J).Y, .Vertex(J).Z
                glVertex3f .Vertex(J + 1).X, .Vertex(J + 1).Y, .Vertex(J + 1).Z
                glVertex3f .Vertex(J + 2).X, .Vertex(J + 2).Y, .Vertex(J + 2).Z
            glEnd
        Next
    End With
End If


Select Case Element.Type_axe
'Cas du magasin d'Tool et de la piece
    Case 4
        For i = 1 To UBound(ToolC)
        'Affichage Tool et porte Tool
        ' n'affiche pas l'Tool current de la machine car il est sur la fraiseuse !!!
            If ToolC(i).Diameter > 0 And i <> Machine.Tool_current Then
                Call Affiche_Tool(ToolC(i), ToolC(i).Origine)
            End If
            
            If POTool(i).NB_point > 0 And i <> Machine.Tool_current Then
                Call Affiche_PoTool(POTool(i), POTool(i).Origine, Machine.MagasinPo.Vecteur)
            End If
        Next
        
        ' Annulation des transformation imposé par axe auxiliaire ( Magasin Tool)
        glRotatef (-1 + Element.Valeur_axe) * (-360 / Element.MaxiAxe), Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z  ' rotation en z
        glTranslatef -Element.Origine.X, -Element.Origine.Y, -Element.Origine.Z

   Case 5
        ' Annulation des transformation pour piece
        glRotatef -Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
        glTranslatef -Element.Origine.X, -Element.Origine.Y, -Element.Origine.Z
End Select

End Sub

' Affiche un element de la machine
'
'
' Render_mode = 1  => Traits cachés
' Render_mode = 2  => Ombrée
' Render_mode = autre => Filaire
'
Sub Affiche_Piece(ByRef Element As Element3D, ByRef Render_mode As Integer)
Dim mode As Integer
Dim J As Integer

' Test si Triangles present dans Element
If Element.STL_def.NmbVertex = 0 Then Exit Sub

        glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
        glRotatef Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
        


    
' mode par defaut = ombrée
mode = bmTriangles

If Render_mode = 2 Then
     mode = bmLineLoop
End If


glColor4f Element.Color.Rouge / 255, Element.Color.Vert / 255, Element.Color.Bleu / 255, ObjectAlpha


With Element.STL_def
    For J = 0 To .NmbVertex - 3 Step 3
        glBegin mode
            glNormal3f .Normal(J / 3).X, .Normal(J / 3).Y, .Normal(J / 3).Z
            glVertex3f .Vertex(J).X, .Vertex(J).Y, .Vertex(J).Z
            glVertex3f .Vertex(J + 1).X, .Vertex(J + 1).Y, .Vertex(J + 1).Z
            glVertex3f .Vertex(J + 2).X, .Vertex(J + 2).Y, .Vertex(J + 2).Z
        glEnd
    Next
End With

    

'trait cachés
If Render_mode = 1 Then
    With Element.STL_def
        For J = 0 To .NmbVertex - 3 Step 3
    
            glColor3f 0, 0, 0
            glBegin bmLines
                glVertex3f .Vertex(J).X, .Vertex(J).Y, .Vertex(J).Z
                glVertex3f .Vertex(J + 1).X, .Vertex(J + 1).Y, .Vertex(J + 1).Z
                glVertex3f .Vertex(J + 2).X, .Vertex(J + 2).Y, .Vertex(J + 2).Z
            glEnd
        Next
    End With
End If



        ' Annulation des transformation pour piece
        glRotatef -Element.Valeur_axe, Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z   ' rotation autour axe
        glTranslatef -Element.Origine.X, -Element.Origine.Y, -Element.Origine.Z


End Sub
Sub Affiche_Magasin(ByRef Element As Element3D, ByRef Render_mode As Integer)
Dim P1 As DecOrigine
Dim i As Integer
Dim mode As Integer

' Test si Triangles present dans Element
If Element.STL_def.NmbVertex = 0 Then Exit Sub

glTranslatef Element.Origine.X, Element.Origine.Y, Element.Origine.Z
glRotatef (-1 + Element.Valeur_axe) * (360 / Element.MaxiAxe), Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z  ' rotation en z


    
' mode par defaut = ombrée
mode = bmTriangles

If Render_mode = 2 Then
     mode = bmLineLoop
End If


glColor4f Element.Color.Rouge / 255, Element.Color.Vert / 255, Element.Color.Bleu / 255, ObjectAlpha


With Element.STL_def
    For i = 0 To .NmbVertex - 3 Step 3
        glBegin mode
            glNormal3f .Normal(i / 3).X, .Normal(i / 3).Y, .Normal(i / 3).Z
            glVertex3f .Vertex(i).X, .Vertex(i).Y, .Vertex(i).Z
            glVertex3f .Vertex(i + 1).X, .Vertex(i + 1).Y, .Vertex(i + 1).Z
            glVertex3f .Vertex(i + 2).X, .Vertex(i + 2).Y, .Vertex(i + 2).Z
        glEnd
    Next
End With

    

'trait cachés
If Render_mode = 1 Then
    With Element.STL_def
        For i = 0 To .NmbVertex - 3 Step 3
    
            glColor3f 0, 0, 0
            glBegin bmLines
                glVertex3f .Vertex(i).X, .Vertex(i).Y, .Vertex(i).Z
                glVertex3f .Vertex(i + 1).X, .Vertex(i + 1).Y, .Vertex(i + 1).Z
                glVertex3f .Vertex(i + 2).X, .Vertex(i + 2).Y, .Vertex(i + 2).Z
            glEnd
        Next
    End With
End If

'Cas du magasin d'Tool et de la piece
        For i = 1 To UBound(ToolC)
        'Affichage Tool et porte Tool
        ' n'affiche pas l'Tool current de la machine car il est sur la fraiseuse !!!
            If ToolC(i).Diameter > 0 And i <> Machine.Tool_current Then
                Call Affiche_Tool(ToolC(i), ToolC(i).Origine)
            End If
            
            If POTool(i).NB_point > 0 And i <> Machine.Tool_current Then
                Call Affiche_PoTool(POTool(i), POTool(i).Origine, Machine.MagasinPo.Vecteur)
            End If
        Next
        
        ' Annulation des transformation imposé par axe auxiliaire ( Magasin Tool)
        glRotatef (-1 + Element.Valeur_axe) * (-360 / Element.MaxiAxe), Element.Vecteur.X, Element.Vecteur.Y, Element.Vecteur.Z  ' rotation en z
        glTranslatef -Element.Origine.X, -Element.Origine.Y, -Element.Origine.Z


End Sub

'Reinit objet
Sub Reinit_Element(Elem As Element3D)

    Call Reinit_Maillage(Elem.STL_def)

End Sub

'Reinit Maillage
Sub Reinit_Maillage(Elem As Maillage)
With Elem
    .NmbNormal = 0
    .NmbVertex = 0
    ReDim .Vertex(0)
    ReDim .Normal(0)
End With
End Sub
