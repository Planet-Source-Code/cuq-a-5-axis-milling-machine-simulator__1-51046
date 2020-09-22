Attribute VB_Name = "PreVisualisation"
Option Explicit

'Initialisation de la previsualisation  Machine
Sub Init_Previsu_Machine(Repertoire As String, Fichier_Def_machine As String, Machine_tempo As Machine3D)
'Dim Fichier_STL As String
Dim i As Integer


'Chargement caractéristique Machine
Call Charger_Machine(Fichier_Def_machine, Machine_tempo)

Call MessageLOG_Print("Preview Machine " & Machine_tempo.name)

'Recuperation de la definition geométrique via fichier STL ascii
For i = 0 To UBound(Machine_tempo.Element)
    Call Reinit_Element(Machine_tempo.Element(i))
    Call ChargeFichierSTL(App.Path + "\Machine_def\" + Repertoire + "\" + Machine_tempo.Element(i).fichier, Machine_tempo.Element(i).STL_def, False)
Next

'Magasin Tool
If Machine_tempo.PositionMagasin Then
    Call Reinit_Element(Machine_tempo.MagasinPo)
    Call ChargeFichierSTL(App.Path + "\Machine_def\" + Repertoire + "\" + Machine_tempo.MagasinPo.fichier, Machine_tempo.MagasinPo.STL_def, False)
End If

End Sub

Sub PrevisuMachine(Pict As PictureBox, Machine_tempo As Machine3D)
'Dim MMatrix(1 To 16) As Double
Dim T As Integer
Dim Z As Integer
Dim i As Integer


On Error Resume Next



'---DEBUT :

    glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
    glMatrixMode mmModelView

    
'---POINT DE VUE : Camera

    gluLookAt 20, 0, 0, 0, 0, 0, 0, 0, 20




'---PETITE INITALISATION :
        
    'glMaterialiv GL_FRONT_AND_BACK, mprAmbientAndDiffuse, GL_SPECULAR
    'glMateriali GL_FRONT_AND_BACK, GL_SHININESS, 100
     glViewport 0, 0, Pict.ScaleWidth, Pict.ScaleHeight

    glScalef Zoom, Zoom, Zoom
    
        glRotatef xm, 0, 1, 0  ' rotation en y
        glRotatef ym, 1, 0, 0  ' rotation en x
        glRotatef zm, 0, 0, 1  ' rotation en z


    T = 50
    Z = 20
    glTranslatef 0, 0, 0
    glBegin bmLines
    glColor3d 1, 1, 1
    glEnable GL_DEPTH_TEST
    For i = -T To T Step 10
        glVertex3f i * Z, -T * Z, 0
        glVertex3f i * Z, T * Z, 0
    Next
    For i = -T To T Step 10
        glVertex3f -T * Z, i * Z, 0
        glVertex3f T * Z, i * Z, 0
    Next
    glEnd

 
'---AFFICHAGE
' Init des traits cachés
glEnable GL_DEPTH_TEST
glEnable GL_POLYGON_OFFSET_FILL
glPolygonOffset 1, 2

glPushMatrix

'Affichage element
For i = 0 To Machine_tempo.PositionMagasin - 1
    'Debug.Print Machine.Element(I).Valeur_axe
    Call Affiche_Element(Machine_tempo.Element(i), 1)
Next

'Affichage magasin
If Machine_tempo.PositionMagasin Then Call Affiche_Element(Machine_tempo.MagasinPo, 1)

        
'Affichage element
For i = Machine_tempo.PositionMagasin To UBound(Machine_tempo.Element)
    Call Affiche_Element(Machine_tempo.Element(i), 1)
Next

glPopMatrix



SwapBuffers Pict.hdc
'reinitialise
glLoadIdentity
End Sub

