Attribute VB_Name = "basRegistry"
Option Explicit


' Options de s�curit� des cl�s de base de registres...
Private Const ERROR_SUCCESS = 0&

' D�claration des variables du processeur.
Public Type SYSTEM_INFO
    wProcessorArchitecture                      As Integer              '
    wReserved                                   As Integer              '
    dwPageSize                                  As Long                 '
    lpMinimumApplicationAddress                 As Long                 '
    lpMaximumApplicationAddress                 As Long                 '
    dwActiveProcessorMask                       As Long                 '
    dwNumberOfProcessors                        As Long                 '
    dwProcessorType                             As Long                 '
    dwAllocationGranulalarity                   As Long                 '
    wProcessorLevel                             As Integer              '
    wProcessorRevision                          As Integer              '
End Type

'Registry Functions
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal HKey As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal HKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal HKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long

'GET SYSTEM INFO
Public Declare Sub GetSystemInfo Lib "kernel32.dll" (lpSystemInfo As SYSTEM_INFO)

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim I As Long                                           ' Compteur de boucle.
    Dim rc As Long                                          ' Code de retour.
    Dim HKey As Long                                        ' Descripteur d'une cl� de base de registres ouverte.
    Dim KeyValType As Long                                  ' Type de donn�es d'une cl� de base de registres.
    Dim tmpVal As String                                    ' Stockage temporaire pour une valeur de cl� de base de registres.
    Dim KeyValSize As Long                                  ' Taille de la variable de la cl� de base de registres.
    '------------------------------------------------------------
    ' Ouvre la cl� de base de registres sous la racine cl� {HKEY_LOCAL_MACHINE...}.
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, HKey) ' Ouvre la cl� de base de registres.
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' G�re l'erreur...
    
    tmpVal = String$(1024, 0)                             ' Alloue de l'espace pour la variable.
    KeyValSize = 1024                                       ' D�finit la taille de la variable.
    
    '------------------------------------------------------------
    ' Extrait la valeur de la cl� de base de registres...
    '------------------------------------------------------------
    rc = RegQueryValueExString(HKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtient/Cr�e la valeur de la cl�.
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' G�re l'erreur.
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 ajoute une cha�ne termin�e par un caract�re nul...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Caract�re nul trouv�, extrait de la cha�ne.
    Else                                                    ' WinNT ne termine pas la cha�ne par un caract�re nul...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Caract�re nul non trouv�, extrait la cha�ne uniquement.
    End If
    '------------------------------------------------------------
    ' D�termine le type de valeur de la cl� pour la conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Recherche les types de donn�es...
    Case REG_SZ                                             ' Type de donn�es cha�ne de la cl� de la base de registres.
        KeyVal = tmpVal                                     ' Copie la valeur de la cha�ne.
    Case REG_DWORD                                          ' Type de donn�es double mot de la cl� de base de registres.
        For I = Len(tmpVal) To 1 Step -1                    ' Convertit chaque bit.
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Construit la valeur caract�re par caract�re.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convertit le mot double en cha�ne.
    End Select
    
    GetKeyValue = True                                      ' Retour avec succ�s.
    rc = RegCloseKey(HKey)                                  ' Ferme la cl� de base de registres
    Exit Function                                           ' Quitte.
    
GetKeyError:      ' R�initialise apr�s qu'une erreur s'est produite...
    KeyVal = ""                                             ' Affecte une cha�ne vide � la valeur de retour.
    GetKeyValue = False                                     ' Retour avec �chec.
    rc = RegCloseKey(HKey)                                  ' Ferme la cl� de base de registres.
End Function


