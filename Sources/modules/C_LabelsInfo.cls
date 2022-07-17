VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_LabelsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Dev")
' ------------------------------------------------------
' Name     : C_LabelsInfo
' ------------------------------------------------------
' Kind     : Module de classe
' Purpose  : Affichage de texte de d'avancement du code dans un formulaire.
' Author   : Laurent
' Sujet    :
' Objectif :
' Date     : 03/07/2022 - 18:01
' DateMod  : 10/07/2022 - 15:07
' ------------------------------------------------------
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
    Private Const LARG_CAR  As Long = 182       '// !!! Largeur 1 Caractèrre, Police Webdings size 8 !!!.
    Private Const CAR_BAR   As String * 1 = "a" '// Defaut caractère a afficher dans la barre (police Webdings).

    Private m_oLblTexte     As Access.Label     '// Label texte d'info.
    Private m_oLblBarre     As Access.Label     '// Label barre.

    Private m_lCompte       As Long             '// Compteur pour la barre d'avancement.
    Private m_CarBarre      As String * 1       '// Caractère à afficher dans la barre (m_oLblBarre).
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::

'//----------------------------------       INIT/TERM       ------------------------------
Private Sub Class_Initialize()
    m_CarBarre = CAR_BAR
End Sub

Private Sub Class_Terminate()
    Set m_oLblTexte = Nothing
    Set m_oLblBarre = Nothing
End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================
Public Property Let LabelTexte(LabelInfo As Access.Label)
     Set m_oLblTexte = LabelInfo
End Property
Public Property Let LabelBarre(LabelBarre As Access.Label)
    Set m_oLblBarre = LabelBarre
End Property
Public Property Let CaractereAfficher(Caractere As String)
    If (Len(Caractere) <> 1) Then Exit Property
    m_CarBarre = Caractere
End Property
'//====================================== END PROP =======================================

'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PUBLIC SUB/FUNC   \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub InitialiseBarre(LongueurBarre As Long)

    If (m_oLblBarre Is Nothing) Then Exit Sub
    If (LongueurBarre < 1) Or (LongueurBarre > 75) Then Exit Sub    '// Definir les limites min et max. (nb de caractère afficher)

    m_oLblBarre.Width = LARG_CAR * LongueurBarre    '// Détermine la dimension de la barre.
    m_lCompte = 1

End Sub

Public Sub ResetLabels()
'// Permet d'utiliser differents labels, ou Afficher autre chose.
    Set m_oLblTexte = Nothing
    Set m_oLblBarre = Nothing
End Sub

'----------------------------------------------------------------
' Procedure Nom   : AfficheBarre
' ----------------------------------------------------------------
' Sujet           : Affiche de texte et barre d'avancement du code.
' Procedure Kind  : Sub
' Procedure Access: Public
' Objectif        :
'
'=== Paramètres ===
' sObjNom (String)  : Texte à afficher.
'==================
'
' Author  : Laurent
' Date    : 03/07/2022 - 17:50
' DateMod :
' ----------------------------------------------------------------
Public Sub UpdateBarre(NewTexte As String)
On Error GoTo ERR_testBarre

    If (Not m_oLblTexte Is Nothing) Then m_oLblTexte.Caption = NewTexte & "....."

    If (Not m_oLblBarre Is Nothing) Then
        m_oLblBarre.Caption = String(m_lCompte, m_CarBarre)
        m_lCompte = m_lCompte + 1
    End If

SORTIE_testBarre:
    Exit Sub

ERR_testBarre:
    MsgBox "Erreur " & Err.Number & vbCrLf & _
            " (" & Err.Description & ")" & vbCrLf & _
            "Dans  TradAccess.MD_Test.testBarre, ligne " & Erl & "."
    Resume SORTIE_testBarre
End Sub
'// \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ END PUB. SUB/FUNC \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
