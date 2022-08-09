VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_TestBarre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Demo")
Option Compare Database
Option Explicit

'//::::::::::::::::::::::::::::::::::    VARIABLES      ::::::::::::::::::::::::::::::::::
    Private LabelsInfo  As C_LabelsInfo     '// Classe affichage info texte et barre d'avancement du code.
'//:::::::::::::::::::::::::::::::::: END VARIABLES ::::::::::::::::::::::::::::::::::::::


'//----------------------------------       INIT/TERM       ------------------------------
Private Sub Class_Initialize()
    If (LabelsInfo Is Nothing) Then Set LabelsInfo = New C_LabelsInfo
End Sub

Private Sub Class_Terminate()
    Set LabelsInfo = Nothing
End Sub
'//---------------------------------------------------------------------------------------

'//==================================       PROP        ==================================

Public Property Get GetClsLabelsInfo() As C_LabelsInfo
    If Not (LabelsInfo Is Nothing) Then Set GetClsLabelsInfo = LabelsInfo
End Property
'//====================================== END PROP =======================================


Public Sub BoucleTestCls(Boucle As Long)
    Dim i As Long
    Dim lRand As Long
    Dim sAvance As String

    For i = 1 To Boucle
        lRand = Int((700 * Rnd) + 1)
        sAvance = "Boucle : " & str(i)
        If LabelsInfo Is Nothing Then Exit For   '// Fermeture du frm démo.
        LabelsInfo.UpdateBarre sAvance   '// Affiche l'état d'avancement du scan...
        DoEvents
        Sleep (lRand)
   Next i
End Sub

Public Sub ResetClsLabelsInfo()
    Set LabelsInfo = Nothing    '// Demo Fermeture du frm démo.
End Sub

'// ################################# END PRIV. SUB/FUNC #################################
