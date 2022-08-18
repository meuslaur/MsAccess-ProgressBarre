VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Demo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Demo")
Option Compare Database
Option Explicit

    Private cLabels  As C_LabelsInfo                        'Code C_LabelsInfo

Private Sub Class_Initialize()
    Set cLabels = New C_LabelsInfo
End Sub

Private Sub Class_Terminate()
    Set cLabels = Nothing                                   'Code C_LabelsInfo destruction de l'objet
End Sub

Public Property Get GetClsLabelsInfo() As C_LabelsInfo
    Set GetClsLabelsInfo = cLabels                          'Code C_LabelsInfo pour le form F_Demo
End Property

Public Sub BoucleTestCls(Boucle As Long)
    Dim i As Long
    Dim lRand As Long
    Dim sAvance As String

    cLabels.SetLengthBarre "a", Boucle                       '// Code C_LabelsInfo

    For i = 1 To Boucle         ' DEMO barre 1
        lRand = Int((700 * Rnd) + 1)

        sAvance = "Boucle : " & CStr(i)
        cLabels.UpdateBarre sAvance                          'Code C_LabelsInfo Affiche l'état d'avancement du scan...
        DoEvents            '// Refresh labels.

        Sleep (lRand)
   Next i

    cLabels.HiddenCurrentLabels                              'Code C_LabelsInfo

End Sub
