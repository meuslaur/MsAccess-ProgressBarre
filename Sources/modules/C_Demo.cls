VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Demo")
Option Compare Database
Option Explicit

    Private cLabels  As C_LabelsInfo                                'Code C_LabelsInfo

Private Sub Class_Initialize()
    If (cLabels Is Nothing) Then Set cLabels = New C_LabelsInfo
End Sub

Private Sub Class_Terminate()
    Set cLabels = Nothing                                           'Code C_LabelsInfo destruction de l'objet
End Sub

Public Property Get GetClsLabelsInfo() As C_LabelsInfo
    If Not (cLabels Is Nothing) Then Set GetClsLabelsInfo = cLabels 'Code C_LabelsInfo pour le form F_Demo
End Property

Public Sub BoucleTestCls(Boucle As Long)
    Dim i As Long
    Dim lRand As Long
    Dim sAvance As String

'DEMO Barre 2 F_DEMO
Dim sCar As String
Dim lBoucle As Long
Dim lDep As Long
lBoucle = Forms("F_Demo")!txtBoucle2.Value
sCar = Forms("F_Demo")!zlCarB2.Value
lDep = Forms("F_Demo")!txtBoucle2Dep.Value
'DEMO

    For i = 1 To Boucle         ' DEMO barre 1
        lRand = Int((700 * Rnd) + 1)

        sAvance = "Boucle 1 : " & CStr(i)
        cLabels.UpdateBarre sAvance                          'Code C_LabelsInfo Affiche l'état d'avancement du scan...
        DoEvents            '// Refresh labels.

        If (i = lDep) Then      ' DEMO barre 2

            cLabels.SetLengthBarre sCar, lBoucle, LabelNum2  'Code C_LabelsInfo
            BoucleTestCls2 lBoucle
            cLabels.ReturnFirstBarre                         'Code C_LabelsInfo

        End If

        Sleep (lRand)
   Next i

    cLabels.HiddenCurrentLabels                              'Code C_LabelsInfo

End Sub

Public Sub BoucleTestCls2(Boucle As Long)
    Dim i As Long
    Dim lRand As Long
    Dim sAvance As String

    For i = 1 To Boucle
        lRand = Int((700 * Rnd) + 1)
        sAvance = "Boucle 2 : " & CStr(i)
        cLabels.UpdateBarre sAvance                         'Code C_LabelsInfo Affiche l'état d'avancement du scan...
        DoEvents            '// Refresh labels.
        Sleep (lRand)
   Next i
End Sub
