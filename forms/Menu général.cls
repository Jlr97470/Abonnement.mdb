Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub OuvriFormulaireAbonnés_Click()
On Error GoTo Err_OuvrirFormulaireAbonnés_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Abonnés"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OuvrirFormulaireAbonnés_Click:
    Exit Sub

Err_OuvrirFormulaireAbonnés_Click:
    MsgBox Err.Description
    Resume Exit_OuvrirFormulaireAbonnés_Click

End Sub

Sub OuvrirDialogueMailing_Click()
On Error GoTo Err_OuvrirDialogueMailing_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "DialogueEtiquettesMailing"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

Exit_OuvrirDialogueMailing_Click:
    Exit Sub

Err_OuvrirDialogueMailing_Click:
    MsgBox Err.Description
    Resume Exit_OuvrirDialogueMailing_Click

End Sub
