Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Sub Annuler_Click()
On Error GoTo Err_Annuler_Click

    DoCmd.Close

Exit_Annuler_Click:
    Exit Sub

Err_Annuler_Click:
    MsgBox Err.Description
    Resume Exit_Annuler_Click

End Sub

Private Sub ApercuEtiquettes_Click()
'Ouvre un état d'étiquettes de mailing avec un filtre, puis ferme la boîte de dialogue.

    Dim strFilter As String, strReportName As String

    'Détermine le jeu d'enregistrements à utiliser.
    Select Case TypeMailing
    Case 1      'Envoyer le journal - Tous les abonnés.
        strFilter = "Echéance >= Date()"
    Case 2      'Abonnements échus dans les trois prochains mois.
        strFilter = "Echéance >= Date() And Echéance < Date () + 90"
    Case 3      'Abonnements expirés ces trois derniers mois.
        strFilter = "Echéance < Date() And Echéance > Date() - 90"
    End Select

    'Détermine les étiquettes à utiliser.
    If TypeEtiquettes = 1 Then
        strReportName = "MailingPetitesEtiquettes"
    Else
        strReportName = "MailingGrandesEtiquettes"
    End If

    DoCmd.OpenReport strReportName, acViewPreview, , strFilter
    DoCmd.Close acForm, "DialogueEtiquettesMailing"

End Sub

Private Sub ImprimerEtiquettes_Click()
'Ouvre un état d'étiquettes de mailing avec un filtre, puis ferme la boîte de dialogue.

    Dim strFilter As String, strReportName As String

    'Détermine le jeu d'enregistrements à utiliser.
    Select Case TypeMailing
    Case 1      'Envoyer le journal - Tous les abonnés.
        strFilter = "Echéance >= Date()"
    Case 2      'Abonnements échus dans les trois prochains mois.
        strFilter = "Echéance >= Date() And Echéance < Date () + 90"
    Case 3      'Abonnements expirés ces trois derniers mois.
        strFilter = "Echéance < Date() And Echéance > Date() - 90"
    End Select

    'Détermine les étiquettes à utiliser.
    If TypeEtiquettes = 1 Then
        strReportName = "MailingPetitesEtiquettes"
    Else
        strReportName = "MailingGrandesEtiquettes"
    End If

    DoCmd.OpenReport strReportName, acViewNormal, , strFilter
    DoCmd.Close acForm, "DialogueEtiquettesMailing"

End Sub
