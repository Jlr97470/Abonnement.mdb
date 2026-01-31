Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_AfterUpdate()
'Met les autres formulaires à jour s'ils sont ouverts.

    If IsOpen("Abonnés") Then
        'Actualise le formulaire Abonnées pour afficher la valeur Echéance.
        Forms!Abonnés.Refresh
    End If

    If IsOpen("HistoriquePaiements") Then
        'Actualise le formulaire indépendant HistoriquePaiements pour afficher le nouveau paiement.
        Forms!HistoriquePaiements.Requery
    End If

End Sub

Private Sub Form_BeforeUpdate(Cancel As Integer)
' Vérifie qu'un mode de paiement est sélectionné, puis actualise le champ
' Echéance (dans la table Abonnés) pour refléter le délai supplémentaire
' dû à ce paiement.

    Dim bytMonths As Byte

    If ModePaiement = 0 Then
        DisplayMessage "Vous devez sélectionner un mode de paiement."
        Cancel = True
        Exit Sub
    End If

    bytMonths = AnnéesPayées * 12

    If IsNull(Echéance) Or Echéance < Date Then
        ' S'il s'agit du premier paiement, définit l'échéance à la date en cours.
        Echéance = DateSerial(Year(Date), Month(Date) + bytMonths, 1)
    Else
        ' Sinon, ajoute les mois supplémentaires à la valeur de l'échéance.
        Echéance = DateSerial(Year(Echéance), Month(Echéance) + _
            bytMonths, 1)
    End If

End Sub

Private Sub Form_Open(Cancel As Integer)
'Définit l'abonné en cours comme le défaut dans le formulaire Abonnés.

    If IsOpen("Abonnés") Then
        RéfAbonné.DefaultValue = Forms!Abonnés!RéfAbonné
    End If

End Sub

Private Sub ModePaiement_AfterUpdate()
' Si l'utilisateur sélectionne carte de crédit, active les champs carte de crédit.

    Const conCreditItem = 3

    If ModePaiement = conCreditItem Then
        CarteCrédit.Enabled = True
        NoCarteCrédit.Enabled = True
        Expire.Enabled = True
        Autorisation.Enabled = True
        CarteCrédit.SetFocus
    Else
        CarteCrédit.Enabled = False
        NoCarteCrédit.Enabled = False
        Expire.Enabled = False
        Autorisation.Enabled = False
    End If

End Sub

Sub Enregistrer_Click()
On Error GoTo Err_Enregistrer_Click

    ' Ouvre un nouvel enregistrement (enregistre l'enregistrement en cours si modifié).
    DoCmd.GoToRecord , , acNewRec

Exit_Enregistrer_Click:
    Exit Sub

Err_Enregistrer_Click:
    MsgBox Err.Description
    Resume Exit_Enregistrer_Click

End Sub

Sub AjouterNouveau_Click()
On Error GoTo Err_AjouterNouveau_Click


    DoCmd.GoToRecord , , acNewRec

Exit_AjouterNouveau_Click:
    Exit Sub

Err_AjouterNouveau_Click:
    MsgBox Err.Description
    Resume Exit_AjouterNouveau_Click

End Sub

Sub Fermer_Click()
On Error GoTo Err_Fermer_Click


    DoCmd.Close

Exit_Fermer_Click:
    Exit Sub

Err_Fermer_Click:
    MsgBox Err.Description
    Resume Exit_Fermer_Click

End Sub
