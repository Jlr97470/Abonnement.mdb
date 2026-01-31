Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub AfficherEnregistrement_Click()
'Recherche l'enregistrement sélectionné puis ferme la boîte de dialogue.

    Dim rst As DAO.Recordset

    'Stocke le jeu d'enregistrement de l'abonné sélectionné.
    Set rst = Forms!Abonnés.RecordsetClone

    'Localise l'enregistrement de l'abonné sélectionné
    rst.FindFirst "RéfAbonné = " & Liste0

    'Définit la propriété Bookmark du formulaire pour se placer sur l'enregistrement.
    Forms!Abonnés.Bookmark = rst.Bookmark

    'Ferme la boîte de dialogue.
    DoCmd.Close acForm, "DialogueAtteindreEnregistrement"

End Sub

Private Sub Liste0_AfterUpdate()
    'Une fois qu'un enregistrement est sélectionné dans la liste, active le bouton AfficherEnregistrement

    AfficherEnregistrement.Enabled = True

End Sub

Private Sub Liste0_DblClick(Cancel As Integer)
'Lorsque l'utilisateur double-clique, agit comme s'il avait cliqué
'sur le bouton AfficherEnregistrement.

    If Not IsNull(Liste0) Then
        AfficherEnregistrement_Click
    End If

End Sub
Private Sub Annuler_Click()
On Error GoTo Err_Annuler_Click


    DoCmd.Close

Exit_Annuler_Click:
    Exit Sub

Err_Annuler_Click:
    MsgBox Err.Description
    Resume Exit_Annuler_Click

End Sub
