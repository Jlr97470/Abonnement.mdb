Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Form_Close()
'Ferme le formulaire indépendant HistoriquePaiements s'il est ouvert.

    If IsOpen("HistoriquePaiements") Then
        DoCmd.Close acForm, "HistoriquePaiements"
    End If

End Sub

Private Sub Form_Current()
'Synchronise le formulaire HistoriquePaiements s'il est ouvert.

    If IsOpen("HistoriquePaiements") Then
        Forms!HistoriquePaiements.Filter = "RéfAbonné = " & Nz(RéfAbonné, 0)
    End If

End Sub
