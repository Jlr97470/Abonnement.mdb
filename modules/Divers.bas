Attribute VB_Name = "Divers"
Option Compare Database
Option Explicit

Const conAppName = "Abonnement"

Public Function IsOpen(ByVal strFormName As String) As Boolean
'Retourne la valeur True si le formulaire spécifié est ouvert en mode Formulaire.

    Const conDesignView = 0
    Const conObjStateClosed = 0

    IsOpen = False
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> _
            conObjStateClosed Then

        If Forms(strFormName).CurrentView <> conDesignView Then
            IsOpen = True
        End If
    End If
End Function

Public Sub DisplayMessage(strMessage As String)
'Affiche un message important à l'utilisateur.

    MsgBox strMessage, vbExclamation, conAppName

End Sub

Public Function Confirm(strMessage As String) As Boolean
'Demande à l'utilisateur de confirmer une action, et retourne True ou False.

    Dim bytChoice As Byte

    bytChoice = MsgBox(strMessage, vbQuestion + vbOKCancel, conAppName)

    If bytChoice = vbOK Then
        Confirm = True
    Else
        Confirm = False
    End If

End Function

Public Function AfficherHistoriquePaiements()
'Ouvre le formulaire HistoriquePaiements et définit le filtre.

    DoCmd.OpenForm "HistoriquePaiements", , , "RéfAbonné" = " & Forms!Abonnés!RéfAbonné"

End Function
