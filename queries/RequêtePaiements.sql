SELECT DISTINCTROW
  Paiements.NoPaiement,
  Paiements.RéfAbonné,
  Paiements.DatePaiement,
  Abonnés.Echéance,
  Paiements.MontantAnnuel,
  Paiements.AnnéesPayées,
  Paiements.ModePaiement,
  Paiements.CarteCrédit,
  Paiements.NoCarteCrédit,
  Paiements.Expire,
  Paiements.Autorisation
FROM
  Abonnés
  INNER JOIN Paiements ON Abonnés.RéfAbonné = Paiements.RéfAbonné;
