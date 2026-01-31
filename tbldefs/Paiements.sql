CREATE TABLE [Paiements] (
  [NoPaiement] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [RéfAbonné] LONG,
  [DatePaiement] DATETIME,
  [MontantAnnuel] CURRENCY,
  [AnnéesPayées] SINGLE,
  [ModePaiement] BYTE,
  [CarteCrédit] VARCHAR (50),
  [NoCarteCrédit] VARCHAR (50),
  [Expire] DATETIME,
  [Autorisation] VARCHAR (50)
)
