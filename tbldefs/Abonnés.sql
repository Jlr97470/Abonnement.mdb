CREATE TABLE [Abonnés] (
  [RéfAbonné] AUTOINCREMENT CONSTRAINT [PrimaryKey] PRIMARY KEY UNIQUE NOT NULL,
  [Prénom] VARCHAR (50),
  [NomFamille] VARCHAR (50),
  [Société] VARCHAR (50),
  [Adresse] VARCHAR (255),
  [Ville] VARCHAR (50),
  [CodePostal] VARCHAR (20),
  [Téléphone] VARCHAR (30),
  [Echéance] DATETIME
)
