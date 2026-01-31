Operation =1
Option =8
Begin InputTables
    Name ="Abonnés"
    Name ="Paiements"
End
Begin OutputColumns
    Expression ="Paiements.NoPaiement"
    Expression ="Paiements.RéfAbonné"
    Expression ="Paiements.DatePaiement"
    Expression ="Abonnés.Echéance"
    Expression ="Paiements.MontantAnnuel"
    Expression ="Paiements.AnnéesPayées"
    Expression ="Paiements.ModePaiement"
    Expression ="Paiements.CarteCrédit"
    Expression ="Paiements.NoCarteCrédit"
    Expression ="Paiements.Expire"
    Expression ="Paiements.Autorisation"
End
Begin Joins
    LeftTable ="Abonnés"
    RightTable ="Paiements"
    Expression ="Abonnés.RéfAbonné = Paiements.RéfAbonné"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderOn" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
Begin
End
Begin
    State =0
    Left =-27
    Top =25
    Right =785
    Bottom =377
    Left =-1
    Top =-1
    Right =805
    Bottom =165
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =113
        Top =0
        Name ="Abonnés"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =113
        Top =0
        Name ="Paiements"
        Name =""
    End
End
