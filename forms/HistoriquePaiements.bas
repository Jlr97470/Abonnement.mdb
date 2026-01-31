Version =19
VersionRequired =19
Begin Form
    AllowFilters = NotDefault
    PopUp = NotDefault
    RecordSelectors = NotDefault
    ShortcutMenu = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Width =5760
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =1755
    Top =1500
    Right =7470
    Bottom =3570
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xcfb9786ef2bae140
    End
    RecordSource ="Paiements"
    Caption ="Historique des paiements"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin ComboBox
            SpecialEffect =2
        End
        Begin FormHeader
            Height =428
            BackColor =12632256
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =288
                    Top =94
                    Width =720
                    Height =236
                    FontWeight =700
                    Name ="Label1"
                    Caption ="Date"
                End
                Begin Label
                    OverlapFlags =85
                    Left =1133
                    Top =94
                    Width =1410
                    Height =240
                    FontWeight =700
                    Name ="Label2"
                    Caption ="Montant annuel"
                End
                Begin Label
                    OverlapFlags =85
                    Left =2721
                    Top =94
                    Width =1380
                    Height =240
                    FontWeight =700
                    Name ="Label3"
                    Caption ="Années payées"
                End
                Begin Label
                    OverlapFlags =85
                    Left =4251
                    Top =94
                    Width =1270
                    Height =249
                    FontWeight =700
                    Name ="Label5"
                    Caption ="Montant total"
                End
            End
        End
        Begin Section
            Height =241
            BackColor =12632256
            Name ="Detail"
            Begin
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =288
                    Top =1
                    Width =864
                    BackColor =12632256
                    Name ="PaymentDate"
                    ControlSource ="DatePaiement"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =1440
                    Top =1
                    Width =1152
                    TabIndex =1
                    BackColor =12632256
                    Name ="MontantAnnuel"
                    ControlSource ="MontantAnnuel"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =2880
                    Top =1
                    Width =1035
                    TabIndex =2
                    BackColor =12632256
                    Name ="Années payées"
                    ControlSource ="AnnéesPayées"
                    EventProcPrefix ="Années_payées"
                End
                Begin TextBox
                    SpecialEffect =0
                    OverlapFlags =85
                    TextAlign =1
                    Left =4176
                    Top =1
                    Width =1152
                    TabIndex =3
                    BackColor =12632256
                    Name ="TotalAmount"
                    ControlSource ="=[MontantAnnuel]*[AnnéesPayées]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
