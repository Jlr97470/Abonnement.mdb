Version =19
VersionRequired =19
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowEdits = NotDefault
    DataEntry = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Width =7632
    DatasheetFontHeight =10
    ItemSuffix =40
    Left =1605
    Top =795
    Right =9240
    Bottom =5100
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xc4afba5af2bae140
    End
    RecordSource ="RequêtePaiements"
    Caption ="Saisie des paiements"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin OptionButton
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BackStyle =0
        End
        Begin TextBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ListBox
            SpecialEffect =2
            BackColor =-2147483643
            BorderColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =3720
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =87
                    ColumnCount =2
                    ListWidth =2880
                    Left =1296
                    Top =240
                    Width =2801
                    Height =261
                    ColumnWidth =2801
                    Name ="RéfAbonné"
                    ControlSource ="RéfAbonné"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Abonnés.RéfAbonné, [NomFamille] & \"  \" & [Prénom] AS Expr1 FROM Abonnés"
                        " ORDER BY [NomFamille] & \"  \" & [Prénom];"
                    ColumnWidths ="0"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =144
                            Top =240
                            Width =1126
                            Height =249
                            Name ="Abonné"
                            Caption ="Abonné"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =1474
                    Top =944
                    Width =1060
                    Height =255
                    ColumnWidth =1060
                    TabIndex =1
                    Name ="DatePaiement"
                    ControlSource ="DatePaiement"
                    DefaultValue ="Date()"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =144
                            Top =960
                            Width =445
                            Height =249
                            Name ="PaymentDate Label"
                            Caption ="Date"
                            EventProcPrefix ="PaymentDate_Label"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =1474
                    Top =1322
                    Width =1584
                    Height =261
                    ColumnWidth =2343
                    TabIndex =2
                    Name ="MontantAnnuel"
                    ControlSource ="MontantAnnuel"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    DefaultValue ="24"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =144
                            Top =1320
                            Width =1185
                            Height =240
                            Name ="Montant annuel"
                            Caption ="Montant annuel"
                            EventProcPrefix ="Montant_annuel"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    Left =1474
                    Top =1700
                    Width =1060
                    Height =261
                    ColumnWidth =1060
                    TabIndex =3
                    Name ="AnnéesPayées"
                    ControlSource ="AnnéesPayées"
                    DefaultValue ="1"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =144
                            Top =1680
                            Width =1170
                            Height =240
                            Name ="Années payées"
                            Caption ="Années payées"
                            EventProcPrefix ="Années_payées"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =5040
                    Top =1920
                    Width =1584
                    Height =261
                    ColumnWidth =2801
                    TabIndex =4
                    Name ="CarteCrédit"
                    ControlSource ="CarteCrédit"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3888
                            Top =1920
                            Width =1065
                            Height =240
                            Name ="NomCarteCrédit"
                            Caption ="Type de carte"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =5040
                    Top =2280
                    Width =2304
                    Height =261
                    ColumnWidth =2801
                    TabIndex =5
                    Name ="NoCarteCrédit"
                    ControlSource ="NoCarteCrédit"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3888
                            Top =2280
                            Width =630
                            Height =240
                            Name ="NuméroCarteCrédit"
                            Caption ="Numéro"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =5040
                    Top =2640
                    Width =863
                    Height =261
                    ColumnWidth =693
                    TabIndex =6
                    Name ="Expire"
                    ControlSource ="Expire"
                    Format ="m/yy"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3888
                            Top =2640
                            Width =838
                            Height =249
                            Name ="DateExpiration"
                            Caption ="Expire le"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    OverlapFlags =93
                    Left =5040
                    Top =3000
                    Width =1872
                    Height =261
                    ColumnWidth =2801
                    TabIndex =7
                    Name ="Autorisation"
                    ControlSource ="Autorisation"
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =3888
                            Top =3000
                            Width =1060
                            Height =249
                            Name ="AutorisationCarteCrédit"
                            Caption ="Autorisation"
                        End
                    End
                End
                Begin OptionGroup
                    SpecialEffect =2
                    OverlapFlags =247
                    Left =3312
                    Top =960
                    Width =4176
                    Height =2531
                    TabIndex =8
                    Name ="ModePaiement"
                    ControlSource ="ModePaiement"
                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =247
                            Left =3456
                            Top =840
                            Width =1590
                            Height =249
                            Name ="Label21"
                            Caption ="Mode de paiement"
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =3456
                            Top =1228
                            OptionValue =1
                            Name ="Option23"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3686
                                    Top =1200
                                    Width =630
                                    Height =240
                                    Name ="Label24"
                                    Caption ="Chèque"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =5472
                            Top =1230
                            OptionValue =2
                            Name ="Option25"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =5702
                                    Top =1200
                                    Width =690
                                    Height =240
                                    Name ="Label26"
                                    Caption ="Espèces"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =119
                            Left =3456
                            Top =1588
                            OptionValue =3
                            Name ="Option27"
                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =3686
                                    Top =1560
                                    Width =1110
                                    Height =240
                                    Name ="Label28"
                                    Caption ="Carte de crédit"
                                End
                            End
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =1474
                    Top =2125
                    Width =1584
                    Height =261
                    TabIndex =9
                    BackColor =13434879
                    Name ="Text29"
                    ControlSource ="=[MontantAnnuel]*[AnnéesPayées]"
                    Format ="#,##0.00\" F\";-#,##0.00\" F\""
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =144
                            Top =2160
                            Width =1100
                            Height =249
                            Name ="Somme"
                            Caption ="Somme totale"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    Left =6480
                    Top =240
                    Width =1008
                    TabIndex =10
                    BackColor =12632256
                    Name ="PaidThrough"
                    ControlSource ="Echéance"
                    Format ="mmm yy"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5184
                            Top =240
                            Width =1165
                            Height =249
                            Name ="Label35"
                            Caption ="Échéance"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =600
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =69
                    Left =288
                    Top =120
                    Width =1872
                    Height =368
                    Name ="Enregistrer"
                    Caption ="&Enregistrer le paiement"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =78
                    Left =2880
                    Top =120
                    Width =2160
                    Height =367
                    TabIndex =1
                    Name ="AjouterNouveau"
                    Caption ="Ajouter &nouveau paiement"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =70
                    Left =5760
                    Top =120
                    Width =1584
                    Height =367
                    TabIndex =2
                    Name ="Fermer"
                    Caption ="&Fermer"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "SaisiePaiements.cls"
