Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ShortcutMenu = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    Width =5616
    DatasheetFontHeight =10
    ItemSuffix =30
    Left =3975
    Top =1860
    Right =10950
    Bottom =6240
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x0a761153d640e140
    End
    Caption ="Impression des étiquettes adresse"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    Begin
        Begin Label
            BackStyle =0
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
        Begin OptionGroup
            SpecialEffect =3
        End
        Begin TextBox
            SpecialEffect =2
            OldBorderStyle =0
        End
        Begin Section
            Height =3000
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =73
                    Left =864
                    Top =2520
                    Height =367
                    TabIndex =1
                    Name ="ImprimerEtiquettes"
                    Caption ="&Imprimer"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    OverlapFlags =85
                    AccessKey =65
                    Left =2448
                    Top =2520
                    Height =367
                    TabIndex =2
                    Name ="ApercuEtiquettes"
                    Caption ="&Aperçu étiquettes"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =4032
                    Top =2520
                    Height =367
                    TabIndex =3
                    Name ="Annuler"
                    Caption ="Annuler"
                    OnClick ="[Event Procedure]"
                End
                Begin Tab
                    OverlapFlags =85
                    Left =150
                    Top =90
                    Width =5325
                    Height =2310
                    Name ="Tabs"
                    Begin
                        Begin Page
                            OverlapFlags =215
                            Left =285
                            Top =495
                            Width =5055
                            Height =1770
                            Name ="Page1"
                            Caption ="Type de mailing"
                            Begin
                                Begin OptionGroup
                                    OverlapFlags =215
                                    Left =288
                                    Top =600
                                    Width =4752
                                    Height =1447
                                    Name ="TypeMailing"
                                    DefaultValue ="1"
                                    Begin
                                        Begin OptionButton
                                            OverlapFlags =215
                                            Left =432
                                            Top =748
                                            OptionValue =1
                                            Name ="Option4"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =662
                                                    Top =720
                                                    Width =3037
                                                    Height =249
                                                    Name ="Label5"
                                                    Caption ="Envoyer le journal (tous les abonnés)"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =215
                                            Left =432
                                            Top =1108
                                            OptionValue =2
                                            Name ="Option6"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =215
                                                    Left =668
                                                    Top =1076
                                                    Width =4254
                                                    Height =248
                                                    Name ="Label7"
                                                    Caption ="Envoyer la facture (échéances dans trois mois)"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =215
                                            Left =432
                                            Top =1468
                                            OptionValue =3
                                            Name ="Option8"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =215
                                                    Left =668
                                                    Top =1442
                                                    Width =4228
                                                    Height =485
                                                    Name ="Label9"
                                                    Caption ="Relance de réabonnement (abonnements échus dans les trois derniers mois)"
                                                End
                                            End
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =285
                            Top =495
                            Width =5055
                            Height =1770
                            Name ="Page2"
                            Caption ="Type étiquettes"
                            Begin
                                Begin OptionGroup
                                    OverlapFlags =247
                                    Left =288
                                    Top =600
                                    Width =2592
                                    Height =840
                                    Name ="TypeEtiquettes"
                                    DefaultValue ="1"
                                    Begin
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =432
                                            Top =748
                                            OptionValue =1
                                            Name ="Option26"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =662
                                                    Top =720
                                                    Width =1846
                                                    Height =249
                                                    Name ="Label27"
                                                    Caption ="Petites étiquettes (par 3)"
                                                End
                                            End
                                        End
                                        Begin OptionButton
                                            OverlapFlags =247
                                            Left =432
                                            Top =1078
                                            OptionValue =2
                                            Name ="Option28"
                                            Begin
                                                Begin Label
                                                    OverlapFlags =247
                                                    Left =662
                                                    Top =1050
                                                    Width =1905
                                                    Height =240
                                                    Name ="Label29"
                                                    Caption ="Grandes étiquettes (par 2)"
                                                End
                                            End
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "DialogueEtiquettesMailing.cls"
