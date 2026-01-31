Version =19
VersionRequired =19
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6372
    DatasheetFontHeight =10
    ItemSuffix =4
    Left =2460
    Top =1845
    Right =9735
    Bottom =6090
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0x128dc7da0dbbe140
    End
    RecordSource ="Abonnés"
    Caption ="Atteindre un enregistrement"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
        End
        Begin ListBox
            SpecialEffect =2
            Width =1701
            Height =1417
            LabelX =-1701
            FontName ="Tahoma"
        End
        Begin Section
            Height =2267
            BackColor =-2147483633
            Name ="Détail"
            Begin
                Begin ListBox
                    OverlapFlags =85
                    ColumnCount =4
                    Left =56
                    Top =388
                    Width =4200
                    Height =1822
                    Name ="Liste0"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Abonnés.RéfAbonné, Abonnés.NomFamille, Abonnés.Prénom, Abonnés.Ville FROM"
                        " Abonnés ORDER BY Abonnés.NomFamille, Abonnés.Prénom;"
                    ColumnWidths ="0;1125;810;1830"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =113
                            Width =1890
                            Height =240
                            Name ="Sélectionnez un abonné :_Étiquette"
                            Caption ="Sélectionnez un abonné :"
                            EventProcPrefix ="Sélectionnez_un_abonné___Étiquette"
                        End
                    End
                End
                Begin CommandButton
                    Default = NotDefault
                    Enabled = NotDefault
                    OverlapFlags =85
                    AccessKey =65
                    Left =4422
                    Top =283
                    Width =1920
                    Height =405
                    TabIndex =1
                    Name ="AfficherEnregistrement"
                    Caption ="&Afficher Enregistrement"
                    OnClick ="[Event Procedure]"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =4478
                    Top =850
                    Width =1875
                    Height =405
                    TabIndex =2
                    Name ="Annuler"
                    Caption ="Annuler"
                    OnClick ="[Event Procedure]"
                End
            End
        End
    End
End
CodeBehindForm
' See "DialogueAtteindreEnregistrement.cls"
