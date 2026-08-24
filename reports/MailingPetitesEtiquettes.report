Version =19
VersionRequired =19
Begin Report
    LayoutForPrint = NotDefault
    DefaultView =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =20
    GridY =24
    Width =3911
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =4035
    Top =705
    DatasheetGridlinesColor =12632256
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xaae611fc30b9e140
    End
    RecordSource ="Abonnés"
    DatasheetFontName ="Arial"
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
        End
        Begin TextBox
            AutoLabel = NotDefault
            OldBorderStyle =0
            FontWeight =300
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="CodePostal"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =1937
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =288
                    Top =240
                    Width =3165
                    Name ="Text2"
                    ControlSource ="=Trim([Prénom] & \" \" & [NomFamille])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =288
                    Top =480
                    Width =3165
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=Trim([Société])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =288
                    Top =720
                    Width =3165
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="=Trim([Adresse])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =288
                    Top =960
                    Width =3165
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="=Trim([CodePostal] & \"  \" & [Ville])"
                End
            End
        End
    End
End
CodeBehindForm
' See "MailingPetitesEtiquettes.cls"
