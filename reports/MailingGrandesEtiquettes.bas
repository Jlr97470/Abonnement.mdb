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
    Width =5346
    DatasheetFontHeight =10
    ItemSuffix =6
    Left =135
    Top =45
    DatasheetGridlinesColor =12632256
    RecSrcDt = Begin
        0xa597dcf930b9e140
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
            Height =300
            FontSize =10
            FontWeight =300
            FontName ="Arial"
        End
        Begin BreakLevel
            ControlSource ="CodePostal"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =2160
            Name ="Detail"
            Begin
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =240
                    Width =5058
                    Name ="Text2"
                    ControlSource ="=Trim([Prénom] & \" \" & [NomFamille])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =540
                    Width =5058
                    TabIndex =1
                    Name ="Text3"
                    ControlSource ="=Trim([Société])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =840
                    Width =5058
                    TabIndex =2
                    Name ="Text4"
                    ControlSource ="=Trim([Adresse])"
                End
                Begin TextBox
                    CanGrow = NotDefault
                    CanShrink = NotDefault
                    TextAlign =1
                    Left =144
                    Top =1140
                    Width =5058
                    TabIndex =3
                    Name ="Text5"
                    ControlSource ="=Trim([CodePostal] & \"  \" & [Ville])"
                End
            End
        End
    End
End
