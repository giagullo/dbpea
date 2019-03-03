Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =5102
    DatasheetFontHeight =11
    ItemSuffix =46
    Top =600
    Right =8910
    Bottom =7620
    DatasheetGridlinesColor =14806254
    Filter ="scenario = '4Q2018'"
    RecSrcDt = Begin
        0xec05a0be6f40e540
    End
    RecordSource ="qry_EditScenario"
    DatasheetFontName ="Calibri"
    AllowDatasheetView =0
    AllowPivotTableView =0
    AllowPivotChartView =0
    AllowPivotChartView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            Width =1701
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin EmptyCell
            Height =240
            GridlineColor =12632256
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin Section
            Height =3911
            Name ="Corpo"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin EmptyCell
                    Left =345
                    Top =345
                    Height =315
                    Name ="CellaVuota27"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =345
                    LayoutCachedTop =345
                    LayoutCachedWidth =1785
                    LayoutCachedHeight =660
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1700
                    Top =1474
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Testo28"
                    ControlSource ="dtInizio"
                    GridlineColor =10921638

                    LayoutCachedLeft =1700
                    LayoutCachedTop =1474
                    LayoutCachedWidth =3401
                    LayoutCachedHeight =1789
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1474
                            Width =795
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta29"
                            Caption ="Inizio"
                            GridlineColor =10921638
                            LayoutCachedTop =1474
                            LayoutCachedWidth =795
                            LayoutCachedHeight =1789
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1700
                    Top =1984
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Testo34"
                    ControlSource ="dtFine"
                    GridlineColor =10921638

                    LayoutCachedLeft =1700
                    LayoutCachedTop =1984
                    LayoutCachedWidth =3401
                    LayoutCachedHeight =2299
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1984
                            Width =795
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta35"
                            Caption ="Fine"
                            GridlineColor =10921638
                            LayoutCachedTop =1984
                            LayoutCachedWidth =795
                            LayoutCachedHeight =2299
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1757
                    Top =566
                    Width =3345
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="CasellaCombinata40"
                    ControlSource ="descProgetto"
                    RowSourceType ="Table/Query"
                    GridlineColor =10921638

                    LayoutCachedLeft =1757
                    LayoutCachedTop =566
                    LayoutCachedWidth =5102
                    LayoutCachedHeight =881
                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =56
                            Top =566
                            Width =1950
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta41"
                            Caption ="Progetto"
                            GridlineColor =10921638
                            LayoutCachedLeft =56
                            LayoutCachedTop =566
                            LayoutCachedWidth =2006
                            LayoutCachedHeight =881
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =1757
                    Top =1077
                    Width =3345
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="CasellaCombinata44"
                    ControlSource ="descTask"
                    RowSourceType ="Table/Query"
                    GridlineColor =10921638

                    LayoutCachedLeft =1757
                    LayoutCachedTop =1077
                    LayoutCachedWidth =5102
                    LayoutCachedHeight =1392
                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =56
                            Top =1077
                            Width =1950
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta45"
                            Caption ="Task"
                            GridlineColor =10921638
                            LayoutCachedLeft =56
                            LayoutCachedTop =1077
                            LayoutCachedWidth =2006
                            LayoutCachedHeight =1392
                        End
                    End
                End
            End
        End
    End
End
