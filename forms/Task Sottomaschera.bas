Version =21
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =10209
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =1020
    Top =3130
    Right =16150
    Bottom =7060
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xaf0b5323b43ae540
    End
    RecordSource ="Task"
    Caption ="Task Sottomaschera"
    DatasheetFontName ="Calibri"
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
        Begin ListBox
            BorderLineStyle =0
            Width =1701
            Height =1417
            LabelX =-1701
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =0
            BackColor =15849926
            Name ="IntestazioneMaschera"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =3854
            Name ="Corpo"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =342
                    Width =7260
                    Height =600
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="codSIPROS"
                    ControlSource ="codSIPROS"
                    StatusBarText ="Codice del task in SIPROS NNN-XXXX-NNN-NNN se presente"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =342
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =942
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =342
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="codSIPROS_Etichetta"
                            Caption ="Codice SIPROS"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =672
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1026
                    Width =7260
                    Height =600
                    ColumnWidth =5520
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="descTask"
                    ControlSource ="descTask"
                    StatusBarText ="Descrizione task"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1026
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =1626
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1026
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="descTask_Etichetta"
                            Caption ="Descrizione Task"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1026
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1356
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2892
                    Top =1710
                    Width =7260
                    Height =600
                    ColumnWidth =3000
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="faseTask"
                    ControlSource ="faseTask"
                    StatusBarText ="Fase cui appartiene il task (studio-realizzazione-stability)"
                    GridlineColor =10921638

                    LayoutCachedLeft =2892
                    LayoutCachedTop =1710
                    LayoutCachedWidth =10152
                    LayoutCachedHeight =2310
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =342
                            Top =1710
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="faseTask_Etichetta"
                            Caption ="Fase"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1710
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =2040
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1360
                    Top =2437
                    Width =1440
                    ColumnWidth =2880
                    TabIndex =3
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="Elenco6"
                    ControlSource ="codPortfolio"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Portfolio].[codPortfolio], [Portfolio].[nomPortfolio] FROM Portfolio ORD"
                        "ER BY [codPortfolio]; "
                    ColumnWidths ="0;1440"
                    GridlineColor =10921638

                    LayoutCachedLeft =1360
                    LayoutCachedTop =2437
                    LayoutCachedWidth =2800
                    LayoutCachedHeight =3854
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =2437
                            Width =1350
                            Height =320
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="nomPortfolio_Etichetta"
                            Caption ="nomPortfolio"
                            GridlineColor =10921638
                            LayoutCachedTop =2437
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =2757
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="PièDiPaginaMaschera"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
