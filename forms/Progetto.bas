Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =7
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x6821b24e6f40e540
    End
    RecordSource ="Progetto"
    Caption ="Progetto"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
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
        Begin Subform
            BorderLineStyle =0
            Width =1701
            Height =1701
            BorderThemeColorIndex =1
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            BorderShade =65.0
            ShowPageHeaderAndPageFooter =1
        End
        Begin FormHeader
            Height =1026
            BackColor =15849926
            Name ="IntestazioneMaschera"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =57
                    Top =57
                    Width =1818
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Etichetta6"
                    Caption ="Progetto"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =1875
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =7781
            Name ="Corpo"
            AutoHeight =1
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
                    Name ="descProgetto"
                    ControlSource ="descProgetto"
                    StatusBarText ="nome progetto come in SIPROS"
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
                            Name ="descProgetto_Etichetta"
                            Caption ="descProgetto"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =342
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =672
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =3234
                    Top =1254
                    Width =8229
                    Height =3164
                    TabIndex =1
                    BorderColor =10921638
                    Name ="Task Sottomaschera1"
                    SourceObject ="Form.Task Sottomaschera1"
                    LinkChildFields ="codSIPROSProg"
                    LinkMasterFields ="codSIPROS"
                    EventProcPrefix ="Task_Sottomaschera1"
                    GridlineColor =10921638

                    LayoutCachedLeft =3234
                    LayoutCachedTop =1254
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =4418
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =684
                            Top =1254
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Task Sottomaschera1_Etichetta"
                            Caption ="Task"
                            EventProcPrefix ="Task_Sottomaschera1_Etichetta"
                            GridlineColor =10921638
                            LayoutCachedLeft =684
                            LayoutCachedTop =1254
                            LayoutCachedWidth =3144
                            LayoutCachedHeight =1584
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =3234
                    Top =4503
                    Width =8229
                    Height =3164
                    TabIndex =2
                    BorderColor =10921638
                    Name ="PianoTask Sottomaschera"
                    SourceObject ="Form.PianoTask Sottomaschera"
                    LinkChildFields ="IDTask"
                    LinkMasterFields ="[Task Sottomaschera1].Form![ID]"
                    EventProcPrefix ="PianoTask_Sottomaschera"
                    GridlineColor =10921638

                    LayoutCachedLeft =3234
                    LayoutCachedTop =4503
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =7667
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =684
                            Top =4503
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="PianoTask Sottomaschera_Etichetta"
                            Caption ="PianoTask"
                            EventProcPrefix ="PianoTask_Sottomaschera_Etichetta"
                            GridlineColor =10921638
                            LayoutCachedLeft =684
                            LayoutCachedTop =4503
                            LayoutCachedWidth =3144
                            LayoutCachedHeight =4833
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
