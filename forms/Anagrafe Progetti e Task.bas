Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =15876
    DatasheetFontHeight =11
    ItemSuffix =9
    Top =1350
    Right =14865
    Bottom =8670
    DatasheetGridlinesColor =14806254
    Filter ="([Progetto].[descProgetto] Like \"*GEPA*\")"
    OrderBy ="[Progetto].[codSIPROS]"
    RecSrcDt = Begin
        0x711f4323b43ae540
    End
    RecordSource ="Progetto"
    Caption ="Anagrafe Progetti e Task"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000201c0000e010000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    AllowDatasheetView =0
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
                    Width =4878
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Etichetta8"
                    Caption ="Anagrafe Progetti e Task"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =4935
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =6315
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
                    Name ="codSIPROS"
                    ControlSource ="codSIPROS"
                    StatusBarText ="codice attribuito in SIPROS o PPPM"
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
                            Caption ="Codice SIPROS progetto"
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
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="descProgetto"
                    ControlSource ="descProgetto"
                    StatusBarText ="nome progetto come in SIPROS"
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
                            Name ="descProgetto_Etichetta"
                            Caption ="Nome progetto"
                            GridlineColor =10921638
                            LayoutCachedLeft =342
                            LayoutCachedTop =1026
                            LayoutCachedWidth =2802
                            LayoutCachedHeight =1356
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =85
                    Left =737
                    Top =2097
                    Width =15139
                    Height =4218
                    TabIndex =2
                    BorderColor =10921638
                    Name ="Task Sottomaschera"
                    SourceObject ="Form.Task Sottomaschera"
                    LinkChildFields ="codSIPROSProg"
                    LinkMasterFields ="codSIPROS"
                    EventProcPrefix ="Task_Sottomaschera"
                    GridlineColor =10921638

                    LayoutCachedLeft =737
                    LayoutCachedTop =2097
                    LayoutCachedWidth =15876
                    LayoutCachedHeight =6315
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
