Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =16110
    DatasheetFontHeight =11
    ItemSuffix =11
    Top =600
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x1d481853753be540
    End
    RecordSource ="SELECT [Progetto].[codSIPROS] AS Progetto_codSIPROS, [Progetto].[descProgetto], "
        "[Task].[codSIPROS] AS Task_codSIPROS, [Task].[descTask] FROM Progetto INNER JOIN"
        " Task ON [Progetto].[codSIPROS] =[Task].[codSIPROSProg]; "
    Caption ="Progetto"
    DatasheetFontName ="Calibri"
    PrtMip = Begin
        0x6801000068010000680100006801000000000000f63e00004a01000001000000 ,
        0x010000006801000000000000a10700000100000001000000
    End
    FilterOnLoad =0
    FitToPage =1
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
            ShowDatePicker =0
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="Progetto_codSIPROS"
        End
        Begin BreakLevel
            ControlSource ="Task_codSIPROS"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =597
            BackColor =15849926
            Name ="IntestazioneReport"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin Label
                    OverlapFlags =81
                    Left =57
                    Top =57
                    Width =1515
                    Height =540
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Etichetta8"
                    Caption ="Progetto"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =1572
                    LayoutCachedHeight =597
                End
            End
        End
        Begin PageHeader
            Height =372
            Name ="SezioneIntestazionePagina"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =81
                    TextAlign =1
                    Left =342
                    Top =57
                    Width =4047
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Progetto_codSIPROS_Etichetta"
                    Caption ="Progetto_codSIPROS"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =4389
                    LayoutCachedHeight =372
                End
                Begin Label
                    OverlapFlags =83
                    TextAlign =1
                    Left =4389
                    Top =57
                    Width =4047
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="descProgetto_Etichetta"
                    Caption ="descProgetto"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4389
                    LayoutCachedTop =57
                    LayoutCachedWidth =8436
                    LayoutCachedHeight =372
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =1
                    Left =8778
                    Top =117
                    Width =2442
                    Height =255
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Task_codSIPROS_Etichetta"
                    Caption ="Task_codSIPROS"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8778
                    LayoutCachedTop =117
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =372
                End
                Begin Label
                    OverlapFlags =81
                    TextAlign =1
                    Left =11340
                    Top =57
                    Width =4721
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="descTask_Etichetta"
                    Caption ="descTask"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =11340
                    LayoutCachedTop =57
                    LayoutCachedWidth =16061
                    LayoutCachedHeight =372
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="IntestazioneGruppo0"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
        Begin Section
            KeepTogether = NotDefault
            Height =330
            Name ="Corpo"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =81
                    IMESentenceMode =3
                    Left =342
                    Width =4047
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Progetto_codSIPROS"
                    ControlSource ="Progetto_codSIPROS"
                    StatusBarText ="codice attribuito in SIPROS o PPPM"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedWidth =4389
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    HideDuplicates = NotDefault
                    OldBorderStyle =0
                    OverlapFlags =83
                    IMESentenceMode =3
                    Left =4389
                    Width =4047
                    Height =330
                    ColumnWidth =5250
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="descProgetto"
                    ControlSource ="descProgetto"
                    StatusBarText ="nome progetto come in SIPROS"
                    GridlineColor =10921638

                    LayoutCachedLeft =4389
                    LayoutCachedWidth =8436
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =81
                    IMESentenceMode =3
                    Left =8778
                    Width =2442
                    Height =330
                    ColumnWidth =3030
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Task_codSIPROS"
                    ControlSource ="Task_codSIPROS"
                    StatusBarText ="Codice del task in SIPROS NNN-XXXX-NNN-NNN se presente"
                    GridlineColor =10921638

                    LayoutCachedLeft =8778
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    OldBorderStyle =0
                    OverlapFlags =243
                    IMESentenceMode =3
                    Left =10770
                    Width =5291
                    Height =330
                    ColumnWidth =3630
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="descTask"
                    ControlSource ="descTask"
                    StatusBarText ="Descrizione task"
                    GridlineColor =10921638

                    LayoutCachedLeft =10770
                    LayoutCachedWidth =16061
                    LayoutCachedHeight =330
                End
            End
        End
        Begin PageFooter
            Height =558
            Name ="SezionePièDiPaginaPagina"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =1
                    IMESentenceMode =3
                    Left =57
                    Top =228
                    Width =5040
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Testo9"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    GridlineColor =10921638

                    LayoutCachedLeft =57
                    LayoutCachedTop =228
                    LayoutCachedWidth =5097
                    LayoutCachedHeight =558
                End
                Begin TextBox
                    OldBorderStyle =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =11021
                    Top =228
                    Width =5040
                    Height =330
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Testo10"
                    ControlSource ="=\"Pagina \" & [Page] & \" di \" & [Pages]"
                    GridlineColor =10921638

                    LayoutCachedLeft =11021
                    LayoutCachedTop =228
                    LayoutCachedWidth =16061
                    LayoutCachedHeight =558
                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="PièDiPaginaReport"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
