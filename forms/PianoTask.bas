﻿Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =11520
    DatasheetFontHeight =11
    ItemSuffix =9
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0xb3e0bad17640e540
    End
    RecordSource ="PianoTask"
    Caption ="PianoTask"
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
                    TextAlign =3
                    Left =342
                    Top =684
                    Width =1425
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="IDTask_Etichetta"
                    Caption ="IDTask"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =684
                    LayoutCachedWidth =1767
                    LayoutCachedHeight =999
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1824
                    Top =684
                    Width =6783
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="scenario_Etichetta"
                    Caption ="scenario"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =1824
                    LayoutCachedTop =684
                    LayoutCachedWidth =8607
                    LayoutCachedHeight =999
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =8664
                    Top =684
                    Width =1539
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="dtInizio_Etichetta"
                    Caption ="dtInizio"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =8664
                    LayoutCachedTop =684
                    LayoutCachedWidth =10203
                    LayoutCachedHeight =999
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =10260
                    Top =684
                    Width =1203
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="dtFine_Etichetta"
                    Caption ="dtFine"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =10260
                    LayoutCachedTop =684
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =999
                End
                Begin Label
                    OverlapFlags =215
                    Left =57
                    Top =57
                    Width =2070
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Etichetta8"
                    Caption ="PianoTask"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =2127
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =714
            Name ="Corpo"
            AutoHeight =1
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =1425
                    Height =330
                    ColumnWidth =1530
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="IDTask"
                    ControlSource ="IDTask"
                    StatusBarText ="punt ai dati statici del task pianificato"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =1767
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1824
                    Top =57
                    Width =6783
                    Height =600
                    ColumnWidth =3000
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="scenario"
                    ControlSource ="scenario"
                    StatusBarText ="Identifica lo scenario"
                    GridlineColor =10921638

                    LayoutCachedLeft =1824
                    LayoutCachedTop =57
                    LayoutCachedWidth =8607
                    LayoutCachedHeight =657
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =8664
                    Top =57
                    Width =1539
                    Height =330
                    ColumnWidth =1620
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="dtInizio"
                    ControlSource ="dtInizio"
                    StatusBarText ="data inizio task nello scenario ipotizzato"
                    GridlineColor =10921638

                    LayoutCachedLeft =8664
                    LayoutCachedTop =57
                    LayoutCachedWidth =10203
                    LayoutCachedHeight =387
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =10260
                    Top =57
                    Width =1203
                    Height =330
                    ColumnWidth =1620
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="dtFine"
                    ControlSource ="dtFine"
                    StatusBarText ="data fine task nello scenario ipotizzato"
                    GridlineColor =10921638

                    LayoutCachedLeft =10260
                    LayoutCachedTop =57
                    LayoutCachedWidth =11463
                    LayoutCachedHeight =387
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
