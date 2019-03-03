Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =12760
    DatasheetFontHeight =11
    ItemSuffix =15
    Left =630
    Top =300
    Right =13680
    Bottom =10785
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x980ff4407040e540
    End
    RecordSource ="qry_EditScenario"
    Caption ="qry_EditScenario"
    OnCurrent ="[Event Procedure]"
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
                    OverlapFlags =93
                    TextAlign =1
                    Left =342
                    Top =684
                    Width =4560
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="descProgetto_Etichetta"
                    Caption ="descProgetto"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =342
                    LayoutCachedTop =684
                    LayoutCachedWidth =4902
                    LayoutCachedHeight =999
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4959
                    Top =684
                    Width =4560
                    Height =315
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="descTask_Etichetta"
                    Caption ="descTask"
                    Tag ="DetachedLabel"
                    GridlineStyleBottom =1
                    GridlineColor =10921638
                    LayoutCachedLeft =4959
                    LayoutCachedTop =684
                    LayoutCachedWidth =9519
                    LayoutCachedHeight =999
                End
                Begin Label
                    OverlapFlags =215
                    Left =57
                    Top =57
                    Width =3420
                    Height =969
                    FontSize =20
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Etichetta8"
                    Caption ="qry_EditScenario"
                    GridlineColor =10921638
                    LayoutCachedLeft =57
                    LayoutCachedTop =57
                    LayoutCachedWidth =3477
                    LayoutCachedHeight =1026
                End
            End
        End
        Begin Section
            Height =1190
            Name ="Corpo"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =342
                    Top =57
                    Width =4560
                    Height =600
                    ColumnWidth =3000
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="descProgetto"
                    ControlSource ="descProgetto"
                    StatusBarText ="nome progetto come in SIPROS"
                    GridlineColor =10921638

                    LayoutCachedLeft =342
                    LayoutCachedTop =57
                    LayoutCachedWidth =4902
                    LayoutCachedHeight =657
                End
                Begin TextBox
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4959
                    Top =57
                    Width =4560
                    Height =600
                    ColumnWidth =3000
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="descTask"
                    ControlSource ="descTask"
                    StatusBarText ="Descrizione task"
                    GridlineColor =10921638

                    LayoutCachedLeft =4959
                    LayoutCachedTop =57
                    LayoutCachedWidth =9519
                    LayoutCachedHeight =657
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6519
                    Top =793
                    Width =1371
                    Height =330
                    ColumnWidth =1620
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="dtInizio"
                    ControlSource ="dtInizio"
                    StatusBarText ="data inizio task nello scenario ipotizzato"
                    GridlineColor =10921638

                    LayoutCachedLeft =6519
                    LayoutCachedTop =793
                    LayoutCachedWidth =7890
                    LayoutCachedHeight =1123
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7993
                    Top =793
                    Width =1134
                    Height =330
                    ColumnWidth =1620
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="dtFine"
                    ControlSource ="dtFine"
                    StatusBarText ="data fine task nello scenario ipotizzato"
                    GridlineColor =10921638

                    LayoutCachedLeft =7993
                    LayoutCachedTop =793
                    LayoutCachedWidth =9127
                    LayoutCachedHeight =1123
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2097
                    Top =793
                    Width =1134
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Testo9"
                    ControlSource ="IDTask"
                    GridlineColor =10921638

                    LayoutCachedLeft =2097
                    LayoutCachedTop =793
                    LayoutCachedWidth =3231
                    LayoutCachedHeight =1108
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =396
                            Top =793
                            Width =690
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta10"
                            Caption ="id task"
                            GridlineColor =10921638
                            LayoutCachedLeft =396
                            LayoutCachedTop =793
                            LayoutCachedWidth =1086
                            LayoutCachedHeight =1108
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5152
                    Top =793
                    Width =1026
                    Height =315
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtScenario"
                    ControlSource ="scenario"
                    GridlineColor =10921638

                    LayoutCachedLeft =5152
                    LayoutCachedTop =793
                    LayoutCachedWidth =6178
                    LayoutCachedHeight =1108
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3458
                            Top =795
                            Width =870
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta12"
                            Caption ="Scenario"
                            GridlineColor =10921638
                            LayoutCachedLeft =3458
                            LayoutCachedTop =795
                            LayoutCachedWidth =4328
                            LayoutCachedHeight =1110
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =9751
                    Top =170
                    Width =1984
                    Height =315
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtCodSIPROS"
                    ControlSource ="Progetto.codSIPROS"
                    GridlineColor =10921638

                    LayoutCachedLeft =9751
                    LayoutCachedTop =170
                    LayoutCachedWidth =11735
                    LayoutCachedHeight =485
                End
            End
        End
        Begin FormFooter
            Height =0
            Name ="PièDiPaginaMaschera"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Current()
    If Me.NewRecord Then
        txtScenario = "1Q2019"
    End If
End Sub
