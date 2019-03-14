Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =8018
    DatasheetFontHeight =11
    ItemSuffix =41
    Left =4665
    Top =3165
    Right =12975
    Bottom =6105
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x11235d41593ee540
    End
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
        Begin CommandButton
            Width =1701
            Height =283
            FontSize =11
            FontWeight =400
            FontName ="Calibri"
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
            UseTheme =1
            Shape =1
            Gradient =12
            BackThemeColorIndex =4
            BackTint =60.0
            BorderLineStyle =0
            BorderColor =16777215
            BorderThemeColorIndex =4
            BorderTint =60.0
            ThemeFontIndex =1
            HoverThemeColorIndex =4
            HoverTint =40.0
            PressedThemeColorIndex =4
            PressedShade =75.0
            HoverForeThemeColorIndex =0
            HoverForeTint =75.0
            PressedForeThemeColorIndex =0
            PressedForeTint =75.0
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
            Height =2948
            Name ="Corpo"
            OnClick ="[Event Procedure]"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin ListBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =1755
                    Top =360
                    Width =1440
                    Height =1410
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstScenario"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Scenario].[codScenario], [Scenario].[descScenario] FROM Scenario; "
                    ColumnWidths ="0;1440"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =1755
                    LayoutCachedTop =360
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =1770
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =360
                            Width =1335
                            Height =1410
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Scenario_Etichetta"
                            Caption ="Scenario"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =1695
                            LayoutCachedHeight =1770
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ListBox
                    OverlapFlags =85
                    MultiSelect =1
                    IMESentenceMode =3
                    ColumnCount =2
                    Left =4755
                    Top =360
                    Width =3225
                    Height =1410
                    TabIndex =1
                    ForeColor =4210752
                    BorderColor =10921638
                    Name ="lstPortfolio"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Portfolio].[codPortfolio], [Portfolio].[nomPortfolio] FROM Portfolio ORD"
                        "ER BY [codPortfolio]; "
                    ColumnWidths ="0;1440"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0

                    LayoutCachedLeft =4755
                    LayoutCachedTop =360
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =1770
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =3255
                            Top =360
                            Width =1440
                            Height =1410
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Portafogli_Etichetta"
                            Caption ="Portafogli"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =3255
                            LayoutCachedTop =360
                            LayoutCachedWidth =4695
                            LayoutCachedHeight =1770
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    RowSourceTypeInt =1
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =1755
                    Top =1950
                    Width =1440
                    Height =315
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="cmbMese"
                    RowSourceType ="Value List"
                    RowSource ="1;2;3;4;5;6;7;8;9;10;11;12"
                    ColumnWidths ="1440"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    AllowValueListEdits =0
                    InheritValueList =0

                    LayoutCachedLeft =1755
                    LayoutCachedTop =1950
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =2265
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =1950
                            Width =1335
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Mese_Etichetta"
                            Caption ="Mese iniziale"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1950
                            LayoutCachedWidth =1695
                            LayoutCachedHeight =2265
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =4755
                    Top =1950
                    Width =3225
                    Height =315
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="txtAnno"
                    InputMask ="9999"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =4755
                    LayoutCachedTop =1950
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =2265
                    RowStart =1
                    RowEnd =1
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =3255
                            Top =1950
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta23"
                            Caption ="Anno"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =3255
                            LayoutCachedTop =1950
                            LayoutCachedWidth =4695
                            LayoutCachedHeight =2265
                            RowStart =1
                            RowEnd =1
                            ColumnStart =2
                            ColumnEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1755
                    Top =2445
                    Width =1440
                    Height =315
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4138256
                    Name ="txtNumMesi"
                    Format ="General Number"
                    ValidationText ="Immettere un valore numerico"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =1755
                    LayoutCachedTop =2445
                    LayoutCachedWidth =3195
                    LayoutCachedHeight =2760
                    RowStart =2
                    RowEnd =2
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    ForeThemeColorIndex =2
                    ForeTint =100.0
                    ForeShade =50.0
                    GroupTable =1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =360
                            Top =2445
                            Width =1335
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Etichetta29"
                            Caption ="Numero mesi"
                            GroupTable =1
                            BottomPadding =150
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2445
                            LayoutCachedWidth =1695
                            LayoutCachedHeight =2760
                            RowStart =2
                            RowEnd =2
                            LayoutGroup =1
                            GroupTable =1
                        End
                    End
                End
                Begin EmptyCell
                    Left =3255
                    Top =2445
                    Height =315
                    Name ="CellaVuota36"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638
                    LayoutCachedLeft =3255
                    LayoutCachedTop =2445
                    LayoutCachedWidth =4695
                    LayoutCachedHeight =2760
                    RowStart =2
                    RowEnd =2
                    ColumnStart =2
                    ColumnEnd =2
                    LayoutGroup =1
                    GroupTable =1
                End
                Begin CommandButton
                    Default = NotDefault
                    OverlapFlags =85
                    Left =4755
                    Top =2445
                    Width =3225
                    Height =315
                    TabIndex =5
                    ForeColor =4210752
                    Name ="cmdOk"
                    Caption ="Ok"
                    OnClick ="[Event Procedure]"
                    GroupTable =1
                    BottomPadding =150
                    GridlineColor =10921638

                    LayoutCachedLeft =4755
                    LayoutCachedTop =2445
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =2760
                    RowStart =2
                    RowEnd =2
                    ColumnStart =3
                    ColumnEnd =3
                    LayoutGroup =1
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    GroupTable =1
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =9
                    Overlaps =1
                End
            End
        End
    End
End
CodeBehindForm
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub cmdOk_Click()
    If IsNull(lstScenario.Value) Then
        MsgBox "Selezionare uno scenario"
        Exit Sub
    End If
    If lstPortfolio.ItemsSelected.Count = 0 Then
        MsgBox "Selezionare almeno un portafoglio"
        Exit Sub
    End If
    If Not Nz(cmbMese.Value, 0) > 0 Then
        MsgBox "Selezionare un mese"
        Exit Sub
    End If
    If Nz(txtAnno, 0) < 2018 Then
        MsgBox "Digitare un anno(dal 2018 in poi)"
        Exit Sub
    End If
    If Nz(txtNumMesi, 0) < 1 Or Nz(txtNumMesi, 0) > 12 Then
        MsgBox "Digitare il numero di mesi (da 1 a 12)"
        Exit Sub
    End If
    
    Debug.Print lstScenario.Value, cmbMese.Value, txtAnno, txtNumMesi
    
    Dim portfolios() As String
    ReDim portfolios(lstPortfolio.ItemsSelected.Count - 1)
    Dim i As Integer
    i = 0
    For Each v In lstPortfolio.ItemsSelected
        Debug.Print lstPortfolio.ItemData(v)
        portfolios(i) = lstPortfolio.ItemData(v)
        i = i + 1
    Next v
    
    
    ret = modExtract_populateDataSheet(lstScenario.Value, portfolios, txtAnno, cmbMese.Value, txtNumMesi)
    If ret Then
        Me.SetFocus
        MsgBox "Scenario " & lstScenario.Value & " generato in excel"
    End If
End Sub

Private Sub Corpo_Click()

End Sub
