Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    Modal = NotDefault
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridY =10
    Width =6292
    DatasheetFontHeight =11
    ItemSuffix =8
    Left =5955
    Top =2910
    Right =12240
    Bottom =7260
    DatasheetGridlinesColor =14806254
    RecSrcDt = Begin
        0x9c28e162a2c9e340
    End
    Caption ="Error"
    OnOpen ="[Event Procedure]"
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
    AllowLayoutView =0
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
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
        Begin Section
            Height =4365
            Name ="Detail"
            AlternateBackColor =15921906
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =566
                    Top =226
                    Width =4980
                    Height =555
                    BorderColor =8355711
                    ForeColor =8355711
                    Name ="Label0"
                    Caption ="An error has occurred in the application, you can now\015\012choose from the fol"
                        "lowing options."
                    GridlineColor =10921638
                    LayoutCachedLeft =566
                    LayoutCachedTop =226
                    LayoutCachedWidth =5546
                    LayoutCachedHeight =781
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =566
                    Top =907
                    Width =4986
                    Height =1140
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtErrorMessage"
                    GridlineColor =10921638

                    LayoutCachedLeft =566
                    LayoutCachedTop =907
                    LayoutCachedWidth =5552
                    LayoutCachedHeight =2047
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =737
                    Top =2211
                    Width =4536
                    Height =568
                    TabIndex =1
                    ForeColor =4210752
                    Name ="cmdTryAgain"
                    Caption ="Try again as you understand the error"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =737
                    LayoutCachedTop =2211
                    LayoutCachedWidth =5273
                    LayoutCachedHeight =2779
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    Enabled = NotDefault
                    OverlapFlags =85
                    Left =737
                    Top =2899
                    Width =4536
                    Height =568
                    TabIndex =2
                    ForeColor =4210752
                    Name ="cmdSkip"
                    Caption ="Skip, as you have been advised to\015\012 skip over this error"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =737
                    LayoutCachedTop =2899
                    LayoutCachedWidth =5273
                    LayoutCachedHeight =3467
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =737
                    Top =3587
                    Width =4536
                    Height =568
                    TabIndex =3
                    ForeColor =4210752
                    Name ="cmdExit"
                    Caption ="Exit, as you are either unsure of what to do\015\012or have been advised to skip"
                        " this error"
                    OnClick ="[Event Procedure]"
                    GridlineColor =10921638

                    LayoutCachedLeft =737
                    LayoutCachedTop =3587
                    LayoutCachedWidth =5273
                    LayoutCachedHeight =4155
                    BackColor =14136213
                    BorderColor =14136213
                    HoverColor =15060409
                    PressedColor =9592887
                    HoverForeColor =4210752
                    PressedForeColor =4210752
                    WebImagePaddingLeft =2
                    WebImagePaddingTop =2
                    WebImagePaddingRight =1
                    WebImagePaddingBottom =1
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
Option Explicit

Private Sub cmdExit_Click()
' set a response to the error and close the form
    global_lngErrorAction = 3
    DoCmd.Close
End Sub

Private Sub cmdSkip_Click()
' set a response to the error and close the form
    global_lngErrorAction = 2
    DoCmd.Close
End Sub

Private Sub cmdTryAgain_Click()
' set a response to the error and close the form
    global_lngErrorAction = 1
    DoCmd.Close
End Sub

Private Sub Form_Open(Cancel As Integer)
    Me.txtErrorMessage = Me.OpenArgs
End Sub
