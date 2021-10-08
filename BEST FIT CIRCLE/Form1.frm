VERSION 5.00
Object = "{D940E4E4-6079-11CE-88CB-0020AF6845F6}#1.6#0"; "cwui.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13125
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   13125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   8160
      TabIndex        =   12
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT PROGRAM"
      Height          =   735
      Left            =   8160
      TabIndex        =   11
      Top             =   7080
      Width           =   3615
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   2040
      TabIndex        =   5
      Top             =   7200
      Width           =   3375
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   8160
      TabIndex        =   4
      Top             =   4080
      Width           =   3375
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   8160
      TabIndex        =   3
      Top             =   3000
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   8160
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.CommandButton BestFitCircle 
      Caption         =   "FIND BEST FIT CIRCLE"
      Height          =   735
      Left            =   8160
      TabIndex        =   1
      Top             =   5040
      Width           =   3615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11640
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CWUIControlsLib.CWGraph CWGraph1 
      Height          =   6180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7980
      _Version        =   393218
      _ExtentX        =   14076
      _ExtentY        =   10901
      _StockProps     =   71
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Reset_0         =   0   'False
      CompatibleVers_0=   393218
      Graph_0         =   1
      ClassName_1     =   "CCWGraphFrame"
      opts_1          =   62
      C[0]_1          =   0
      Event_1         =   2
      ClassName_2     =   "CCWGFPlotEvent"
      Owner_2         =   1
      Plots_1         =   3
      ClassName_3     =   "CCWDataPlots"
      Array_3         =   2
      Editor_3        =   4
      ClassName_4     =   "CCWGFPlotArrayEditor"
      Owner_4         =   1
      Array[0]_3      =   5
      ClassName_5     =   "CCWDataPlot"
      opts_5          =   4194367
      Name_5          =   "Plot-1"
      C[0]_5          =   65280
      C[1]_5          =   65535
      C[2]_5          =   16711680
      C[3]_5          =   16776960
      Event_5         =   2
      X_5             =   6
      ClassName_6     =   "CCWAxis"
      opts_6          =   1599
      Name_6          =   "XAxis"
      Orientation_6   =   2944
      format_6        =   7
      ClassName_7     =   "CCWFormat"
      Scale_6         =   8
      ClassName_8     =   "CCWScale"
      opts_8          =   90112
      rMin_8          =   38
      rMax_8          =   508
      dMin_8          =   1
      dMax_8          =   10
      discInterval_8  =   1
      discBase_8      =   1
      Radial_6        =   0
      Enum_6          =   9
      ClassName_9     =   "CCWEnum"
      Editor_9        =   10
      ClassName_10    =   "CCWEnumArrayEditor"
      Owner_10        =   6
      Font_6          =   0
      tickopts_6      =   2711
      major_6         =   1
      minor_6         =   0.5
      Caption_6       =   11
      ClassName_11    =   "CCWDrawObj"
      opts_11         =   62
      C[0]_11         =   -2147483640
      Image_11        =   12
      ClassName_12    =   "CCWTextImage"
      font_12         =   0
      Animator_11     =   0
      Blinker_11      =   0
      Y_5             =   13
      ClassName_13    =   "CCWAxis"
      opts_13         =   1599
      Name_13         =   "YAxis-1"
      Orientation_13  =   2067
      format_13       =   14
      ClassName_14    =   "CCWFormat"
      Scale_13        =   15
      ClassName_15    =   "CCWScale"
      opts_15         =   122880
      rMin_15         =   12
      rMax_15         =   384
      dMax_15         =   10
      discInterval_15 =   1
      Radial_13       =   0
      Enum_13         =   16
      ClassName_16    =   "CCWEnum"
      Editor_16       =   17
      ClassName_17    =   "CCWEnumArrayEditor"
      Owner_17        =   13
      Font_13         =   0
      tickopts_13     =   2711
      major_13        =   1
      minor_13        =   0.5
      Caption_13      =   18
      ClassName_18    =   "CCWDrawObj"
      opts_18         =   62
      C[0]_18         =   -2147483640
      Image_18        =   19
      ClassName_19    =   "CCWTextImage"
      font_19         =   0
      Animator_18     =   0
      Blinker_18      =   0
      PointStyle_5    =   10
      LineStyle_5     =   1
      LineWidth_5     =   1
      BasePlot_5      =   0
      DefaultXInc_5   =   1
      DefaultPlotPerRow_5=   -1  'True
      Array[1]_3      =   20
      ClassName_20    =   "CCWDataPlot"
      opts_20         =   4194367
      Name_20         =   "Plot-2"
      C[0]_20         =   16711680
      C[1]_20         =   16711680
      C[2]_20         =   16711680
      C[3]_20         =   16776960
      Event_20        =   2
      X_20            =   6
      Y_20            =   13
      LineStyle_20    =   1
      LineWidth_20    =   2
      BasePlot_20     =   0
      DefaultXInc_20  =   1
      DefaultPlotPerRow_20=   -1  'True
      Axes_1          =   21
      ClassName_21    =   "CCWAxes"
      Array_21        =   2
      Editor_21       =   22
      ClassName_22    =   "CCWGFAxisArrayEditor"
      Owner_22        =   1
      Array[0]_21     =   6
      Array[1]_21     =   13
      DefaultPlot_1   =   23
      ClassName_23    =   "CCWDataPlot"
      opts_23         =   4194367
      Name_23         =   "[Template]"
      C[0]_23         =   65280
      C[1]_23         =   255
      C[2]_23         =   16711680
      C[3]_23         =   16776960
      Event_23        =   2
      X_23            =   6
      Y_23            =   13
      LineStyle_23    =   1
      LineWidth_23    =   1
      BasePlot_23     =   0
      DefaultXInc_23  =   1
      DefaultPlotPerRow_23=   -1  'True
      Cursors_1       =   24
      ClassName_24    =   "CCWCursors"
      Editor_24       =   25
      ClassName_25    =   "CCWGFCursorArrayEditor"
      Owner_25        =   1
      TrackMode_1     =   2
      GraphFrameStyle_1=   1
      GraphBackground_1=   0
      GraphFrame_1    =   26
      ClassName_26    =   "CCWDrawObj"
      opts_26         =   62
      Image_26        =   27
      ClassName_27    =   "CCWPictImage"
      opts_27         =   1280
      Rows_27         =   1
      Cols_27         =   1
      Pict_27         =   450
      F_27            =   -2147483633
      B_27            =   -2147483633
      ColorReplaceWith_27=   8421504
      ColorReplace_27 =   8421504
      Tolerance_27    =   2
      Animator_26     =   0
      Blinker_26      =   0
      PlotFrame_1     =   28
      ClassName_28    =   "CCWDrawObj"
      opts_28         =   62
      C[1]_28         =   0
      Image_28        =   29
      ClassName_29    =   "CCWPictImage"
      opts_29         =   1280
      Rows_29         =   1
      Cols_29         =   1
      Pict_29         =   1
      F_29            =   -2147483633
      B_29            =   0
      ColorReplaceWith_29=   8421504
      ColorReplace_29 =   8421504
      Tolerance_29    =   2
      Animator_28     =   0
      Blinker_28      =   0
      Caption_1       =   30
      ClassName_30    =   "CCWDrawObj"
      opts_30         =   62
      C[0]_30         =   -2147483640
      Image_30        =   31
      ClassName_31    =   "CCWTextImage"
      font_31         =   0
      Animator_30     =   0
      Blinker_30      =   0
      DefaultXInc_1   =   1
      DefaultPlotPerRow_1=   -1  'True
      Bindings_1      =   32
      ClassName_32    =   "CCWBindingHolderArray"
      Editor_32       =   33
      ClassName_33    =   "CCWBindingHolderArrayEditor"
      Owner_33        =   1
      Annotations_1   =   34
      ClassName_34    =   "CCWAnnotations"
      Editor_34       =   35
      ClassName_35    =   "CCWAnnotationArrayEditor"
      Owner_35        =   1
      AnnotationTemplate_1=   36
      ClassName_36    =   "CCWAnnotation"
      opts_36         =   63
      Name_36         =   "[Template]"
      Plot_36         =   23
      Text_36         =   "[Template]"
      TextXPoint_36   =   6.7
      TextYPoint_36   =   6.7
      TextColor_36    =   16777215
      TextFont_36     =   37
      ClassName_37    =   "CCWFont"
      bFont_37        =   -1  'True
      BeginProperty Font_37 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShapeXPoints_36 =   38
      ClassName_38    =   "CDataBuffer"
      Type_38         =   5
      m_cDims;_38     =   1
      m_cElts_38      =   1
      Element[0]_38   =   3.3
      ShapeYPoints_36 =   39
      ClassName_39    =   "CDataBuffer"
      Type_39         =   5
      m_cDims;_39     =   1
      m_cElts_39      =   1
      Element[0]_39   =   3.3
      ShapeFillColor_36=   16777215
      ShapeLineColor_36=   16777215
      ShapeLineWidth_36=   1
      ShapeLineStyle_36=   1
      ShapePointStyle_36=   10
      ShapeImage_36   =   40
      ClassName_40    =   "CCWDrawObj"
      opts_40         =   62
      Image_40        =   41
      ClassName_41    =   "CCWPictImage"
      opts_41         =   1280
      Rows_41         =   1
      Cols_41         =   1
      Pict_41         =   7
      F_41            =   -2147483633
      B_41            =   -2147483633
      ColorReplaceWith_41=   8421504
      ColorReplace_41 =   8421504
      Tolerance_41    =   2
      Animator_40     =   0
      Blinker_40      =   0
      ArrowVisible_36 =   -1  'True
      ArrowColor_36   =   16777215
      ArrowWidth_36   =   1
      ArrowLineStyle_36=   1
      ArrowHeadStyle_36=   1
   End
   Begin VB.Label Label6 
      Caption         =   "Qty of DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   13
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "Expressed as % of Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   6720
      Width           =   7335
   End
   Begin VB.Label Label4 
      Caption         =   "Root Mean Square Deviation of Points from circle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   6360
      Width           =   7335
   End
   Begin VB.Label Label3 
      Caption         =   "R - Radius"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   3720
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Y axis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   2640
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "X axis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8160
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOPEN 
         Caption         =   "Open"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private Sub BestFitCircle_Click()
    On Error Resume Next
    A = 0: B = 0: C = 0: D = 0: E = 0: F = 0: K = 0: L = 0: M = 0

    
    For I = 1 To N
        A = A + 2 * X(I) ^ 2
        B = B + 2 * X(I) * Y(I)
        C = C - X(I)
        D = D + 2 * Y(I) ^ 2
        E = E - Y(I)
        F = F + 0.5
        K = K + X(I) ^ 3 + X(I) * Y(I) ^ 2
        L = L + X(I) ^ 2 * Y(I) + Y(I) ^ 3
        M = M - 0.5 * (X(I) ^ 2 + Y(I) ^ 2)
    Next I
    
    DET = A * D * F + 2 * B * C * E - A * E * E - F * B * B - D * C * C
    
    CC = (K * (D * F - E * E) - L * (B * F - E * C) + M * (B * E - D * C)) / DET
    
    DC = (-K * (B * F - C * E) + L * (A * F - C * C) - M * (A * E - B * C)) / DET
    
    BC = (K * (B * E - D * C) - L * (A * E - B * C) + M * (A * D - B * B)) / DET
    
    Root = (CC * CC + DC * DC - BC)
    R = Sqr(Root)
    
    Text1.Text = Format(CC, "0.0###")
    Text2.Text = Format(DC, "0.0###")
    Text3.Text = Format(R, "0.0###")
    
    DISTCENT = 0: TSVAR = 0: RMSV = 0
    
    For I = 1 To N
        DISTCENT = Sqr((X(I) - CC) ^ 2 + (Y(I) - DC) ^ 2)
        TSVAR = TSVAR + (DISTCENT - R) ^ 2
    Next I
    RMSV = Sqr(TSVAR / N) / R * 100
    
    Text4.Text = Format(RMSV, "0.0###")
    
    For AA = 1 To 360
        XX(AA) = R * Cos(AA * 3.1415926 / 180)
        YY(AA) = R * Sin(AA * 3.1415926 / 180)
    Next AA

    CWGraph1.Plots(2).PlotXvsY XX, YY
    
    
    For Each Axis In CWGraph1.Axes
        Axis.AutoScaleNow
    Next
End Sub

Private Sub Command1_Click()
    Unload Me
    End
End Sub

Private Sub mnuOPEN_Click()
    loadF
    
    Text5.Text = N
    CWGraph1.ClearData
    
    CWGraph1.BackColor = vbYellow
    CWGraph1.CaptionColor = vbBlack
    CWGraph1.Caption = " BEST FIT CICLE "

    CWGraph1.Plots(1).PlotXvsY X, Y
    
    For Each Axis In CWGraph1.Axes
        Axis.AutoScaleNow
    Next
    
    For I = 1 To N
        CWGraph1.Annotations.Add
        With CWGraph1.Annotations(I)
            .CoordinateType = cwAxesCoordinates
            .Arrow.Visible = False
            .Shape.Type = cwShapePoint
            .Caption.Text = X(I) & " , " & Y(I)
            .Caption.SetCoordinates X(I), Y(I)
        End With
        
    Next I
        
End Sub
