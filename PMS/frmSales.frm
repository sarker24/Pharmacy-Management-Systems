VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSales 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Sales Information [Doctors Clinic Unit - 2]"
   ClientHeight    =   11010
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   13575
   Icon            =   "frmSales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   13575
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7800
      TabIndex        =   57
      Top             =   9480
      Width           =   975
   End
   Begin VB.CommandButton cmdSelectItems 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Item ..."
      Height          =   330
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Browse Items"
      Top             =   9600
      Width           =   1500
   End
   Begin VB.CommandButton cmAutoPaid 
      BackColor       =   &H00C0C000&
      Caption         =   ":: Auto paid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9600
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   6375
      Left            =   240
      TabIndex        =   48
      Top             =   2880
      Width           =   14895
      Begin VB.CommandButton cmdLDelete 
         BackColor       =   &H00C0B4A9&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   14400
         Picture         =   "frmSales.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   480
      End
      Begin VSFlex7LCtl.VSFlexGrid fgSales 
         Height          =   6015
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   14175
         _cx             =   25003
         _cy             =   10610
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   8421376
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSales.frx":0894
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Height          =   2175
      Left            =   240
      TabIndex        =   24
      Top             =   600
      Width           =   14895
      Begin VB.TextBox txtReg 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   4560
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtAdvance 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7800
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   9
         Text            =   " "
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpCurrentDate 
         Height          =   375
         Left            =   10800
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   49938435
         CurrentDate     =   40924
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1695
         Left            =   12600
         TabIndex        =   33
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtDiscountAmt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9240
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPost 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   10680
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Text            =   "frmSales.frx":0A20
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox txtDiscuntPer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7800
         Locked          =   -1  'True
         TabIndex        =   31
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtTDue 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7800
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtCAddress 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Input your Customer Address."
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtTotalAmt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtCMID 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   4560
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   12000
         Top             =   240
      End
      Begin VB.TextBox txtTime 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox txtNAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker SDate 
         Height          =   345
         Left            =   4560
         TabIndex        =   34
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   49676291
         CurrentDate     =   38465
      End
      Begin MSForms.ComboBox CmbCName 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4335
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7646;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbBedCabin 
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   1560
         Width           =   1575
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "2778;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.ComboBox cmbPType 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   1560
         Width           =   1935
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "3413;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label lblAdvance 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Advance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6960
         TabIndex        =   53
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Data Edited by"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   12600
         TabIndex        =   47
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10680
         TabIndex        =   46
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label lblTotalPaid 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Paid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6360
         TabIndex        =   45
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblTotalAmount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6360
         TabIndex        =   44
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblDiscountTk 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount (Tk)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   43
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblDiscountper 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount (%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   42
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label lblPurchaseDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   41
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label lblRegNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Bed No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         TabIndex        =   40
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblDue 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Due"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7800
         TabIndex        =   38
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Address/Contact No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblSalesID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Invoice No "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         TabIndex        =   36
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblNAmount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Net Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.TextBox txtUName 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   9240
      Width           =   2295
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   9600
      Width           =   735
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdEdit 
      Height          =   495
      Left            =   1080
      TabIndex        =   16
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdDelete 
      Height          =   495
      Left            =   1920
      TabIndex        =   19
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdFind 
      Height          =   495
      Left            =   2760
      TabIndex        =   20
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Find"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdCancel 
      Height          =   495
      Left            =   3600
      TabIndex        =   21
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdClose 
      Height          =   495
      Left            =   6240
      TabIndex        =   22
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Quit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPost 
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "P&ost"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdChange 
      Height          =   495
      Left            =   8760
      TabIndex        =   18
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Bill"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdAddItem 
      Height          =   495
      Left            =   11520
      TabIndex        =   3
      Top             =   9240
      Width           =   1095
      _Version        =   196612
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   " Add &Item"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   32896
      WordWrap        =   0   'False
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPreview 
      Height          =   495
      Left            =   5280
      TabIndex        =   23
      Top             =   10200
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Preview"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   10680
      Top             =   10200
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Driver={SQL Server};Server=MAS;Database=KG;Trusted_Connection=yes"
      OLEDBString     =   "Driver={SQL Server};Server=MAS;Database=KG;Trusted_Connection=yes"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "tblCashMaster"
      Caption         =   "Record Search"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Print"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdRefund1 
      Height          =   495
      Left            =   9600
      TabIndex        =   17
      Top             =   10200
      Visible         =   0   'False
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Refund"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdDue 
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Due"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12629161
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin MSAdodcLib.Adodc dcIndoor 
      Height          =   330
      Left            =   10680
      Top             =   10560
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Patient Information"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "MEDICINCE SALES INFORMATION    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   51
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 Option Explicit
 Private rsItemMaster                    As ADODB.Recordset
 Private rsItemDetail                    As ADODB.Recordset
 Private rsCustomerMaster                As New ADODB.Recordset
 Private rs                              As ADODB.Recordset
 Private rsTemp                         As ADODB.Recordset
 Private rsTemp1                         As ADODB.Recordset
 Private rsTemp2                         As ADODB.Recordset
 
 Private bRecordExists                   As Boolean
 Dim str As String
 Dim str1 As String
 Private rs1                              As ADODB.Recordset
 
 
'-----------Get computer name-------------------------
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'-----------End to get computer name-------------------
'---------------------------------------------------------------------------
'------------------------Add For Reporting Perpose--------------------------
'---------------------------------------------------------------------------

Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition


Private objReportSub                        As CRPEAuto.Report 'sub
Private objReportDatabaseSub                As CRPEAuto.Database 'sub
Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition

                          
Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private rsDailyRpt                          As ADODB.Recordset
Private Tracer                              As Integer
Private strGroupName                        As String
Private temp As Double
Private temp1 As Double
Private temp2 As Double
'Private temp3 As Double
'Private temp4 As Double
''--------------------------------------------------------------------------------

Private Sub cmAutoPaid_Click()
txtPaid.text = txtNAmount.text
TxtTDue.text = txtNAmount.text - txtPaid.text
txtAdvance.text = txtPaid.text
End Sub

Private Sub cmAutoPaid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbPType_Change()
If cmbPType.text = "Admitted" Then
txtReg.Enabled = True
CmbCName.Enabled = True
cmbBedCabin.Enabled = True
Call BedCabin
Call RegNo
ElseIf cmbPType.text = "Outdoor" Then
txtReg.Enabled = False
cmbBedCabin.Enabled = False
CmbCName.Clear
End If
End Sub

Private Sub cmbPType_LostFocus()
If cmbPType.text = "Admitted" Then
txtReg.Enabled = True
CmbCName.Enabled = True
CmbCName.SetFocus

    End If
    
End Sub

'----------------------- Indoor Customer Related -------------------------------------------------------

Private Sub CmbCName_DropDown()
CmbCName.Refresh
End Sub

Private Sub CmbCName_LostFocus()
Call PRelated
End Sub

Private Sub PRelated()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT PID,RegNo,PName,PhoneNo FROM PRegtration where RegNo='" & parseQuotes(CmbCName.text) & "'"), cn, adOpenStatic
    
    If rsTemp.RecordCount > 0 Then
        CmbCName = rsTemp!PName
        txtReg = rsTemp!RegNo
        txtCAddress = rsTemp!PhoneNo
 End If
    rsTemp.Close
End Sub

Private Sub BedCabin()

Dim rsTemp As New ADODB.Recordset

     rsTemp.Open ("SELECT DISTINCT Name FROM [Beds And Cabins] ORDER BY Name ASC"), cn, adOpenStatic

    While Not rsTemp.EOF
        cmbBedCabin.AddItem rsTemp("Name")
        rsTemp.MoveNext
    Wend
    rsTemp.Close

End Sub

Private Sub RegNo()

Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT RegNo FROM PRegtration ORDER BY RegNo DESC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        CmbCName.AddItem rsTemp2("RegNo")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

'----------------------------Due Collection -------------------------
Private Sub cmdDue_Click()
If TxtTDue.text = 0 Then Exit Sub
If txtPost.text = "Posted" Then
If TxtTDue.text > 0 Then
    Call Duecollection
    Else
    MsgBox "Please posting the Bill before Due Collection."
    End If
    End If
    
End Sub

Private Sub Duecollection()
On Error Resume Next
Dim Due As String
Dim str As String
'------------------Define Due Date-------------------------
Dim D1, D2 As Date
D1 = SDate
D2 = dtpCurrentDate
Dim Particulars As String
Dim Status As String
Dim PartialAmount As Integer

  If D1 = D2 Then
                Particulars = "Current Due Amount"
                Status = "Advance"
          Else
                Particulars = "Previous Due Amount"
                Status = "Due"
  End If
'---------------End Define Due Date-----------------------

'---------Define Computer Name-----------------------------
 Dim dwLen As Long
 Dim strString As String
 dwLen = MAX_COMPUTERNAME_LENGTH + 1
 strString = String(dwLen, "X")
 GetComputerName strString, dwLen
 strString = Left(strString, dwLen)

'---------End Define Computer Name-------------------------
Due = MsgBox("Do you want to all Clear Dues", vbYesNo)

If Due = 0 Then Exit Sub
If Due = vbYes Then

           cn.Execute "insert into Medicine_Payment (CDate,CTime,ComputerName,Amount,CMID,RegNo,Post,Particulars,UName,Status) " & _
                        " values('" & Format(dtpCurrentDate, "dd-mmm-yyyy") & "',    " & _
                        "   '" & txtTime & "','" & parseQuotes(strString) & "',         " & _
                        "   " & (TxtTDue) & "," & (txtCMID) & ",                     " & _
                        "   '" & parseQuotes(txtReg) & "',                             " & _
                        "   '" & parseQuotes(txtPost) & "',                             " & _
                        "   '" & (Particulars) & "','" & parseQuotes(frmLogin.txtUName) & "', " & _
                        "   '" & (Status) & "')"

ElseIf Due = vbNo Then
      PartialAmount = InputBox("Enter amount of Due payment", "Partial Due Payment")
      If TxtTDue >= PartialAmount Then
             
             cn.Execute "insert into Medicine_Payment (CDate,CTime,ComputerName,Amount,CMID,RegNo,Post,Particulars,UName,Status) " & _
                        " values('" & Format(dtpCurrentDate, "dd-mmm-yyyy") & "',    " & _
                        "   '" & txtTime & "','" & parseQuotes(strString) & "',         " & _
                        "   " & (PartialAmount) & "," & (txtCMID) & ",                     " & _
                        "   '" & parseQuotes(txtReg) & "',                             " & _
                        "   '" & parseQuotes(txtPost) & "',                             " & _
                        "   '" & (Particulars) & "','" & parseQuotes(frmLogin.txtUName) & "', " & _
                        "   '" & (Status) & "')"
'                End If
     Else
           Exit Sub
       End If
End If

    Call DueCheck

    Call Calculation1
    Call Dueupdate

MsgBox "Record added Successfully", vbInformation, "Confirmation"

End Sub

Private Sub Calculation1()
If Val(rs1!Tpaid) = 0 Then
TxtTDue = CDbl(txtNAmount) - Val(rs!Amt)
Else

TxtTDue = CDbl(txtNAmount) - Val(rs!Amt)
End If
txtPaid = (Val(rs!Amt) + Val(rs1!Tpaid))
End Sub

Private Sub DueCheck()
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "Select ISNULL(sum(Amount),0) as Amt from Medicine_Payment where CMID='" & parseQuotes(Me.txtCMID) & "'"
           rs.Open str, cn, adOpenStatic, adLockReadOnly

End Sub

Private Sub TotalPaid()

Set rs1 = New ADODB.Recordset
If rs1.State <> 0 Then rs.Close
           str1 = "Select ISNULL(sum(Tpaid),0) as Tpaid from SalesMaster where CMID='" & parseQuotes(Me.txtCMID) & "'"
           rs1.Open str1, cn, adOpenStatic, adLockReadOnly
End Sub

Private Sub Dueupdate()
 cn.Execute "UPDATE SalesMaster SET SDate = '" & Format(SDate, "dd-mmm-yyyy") & "',CName='" & parseQuotes(CmbCName.text) & "', " & _
               "CAddress='" & parseQuotes(txtCAddress.text) & "',RegNo='" & parseQuotes(txtReg.text) & "', " & _
               "Tamount=" & Val(txtTotalAmt.text) & ", DiscuntPer=" & Val(txtDiscuntPer.text) & ", " & _
               "DiscuntTaka = " & Val(txtDiscountAmt.text) & ",Ttime = '" & txtTime & "',Tpaid=" & Val(txtPaid.text) & ", " & _
               "TDue=" & Val(TxtTDue.text) & ", Advance=" & Val(txtAdvance.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "', " & _
               "Admitted='" & parseQuotes(cmbPType.text) & "'  WHERE CMID = '" & parseQuotes(txtCMID) & "'"

End Sub

Private Sub Datediff()
Dim D1, D2 As Date
D1 = SDate
D2 = dtpCurrentDate
Dim Particulars As String

  If D1 = D2 Then
                Particulars = "Current Due Amount"
          Else
                Particulars = "Previous Due Amount"
  End If
End Sub

Private Sub ComputerName()
Dim dwLen As Long
 Dim strString As String
 dwLen = MAX_COMPUTERNAME_LENGTH + 1
 strString = String(dwLen, "X")
 'Get the computer name
 GetComputerName strString, dwLen
 'get only the actual data
 strString = Left(strString, dwLen)
End Sub

Private Sub cmdFirst_Click()
'Call Calculation
Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MoveFirst
If Adodc2.Recordset.EOF = True Then
          MsgBox "end of file"
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

        txtCMID = Adodc2.Recordset!CMID
        SDate = Adodc2.Recordset!SDate
        CmbCName = Adodc2.Recordset!CName
        txtCAddress = Adodc2.Recordset!CAddress
        txtReg = Adodc2.Recordset!RegNo
        cmbBedCabin = Adodc2.Recordset!BedNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtDiscuntPer = Adodc2.Recordset!DiscuntPer
        txtDiscountAmt.text = Adodc2.Recordset!DiscuntTaka
        txtTime = Adodc2.Recordset!Ttime
        txtPaid.text = Adodc2.Recordset!Tpaid
        txtAdvance.text = Adodc2.Recordset!Advance
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        cmbPType.text = Adodc2.Recordset!Admitted
        txtRemarks.text = Adodc2.Recordset!Remark
        '---------------------------------------------


        ' show calculation------------------------
'         Call Calculation
    txtNAmount = CDbl(txtTotalAmt) - CDbl(txtDiscountAmt)
   TxtTDue = CDbl(txtNAmount) - CDbl(txtPaid)




        fgSales.Rows = 1
  strPaymentDetail = "SELECT  CMID, SName, MID,MCatagory, MName, Qty, SRate, Amount, PRate,Posted, UName, Discount, SDate" & _
                     " FROM SalesDetail " & _
                     "WHERE CMID='" & parseQuotes(Me.txtCMID) & "' order by SerialNo "

'strPaymentDetail = "Exec  Sms_SalesDetails_Find1 '" & parseQuotes(Me.txtCMID) & "'"
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgSales.Rows = rsItemDetail.RecordCount + 1

 For i = 1 To rsItemDetail.RecordCount
            fgSales.TextMatrix(i, 1) = rsItemDetail("CMID")
            fgSales.TextMatrix(i, 2) = rsItemDetail("SName")
            fgSales.TextMatrix(i, 3) = rsItemDetail("MID")
            fgSales.TextMatrix(i, 4) = rsItemDetail("MCatagory")
            fgSales.TextMatrix(i, 5) = rsItemDetail("MName")
            fgSales.TextMatrix(i, 6) = rsItemDetail("Qty")
            fgSales.TextMatrix(i, 7) = rsItemDetail("SRate")
            fgSales.TextMatrix(i, 8) = rsItemDetail("Amount")
            fgSales.TextMatrix(i, 9) = rsItemDetail("PRate")
            fgSales.TextMatrix(i, 10) = rsItemDetail("Discount")
            fgSales.TextMatrix(i, 11) = rsItemDetail("Posted")
            fgSales.TextMatrix(i, 12) = rsItemDetail("UName")
            fgSales.TextMatrix(i, 13) = rsItemDetail("SDate")

        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

  End If
End Sub

Private Sub cmdLast_Click()
'    Call Calculation
Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MoveLast
If Adodc2.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

        txtCMID = Adodc2.Recordset!CMID
        SDate = Adodc2.Recordset!SDate
        CmbCName = Adodc2.Recordset!CName
        txtCAddress = Adodc2.Recordset!CAddress
        txtReg = Adodc2.Recordset!RegNo
        cmbBedCabin = Adodc2.Recordset!BedNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtDiscuntPer = Adodc2.Recordset!DiscuntPer
        txtDiscountAmt.text = Adodc2.Recordset!DiscuntTaka
        txtTime = Adodc2.Recordset!Ttime
        txtPaid.text = Adodc2.Recordset!Tpaid
        txtAdvance.text = Adodc2.Recordset!Advance
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        cmbPType.text = Adodc2.Recordset!Admitted
        txtRemarks.text = Adodc2.Recordset!Remark

' show calculation------------------------
'    Call Calculation
   txtNAmount = CDbl(txtTotalAmt) - CDbl(txtDiscountAmt)
   TxtTDue = CDbl(txtNAmount) - CDbl(txtPaid)


        fgSales.Rows = 1
  strPaymentDetail = "SELECT  CMID, SName, MID,MCatagory, MName, Qty, SRate, Amount, PRate,Posted, UName, Discount, SDate" & _
                     " FROM SalesDetail " & _
                     "WHERE CMID='" & parseQuotes(Me.txtCMID) & "' order by SerialNo "

'strPaymentDetail = "Exec  Sms_SalesDetails_Find1 '" & parseQuotes(Me.txtCMID) & "'"
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgSales.Rows = rsItemDetail.RecordCount + 1

 For i = 1 To rsItemDetail.RecordCount
            fgSales.TextMatrix(i, 1) = rsItemDetail("CMID")
            fgSales.TextMatrix(i, 2) = rsItemDetail("SName")
            fgSales.TextMatrix(i, 3) = rsItemDetail("MID")
            fgSales.TextMatrix(i, 4) = rsItemDetail("MCatagory")
            fgSales.TextMatrix(i, 5) = rsItemDetail("MName")
            fgSales.TextMatrix(i, 6) = rsItemDetail("Qty")
            fgSales.TextMatrix(i, 7) = rsItemDetail("SRate")
            fgSales.TextMatrix(i, 8) = rsItemDetail("Amount")
            fgSales.TextMatrix(i, 9) = rsItemDetail("PRate")
            fgSales.TextMatrix(i, 10) = rsItemDetail("Discount")
            fgSales.TextMatrix(i, 11) = rsItemDetail("Posted")
            fgSales.TextMatrix(i, 12) = rsItemDetail("UName")
            fgSales.TextMatrix(i, 13) = rsItemDetail("SDate")

        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close



End If
End Sub

Private Sub cmdNext_Click()
'Call Calculation
Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MoveNext
If Adodc2.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

        txtCMID = Adodc2.Recordset!CMID
        SDate = Adodc2.Recordset!SDate
        CmbCName = Adodc2.Recordset!CName
        txtCAddress = Adodc2.Recordset!CAddress
        txtReg = Adodc2.Recordset!RegNo
        cmbBedCabin = Adodc2.Recordset!BedNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtDiscuntPer = Adodc2.Recordset!DiscuntPer
        txtDiscountAmt.text = Adodc2.Recordset!DiscuntTaka
        txtTime = Adodc2.Recordset!Ttime
        txtPaid.text = Adodc2.Recordset!Tpaid
        txtAdvance.text = Adodc2.Recordset!Advance
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        cmbPType.text = Adodc2.Recordset!Admitted
        txtRemarks.text = Adodc2.Recordset!Remark


' show calculation------------------------
'    Call Calculation
'    Text1 = CDbl(txtTotalAmt) - (CDbl(TxtAdvance) + Val(rs1!Tpaid))

    txtNAmount = CDbl(txtTotalAmt) - CDbl(txtDiscountAmt)
   TxtTDue = CDbl(txtNAmount) - CDbl(txtPaid)




        fgSales.Rows = 1
 strPaymentDetail = "SELECT  CMID, SName, MID,MCatagory, MName, Qty, SRate, Amount, PRate,Posted, UName, Discount, SDate" & _
                     " FROM SalesDetail " & _
                     "WHERE CMID='" & parseQuotes(Me.txtCMID) & "' order by SerialNo "

'strPaymentDetail = "Exec  Sms_SalesDetails_Find1 '" & parseQuotes(Me.txtCMID) & "'"
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgSales.Rows = rsItemDetail.RecordCount + 1

 For i = 1 To rsItemDetail.RecordCount
            fgSales.TextMatrix(i, 1) = rsItemDetail("CMID")
            fgSales.TextMatrix(i, 2) = rsItemDetail("SName")
            fgSales.TextMatrix(i, 3) = rsItemDetail("MID")
            fgSales.TextMatrix(i, 4) = rsItemDetail("MCatagory")
            fgSales.TextMatrix(i, 5) = rsItemDetail("MName")
            fgSales.TextMatrix(i, 6) = rsItemDetail("Qty")
            fgSales.TextMatrix(i, 7) = rsItemDetail("SRate")
            fgSales.TextMatrix(i, 8) = rsItemDetail("Amount")
            fgSales.TextMatrix(i, 9) = rsItemDetail("PRate")
            fgSales.TextMatrix(i, 10) = rsItemDetail("Discount")
            fgSales.TextMatrix(i, 11) = rsItemDetail("Posted")
            fgSales.TextMatrix(i, 12) = rsItemDetail("UName")
            fgSales.TextMatrix(i, 13) = rsItemDetail("SDate")
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
End If
End Sub

Private Sub cmdPreview_Click()
    Tracer = 0
    If txtPost.text = "Posted" Then
    Call printReport
    Else
    MsgBox "Please Post the Bill "
    End If
End Sub

Private Sub cmdAddItem_Click()
frmSalesSearch.Show vbModal
Call Calculation
'Call allenable
End Sub


Private Sub cmdCancel_Click()

    cmdCancel.Enabled = True
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdFind.Enabled = True
    cmdPost.Caption = "&Post"
    CmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
    cmdPost.Enabled = True
    cmbPType.Enabled = True
    cmbPType.Enabled = True
    cmdChange.Enabled = True
    Call Clear
    Call alldisable
'    If Not rsItemMaster.EOF Then FindRecord

End Sub


Private Sub cmdClose_Click()
    Unload Me
'Call Delete_Duplicates
End Sub


Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If txtUName.text = "Admin" Then
    If idelete = vbYes Then
            cn.Execute "Delete From SalesMaster Where CMID ='" & parseQuotes(txtCMID) & "'"
            cn.Execute "DELETE FROM SalesDetail WHERE CMID='" & parseQuotes(txtCMID) & "'"
            Call Clear
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
     End If
End Sub

Private Sub cmdEdit_Click()
 
'-----------------Admin Check--------
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Password,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
' ----------------Check End------

If rs!Privilegegroup = 0 Then

 If txtPost.text = "Not Posted" Then
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdFind.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdLDelete.Enabled = True
        fgSales.Editable = flexEDKbdMouse
        txtCMID.Enabled = False
        cmdPost.Enabled = False
         Call Calculation
        
    ElseIf cmdEdit.Caption = "&Update" Then
'          Call duplicate
        Call Calculation
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdFind.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                fgSales.Editable = flexEDNone
                Call alldisable
                
                rsItemMaster.Requery
'                Dim s As String
'
'                s = txtCMID
'                rsItemMaster.MoveFirst
'                rsItemMaster.Find "CMID='" & parseQuotes(s) & "'"
'                FindRecord
            End If
        End If
   End If
 End If

Else
    
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        cmdFind.Enabled = False
        CmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdLDelete.Enabled = True
        fgSales.Editable = flexEDKbdMouse
        txtCMID.Enabled = False
        cmdPost.Enabled = False
        cmdChange.Enabled = False
        Call Calculation
        
    ElseIf cmdEdit.Caption = "&Update" Then
          Call Calculation
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdFind.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                CmdDelete.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                cmdChange.Enabled = True
                fgSales.Editable = flexEDNone
                Call alldisable
                
                rsItemMaster.Requery
''                Dim s As String
'                s = txtCMID
'                rsItemMaster.MoveFirst
'                rsItemMaster.Find "CMID='" & parseQuotes(s) & "'"
'                FindRecord
                
            End If
        End If
    End If

End If
End Sub

Private Sub cmdLAdd_Click()
With fgSales
        If .Row = -1 Or .Row = 0 Then
            .AddItem ""
            Exit Sub
        End If
        If .Row > 0 Then
                .AddItem "", .Row + 1
        End If
    End With
    
End Sub

Private Sub CmdPost_Click()
Dim s As String
cmdPost.Caption = "&Posting"
fgSales.Editable = flexEDKbdMouse
'Call postedCheck


If cmdPost.Caption = "&Posting" Then
     If txtPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 fgSales.Enabled = False
                 cmdFind.Enabled = True
'                 chameleonButton1.Enabled = True
                 CmdDelete.Enabled = True
'                 cmdChange.Enabled = True
'                 txtBillSerialNo.Enabled = False
                 Call alldisable
           End If
        End If
      End If
Else
'    cmdtemSelected.Enabled = False
    cmdLDelete.Enabled = False
 End If
cmdPost.Caption = "&Post"
 
End Sub



Private Sub cmdPrevious_Click()
'Call Calculation
Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset

Adodc2.Recordset.MovePrevious
If Adodc2.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdPrevious.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

    txtCMID = Adodc2.Recordset!CMID
        SDate = Adodc2.Recordset!SDate
        CmbCName = Adodc2.Recordset!CName
        txtCAddress = Adodc2.Recordset!CAddress
        txtReg = Adodc2.Recordset!RegNo
        cmbBedCabin = Adodc2.Recordset!BedNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtDiscuntPer = Adodc2.Recordset!DiscuntPer
        txtDiscountAmt.text = Adodc2.Recordset!DiscuntTaka
        txtTime = Adodc2.Recordset!Ttime
        txtPaid.text = Adodc2.Recordset!Tpaid
        txtAdvance.text = Adodc2.Recordset!Advance
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        cmbPType.text = Adodc2.Recordset!Admitted
        txtRemarks.text = Adodc2.Recordset!Remark


        ' show calculation------------------------
'          Call Calculation
    txtNAmount = CDbl(txtTotalAmt) - CDbl(txtDiscountAmt)
   TxtTDue = CDbl(txtNAmount) - CDbl(txtPaid)




        fgSales.Rows = 1
  strPaymentDetail = "SELECT  CMID, SName, MID, MCatagory, MName, Qty, SRate, Amount, PRate,Posted, UName, Discount, SDate" & _
                     " FROM SalesDetail " & _
                     "WHERE CMID='" & parseQuotes(Me.txtCMID) & "' order by SerialNo "

'strPaymentDetail = "Exec  Sms_SalesDetails_Find1 '" & parseQuotes(Me.txtCMID) & "'"
    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgSales.Rows = rsItemDetail.RecordCount + 1

 For i = 1 To rsItemDetail.RecordCount
            fgSales.TextMatrix(i, 1) = rsItemDetail("CMID")
            fgSales.TextMatrix(i, 2) = rsItemDetail("SName")
            fgSales.TextMatrix(i, 3) = rsItemDetail("MID")
            fgSales.TextMatrix(i, 4) = rsItemDetail("MCatagory")
            fgSales.TextMatrix(i, 5) = rsItemDetail("MName")
            fgSales.TextMatrix(i, 6) = rsItemDetail("Qty")
            fgSales.TextMatrix(i, 7) = rsItemDetail("SRate")
            fgSales.TextMatrix(i, 8) = rsItemDetail("Amount")
            fgSales.TextMatrix(i, 9) = rsItemDetail("PRate")
            fgSales.TextMatrix(i, 10) = rsItemDetail("Discount")
            fgSales.TextMatrix(i, 11) = rsItemDetail("Posted")
            fgSales.TextMatrix(i, 12) = rsItemDetail("UName")
            fgSales.TextMatrix(i, 13) = rsItemDetail("SDate")
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
End If
End Sub

Private Sub cmdPrint_Click()

Dim s As String
If cmdPrint.Caption = "&Print" Then
cmdPrint.Caption = "&Printing"
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Enabled = True
                cmdCancel.Enabled = True
                cmdClose.Enabled = True
                fgSales.Enabled = False
                cmdFind.Enabled = True
                cmdPreview.Enabled = True
                cmdChange.Enabled = True
                txtCMID.Enabled = False
                Call alldisable

            End If
        End If
    End If
    
Tracer = 1
Screen.MousePointer = vbHourglass
If txtPost.text = "Posted" Then
Call printReport
'Call CashCopy
'Call GuestCopy
End If
Screen.MousePointer = vbDefault

cmdPrint.Caption = "&Print"

End Sub

Private Sub cmdChange_Click()
frmIndoorBill.Show vbModal
End Sub


Private Sub cmdLDelete_Click()
   
    If fgSales.Rows = 1 Then Exit Sub

     If fgSales.Row >= 1 Then
      If MsgBox("Are you sure to delete the selected record", vbYesNo, "General Setup") = vbYes Then fgSales.RemoveItem fgSales.Row
     Else
      MsgBox "You have to select a row to delete.", vbInformation, "General"
    End If
    
Call Calculation
End Sub

Private Sub cmdNew_Click()

'-----------------Admin Check--------
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Password,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
' ----------------Check End------

If rs!Privilegegroup = "0" Then

If cmdNew.Caption = "&New" Then
        
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdFind.Enabled = False
        CmdDelete.Enabled = False
        cmdFind.Enabled = False
        cmdClose.Enabled = False
        cmdPost.Enabled = False
        cmdChange.Enabled = False
        cmbPType.Enabled = True
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmbPType.Enabled = False
        TextClear Me
        Call Clear
         
        fgSales.Rows = 1
        fgSales.Editable = flexEDKbdMouse
        Call allenable
        txtPost.text = "Not Posted"
        txtUName.text = frmLogin.txtUName.text
        cmbPType.SetFocus

        
    ElseIf cmdNew.Caption = "&Save" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdFind.Enabled = True
                cmdCancel.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
                cmdChange.Enabled = True
                cmbPType.Enabled = True
                
'                Call Calculation
                
                Call alldisable
            End If
        End If
    End If
    
Else

If cmdNew.Caption = "&New" Then
        
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdFind.Enabled = False
        CmdDelete.Enabled = False
        cmdFind.Enabled = False
        cmdClose.Enabled = False
        cmdPost.Enabled = False
        cmdChange.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdLDelete.Enabled = True
        TextClear Me
        Call Clear
         
        fgSales.Rows = 1
        fgSales.Editable = flexEDKbdMouse
        Call allenable
        txtPost.text = "Not Posted"
        txtUName.text = frmLogin.txtUName.text
        cmbPType.SetFocus

        
    ElseIf cmdNew.Caption = "&Save" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                cmdFind.Enabled = True
                cmdCancel.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
                cmdChange.Enabled = True
                
                Call alldisable
            
            End If
        End If
    End If
End If
    
End Sub

Private Sub Clear()
     txtCMID.text = ""
     cmbPType.text = ""
     SDate.Enabled = True
     SDate.Value = Date
     CmbCName.text = ""
     CmbCName.Clear
     cmbBedCabin.text = ""
     cmbBedCabin.Clear
     txtCAddress.text = ""
     txtReg.text = ""
     txtTime.Enabled = True
     txtTotalAmt.text = ""
     txtNAmount.text = ""
     txtDiscuntPer.text = ""
     txtDiscountAmt.text = ""
     txtPaid.text = ""
     txtAdvance.text = ""
    
End Sub

Private Sub allenable()
    txtCMID.Enabled = True
     SDate.Enabled = True
     cmdLDelete.Enabled = True
     fgSales.Enabled = True
     CmbCName.Enabled = True
     cmbBedCabin.Enabled = True
     txtCAddress.Enabled = True
     txtReg.Enabled = True
     txtTime.Enabled = True
     txtTotalAmt.Enabled = True
     txtDiscuntPer.Enabled = True
     txtDiscountAmt.Enabled = True
     txtPaid.Enabled = True
     txtAdvance.Enabled = True
     cmdAddItem.Enabled = True
     cmbPType.Enabled = True
     cmAutoPaid.Enabled = True
    End Sub

Private Sub alldisable()
     txtCMID.Enabled = False
     txtReg.Enabled = False
     cmbBedCabin.Enabled = False
     SDate.Enabled = False
     cmdLDelete.Enabled = False
     fgSales.Enabled = False
     CmbCName.Enabled = False
     txtCAddress.Enabled = False
     txtTime.Enabled = False
     txtTotalAmt.Enabled = False
     txtDiscuntPer.Enabled = False
     txtDiscountAmt.Enabled = False
     txtNAmount.Enabled = False
     txtPaid.Enabled = False
     txtAdvance.Enabled = False
     cmdAddItem.Enabled = False
     cmbPType.Enabled = False
     cmAutoPaid.Enabled = False

End Sub

Private Sub cmdFind_Click()
    frmSalesMasterSearch.Show vbModal
    Call Calculation
    cmdFind.Enabled = True
    cmdCancel.Enabled = True
        
End Sub
    
Private Sub fgSales_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Col
   Case 3, 4, 6, 7

      Cancel = True
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub
 
 Private Sub Form_Load()
         Call Connect
     ModFunction.StartUpPosition Me
     txtUName.text = frmLogin.txtUName.text
       
       Call alldisable
       
   Set rsItemMaster = New ADODB.Recordset
 
  If rsItemMaster.State <> 0 Then rsItemMaster.Close
     rsItemMaster.Open "select TOP 1 * FROM SalesMaster ORDER BY CMID DESC", cn, adOpenStatic, adLockReadOnly
'rsItemMaster.Open "exec Sms_Sales_Master", cn, adOpenStatic, adLockReadOnly
  If rsItemMaster.RecordCount > 0 Then
      rsItemMaster.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
    txtPost.text = "Not Posted"
    
    If Not rsItemMaster.EOF Then FindRecord

'-----------------For Record Search----------
Adodc2.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc2.CommandType = adCmdTable
  Adodc2.RecordSource = "SalesMaster"

  Adodc2.Refresh
  
  cmbPType.AddItem "Outdoor"
  cmbPType.AddItem "Admitted"
  cmbPType.AddItem "Released"
'-------------------End Record Search---------

dtpCurrentDate = Now

Call Calculation
CmbCName.Enabled = False
txtReg.Enabled = False

End Sub

'Private Sub CmbCName_Click()
'
'If cmbPType.text = "Admitted" Then
'
'Set rsItemMaster = New ADODB.Recordset
'
'    If rsItemMaster.State <> 0 Then rsItemMaster.Close
'       rsItemMaster.Open "select PID,RegNo,PName,PhoneNo from PRegtration where RegNo ='" & CmbCName & "' ", cn, adOpenStatic, adLockReadOnly
'
'   If rsItemMaster.RecordCount > 0 Then
'      rsItemMaster.MoveFirst
'    End If
'
'    If Not rsItemMaster.EOF Then FindRecord2
'
'    End If
'End Sub
'
'Private Sub CmbCName_LostFocus()
'
'If cmbPType.text = "Admitted" Then
'
'Set rsItemMaster = New ADODB.Recordset
'
'    If rsItemMaster.State <> 0 Then rsItemMaster.Close
'       rsItemMaster.Open "select PID,RegNo,PName,PhoneNo from PRegtration where PName ='" & CmbCName & "' ", cn, adOpenStatic, adLockReadOnly
'
'   If rsItemMaster.RecordCount > 0 Then
'      rsItemMaster.MoveFirst
'    End If
'
'    If Not rsItemMaster.EOF Then FindRecord2
'    CmbCName.text = StrConv(CmbCName.text, vbProperCase)
'
'    End If
'End Sub
'
'
'Private Sub CmbCName_DropDown()
'CmbCName.Refresh
'End Sub
'
'Private Sub FindRecord2()
'    CmbCName = rsItemMaster!PName
'    txtCAddress = rsItemMaster!PhoneNo
'    txtReg = rsItemMaster!RegNo
''    txtVAT = rsItemMaster!RPoint
'End Sub


Private Sub Calculation()

  Dim j As Integer
       temp = 0
       temp1 = 0
       temp2 = 0
  
  Call DueCheck
  
  Call TotalPaid
    
    For j = 1 To fgSales.Rows - 1
    
    Val (fgSales.TextMatrix(j, 8) = (CDbl(Val(fgSales.TextMatrix(j, 6)) * CDbl(Val(fgSales.TextMatrix(j, 7))))))

        temp = temp + CDbl(Val(fgSales.TextMatrix(j, 6)) * CDbl(Val(fgSales.TextMatrix(j, 7))))
        temp2 = temp2 + CDbl(Val(fgSales.TextMatrix(j, 8)) * CDbl(Val(fgSales.TextMatrix(j, 10)))) / 100
   
   Next

txtTotalAmt = temp
txtDiscountAmt = temp2

temp1 = (temp * CDbl(Val(txtDiscuntPer) / 100))
txtNAmount = CDbl(Val(txtTotalAmt)) - CDbl(Val(temp1 + temp2))

TxtTDue = CDbl(txtNAmount) - CDbl(Val(txtPaid))

End Sub

Private Sub default()
txtDiscountAmt = "0"
txtTotalAmt = "0"
txtPaid = "0"
End Sub


 Private Function rcupdate() As Boolean

    Dim strSQL As String
    Dim iRow As Integer
    Dim j As Integer
    Dim i As Integer
    Dim blnAlarm As Boolean
    Dim strDeliveryDate As String
    Dim str As String
    Set rs = New ADODB.Recordset
    Dim iPost
    Dim strExpDate As String
'-------------------------------Group Permission------------
str = "select SerialNo,Password,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
'           If rs.RecordCount = 0 Then Exit Sub

'-------------------------------------Group permission end-------------------
    If rs!Privilegegroup = "0" Then
     
     If cmdNew.Caption = "&Save" Then
        
     strSQL = "INSERT INTO SalesMaster (SDate, CName, CAddress, RegNo, BedNo, Tamount, DiscuntPer, DiscuntTaka, Ttime, Tpaid, Tdue, Advance, Posted, UName, Admitted, Remark) " & _
              "VALUES ('" & Format(SDate, "dd-mmm-yyyy") & "','" & parseQuotes(CmbCName.text) & "','" & parseQuotes(txtCAddress.text) & "', " & _
              " '" & parseQuotes(txtReg.text) & "','" & parseQuotes(cmbBedCabin.text) & "'," & Val(txtTotalAmt.text) & "," & Val(txtDiscuntPer.text) & "," & Val(txtDiscountAmt.text) & "," & _
              " '" & txtTime & "'," & Val(txtPaid.text) & "," & Val(TxtTDue.text) & "," & Val(txtAdvance.text) & ",'" & txtPost & "'," & _
              "'" & txtUName.text & "','" & parseQuotes(cmbPType.text) & "','" & parseQuotes(txtRemarks.text) & "')"
     
      cn.Execute strSQL
      rcupdate = True
     
'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(CMID),1) as InvNo from SalesMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtCMID = Val(rs!InvNo)
'-------------------------------------------------------------------------
            j = 0
            For j = 1 To fgSales.Rows - 1
            
        cn.Execute "INSERT INTO SalesDetail (CMID,SName,MID,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
                        "UName,SDate) " & _
                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 5)) & "', " & _
                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & "," & _
                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 10) = "", "0", fgSales.TextMatrix(j, 10)) & ", " & _
                        "'" & parseQuotes(txtPost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"

               Next
        
        rcupdate = True
        
'        MsgBox "Record added Successfully", vbInformation, "Confirmation"
        
ElseIf (cmdEdit.Caption = "&Update") Then
            
 cn.Execute "UPDATE SalesMaster SET SDate = '" & Format(SDate, "dd-mmm-yyyy") & "',CName='" & parseQuotes(CmbCName.text) & "', " & _
            "CAddress='" & parseQuotes(txtCAddress.text) & "',RegNo='" & parseQuotes(txtReg.text) & "',BedNo='" & parseQuotes(cmbBedCabin.text) & "', " & _
            "Tamount=" & Val(txtTotalAmt.text) & ", DiscuntPer=" & Val(txtDiscuntPer.text) & ", " & _
            "DiscuntTaka = " & Val(txtDiscountAmt.text) & ",Ttime = '" & txtTime & "',Tpaid=" & Val(txtPaid.text) & ", " & _
            "TDue=" & Val(TxtTDue.text) & ", Advance=" & Val(txtAdvance.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "', " & _
            "Admitted='" & parseQuotes(cmbPType.text) & "'  WHERE CMID = '" & parseQuotes(txtCMID) & "'"

         cn.Execute "DELETE FROM SalesDetail WHERE CMID='" & parseQuotes(txtCMID) & "'"

           j = 0
              For j = 1 To fgSales.Rows - 1
            
     
cn.Execute "INSERT INTO SalesDetail (CMID,SName,MID,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
                        "UName,SDate) " & _
                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 5)) & "', " & _
                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & "," & _
                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 10) = "", "0", fgSales.TextMatrix(j, 10)) & ", " & _
                        "'" & parseQuotes(txtPost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"
               Next
        
        rcupdate = True
        
'        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
        
'----------------------------------------------Printing Start--------------------------
'  ElseIf cmdPrint.Caption = "&Printing" Then


'Dim iprint
'
'iprint = MsgBox("Do you want to Print this bill?", vbYesNo)
'
'If iprint = vbYes Then
'
'         txtpost.text = "Posted"
'
'  cn.Execute "UPDATE SalesMaster SET SDate = '" & Format(SDate, "dd-mmm-yyyy") & "',CName='" & parseQuotes(CmbCName.text) & "', " & _
'               "CAddress='" & parseQuotes(txtCAddress.text) & "',RegNo='" & parseQuotes(txtReg.text) & "', " & _
'               "Tamount=" & Val(txtTotalAmt.text) & ", DiscuntPer=" & Val(txtDiscuntPer.text) & ", " & _
'               "DiscuntTaka = " & Val(txtDiscountAmt.text) & ",Ttime = '" & txtTime & "',Tpaid=" & Val(txtPaid.text) & ", " & _
'               "Advance=" & Val(txtAdvance.text) & ",Posted='" & txtpost & "', UName='" & txtUName.text & "', " & _
'               "Admitted='" & parseQuotes(cmbPType.text) & "'  WHERE CMID = '" & parseQuotes(txtCMID) & "'"
'
'  cn.Execute "DELETE FROM SalesDetail WHERE CMID='" & parseQuotes(txtCMID) & "'"
'
'                j = 0
'            For j = 1 To fgSales.Rows - 1
'
'           cn.Execute "INSERT INTO SalesDetail (CMID,SName,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
'                        "UName,SDate) " & _
'                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
'                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
'                        IIf(fgSales.TextMatrix(j, 5) = "", "0", fgSales.TextMatrix(j, 5)) & "," & _
'                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & ", " & _
'                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
'                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
'                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
'                        "'" & parseQuotes(txtpost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"
'               Next
'
'        rcupdate = True
'
'        End If
'------------------------------Printing End-----------------------------
               
'  --------------------------------Posting Information-----------------------------------------
    
ElseIf cmdPost.Caption = "&Posting" Then
   
iPost = MsgBox("Do you want to Post this bill?", vbYesNo)

If iPost = vbYes Then
     
     txtPost.text = "Posted"
   
  cn.Execute "UPDATE SalesMaster SET SDate = '" & Format(SDate, "dd-mmm-yyyy") & "',CName='" & parseQuotes(CmbCName.text) & "', " & _
               "CAddress='" & parseQuotes(txtCAddress.text) & "',RegNo='" & parseQuotes(txtReg.text) & "',BedNo='" & parseQuotes(cmbBedCabin.text) & "', " & _
               "Tamount=" & Val(txtTotalAmt.text) & ", DiscuntPer=" & Val(txtDiscuntPer.text) & ", " & _
               "DiscuntTaka = " & Val(txtDiscountAmt.text) & ",Ttime = '" & txtTime & "',Tpaid=" & Val(txtPaid.text) & ", " & _
               "Advance=" & Val(txtAdvance.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "', " & _
               "Admitted='" & parseQuotes(cmbPType.text) & "'  WHERE CMID = '" & parseQuotes(txtCMID) & "'"

        cn.Execute "DELETE FROM SalesDetail WHERE CMID='" & parseQuotes(txtCMID) & "'"


        j = 0
            For j = 1 To fgSales.Rows - 1
            
     
cn.Execute "INSERT INTO SalesDetail (CMID,SName,MID,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
                        "UName,SDate) " & _
                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 5)) & "', " & _
                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & "," & _
                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 10) = "", "0", fgSales.TextMatrix(j, 10)) & ", " & _
                        "'" & parseQuotes(txtPost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"
        Next
        
        rcupdate = True
        
'        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
            End If
        End If
'    End If
   
   Else
   '------------------------------Admin group Start------------------

   If cmdNew.Caption = "&Save" Then
        
     strSQL = "INSERT INTO SalesMaster (SDate, CName, CAddress,RegNo,BedNo,Tamount,DiscuntPer,DiscuntTaka,Ttime,Tpaid,Tdue,Advance,Posted," & _
              " UName,Admitted,Remark) " & _
              "VALUES ('" & Format(SDate, "dd-mmm-yyyy") & "','" & parseQuotes(CmbCName.text) & "','" & parseQuotes(txtCAddress.text) & "', " & _
              " '" & parseQuotes(txtReg.text) & "','" & parseQuotes(cmbBedCabin.text) & "'," & Val(txtTotalAmt.text) & "," & Val(txtDiscuntPer.text) & "," & Val(txtDiscountAmt.text) & "," & _
              " '" & txtTime & "'," & Val(txtPaid.text) & "," & Val(TxtTDue.text) & "," & Val(txtAdvance.text) & ",'" & txtPost & "'," & _
              "'" & txtUName.text & "','" & parseQuotes(cmbPType.text) & "','" & parseQuotes(txtRemarks.text) & "')"
     
      cn.Execute strSQL
      rcupdate = True
     
'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(CMID),1) as InvNo from SalesMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtCMID = Val(rs!InvNo)
'-------------------------------------------------------------------------
            j = 0
            For j = 1 To fgSales.Rows - 1
            
            cn.Execute "INSERT INTO SalesDetail (CMID,SName,MID,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
                        "UName,SDate) " & _
                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 5)) & "', " & _
                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & "," & _
                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 10) = "", "0", fgSales.TextMatrix(j, 10)) & ", " & _
                        "'" & parseQuotes(txtPost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"
               Next
        
        rcupdate = True
        
'        MsgBox "Record added Successfully", vbInformation, "Confirmation"
        
ElseIf (cmdEdit.Caption = "&Update") Then
            
 cn.Execute "UPDATE SalesMaster SET SDate = '" & Format(SDate, "dd-mmm-yyyy") & "',CName='" & parseQuotes(CmbCName.text) & "', " & _
               "CAddress='" & parseQuotes(txtCAddress.text) & "',RegNo='" & parseQuotes(txtReg.text) & "',BedNo='" & parseQuotes(cmbBedCabin.text) & "', " & _
               "Tamount=" & Val(txtTotalAmt.text) & ", DiscuntPer=" & Val(txtDiscuntPer.text) & ", " & _
               "DiscuntTaka = " & Val(txtDiscountAmt.text) & ",Ttime = '" & txtTime & "',Tpaid=" & Val(txtPaid.text) & ", " & _
               "TDue=" & Val(TxtTDue.text) & ", Advance=" & Val(txtAdvance.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "', " & _
               "Admitted='" & parseQuotes(cmbPType.text) & "'  WHERE CMID = '" & parseQuotes(txtCMID) & "'"

         cn.Execute "DELETE FROM SalesDetail WHERE CMID='" & parseQuotes(txtCMID) & "'"

           j = 0
              For j = 1 To fgSales.Rows - 1
            
     
cn.Execute "INSERT INTO SalesDetail (CMID,SName,MID,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
                        "UName,SDate) " & _
                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 5)) & "', " & _
                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & "," & _
                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 10) = "", "0", fgSales.TextMatrix(j, 10)) & ", " & _
                        "'" & parseQuotes(txtPost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"
               Next
        
        rcupdate = True
        
'        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
   
ElseIf cmdPost.Caption = "&Posting" Then
  
'  Dim iPost
     
iPost = MsgBox("Do you want to Post this bill?", vbYesNo)

If iPost = vbYes Then
     
     txtPost.text = "Posted"
   
  cn.Execute "UPDATE SalesMaster SET SDate = '" & Format(SDate, "dd-mmm-yyyy") & "',CName='" & parseQuotes(CmbCName.text) & "', " & _
               "CAddress='" & parseQuotes(txtCAddress.text) & "',RegNo='" & parseQuotes(txtReg.text) & "',BedNo='" & parseQuotes(cmbBedCabin.text) & "', " & _
               "Tamount=" & Val(txtTotalAmt.text) & ", DiscuntPer=" & Val(txtDiscuntPer.text) & ", " & _
               "DiscuntTaka = " & Val(txtDiscountAmt.text) & ",Ttime = '" & txtTime & "',Tpaid=" & Val(txtPaid.text) & ", " & _
               "Advance=" & Val(txtAdvance.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "', " & _
               "Admitted='" & parseQuotes(cmbPType.text) & "'  WHERE CMID = '" & parseQuotes(txtCMID) & "'"

        cn.Execute "DELETE FROM SalesDetail WHERE CMID='" & parseQuotes(txtCMID) & "'"


        j = 0
            For j = 1 To fgSales.Rows - 1
            
     
cn.Execute "INSERT INTO SalesDetail (CMID,SName,MID,MCatagory,MName,Qty,SRate,Amount,PRate,Discount,Posted, " & _
                        "UName,SDate) " & _
                        "Values ('" & parseQuotes(txtCMID) & "','" & parseQuotes(fgSales.TextMatrix(j, 2)) & "','" & parseQuotes(fgSales.TextMatrix(j, 3)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 4)) & "', " & _
                        "'" & parseQuotes(fgSales.TextMatrix(j, 5)) & "', " & _
                        IIf(fgSales.TextMatrix(j, 6) = "", "0", fgSales.TextMatrix(j, 6)) & "," & _
                        IIf(fgSales.TextMatrix(j, 7) = "", "0", fgSales.TextMatrix(j, 7)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 8) = "", "0", fgSales.TextMatrix(j, 8)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 9) = "", "0", fgSales.TextMatrix(j, 9)) & ", " & _
                        IIf(fgSales.TextMatrix(j, 10) = "", "0", fgSales.TextMatrix(j, 10)) & ", " & _
                        "'" & parseQuotes(txtPost.text) & "','" & parseQuotes(txtUName.text) & "','" & Format(SDate, "dd-mmm-yyyy") & "')"
        Next
        
        rcupdate = True
        
'        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
            End If
        
        End If
        End If
        
    Exit Function
    
   
   If Err.Number = -2147217874 Then
    MsgBox "You can't Insert same Medicine Name."
   End If
'            MsgBox cn.Errors(0).NativeError & " : " & cn.Errors(0).Description
'    End Select
End Function

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    If Trim(cmbPType) = "" Then
        MsgBox "Entry Valid Patient Type.", vbInformation
        cmbPType.SetFocus
        IsValidRecord = False
        Exit Function
    
'    If txtCAddress = "" Then
'        MsgBox "Input Patient Phone No.", vbInformation
'        txtCAddress.SetFocus
'        IsValidRecord = False
'        Exit Function
           
     ElseIf Trim(fgSales.Rows) < 2 Then
        MsgBox "Please input Medicine Name", vbInformation
        cmdAddItem.SetFocus
        IsValidRecord = False
        Exit Function
 
        
  
     txtDiscuntPer.Enabled = False
     txtDiscountAmt.Enabled = False
     txtPaid.Enabled = False
     txtAdvance.Enabled = False
     cmdAddItem.Enabled = True
'-----------------------------------------------------------------------
'    Else
End If
'End If
        
  Exit Function
'     End If
'     End If
    End Function
    
Private Sub FindRecord()

    Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset
    
    txtCMID = rsItemMaster!CMID
    SDate = rsItemMaster!SDate
    CmbCName = rsItemMaster!CName
    txtCAddress = rsItemMaster!CAddress
    txtReg = rsItemMaster!RegNo
    cmbBedCabin = rsItemMaster!BedNo
    txtTotalAmt = rsItemMaster!TAmount
    txtDiscuntPer = rsItemMaster!DiscuntPer
    txtDiscountAmt = rsItemMaster!DiscuntTaka
    txtTime = rsItemMaster!Ttime
    txtPaid = rsItemMaster!Tpaid
    txtAdvance = rsItemMaster!Advance
    txtPost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName
    cmbPType = rsItemMaster!Admitted
    txtRemarks = rsItemMaster!Remark


    ' show calculation------------------------
    Call Calculation

    fgSales.Rows = 1
    strPaymentDetail = "SELECT  CMID, SName, MID, MCatagory, MName, Qty, SRate, Amount,PRate,Discount, Posted, UName, SDate " & _
                       "FROM SalesDetail " & _
                       "WHERE CMID='" & parseQuotes(Me.txtCMID) & "' order by SerialNo "

'strPaymentDetail = "Exec  Sms_SalesDetails_Find1 '" & parseQuotes(Me.txtCMID) & "'"

    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgSales.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgSales.TextMatrix(i, 1) = rsItemDetail("CMID")
            fgSales.TextMatrix(i, 2) = rsItemDetail("SName")
            fgSales.TextMatrix(i, 3) = rsItemDetail("MID")
            fgSales.TextMatrix(i, 4) = rsItemDetail("MCatagory")
            fgSales.TextMatrix(i, 5) = rsItemDetail("MName")
            fgSales.TextMatrix(i, 6) = rsItemDetail("Qty")
            fgSales.TextMatrix(i, 7) = rsItemDetail("SRate")
            fgSales.TextMatrix(i, 8) = rsItemDetail("Amount")
            fgSales.TextMatrix(i, 9) = rsItemDetail("PRate")
            fgSales.TextMatrix(i, 10) = rsItemDetail("Discount")
            fgSales.TextMatrix(i, 11) = rsItemDetail("Posted")
            fgSales.TextMatrix(i, 12) = rsItemDetail("UName")
            fgSales.TextMatrix(i, 13) = rsItemDetail("SDate")
           
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close
End Sub


Public Sub printReport()

On Error GoTo ErrH
    Dim strPath    As String
    Dim strSQL     As String
    Dim temp       As Double
    If rsItemMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation, "Confarmation"
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\SalesPreview.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        


    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close


                      
        strSQL = " SELECT SalesMaster.CMID, SalesMaster.SDate, SalesMaster.CName, SalesMaster.CAddress, " & _
                 " SalesMaster.RegNo,SalesMaster.BedNo, SalesMaster.Tamount,SalesMaster.DiscuntPer, SalesMaster.DiscuntTaka, SalesMaster.Ttime, " & _
                 " SalesMaster.Tpaid, SalesMaster.Advance, SalesMaster.Posted, SalesMaster.UName, SalesMaster.Admitted, SalesMaster.Remark, " & _
                 " SalesDetail.MCatagory,SalesDetail.MName, SalesDetail.Qty, SalesDetail.SRate, SalesDetail.Amount,SalesDetail.Discount " & _
                 " FROM SalesMaster INNER JOIN " & _
                 " SalesDetail ON SalesMaster.CMID = SalesDetail.CMID where SalesMaster.CMID= '" & Me.txtCMID & "'"


                 rsDailyRpt.Open strSQL, cn, adOpenStatic
'        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"


        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        
        If Tracer = 0 Then
         objReport.Preview "Sales Privew Report", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut (False)
        End If
       
    
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub

ErrH:

    Select Case Err.Number
      Case 20545
            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Item catagory Report"
    End Select
End Sub


Private Sub duplicate()
   Dim j As Integer
        
         For j = 1 To fgSales.Rows - 2
        
        If Val(fgSales.TextMatrix(j, 4)) = Val(fgSales.TextMatrix(j + 1, 4)) Then
        MsgBox "Duplicate Item Code Number.", vbInformation
         fgSales.TextMatrix(j, 4) = ""
         End If

         Next

End Sub

Public Sub PopulateForm(StrID As String)
    rsItemMaster.Close
    Text1.text = frmSalesMasterSearch.txtSearch.text
'    rsItemMaster.Open "select * from SalesMaster", cn, adOpenStatic, adLockReadOnly
     rsItemMaster.Open "select CMID,SDate,CName,CAddress,RegNo,BedNo,Tamount,DiscuntPer,DiscuntTaka,Ttime,Tpaid,Advance,Posted,UName,Admitted,Remark from SalesMaster where CMID= '" & Me.Text1 & "'", cn, adOpenStatic, adLockReadOnly

    rsItemMaster.MoveFirst
    rsItemMaster.Find "CMID=" & parseQuotes(StrID)
    If rsItemMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


Private Sub fgSales_AfterEdit(ByVal Row As Long, ByVal Col As Long)

    If Col = 7 Then
          Dim j As Integer

        For j = 1 To fgSales.Rows - 1
        fgSales.TextMatrix(j, 7) = fgSales.TextMatrix(Row, 5) * fgSales.TextMatrix(Row, 6)
    
       Next
 End If
Call Calculation
End Sub


Private Sub Timer1_Timer()
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub


Private Sub txtCAddress_LostFocus()
'If Len(txtCAddress) < 11 Or Len(txtCAddress) > 11 Then
''If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
'MsgBox "Input valid Phone no."
'txtCAddress.SetFocus
'End If
End Sub

Private Sub txtPaid_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
'Call Advance
End Sub

Private Sub Advance()
txtAdvance = txtPaid
End Sub

Private Sub txtCAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
End Sub


Private Sub txtDiscountAmt_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
Call Calculation
End Sub

