VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSalesDueinfo 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Sales Dues Collection"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9600
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   9600
      Width           =   735
   End
   Begin VB.TextBox txtUserName 
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
      TabIndex        =   35
      Top             =   9240
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Height          =   2175
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   14895
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
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
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
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   1335
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   12000
         Top             =   240
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   4560
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
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
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtReg 
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
         Left            =   1920
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtCustomerAddress 
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
         TabIndex        =   13
         ToolTipText     =   "Input your Customer Address."
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtCustomerName 
         Height          =   345
         Left            =   4800
         TabIndex        =   12
         ToolTipText     =   "Input your buyer Name"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtBillNo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   7440
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Top             =   1680
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox TxtDue 
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
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtDiscountPer 
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
         TabIndex        =   9
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Height          =   345
         Left            =   6360
         TabIndex        =   8
         Text            =   " "
         Top             =   1080
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
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "frmSalesDueinfo.frx":0000
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CheckBox chkOutdoor 
         BackColor       =   &H00808000&
         Caption         =   "Outdoor"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Value           =   2  'Grayed
         Width           =   1095
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
         TabIndex        =   5
         Text            =   " "
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtRemarks 
         Height          =   1695
         Left            =   12600
         TabIndex        =   4
         Top             =   360
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPSalesDate 
         Height          =   345
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   66977795
         CurrentDate     =   38465
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo CmbCName 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Input Customer Name"
         Top             =   360
         Width           =   4335
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   7646
         _ExtentY        =   661
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   34
         Top             =   840
         Width           =   1215
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
         TabIndex        =   33
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label lblBillno 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Bill No"
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
         Left            =   6840
         TabIndex        =   32
         Top             =   1680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblAddress 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Address"
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
         TabIndex        =   31
         Top             =   840
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
         TabIndex        =   30
         Top             =   840
         Width           =   1215
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
         TabIndex        =   29
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblRegNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Reg No "
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
         Left            =   1200
         TabIndex        =   28
         Top             =   1560
         Width           =   855
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
         TabIndex        =   27
         Top             =   120
         Width           =   1695
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   120
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
         TabIndex        =   24
         Top             =   120
         Width           =   1215
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
         TabIndex        =   23
         Top             =   840
         Width           =   1215
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
         TabIndex        =   22
         Top             =   120
         Width           =   1575
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
         TabIndex        =   21
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   6375
      Left            =   240
      TabIndex        =   0
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
         Picture         =   "frmSalesDueinfo.frx":0004
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   120
         Width           =   480
      End
      Begin VSFlex7LCtl.VSFlexGrid fgSales 
         Height          =   6015
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   12615
         _cx             =   22251
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
         GridColor       =   -2147483633
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSalesDueinfo.frx":058E
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
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   495
      Left            =   240
      TabIndex        =   40
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
      TabIndex        =   41
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
      TabIndex        =   42
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
      TabIndex        =   43
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
      TabIndex        =   44
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
      TabIndex        =   45
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
      TabIndex        =   46
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
   Begin SSDataWidgets_A.SSDBCommand cmdUndoPost 
      Height          =   495
      Left            =   8760
      TabIndex        =   47
      Top             =   10200
      Width           =   855
      _Version        =   196612
      _ExtentX        =   1508
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Modify"
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
      TabIndex        =   48
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
      TabIndex        =   49
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
      TabIndex        =   50
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
      TabIndex        =   51
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
      TabIndex        =   52
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
      Caption         =   "MEDICINCE SALES DUES COLLECTION INFORMATION    "
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
      TabIndex        =   53
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmSalesDueinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
