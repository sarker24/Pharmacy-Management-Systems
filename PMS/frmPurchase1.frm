VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmtest 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Purchase Informations"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPurchase.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   12720
      Picture         =   "frmPurchase.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2760
      Width           =   420
   End
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00C0C0C0&
      Height          =   495
      Left            =   6000
      TabIndex        =   38
      Top             =   8160
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Left            =   10440
      Top             =   8280
   End
   Begin SSDataWidgets_A.SSDBCommand cmdAddItem 
      Height          =   495
      Left            =   4560
      TabIndex        =   32
      Top             =   8160
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   615
      Left            =   630
      TabIndex        =   0
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
   Begin SSDataWidgets_A.SSDBCommand chameleonButton1 
      Height          =   615
      Index           =   0
      Left            =   6390
      TabIndex        =   24
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Pre&view"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Purchase Details Information"
      Height          =   5655
      Left            =   120
      TabIndex        =   14
      Top             =   2520
      Width           =   13095
      Begin VSFlex7LCtl.VSFlexGrid fgPurchase 
         Height          =   5175
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   12375
         _cx             =   21828
         _cy             =   9128
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         FormatString    =   $"frmPurchase.frx":0E54
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Purchase Master Information"
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   13095
      Begin VB.TextBox txtpost 
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
         Height          =   495
         Left            =   10920
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   36
         Text            =   "frmPurchase.frx":0FC3
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtDiscountAmt 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9240
         TabIndex        =   8
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   10920
         TabIndex        =   34
         Text            =   " "
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtTotalAmt 
         BackColor       =   &H00C0E0FF&
         Height          =   420
         Left            =   4320
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPaid 
         BackColor       =   &H00C0C000&
         Height          =   420
         Left            =   5880
         TabIndex        =   6
         Text            =   " "
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtVatAmount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9240
         TabIndex        =   3
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtVatPercentage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7440
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
      Begin SSDataWidgets_B.SSDBCombo cmbSupplierName 
         Height          =   420
         Left            =   2520
         TabIndex        =   1
         Top             =   600
         Width           =   4815
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Columns(0).Width=   3200
         _ExtentX        =   8493
         _ExtentY        =   741
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtDiscountPer 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7440
         TabIndex        =   7
         Text            =   " "
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtSupplierBill 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2520
         TabIndex        =   5
         Text            =   " "
         Top             =   1320
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPPurchaseDate 
         Height          =   420
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   741
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   20840451
         CurrentDate     =   38465
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   420
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   37
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Time"
         Height          =   255
         Left            =   10800
         TabIndex        =   35
         Top             =   360
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
         Left            =   5880
         TabIndex        =   22
         Top             =   1080
         Width           =   1455
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
         Left            =   4320
         TabIndex        =   21
         Top             =   1080
         Width           =   1455
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
         TabIndex        =   20
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label lblVatTk 
         BackColor       =   &H00C0B4A9&
         Caption         =   "VAT (Tk)"
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
         Left            =   9240
         TabIndex        =   19
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label lblVatPer 
         BackColor       =   &H00C0B4A9&
         Caption         =   "VAT (%)"
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
         Left            =   7440
         TabIndex        =   18
         Top             =   360
         Width           =   1575
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
         Left            =   7440
         TabIndex        =   17
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label lblSupplierBill 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Bill "
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
         Left            =   2520
         TabIndex        =   16
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblPurchaseDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Purchase Date "
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
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblSupplierName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplair Name "
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
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblChallanNo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Challan No "
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
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSAdodcLib.Adodc dcSupplierName 
      Height          =   480
      Left            =   10920
      Top             =   8880
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
      _ExtentY        =   847
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
      Caption         =   "dcCatagory"
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
   Begin SSDataWidgets_A.SSDBCommand cmdEdit 
      Height          =   615
      Left            =   1590
      TabIndex        =   25
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
      Height          =   615
      Left            =   2550
      TabIndex        =   26
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
   Begin SSDataWidgets_A.SSDBCommand cmdOpen 
      Height          =   615
      Left            =   3510
      TabIndex        =   27
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
      Height          =   615
      Left            =   4470
      TabIndex        =   9
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   615
      Index           =   1
      Left            =   5430
      TabIndex        =   28
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   615
      Left            =   7350
      TabIndex        =   29
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
      Height          =   615
      Left            =   8310
      TabIndex        =   30
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
      Height          =   615
      Left            =   9270
      TabIndex        =   31
      Top             =   8760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
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
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'Option Explicit
' Private rsItemMaster                    As ADODB.Recordset
' Private rsItemDetail                    As ADODB.Recordset
' Private rs                              As ADODB.Recordset
' Private bRecordExists                   As Boolean
' Dim str As String
''---------------------------------------------------------------------------
''---------------------------------------------------------------------------
''----Add For Reporting Perpose----------------------------------------------
'Private objReportApp                        As CRPEAuto.Application
'Private objReport                           As CRPEAuto.Report
'Private objReportDatabase                   As CRPEAuto.Database
'Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
'Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
'
'
'Private objReportSub                        As CRPEAuto.Report 'sub
'Private objReportDatabaseSub                As CRPEAuto.Database 'sub
'Private objReportDatabaseTablesSub          As CRPEAuto.DatabaseTables 'sub
'Private objReportDatabaseTableSub           As CRPEAuto.DatabaseTable 'sub
'Private objReportFormulaFieldDefinationsSub    As CRPEAuto.FormulaFieldDefinitions
'Private objReportFFSub                         As CRPEAuto.FormulaFieldDefinition
'
'
'Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
'Private rsDailyRpt                          As ADODB.Recordset
'Private Tracer                              As Integer
'Private strGroupName                        As String
'Private temp As Double
'Private temp1 As Double
''--------------------------------------------------------------------------------
'
'
''Private Sub chameleonButton1_Click()
''    Call printReport
''End Sub
'
''Private Sub Check1_Click()
''Dim iRows As Integer
''Dim f As Integer
''      f = fgPurchase.Rows - 1
''  If Check1.Value = 1 Then
''       Dim i As Integer
''    For i = 1 To f
''      fgPurchase.Cell(flexcpChecked, i, 10) = flexChecked
''
''      Next i
''Else
''     For i = 1 To f
''
''        fgPurchase.Cell(flexcpChecked, i, 10) = flexUnchecked
''
''    Next i
''    End If
''End Sub
'
''Private Sub chkAutoposting_Click()
''
''      Dim f As Integer
''      f = fgPurchase.Rows - 1
''      If chkAutoposting.Value = 1 Then
''      Dim i As Integer
''For i = 1 To f
''    fgPurchase.Cell(flexcpChecked, i, 12) = flexChecked
''Next i
''Else
''For i = 1 To f
''    fgPurchase.Cell(flexcpChecked, i, 12) = flexUnchecked
''
''    Next i
''End If
''
''End Sub
'
'Private Sub cmdAddItem_Click()
''frmItemInputReceiving.Show vbModal
'End Sub
'
'
'Private Sub postedCheck()
'      Dim f As Integer
'      Dim i As Integer
'      f = fgPurchase.Rows - 1
''      If chkAutoposting.Value = 1 Then
'
'For i = 1 To f
'    fgPurchase.Cell(flexcpChecked, i, 12) = flexChecked
'Next i
'
'
'End Sub
'Private Sub cmdCancel_Click()
'
'Set rs = New ADODB.Recordset
'If rs.State <> 0 Then rs.Close
'str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.TxtUName.text & "'"
'         If rs.State <> 0 Then rs.Close
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
'
' If rs!Privilegegroup = 0 Then
'    cmdCancel.Enabled = False
'    cmdNew.Enabled = True
'    cmdEdit.Caption = "&Edit"
'    cmdNew.Caption = "&New"
''    cmdClose.Enabled = True
'    cmdEdit.Enabled = True
'    cmdOpen.Enabled = True
'    cmdPost.Caption = "&Post"
''    cmdDelete.Enabled = True
''    chameleonButton1.Enabled = True
'    cmdPost.Enabled = True
'    Call alldisable
''    If Not rsItemMaster.EOF Then FindRecord
'Else
'cmdCancel.Enabled = False
'    cmdNew.Enabled = True
'    cmdEdit.Caption = "&Edit"
'    cmdNew.Caption = "&New"
''    cmdClose.Enabled = True
'    cmdEdit.Enabled = True
'    cmdOpen.Enabled = True
'    cmdPost.Caption = "&Post"
'    cmdDelete.Enabled = True
''    chameleonButton1.Enabled = True
'    cmdPost.Enabled = True
'    cmdUndoPost.Enabled = True
'    Call alldisable
''    If Not rsItemMaster.EOF Then FindRecord
'End If
'
'End Sub
'
'
'Private Sub cmdClose_Click()
'    Unload Me
''Call Delete_Duplicates
'End Sub
'
'
'Private Sub cmdDelete_Click()
'On Error GoTo ErrHandler
'     Dim idelete As Integer
'    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
'    If idelete = vbYes Then
'            cn.Execute "Delete From SMSStockMaster Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
'            cn.Execute "DELETE FROM SMSStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'            Call Clear
'    End If
'ErrHandler:
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
'     End Select
'End Sub
'
'Private Sub cmdEdit_Click()
'
''-----------------Admin Check--------
'Set rs = New ADODB.Recordset
'If rs.State <> 0 Then rs.Close
'str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.TxtUName.text & "'"
'         If rs.State <> 0 Then rs.Close
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
'' ----------------Check End------
'
'If rs!Privilegegroup = 0 Then
'
' If txtpost.text = "Not Posted" Then
'    If cmdEdit.Caption = "&Edit" Then
'        cmdNew.Enabled = False
'        Call allenable
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
''        cmdClose.Enabled = False
'        cmdOpen.Enabled = False
''        cmdDelete.Enabled = False
''        chameleonButton1.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
'        fgPurchase.Editable = flexEDKbdMouse
'        txtSerialNo.Enabled = False
'        cmdPost.Enabled = False
'
'    ElseIf cmdEdit.Caption = "&Update" Then
''          Call duplicate
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
''                cmdClose.Enabled = True
'                cmdOpen.Enabled = True
'                chameleonButton1.Enabled = True
''                cmdDelete.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                fgPurchase.Editable = flexEDNone
'                Call alldisable
'
'                rsItemMaster.Requery
'                Dim s As String
'                s = txtSerialNo
'                rsItemMaster.MoveFirst
'                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
'                FindRecord
'            End If
'        End If
'   End If
' End If
'
'Else
'' If txtpost.text = "Not Posted" Then
'    If cmdEdit.Caption = "&Edit" Then
'        cmdNew.Enabled = False
'        Call allenable
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdOpen.Enabled = False
'        cmdDelete.Enabled = False
'        chameleonButton1.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
'        fgPurchase.Editable = flexEDKbdMouse
'        txtSerialNo.Enabled = False
'        cmdPost.Enabled = False
'        cmdUndoPost.Enabled = False
'
'    ElseIf cmdEdit.Caption = "&Update" Then
''          Call duplicate
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdOpen.Enabled = True
'                chameleonButton1.Enabled = True
'                cmdDelete.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                cmdUndoPost.Enabled = True
'                fgPurchase.Editable = flexEDNone
'                Call alldisable
'
'                rsItemMaster.Requery
''                Dim s As String
'                s = txtSerialNo
'                rsItemMaster.MoveFirst
'                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
'                FindRecord
'
'            End If
'        End If
'    End If
''  End If
'
'End If
'
'End Sub
'
'Private Sub cmdLAdd_Click()
'With fgPurchase
'        If .Row = -1 Or .Row = 0 Then
'            .AddItem ""
'            Exit Sub
'        End If
'        If .Row > 0 Then
'                .AddItem "", .Row + 1
'        End If
'    End With
'
'End Sub
'
'Private Sub CmdPost_Click()
'Dim s As String
'cmdPost.Caption = "&Posting"
'fgPurchase.Editable = flexEDKbdMouse
'Call postedCheck
'
'
'If cmdPost.Caption = "&Posting" Then
'     If txtpost.text = "Not Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 cmdNew.Caption = "&New"
'                 cmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 fgPurchase.Enabled = False
'                 cmdOpen.Enabled = True
'                 chameleonButton1.Enabled = True
'                 cmdDelete.Enabled = True
''                 cmdChange.Enabled = True
''                 txtBillSerialNo.Enabled = False
'                 Call alldisable
'           End If
'        End If
'      End If
'Else
''    cmdtemSelected.Enabled = False
'    cmdLDelete.Enabled = False
' End If
'cmdPost.Caption = "&Post"
'
'End Sub
'
'Private Sub cmdUndoPost_Click()
'Dim s As String
'cmdUndoPost.Caption = "&Undo Posting"
'fgPurchase.Editable = flexEDKbdMouse
'Call postedCheck
'
'
'If cmdUndoPost.Caption = "&Undo Posting" Then
'     If txtpost.text = "Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 cmdNew.Caption = "&New"
'                 cmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 fgPurchase.Enabled = False
'                 cmdOpen.Enabled = True
'                 chameleonButton1.Enabled = True
'                 cmdDelete.Enabled = True
''                 cmdChange.Enabled = True
''                 txtBillSerialNo.Enabled = False
'                 Call alldisable
'           End If
'        End If
'      End If
'Else
''    cmdtemSelected.Enabled = False
'    cmdLDelete.Enabled = False
' End If
'cmdUndoPost.Caption = "&Undo Post"
'End Sub
'
'Private Sub fgPurchase_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'   Dim pt As POINTAPI
'
'    ' get popup window position
'    pt.X = fgPurchase.ColPos(Col) \ Screen.TwipsPerPixelX
'    pt.Y = (fgPurchase.RowPos(Row) + fgPurchase.RowHeight(Row)) \ Screen.TwipsPerPixelY
'    ClientToScreen fgPurchase.hwnd, pt
'
'    ' show date popup
'    If fgPurchase.ColDataType(Col) = flexDTDate Then
''      If Col = 9 Then
'        With frmDate
'            .lblRow = Row
'            .lblCol = Col
'            Set rsServerDate = New ADODB.Recordset
'            rsServerDate.Open "select getdate()", cn, adOpenStatic, adLockReadOnly
'            rsServerDate.Requery
'            .Tag = IIf(fgPurchase.Cell(flexcpText, Row, Col) = "", rsServerDate(0), fgPurchase.Cell(flexcpText, Row, Col))
'            strCallingForm = LCase("frmStock")
'            .Move pt.X * Screen.TwipsPerPixelX, pt.Y * Screen.TwipsPerPixelY
'            .Show vbModal
'        End With
'        Exit Sub
''       End If
'    End If
'End Sub
'
'Private Sub cmdLDelete_Click()
''               With fgPurchase
''        If .Row = 0 Or .Row = -1 Then Exit Sub
''
''        If .Rows > 1 Then .RemoveItem .Row
''    End With
'
'
'    If fgPurchase.Rows = 1 Then Exit Sub
'
'     If fgPurchase.Row >= 1 Then
'      If MsgBox("Are you sure to delete the selected record", vbYesNo, "General Setup") = vbYes Then fgPurchase.RemoveItem fgPurchase.Row
'     Else
'      MsgBox "You have to select a row to delete.", vbInformation, "General"
'    End If
'
'
'End Sub
'
'Private Sub cmdNew_Click()
'
''-----------------Admin Check--------
'Set rs = New ADODB.Recordset
'If rs.State <> 0 Then rs.Close
'str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.TxtUName.text & "'"
'         If rs.State <> 0 Then rs.Close
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
'' ----------------Check End------
'
'   '   Dim rs As String
'If rs!Privilegegroup = "USER" Then
'
''    Set rs = New ADODB.Recordset
'If cmdNew.Caption = "&New" Then
'
'        cmdNew.Caption = "&Save"
'        cmdEdit.Enabled = False
'        cmdCancel.Enabled = True
'        cmdOpen.Enabled = False
''        cmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'        cmdClose.Enabled = False
'        cmdPost.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
'        chameleonButton1.Enabled = False
'        TextClear Me
'        Call Clear
'
'        fgPurchase.Rows = 1
'        fgPurchase.Editable = flexEDKbdMouse
'        Call allenable
'        txtpost.text = "Not Posted"
'        txtUserName.text = frmLogin.TxtUName.text
''        cmbItemCatagory.SetFocus
'
'
'    ElseIf cmdNew.Caption = "&Save" Then
''        Dim rs As String
''        Call duplicate1
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdNew.Caption = "&New"
'                cmdEdit.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
''                cmdDelete.Enabled = True
'                cmdOpen.Enabled = True
'                cmdCancel.Enabled = True
'                chameleonButton1.Enabled = True
'                cmdPost.Enabled = True
'
'                Call alldisable
'            End If
'        End If
'    End If
'
'Else
'
'' Set rs = New ADODB.Recordset
'If cmdNew.Caption = "&New" Then
'
'        cmdNew.Caption = "&Save"
'        cmdEdit.Enabled = False
'        cmdCancel.Enabled = True
'        cmdOpen.Enabled = False
'        cmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'        cmdClose.Enabled = False
'        cmdPost.Enabled = False
'        cmdUndoPost.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
'        chameleonButton1.Enabled = False
'        TextClear Me
'        Call Clear
'
'        fgPurchase.Rows = 1
'        fgPurchase.Editable = flexEDKbdMouse
'        Call allenable
'        txtpost.text = "Not Posted"
'        txtUserName.text = frmLogin.TxtUName.text
''        cmbItemCatagory.SetFocus
'
'
'    ElseIf cmdNew.Caption = "&Save" Then
''        Dim rs As String
''        Call duplicate1
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdNew.Caption = "&New"
'                cmdEdit.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdDelete.Enabled = True
'                cmdOpen.Enabled = True
'                cmdCancel.Enabled = True
'                chameleonButton1.Enabled = True
'                cmdPost.Enabled = True
'                cmdUndoPost.Enabled = True
'
'                Call alldisable
'            End If
'        End If
'    End If
'End If
'
'End Sub
'
'Private Sub Clear()
'    txtSerialNo.text = ""
'    DTPPurchaseDate.Enabled = True
'    cmbSupplierName.text = ""
'    txtSupplierBill.text = ""
'    txtVatPercentage.text = ""
'     txtVatAmount.text = ""
'    txtTotalAmt.text = ""
'    txtPaid.text = ""
'    txtDiscountPer.text = ""
'    txtDiscountAmt.text = ""
'
'End Sub
'
'Private Sub allenable()
'     txtSerialNo.Enabled = True
'     cmdLDelete.Enabled = True
'     fgPurchase.Enabled = True
'     DTPPurchaseDate.Enabled = True
'     DTPPurchaseDate.Value = Date
'     cmbSupplierName.Enabled = True
'     txtVatPercentage.Enabled = True
'     txtVatAmount.Enabled = True
'     txtTime.Enabled = True
'     txtSupplierBill.Enabled = True
'     txtTotalAmt.Enabled = True
'     txtPaid.Enabled = True
'     txtDiscountPer.Enabled = True
'     txtDiscountAmt.Enabled = True
'     cmdAddItem.Enabled = True
'    End Sub
'
'Private Sub alldisable()
'     txtSerialNo.Enabled = False
'     cmdLDelete.Enabled = False
'     fgPurchase.Enabled = False
'     DTPPurchaseDate.Enabled = False
'     DTPPurchaseDate.Value = Date
'     cmbSupplierName.Enabled = False
'     txtVatPercentage.Enabled = False
'     txtVatAmount.Enabled = False
'     txtTime.Enabled = False
'     txtSupplierBill.Enabled = False
'     txtTotalAmt.Enabled = False
'     txtPaid.Enabled = False
'     txtDiscountPer.Enabled = False
'     txtDiscountAmt.Enabled = False
'     cmdAddItem.Enabled = False
'
'
'End Sub
'
'Private Sub cmdOpen_Click()
'    frmStockSearch.Show vbModal
'    cmdOpen.Enabled = True
'    cmdCancel.Enabled = True
'
'End Sub
'
'Private Sub Command1_Click()
'frmCatagory.Show vbModal
'End Sub
'
'
' Private Sub Form_Load()
'         Call Connect
'     ModFunction.StartUpPosition Me
'     txtUserName.text = frmLogin.TxtUName.text
'
'       Call alldisable
'       Call SupplierName
'
'   Set rsItemMaster = New ADODB.Recordset
'
'  If rsItemMaster.State <> 0 Then rsItemMaster.Close
'     rsItemMaster.Open "select * FROM PurchaseMaster", cn, adOpenStatic, adLockReadOnly
'
'  If rsItemMaster.RecordCount > 0 Then
'      rsItemMaster.MoveFirst
'        bRecordExists = True
'    Else
'        bRecordExists = False
'    End If
'    txtpost.text = "Not Posted"
'
''     fgPurchase.ColDataType(10) = flexDTBoolean
''    If Not rsItemMaster.EOF Then FindRecord
''    txtSerialNo.Enabled = False
''    ReceivingDate.Value = Null
'End Sub
'
'
'Private Sub SupplierName()
'    dcSupplierName.CursorLocation = adUseClient
'    dcSupplierName.ConnectionString = cn.ConnectionString
'    dcSupplierName.LockType = adLockReadOnly
'    dcSupplierName.RecordSource = "SELECT SupplierName as SupplierName , Address as Address FROM Suppliers ORDER BY Address"
'    cmbSupplierName.DataMode = ssDataModeBound
'    Set cmbSupplierName.DataSource = dcSupplierName
'    cmbSupplierName.DataSourceList = dcSupplierName
'    cmbSupplierName.DataFieldList = "SupplierName"
'    cmbSupplierName.BackColorOdd = &HFFFF00
'    cmbSupplierName.BackColorEven = &HFFC0C0
'    cmbSupplierName.ForeColorEven = &H80000008
'End Sub
'
'
'
' Private Function rcupdate() As Boolean
'
'On Error GoTo ErrHandler
'    Dim strSQL As String
'    Dim iRow As Integer
'    Dim j As Integer
'    Dim i As Integer
'    Dim blnAlarm As Boolean
'    Dim strDeliveryDate As String
'    Dim str As String
'    Set rs = New ADODB.Recordset
'    Dim iPost
'    Dim strExpDate As String
''-------------------------------Group Permission------------
'str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.TxtUName.text & "'"
'         If rs.State <> 0 Then rs.Close
'            rs.Open str, cn, adOpenStatic, adLockReadOnly
''           If rs.RecordCount = 0 Then Exit Sub
'
'
''--------------------------------------------------------------Group permission end-------------------
'    If rs!Privilegegroup = 0 Then
'
'     cn.BeginTrans
'     If cmdNew.Caption = "&Save" Then
'
'    'General Information for Payment Master
'     strSQL = "INSERT INTO SMSStockMaster (ReceivingDate,SupplierName, SupplierBill, BudgetHead,ConPpsted,UserName " & _
'                ") " & _
'                "VALUES ('" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                " '" & cmbSupplierName.Columns(0).text & "','" & txtSupplierBill & "','" & cmbBudgetHead.Columns(0).text & "','" & txtpost & "','" & txtUserName.text & "')"
'     cn.Execute strSQL
'      rcupdate = True
''     cn.CommitTrans
'
''     -------------For primary key and foreign key relation------------
'         If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(SerialNo),1) as InvNo from SMSStockMaster"
'           rs.Open str, cn, adOpenStatic, adLockReadOnly
'           txtSerialNo = Val(rs!InvNo)
'
''------------------------
'
''     '" & parseQuotes(txtSerialNo) & "',
'    'payment Detail Information Enter This table
''    strDeliveryDate = "'" & parseQuotes(fgPurchase.TextMatrix(j, 11)) & "'"
'            j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'
'             cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'
'        rcupdate = True
'
'
'
'        cn.CommitTrans
'        MsgBox "Record added Successfully", vbInformation, "Confirmation"
'
'    ' Update Information
'
'
'
'ElseIf (cmdEdit.Caption = "&Update") Then
'
''            If txtpost.text = "Not Posted" Then
'
'                   cn.Execute "UPDATE SMSStockMaster SET  ReceivingDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              "SupplierName='" & cmbSupplierName.Columns(0).text & "',SupplierBill='" & txtSupplierBill & "', " & _
'                              "BudgetHead='" & cmbBudgetHead.Columns(0).text & "',ConPpsted='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SMSStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'        j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'
'             cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
'
''        End If
'
'
''  --------------------------------Posting Information-----------------------------------------
'
'
'
'   ElseIf cmdPost.Caption = "&Posting" Then
'
'
''''     Dim iPost
'     txtpost.text = "Posted"
'
'
'
'iPost = MsgBox("Do you want to Post this bill?", vbYesNo)
'
'If iPost = vbYes Then
'
'     txtpost.text = "Posted"
'     cn.Execute "UPDATE SMSStockMaster SET  ReceivingDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              "SupplierName='" & cmbSupplierName.Columns(0).text & "',SupplierBill='" & txtSupplierBill & "', " & _
'                              "BudgetHead='" & cmbBudgetHead.Columns(0).text & "',ConPpsted='" & txtpost & "',UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SMSStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'       j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'
'             cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
'
'        End If
'
'
'
'    End If
'
'Else
'
''-------------Admin group------------------
'
'  cn.BeginTrans
'     If cmdNew.Caption = "&Save" Then
'
'    'General Information for Payment Master
'     strSQL = "INSERT INTO SMSStockMaster (ReceivingDate,SupplierName, SupplierBill, BudgetHead,ConPpsted,UserName " & _
'                ") " & _
'                "VALUES ('" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                " '" & cmbSupplierName.Columns(0).text & "','" & txtSupplierBill & "','" & cmbBudgetHead.Columns(0).text & "','" & txtpost & "','" & txtUserName.text & "')"
'     cn.Execute strSQL
'      rcupdate = True
''     cn.CommitTrans
'
''     -------------For primary key and foreign key relation------------
'         If rs.State <> 0 Then rs.Close
'           str = "Select ISNULL(max(SerialNo),1) as InvNo from SMSStockMaster"
'           rs.Open str, cn, adOpenStatic, adLockReadOnly
'           txtSerialNo = Val(rs!InvNo)
'
''------------------------
'
''     '" & parseQuotes(txtSerialNo) & "',
'    'payment Detail Information Enter This table
''    strDeliveryDate = "'" & parseQuotes(fgPurchase.TextMatrix(j, 11)) & "'"
'            j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'             If fgPurchase.TextMatrix(i, 11) = "" Then
'            strExpDate = "null"
'            End If
'
''          cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
''                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
''                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
''                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
''                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
''                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
''                           strExpDate & ", " & _
''                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
''               Next
'
'                cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'        rcupdate = True
'
'
'
'        cn.CommitTrans
'        MsgBox "Record added Successfully", vbInformation, "Confirmation"
'
'    ' Update Information
'
'
'
'ElseIf (cmdEdit.Caption = "&Update") Then
'
''            If txtpost.text = "Not Posted" Then
'
'                   cn.Execute "UPDATE SMSStockMaster SET  ReceivingDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              "SupplierName='" & cmbSupplierName.Columns(0).text & "',SupplierBill='" & txtSupplierBill & "', " & _
'                              "BudgetHead='" & cmbBudgetHead.Columns(0).text & "',ConPpsted='" & txtpost & "', UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SMSStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'        j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'
'              cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
'
''        End If
'
'
''  --------------------------------Posting Information-----------------------------------------
'
'
'
'   ElseIf cmdPost.Caption = "&Posting" Then
'
'
''''     Dim iPost
'     txtpost.text = "Posted"
'
'
'
'iPost = MsgBox("Do you want to Post this bill?", vbYesNo)
'
'If iPost = vbYes Then
'
'     txtpost.text = "Posted"
'     cn.Execute "UPDATE SMSStockMaster SET  ReceivingDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              "SupplierName='" & cmbSupplierName.Columns(0).text & "',SupplierBill='" & txtSupplierBill & "', " & _
'                              "BudgetHead='" & cmbBudgetHead.Columns(0).text & "',ConPpsted='" & txtpost & "',UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SMSStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'       j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'
'       cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
'
'        End If
'
'
'
'
''  -----------Undo Posting-------
'
'ElseIf cmdUndoPost.Caption = "&Undo Posting" Then
'
'
''''     Dim iPost
'     txtpost.text = "Not Posted"
'
'
'
'iPost = MsgBox("Do you want to Undo post this bill?", vbYesNo)
'
'If iPost = vbYes Then
'
'     txtpost.text = "Not Posted"
'     cn.Execute "UPDATE SMSStockMaster SET  ReceivingDate = '" & Format(ReceivingDate, "dd-mmm-yyyy") & "', " & _
'                              "SupplierName='" & cmbSupplierName.Columns(0).text & "',SupplierBill='" & txtSupplierBill & "', " & _
'                              "BudgetHead='" & cmbBudgetHead.Columns(0).text & "',ConPpsted='" & txtpost & "',UserName='" & txtUserName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"
'
'
'        cn.Execute "DELETE FROM SMSStockDetails WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
'
'
'       j = 0
'            For j = 1 To fgPurchase.Rows - 1
'
'            If fgPurchase.Cell(flexcpChecked, j, 12) = flexChecked Then
'               blnAlarm = True
'            Else
'                blnAlarm = False
'            End If
'
'
'      cn.Execute "INSERT INTO SMSStockDetails (SerialNo,ReceivingDate,SupplierName,ProductCatagory,SubCode,ItemName,Quentity,Rate, " & _
'                           "Amount,Rol,ExpDate,Posted,Warrenty,Remarks, ConPpsted,Unit) " & _
'                           "Values ('" & parseQuotes(txtSerialNo) & "','" & Format(ReceivingDate, "dd-mmm-yyyy") & "','" & cmbSupplierName.Columns(0).text & "','" & parseQuotes(fgPurchase.TextMatrix(j, 4)) & "', " & _
'                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
'                           IIf(fgPurchase.TextMatrix(j, 7) = "", "0", fgPurchase.TextMatrix(j, 7)) & "," & IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & "," & IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
'                           IIf(fgPurchase.TextMatrix(j, 11) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 11) & "' ") & ", " & _
'                           IIf(blnAlarm, 1, 0) & ",'" & parseQuotes(fgPurchase.TextMatrix(j, 13)) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 14)) & "','" & parseQuotes(txtpost.text) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 16)) & "')"
'               Next
'
'        rcupdate = True
'        cn.CommitTrans
'        MsgBox "Record Undo Posted Successfully", vbInformation, "Confirmation"
'
'        End If
'
'
''-------------Undo Posting End----
'
'
'
'
'
'
'    End If
'
'
'
'
''-------------Admin group end--------------
'
'
'End If
'
''    cn.CommitTrans
'
'    Exit Function
'
'ErrHandler:
'
'    cn.RollbackTrans
'    Select Case Err.Number
'        Case -2147217900
'            MsgBox "Please select Numeric number in ROL field", vbInformation, "Confirmation"
'
'   End Select
'
'
''   If Err.Number = -2147217874 Then
''    MsgBox "You can't Insert same item from same style multiple times in one BTB LC."
'''   End If
''            MsgBox cn.Errors(0).NativeError & " : " & cn.Errors(0).Description
''    End Select
'End Function
'
'Private Function IsValidRecord() As Boolean
'    IsValidRecord = True
'
'    If Trim(ReceivingDate) = "" Then
'        MsgBox "Your are missing Receiving Information.", vbInformation
'        ReceivingDate.SetFocus
'        IsValidRecord = False
'        Exit Function
'
'
'  ElseIf Trim(cmbSupplierName) = "" Then
'        MsgBox "Your are missing Supplier Name Information.", vbInformation
''        cmbSupplierName.SetFocus
'        IsValidRecord = False
'        Exit Function
'
'
' ElseIf Trim(txtSupplierBill) = "" Then
'        MsgBox "Your are missing Supplier Bill No.", vbInformation
''        cmbSupplierName.SetFocus
'        IsValidRecord = False
'        Exit Function
'
'
''    ---------------------------------------------------
'
''ElseIf cmdNew.Caption = "&Save" Or cmdEdit.Caption = "&Update" Then
''         Dim k As Integer
''         If rsItemDetail.RecordCount > 0 Then
''         If rsItemDetail.State <> 0 Then rsItemDetail.Close
''            rsItemDetail.Open "select * from tblItemDetail where ItemCode='" & fgItem.TextMatrix(Row, 4) & "'", cn, adOpenStatic, adLockReadOnly
''
''             If Not rsItemDetail.EOF Then
''        MsgBox "This Record exists Duplicate ItemCode No.", vbInformation, Me.Caption & " - " & App.Title
''          fgItem.TextMatrix(k, 4).SetFocus
''          IsValidRecord = False
''         Exit Function
''            End If
''         End If
''         End If
''    Exit Function
''-----------------------------------------------------------------------
'
''-----------------------------------------------------------------------
'    Else
'
''        Dim j As Integer
''
''         For j = 1 To fgItem.Rows - 2
''
''        If Not IsNumeric(fgItem.TextMatrix(j, 6)) Then
''        MsgBox "Select Numeric value in ROL field.", vbInformation
'''         fgItem.TextMatrix(j, 4) = ""
'''         fgItem.RemoveItem fgItem.Row
''        IsValidRecord = False
''
''        End If
''
''       Next
'
'       Exit Function
'     End If
'    End Function
'
'Private Sub FindRecord()
'
'    Dim i As Integer
'    Dim strPaymentDetail As String
'    Set rsItemDetail = New ADODB.Recordset
'    txtSerialNo = rsItemMaster!SerialNo
'    cmbSupplierName = rsItemMaster!SupplierName
'    ReceivingDate = rsItemMaster!ReceivingDate
'    txtSupplierBill = rsItemMaster!SupplierBill
'    cmbBudgetHead = rsItemMaster!BudgetHead
''    chkAutoposting = rsItemMaster!Autoposting
'    txtpost = rsItemMaster!ConPpsted
'    txtUserName = rsItemMaster!UserName
'
'
'    fgPurchase.Rows = 1
'    strPaymentDetail = "SELECT  SerialNo, ReceivingDate, SupplierName,ProductCatagory ,SubCode,ItemName,Quentity, " & _
'                "Rate,Amount,Rol,ExpDate,Posted,Warrenty,Remarks,ConPpsted,Unit FROM SMSStockDetails " & _
'                "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
'    rsItemDetail.CursorLocation = adUseClient
'    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly
'
'
'    If rsItemDetail.RecordCount <> 0 Then
'
'        fgPurchase.Rows = rsItemDetail.RecordCount + 1
''                i = 0
'        For i = 1 To rsItemDetail.RecordCount
'            fgPurchase.TextMatrix(i, 1) = rsItemDetail("SerialNo")
'            fgPurchase.TextMatrix(i, 2) = rsItemDetail("ReceivingDate")
'            fgPurchase.TextMatrix(i, 3) = rsItemDetail("SupplierName")
'            fgPurchase.TextMatrix(i, 4) = rsItemDetail("ProductCatagory")
'            fgPurchase.TextMatrix(i, 5) = rsItemDetail("SubCode")
'            fgPurchase.TextMatrix(i, 6) = rsItemDetail("ItemName")
'            fgPurchase.TextMatrix(i, 7) = rsItemDetail("Quentity")
'            fgPurchase.TextMatrix(i, 8) = rsItemDetail("Rate")
'            fgPurchase.TextMatrix(i, 9) = rsItemDetail("Amount")
'            fgPurchase.TextMatrix(i, 10) = rsItemDetail("Rol")
'            fgPurchase.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("ExpDate")), "", Format(rsItemDetail("ExpDate"), "dd-mmm-yyyy"))
'            fgPurchase.TextMatrix(i, 12) = rsItemDetail("Posted")
'            fgPurchase.TextMatrix(i, 13) = rsItemDetail("Warrenty")
'            fgPurchase.TextMatrix(i, 14) = rsItemDetail("Remarks")
'            fgPurchase.TextMatrix(i, 15) = rsItemDetail("ConPpsted")
'            fgPurchase.TextMatrix(i, 16) = rsItemDetail("Unit")
'        rsItemDetail.MoveNext
'        Next
'      End If
'        rsItemDetail.Close
'End Sub
'
'
'Public Sub printReport()
'
'On Error GoTo ErrH
'    Dim strPath    As String
'    Dim strSQL     As String
'    Dim temp       As Double
'    If rsItemMaster.RecordCount = 0 Then
'        MsgBox "Data not available", vbInformation, "Confarmation"
'        Exit Sub
'    End If
'
'
'        strPath = App.Path + "\reports\ReceivingReceipt.rpt"
'        Set objReportApp = CreateObject("Crystal.CRPE.Application")
'        Set objReport = objReportApp.OpenReport(strPath)
'        Set objReportDatabase = objReport.Database
'        Set objReportDatabaseTables = objReportDatabase.Tables
'        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'        Set ObjPrinterSetting = objReport.PrintWindowOptions
'        Set objReportFormulaFieldDefinations = objReport.FormulaFields
'
'
'
'    Set rsDailyRpt = New ADODB.Recordset
'If rsDailyRpt.State <> 0 Then rsDailyRpt.Close
'
'
'
'
'            strSQL = "SELECT SMSStockMaster.SerialNo, SMSStockMaster.ReceivingDate,SMSStockMaster.SupplierName,SMSStockMaster.SupplierBill, " & _
'                      "SMSStockMaster.BudgetHead, SMSStockDetails.ProductCatagory, SMSStockDetails.SubCode, SMSStockDetails.ItemName, " & _
'                      "SMSStockDetails.Quentity,SMSStockDetails.Rate, SMSStockDetails.Amount, SMSStockDetails.ExpDate, " & _
'                      "SMSStockDetails.Posted , SMSStockDetails.Warrenty, SMSStockDetails.Remarks,SMSStockMaster.UserName " & _
'                      "FROM SMSStockMaster INNER JOIN " & _
'                      "SMSStockDetails ON SMSStockMaster.SerialNo = SMSStockDetails.SerialNo and SMSStockMaster.SerialNo ='" & Me.txtSerialNo & "' ORDER BY SMSStockDetails.ProductCatagory "
'
'                      rsDailyRpt.Open strSQL, cn, adOpenStatic
''        Set objReportFF = objReportFormulaFieldDefinations.Item(1)
''            objReportFF.text = "'" + parseQuotes(txtWords.text) + " '"
'
'
'        objReportDatabaseTable.SetPrivateData 3, rsDailyRpt
'
'        ObjPrinterSetting.HasPrintSetupButton = True
'        ObjPrinterSetting.HasRefreshButton = True
'        ObjPrinterSetting.HasSearchButton = True
'        ObjPrinterSetting.HasZoomControl = True
'
'        objReport.DiscardSavedData
'        objReport.Preview "Menu Item List Report", , , , , 16777216 Or 524288 Or 65536
'
'
'        Set objReport = Nothing
'        Set objReportDatabase = Nothing
'        Set objReportDatabaseTables = Nothing
'        Set objReportDatabaseTable = Nothing
'    Exit Sub
'
'ErrH:
'
'    Select Case Err.Number
'        Case -2147217913
'            MsgBox "You need to select record first", vbInformation, "Item Catagory List Report"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Item catagory Report"
'    End Select
'End Sub
'
'
'Private Sub duplicate()
'   Dim j As Integer
'
'         For j = 1 To fgPurchase.Rows - 2
'
'        If Val(fgPurchase.TextMatrix(j, 4)) = Val(fgPurchase.TextMatrix(j + 1, 4)) Then
'        MsgBox "Duplicate Item Code Number.", vbInformation
'         fgPurchase.TextMatrix(j, 4) = ""
'         End If
'
'         Next
'
'End Sub
'
'Public Sub PopulateForm(StrID As String)
'    rsItemMaster.Close
'    rsItemMaster.Open "select * from SMSStockMaster", cn, adOpenStatic, adLockReadOnly
'    rsItemMaster.MoveFirst
'    rsItemMaster.Find "SerialNo=" & parseQuotes(StrID)
'    If rsItemMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord
'
'End Sub
'
'
'Private Sub fgPurchase_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'
''         Dim k As Integer
''
''         Set rsItemDetail = New ADODB.Recordset
''
''                 If rsItemDetail.State <> 0 Then rsItemDetail.Close
''         If col = 5 Then
''            rsItemDetail.Open "select * from SMSSubCatagoryDetail where SubCatagoryCode='" & fgPurchase.TextMatrix(row, 4) & "'", cn, adOpenStatic, adLockReadOnly
''
''             If Not rsItemDetail.EOF Then
''        MsgBox "This Record exists Duplicate Item Code No.", vbInformation, Me.Caption & " - " & App.Title
''            End If
''         End If
''
''
''' .-------------------------------
''
'    If Col = 9 Then
'          Dim j As Integer
'
'        For j = 1 To fgPurchase.Rows - 1
'
'        fgPurchase.TextMatrix(j, 9) = fgPurchase.TextMatrix(Row, 7) * fgPurchase.TextMatrix(Row, 8)
'
'
'
'       Next
' End If
''
''
''' ---------------------------------------
''
''
''If fgPurchase.Rows > 2 Then
''        For j = 1 To fgPurchase.Rows - 1
''            If fgPurchase.TextMatrix(j, 5) = fgPurchase.TextMatrix(row, 5) And j <> fgPurchase.row Then
''                MsgBox "This charge already selected.", vbInformation
''                fgPurchase.TextMatrix(row, 5) = ""
''            End If
''        Next
''    End If
''
'''               --------------------------
''
''
''
''End Sub
''
'''Private Sub fgPurchase_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'''Dim j As Integer
'''
'''If fgPurchase.Rows > 2 Then
'''        For j = 1 To fgPurchase.Rows - 1
'''            If fgPurchase.TextMatrix(j, 5) = fgPurchase.TextMatrix(Row, 5) And j <> fgPurchase.Row Then
'''                MsgBox "This charge already selected.", vbInformation
'''                fgPurchase.TextMatrix(Row, 5) = ""
'''            End If
'''        Next
'''    End If
'''End Sub
''
''
''
''Private Sub check()
''
''Dim i As Integer
''Dim j As Integer
''Dim sSQL As String
''For i = fgPurchase.Rows - 1 To 1 Step -1
''  sSQL = ""
''  For j = 0 To fgPurchase.Cols - 1
''     sSQL = sSQL & Trim(fgPurchase.TextMatrix(i, j))
''  Next
''  If Trim(sSQL) = "" Then
''      fgPurchase.RemoveItem i
''  End If
''  If fgPurchase.Rows <= 1 Then Exit For
''Next
''
''
''
'''    If Col = 5 Then
'''        If CDbl(fgPurchase.TextMatrix(Row, 11)) > CDbl(fgPurchase.TextMatrix(Row, 10)) Then
'''            MsgBox "Rejected quantity can't be greater than Receive Quantity.", vbInformation
'''            fgPurchase.Select Row, 11
'''        End If
'''    End If
''End Sub
''
''
''
'''Private Sub deleteRow()
'''If fgPurchase.Rows = 1 Then fgPurchase.AddItem ""
'''        On Error Resume Next
'''        If fgPurchase.Rows <= 1 Then Exit Sub
'''Dim i As Integer
'''Dim j As Integer
'''    For i = 0 To fgPurchase.Rows - 1
'''             j = 1
'''             For j = 1 To fgPurchase.Rows - 1
'''
'''                If fgPurchase.TextMatrix(i, 5) = fgPurchase.TextMatrix(j, 5) Then
'''                    fgPurchase.RemoveItem j
'''                End If
'''             Next
'''    Next
'''
'''End Sub
''
''
'
''----------------------------------------------------------------------------------
'
''Option Explicit
''
''Private Sub Check1_Click()
'''Dim iRows As Integer
'''Dim i As Integer
''' For i = 1 To fgPurchase.FixedRows - 1
'''       fgPurchase.Cell(flexcpChecked, i, 10) = flexChecked
'''       Next
''      Dim f As Integer
''      f = fgPurchase.Rows - 1
''  If Check1.Value = 1 Then
''
'''f = fgPurchase.Rows - 1
''Dim i As Integer
''For i = 1 To f
''    fgPurchase.Cell(flexcpChecked, i, 10) = flexChecked
'''    fgPurchase.TextMatrix(i, 2) = "Descripción"
'''    fgPurchase.TextMatrix(i, 3) = "Precio"
''Next i '
'''Dim f As Integer
'''f = fgPurchase.Rows - 1
'''Dim i As Integer
''Else
''For i = 1 To f
''    fgPurchase.Cell(flexcpChecked, i, 1, i, 1) = flexUnchecked
'''    fgPurchase.TextMatrix(i, 2) = "Descripción"
'''    fgPurchase.TextMatrix(i, 3) = "Precio"
''    Next i
''
''
''End If
''
''
''
''End Sub
''
''Private Sub cmdAddItem_Click()
''frmItemInputReceiving.Show vbModal
''End Sub
''
''Private Sub cmdNew_Click()
''' Dim r&
'''
'''        For r = fgPurchase.FixedRows To fgPurchase.Rows - 1
'''
'''            If fgPurchase.Cell(flexcpChecked, r, 1) = flexChecked Then
'''
'''                Debug.Print fgPurchase.TextMatrix(r, 1); " is Checked"
'''
'''            End If
'''
'''        Next
''
''Dim f As Integer
''f = fgPurchase.Rows - 1
''Dim i As Integer
''For i = 1 To f
''    fgPurchase.Cell(flexcpChecked, i, 1, i, 1) = flexChecked
'''    fgPurchase.TextMatrix(i, 2) = "Descripción"
'''    fgPurchase.TextMatrix(i, 3) = "Precio"
''Next i
''
''
''End Sub
''
''Private Sub fgPurchase_Click()
''fgPurchase.Editable = flexEDKbdMouse
''End Sub
''
''Private Sub Form_Load()
'' fgPurchase.ColDataType(10) = flexDTBoolean
''' fgPurchase.ColDataType(1) = flexDTBoolean
'End Sub
'
''------------------------------------------
'
'
