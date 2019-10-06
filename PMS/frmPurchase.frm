VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmPurchase 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Purchase Informations [Doctors Clinic Unit - 2]"
   ClientHeight    =   10515
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   15240
   Icon            =   "frmPurchase.frx":0000
   ScaleHeight     =   10515
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   9240
      Width           =   1575
   End
   Begin VB.TextBox txtUName 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      Locked          =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Purchase Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   14895
      Begin VSFlex7LCtl.VSFlexGrid fgPurchase 
         Height          =   6135
         Left            =   120
         TabIndex        =   43
         Top             =   360
         Width           =   14775
         _cx             =   26061
         _cy             =   10821
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
         BackColorAlternate=   16761024
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
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPurchase.frx":000C
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   480
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
         Left            =   14280
         Picture         =   "frmPurchase.frx":01CF
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   1440
         Width           =   480
      End
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
         Height          =   420
         Left            =   10920
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1695
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
         Height          =   420
         Left            =   9000
         TabIndex        =   17
         Text            =   " "
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtTotalAmt 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtPaid 
         BackColor       =   &H00C0C000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   9000
         TabIndex        =   15
         Text            =   " "
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtVatAmount 
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
         Left            =   5760
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtVatPercentage 
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
         Left            =   4320
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
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
         Height          =   420
         Left            =   7320
         TabIndex        =   12
         Text            =   " "
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txtSupplierBill 
         Height          =   420
         Left            =   2520
         TabIndex        =   1
         Text            =   " "
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtSerialNo 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   600
         Width           =   2295
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   13680
         Top             =   960
      End
      Begin VB.TextBox txtTime 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   420
         Left            =   10920
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbMSupplierName 
         Height          =   315
         Left            =   2520
         TabIndex        =   0
         Top             =   600
         Width           =   4575
      End
      Begin MSComCtl2.DTPicker DTPPurchaseDate 
         Height          =   420
         Left            =   120
         TabIndex        =   19
         Top             =   1320
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   741
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   16449539
         CurrentDate     =   38465
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Time"
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
         Left            =   10920
         TabIndex        =   30
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
         Left            =   9000
         TabIndex        =   29
         Top             =   360
         Width           =   1695
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
         Left            =   7320
         TabIndex        =   28
         Top             =   360
         Width           =   1575
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
         Left            =   9000
         TabIndex        =   27
         Top             =   1080
         Width           =   1695
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
         Left            =   5760
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
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
         Left            =   4320
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
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
         Left            =   7320
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         Height          =   300
         Left            =   2520
         TabIndex        =   21
         Top             =   360
         Width           =   3975
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
         TabIndex        =   20
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9240
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   9240
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9240
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   9240
      Width           =   735
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPost 
      Height          =   615
      Left            =   8040
      TabIndex        =   7
      Top             =   9840
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
   Begin SSDataWidgets_A.SSDBCommand cmdCancel 
      Height          =   615
      Left            =   3240
      TabIndex        =   8
      Top             =   9840
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
   Begin SSDataWidgets_A.SSDBCommand cmdAddItem 
      Height          =   495
      Left            =   11640
      TabIndex        =   2
      Top             =   9720
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
      Left            =   360
      TabIndex        =   33
      Top             =   9840
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
   Begin SSDataWidgets_A.SSDBCommand cmdEdit 
      Height          =   615
      Left            =   1320
      TabIndex        =   34
      Top             =   9840
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
      Left            =   2280
      TabIndex        =   35
      Top             =   9840
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
      Left            =   7080
      TabIndex        =   36
      Top             =   9840
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
   Begin SSDataWidgets_A.SSDBCommand cmdClose 
      Height          =   615
      Left            =   4170
      TabIndex        =   37
      Top             =   9840
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
   Begin SSDataWidgets_A.SSDBCommand cmdUndoPost 
      Height          =   615
      Left            =   9000
      TabIndex        =   38
      Top             =   9840
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
   Begin SSDataWidgets_A.SSDBCommand cmdPreview 
      Height          =   615
      Left            =   5160
      TabIndex        =   39
      Top             =   9840
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
      Font3D          =   4
      CaptionAlignment=   7
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   9960
      Top             =   9360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Height          =   615
      Left            =   6120
      TabIndex        =   40
      Top             =   9840
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "P&rint"
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MEDICINE PURCHASE INFORMATION"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   15255
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 Private rsItemMaster                    As ADODB.Recordset
 Private rsItemDetail                    As ADODB.Recordset
 Private rs                              As ADODB.Recordset
 Private bRecordExists                   As Boolean
 Dim str                                 As String
 Private rsTemp2                         As ADODB.Recordset
'---------------------------------------------------------------------------
'--------------------Add For Reporting Perpose------------------------------
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
Private temp3 As Double
Private temp4 As Double
'--------------------------------------------------------------------------------

Private Sub cmAutoPaid_Click()
txtPaid = txtTotalAmt
End Sub

Private Sub cmbMSupplierName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMSupplierName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbMSupplierName_GotFocus()
cmbMSupplierName.BackColor = &HFFFFC0
End Sub

Private Sub cmbMSupplierName_LostFocus()
    cmbMSupplierName.BackColor = vbWhite
End Sub

Private Sub cmdFirst_Click()
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
       
       txtSerialNo = Adodc2.Recordset!SerialNo
        cmbMSupplierName = Adodc2.Recordset!SName
        txtVatPercentage = Adodc2.Recordset!VatP
        txtVatAmount = Adodc2.Recordset!VatT
        txtTime = Adodc2.Recordset!Ptime
        DTPPurchaseDate = Adodc2.Recordset!PDate
        txtSupplierBill = Adodc2.Recordset!SbillNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtPaid = Adodc2.Recordset!Tpaid
        txtDiscountPer.text = Adodc2.Recordset!Discountp
        txtDiscountAmt.text = Adodc2.Recordset!Discount
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        
        fgPurchase.Rows = 1
        
        strPaymentDetail = "SELECT  SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, Qty, PRate, Amount, SRate, DateExp, DateM, Posted FROM PurchaseDetail " & _
                            "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "


'    strPaymentDetail = "Exec  Sms_PurchaseDetails_Find '" & parseQuotes(Me.txtSerialNo) & "'"

    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgPurchase.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgPurchase.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgPurchase.TextMatrix(i, 2) = rsItemDetail("Sname")
            fgPurchase.TextMatrix(i, 3) = IIf(IsNull(rsItemDetail("Pdate")), "", Format(rsItemDetail("Pdate"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 4) = rsItemDetail("SbillNo")
            fgPurchase.TextMatrix(i, 5) = rsItemDetail("MCatagory")
            fgPurchase.TextMatrix(i, 6) = rsItemDetail("Mname")
            fgPurchase.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgPurchase.TextMatrix(i, 8) = rsItemDetail("PRate")
            fgPurchase.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgPurchase.TextMatrix(i, 10) = rsItemDetail("SRate")
            fgPurchase.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("DateExp")), "", Format(rsItemDetail("DateExp"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 12) = IIf(IsNull(rsItemDetail("DateM")), "", Format(rsItemDetail("DateM"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 13) = rsItemDetail("Posted")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

        
        
End If

End Sub

Private Sub cmdLast_Click()
'On Error GoTo ErrorHandler
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
       
       txtSerialNo = Adodc2.Recordset!SerialNo
        cmbMSupplierName = Adodc2.Recordset!SName
        txtVatPercentage = Adodc2.Recordset!VatP
        txtVatAmount = Adodc2.Recordset!VatT
        txtTime = Adodc2.Recordset!Ptime
        DTPPurchaseDate = Adodc2.Recordset!PDate
        txtSupplierBill = Adodc2.Recordset!SbillNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtPaid = Adodc2.Recordset!Tpaid
        txtDiscountPer.text = Adodc2.Recordset!Discountp
        txtDiscountAmt.text = Adodc2.Recordset!Discount
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        
        
        fgPurchase.Rows = 1
        
         strPaymentDetail = "SELECT  SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, Qty, PRate, Amount, SRate, DateExp, DateM, Posted FROM PurchaseDetail " & _
                            "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "


'    strPaymentDetail = "Exec  Sms_PurchaseDetails_Find '" & parseQuotes(Me.txtSerialNo) & "'"

    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgPurchase.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
             fgPurchase.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgPurchase.TextMatrix(i, 2) = rsItemDetail("Sname")
            fgPurchase.TextMatrix(i, 3) = IIf(IsNull(rsItemDetail("Pdate")), "", Format(rsItemDetail("Pdate"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 4) = rsItemDetail("SbillNo")
            fgPurchase.TextMatrix(i, 5) = rsItemDetail("MCatagory")
            fgPurchase.TextMatrix(i, 6) = rsItemDetail("Mname")
            fgPurchase.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgPurchase.TextMatrix(i, 8) = rsItemDetail("PRate")
            fgPurchase.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgPurchase.TextMatrix(i, 10) = rsItemDetail("SRate")
            fgPurchase.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("DateExp")), "", Format(rsItemDetail("DateExp"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 12) = IIf(IsNull(rsItemDetail("DateM")), "", Format(rsItemDetail("DateM"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 13) = rsItemDetail("Posted")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

        
        
End If

'ErrorHandler:
'    If Err = 3021 Then    ' no current record
'        Resume Next
'    Else
'        MsgBox "No record found"
'        Resume ErrorHandlerExit
'    End If
End Sub

Private Sub cmdNext_Click()

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
       
       txtSerialNo = Adodc2.Recordset!SerialNo
        cmbMSupplierName = Adodc2.Recordset!SName
        txtVatPercentage = Adodc2.Recordset!VatP
        txtVatAmount = Adodc2.Recordset!VatT
        txtTime = Adodc2.Recordset!Ptime
        DTPPurchaseDate = Adodc2.Recordset!PDate
        txtSupplierBill = Adodc2.Recordset!SbillNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtPaid = Adodc2.Recordset!Tpaid
        txtDiscountPer.text = Adodc2.Recordset!Discountp
        txtDiscountAmt.text = Adodc2.Recordset!Discount
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
            
        fgPurchase.Rows = 1
        
        strPaymentDetail = "SELECT  SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, Qty, PRate, Amount, SRate, DateExp, DateM, Posted FROM PurchaseDetail " & _
                            "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "


'    strPaymentDetail = "Exec  Sms_PurchaseDetails_Find '" & parseQuotes(Me.txtSerialNo) & "'"

    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgPurchase.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
             fgPurchase.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgPurchase.TextMatrix(i, 2) = rsItemDetail("Sname")
            fgPurchase.TextMatrix(i, 3) = IIf(IsNull(rsItemDetail("Pdate")), "", Format(rsItemDetail("Pdate"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 4) = rsItemDetail("SbillNo")
            fgPurchase.TextMatrix(i, 5) = rsItemDetail("MCatagory")
            fgPurchase.TextMatrix(i, 6) = rsItemDetail("Mname")
            fgPurchase.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgPurchase.TextMatrix(i, 8) = rsItemDetail("PRate")
            fgPurchase.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgPurchase.TextMatrix(i, 10) = rsItemDetail("SRate")
            fgPurchase.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("DateExp")), "", Format(rsItemDetail("DateExp"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 12) = IIf(IsNull(rsItemDetail("DateM")), "", Format(rsItemDetail("DateM"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 13) = rsItemDetail("Posted")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

        
        
End If
End Sub

Private Sub cmdPreview_Click()
    Tracer = 0
    Call printReport
End Sub

Private Sub cmdAddItem_Click()
frmPurchaseSearch.Show vbModal
Call Calculation
End Sub


Private Sub cmdCancel_Click()
    
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Password,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
            
 If rs!Privilegegroup = 0 Then
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    cmdPost.Caption = "&Post"
    CmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
    cmdPost.Enabled = True
    Call alldisable
    If Not rsItemMaster.EOF Then FindRecord
Else
cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    cmdPost.Caption = "&Post"
    CmdDelete.Enabled = True
    cmdPreview.Enabled = True
'    cmdPrint.Enabled = True
    cmdPost.Enabled = True
    cmdUndoPost.Enabled = True
    Call alldisable
    If Not rsItemMaster.EOF Then FindRecord
End If
    
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
            cn.Execute "Delete From PurchaseMaster Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
            cn.Execute "DELETE FROM PurchaseDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
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
        cmdOpen.Enabled = False
'        cmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
'        cmdLAdd.Enabled = True
        cmdLDelete.Enabled = True
        fgPurchase.Editable = flexEDKbdMouse
        txtSerialNo.Enabled = False
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
                cmdOpen.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
'                cmdDelete.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                fgPurchase.Editable = flexEDNone
                Call alldisable
                
                rsItemMaster.Requery
                Dim s As String
                s = txtSerialNo
                rsItemMaster.MoveFirst
                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord
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
        cmdOpen.Enabled = False
        CmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
'        cmdLAdd.Enabled = True
        cmdLDelete.Enabled = True
        fgPurchase.Editable = flexEDKbdMouse
        txtSerialNo.Enabled = False
        cmdPost.Enabled = False
        cmdUndoPost.Enabled = False
        Call Calculation
        
    ElseIf cmdEdit.Caption = "&Update" Then
          Call Calculation
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                CmdDelete.Enabled = True
                cmdClose.Enabled = True
                cmdPost.Enabled = True
                cmdUndoPost.Enabled = True
                fgPurchase.Editable = flexEDNone
                Call alldisable
                
                rsItemMaster.Requery
'                Dim s As String
                s = txtSerialNo
                rsItemMaster.MoveFirst
                rsItemMaster.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord
                
            End If
        End If
    End If
'  End If

End If

End Sub

Private Sub cmdLAdd_Click()
With fgPurchase
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
fgPurchase.Editable = flexEDKbdMouse
'Call postedCheck


If cmdPost.Caption = "&Posting" Then
     If txtPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 fgPurchase.Enabled = False
                 cmdOpen.Enabled = True
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
       
       txtSerialNo = Adodc2.Recordset!SerialNo
        cmbMSupplierName = Adodc2.Recordset!SName
        txtVatPercentage = Adodc2.Recordset!VatP
        txtVatAmount = Adodc2.Recordset!VatT
        txtTime = Adodc2.Recordset!Ptime
        DTPPurchaseDate = Adodc2.Recordset!PDate
        txtSupplierBill = Adodc2.Recordset!SbillNo
        txtTotalAmt = Adodc2.Recordset!TAmount
        txtPaid = Adodc2.Recordset!Tpaid
        txtDiscountPer.text = Adodc2.Recordset!Discountp
        txtDiscountAmt.text = Adodc2.Recordset!Discount
        txtPost.text = Adodc2.Recordset!Posted
        txtUName.text = Adodc2.Recordset!UName
        
        fgPurchase.Rows = 1
        
        strPaymentDetail = "SELECT  SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, Qty, PRate, Amount, SRate, DateExp, DateM, Posted FROM PurchaseDetail " & _
                            "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "


'    strPaymentDetail = "Exec  Sms_PurchaseDetails_Find '" & parseQuotes(Me.txtSerialNo) & "'"

    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgPurchase.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgPurchase.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgPurchase.TextMatrix(i, 2) = rsItemDetail("Sname")
            fgPurchase.TextMatrix(i, 3) = IIf(IsNull(rsItemDetail("Pdate")), "", Format(rsItemDetail("Pdate"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 4) = rsItemDetail("SbillNo")
            fgPurchase.TextMatrix(i, 5) = rsItemDetail("MCatagory")
            fgPurchase.TextMatrix(i, 6) = rsItemDetail("Mname")
            fgPurchase.TextMatrix(i, 7) = rsItemDetail("Qty")
            fgPurchase.TextMatrix(i, 8) = rsItemDetail("PRate")
            fgPurchase.TextMatrix(i, 9) = rsItemDetail("Amount")
            fgPurchase.TextMatrix(i, 10) = rsItemDetail("SRate")
            fgPurchase.TextMatrix(i, 11) = IIf(IsNull(rsItemDetail("DateExp")), "", Format(rsItemDetail("DateExp"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 12) = IIf(IsNull(rsItemDetail("DateM")), "", Format(rsItemDetail("DateM"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 13) = rsItemDetail("Posted")
            
        rsItemDetail.MoveNext
        Next
      End If
        rsItemDetail.Close

        
        
End If
End Sub


Private Sub cmdPrint_Click()
Tracer = 1
'    Call FetchData
   Call printReport
End Sub

Private Sub cmdUndoPost_Click()
Dim s As String
cmdUndoPost.Caption = "&Undo Posting"
fgPurchase.Editable = flexEDKbdMouse
'Call postedCheck

 Call Calculation
If cmdUndoPost.Caption = "&Undo Posting" Then
   Call Calculation
     If txtPost.text = "Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = True
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 fgPurchase.Enabled = False
                 cmdOpen.Enabled = True
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
cmdUndoPost.Caption = "&Modify"
End Sub


Private Sub cmdLDelete_Click()
    
    
    If fgPurchase.Rows = 1 Then Exit Sub

     If fgPurchase.Row >= 1 Then
      If MsgBox("Are you sure to delete the selected record", vbYesNo, "General Setup") = vbYes Then fgPurchase.RemoveItem fgPurchase.Row
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
'str = "select SerialNo,Password,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='USER'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
' ----------------Check End------
            
If rs!Privilegegroup = "0" Then

If cmdNew.Caption = "&New" Then
        
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdOpen.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdClose.Enabled = False
        cmdPost.Enabled = False
        cmdUndoPost.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        TextClear Me
        Call Clear
         
        fgPurchase.Rows = 1
        fgPurchase.Editable = flexEDKbdMouse
        Call allenable
        txtPost.text = "Not Posted"
        txtUName.text = frmLogin.txtUName.text
        cmbMSupplierName.SetFocus

        
    ElseIf cmdNew.Caption = "&Save" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
'                cmdDelete.Enabled = True
                cmdOpen.Enabled = True
                cmdCancel.Enabled = True
                cmdPrint.Enabled = True
                cmdPost.Enabled = True
                
                Call alldisable
            End If
        End If
    End If
    
Else

If cmdNew.Caption = "&New" Then
        
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdOpen.Enabled = False
        CmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdClose.Enabled = False
        cmdPost.Enabled = False
        cmdUndoPost.Enabled = False
        cmdLDelete.Enabled = True
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        TextClear Me
        Call Clear
         
        fgPurchase.Rows = 1
        fgPurchase.Editable = flexEDKbdMouse
        Call allenable
        txtPost.text = "Not Posted"
        txtUName.text = frmLogin.txtUName.text
        cmbMSupplierName.SetFocus

        
    ElseIf cmdNew.Caption = "&Save" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                CmdDelete.Enabled = True
                cmdOpen.Enabled = True
                cmdCancel.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
                cmdUndoPost.Enabled = True
                
                Call alldisable
            End If
        End If
    End If
End If
    
End Sub

Private Sub Clear()
    txtSerialNo.text = ""
    DTPPurchaseDate.Enabled = True
    cmbMSupplierName.text = ""
    txtSupplierBill.text = ""
    txtVatPercentage.text = "0"
    txtVatAmount.text = "0"
    txtPaid.text = "0"
    txtDiscountPer.text = "0"
    txtDiscountAmt.text = "0"
    txtTotalAmt.text = ""
End Sub

Private Sub allenable()
     txtSerialNo.Enabled = True
     cmdLDelete.Enabled = True
     fgPurchase.Enabled = True
     DTPPurchaseDate.Enabled = True
     DTPPurchaseDate.Value = Date
     cmbMSupplierName.Enabled = True
     txtVatPercentage.Enabled = True
     txtVatAmount.Enabled = True
     txtTime.Enabled = True
     txtSupplierBill.Enabled = True
     txtTotalAmt.Enabled = True
     txtPaid.Enabled = True
     txtDiscountPer.Enabled = True
     txtDiscountAmt.Enabled = True
     cmdAddItem.Enabled = True
    End Sub

Private Sub alldisable()
     txtSerialNo.Enabled = False
     cmdLDelete.Enabled = False
     fgPurchase.Enabled = False
     DTPPurchaseDate.Enabled = False
     DTPPurchaseDate.Value = Date
     cmbMSupplierName.Enabled = False
     txtVatPercentage.Enabled = False
     txtVatAmount.Enabled = False
     txtTime.Enabled = False
     txtSupplierBill.Enabled = False
     txtTotalAmt.Enabled = False
     txtPaid.Enabled = False
     txtDiscountPer.Enabled = False
     txtDiscountAmt.Enabled = False
     cmdAddItem.Enabled = False

    
End Sub

Private Sub cmdOpen_Click()
    frmMPFind.Show vbModal
    Call Calculation
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
        
End Sub

Private Sub fgPurchase_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
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
       Call MSupplierName
       
   Set rsItemMaster = New ADODB.Recordset
 
  If rsItemMaster.State <> 0 Then rsItemMaster.Close
  rsItemMaster.Open "select TOP 1 * FROM PurchaseMaster ORDER BY SerialNo DESC", cn, adOpenStatic, adLockReadOnly
'  rsItemMaster.Open "exec Sms_Purchase_Master", cn, adOpenStatic, adLockReadOnly
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
  Adodc2.RecordSource = "PurchaseMaster"

  Adodc2.Refresh
'-------------------End Record Search---------
 Call Calculation
 Call ModifyVisible
End Sub

Private Sub ModifyVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UName,Password,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
'           If rs!Name = "Admin" And rs!Password = "123" Then
            If rs!Name = "Admin" Then
              cmdUndoPost.Visible = True
            
        ElseIf rs!Name = "BORHAN" And rs!Password = "01920468031" Then
        
              cmdUndoPost.Visible = True
           Else
               cmdUndoPost.Visible = False
               
           End If
End Sub

Private Sub Calculation()
  Dim j As Integer
       temp = 0
       temp1 = 0
       temp2 = 0
       temp3 = 0
       temp4 = 0
    For j = 1 To fgPurchase.Rows - 1

        temp = temp + CDbl(Val(fgPurchase.TextMatrix(j, 8)) * CDbl(Val(fgPurchase.TextMatrix(j, 9))))

   Next
   
   
'txtTotalAmt = temp + (CDbl(Val(txtVatPercentage) / 100) + CDbl(txtVatAmount)) - ((CDbl(Val(txtDiscountPer) / 100) + CDbl(txtDiscountAmt)))
temp1 = (temp * CDbl(Val(txtVatPercentage) / 100))
'If txtVatAmount = "" Then
'   txtVatAmount = 0
'End If
temp2 = CDbl(Val(txtVatAmount))
temp3 = (temp * CDbl(Val(txtDiscountPer) / 100))
If txtDiscountAmt = Empty Then
   txtDiscountAmt = 0
End If
temp4 = CDbl(Val(txtDiscountAmt))
txtTotalAmt = (temp + temp1 + temp2) - (temp3 + temp4)

'txtTotalAmt = (temp + (temp * CDbl(Val(txtVatPercentage) / 100)) + CDbl(txtVatAmount) + CDbl(txtVatAmount))


'txtTotalSCharge = temp * CDbl(Val(txtServiceCharge) / 100)
'txtTotalDiscount = temp * CDbl(Val(txtDiscount) / 100)
'txtNPayable = CDbl(txtTotalBill) + CDbl(txtTotalVat) + CDbl(txtTotalSCharge) - CDbl(txtTotalDiscount)
'TxtDue = (CDbl(txtTotalBill) + CDbl(txtTotalVat) + CDbl(txtTotalSCharge) - CDbl(txtTotalDiscount)) - CDbl(Val(txtPaid))
'txtWords = InWords(txtNPayable.text)

End Sub

Private Sub MSupplierName()
    Dim rsTemp2 As New ADODB.Recordset
          
     rsTemp2.Open ("SELECT DISTINCT SName FROM Suppliers ORDER BY SName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbMSupplierName.AddItem rsTemp2("SName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub



 Private Function rcupdate() As Boolean

On Error GoTo ErrHandler
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
'--------------------------------------------------------------Group permission end-------------------
    If rs!Privilegegroup = "0" Then
     
     If cmdNew.Caption = "&Save" Then
        
     strSQL = "INSERT INTO PurchaseMaster (Sname,VatP, VatT, Ptime,Pdate,SbillNo,Tamount,Tpaid,Discountp,Discount,Posted,UName " & _
                ") " & _
                "VALUES ('" & cmbMSupplierName.text & "'," & Val(txtVatPercentage.text) & "," & Val(txtVatAmount.text) & ",'" & txtTime & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "', " & _
                " '" & txtSupplierBill & "'," & Val(txtTotalAmt.text) & "," & Val(txtPaid.text) & "," & Val(txtDiscountPer.text) & "," & Val(txtDiscountAmt.text) & ",'" & txtPost & "','" & txtUName.text & "')"
     cn.Execute strSQL
      rcupdate = True
     
'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),1) as InvNo from PurchaseMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo = Val(rs!InvNo)
           
           j = 0
            For j = 1 To fgPurchase.Rows - 1
            
     
             cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"
               Next
        
        rcupdate = True
        
        MsgBox "Record added Successfully", vbInformation, "Confirmation"
     
        ElseIf (cmdEdit.Caption = "&Update") Then
           
        cn.Execute "UPDATE PurchaseMaster SET Sname='" & cmbMSupplierName.text & "', " & _
                   "VatP=" & Val(txtVatPercentage.text) & ",VatT=" & Val(txtVatAmount.text) & ",Ptime='" & txtTime & "', " & _
                   "Pdate = '" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "',SbillNo='" & txtSupplierBill & "', Tamount=" & Val(txtTotalAmt.text) & ", Tpaid = " & Val(txtPaid.text) & ", " & _
                   "Discountp = " & Val(txtDiscountPer.text) & ",Discount = " & Val(txtDiscountAmt.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM PurchaseDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"

           j = 0
              For j = 1 To fgPurchase.Rows - 1
            
     
          cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"

               Next
        
        rcupdate = True
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
        
'  --------------------------------Posting Information-----------------------------------------

   ElseIf cmdPost.Caption = "&Posting" Then
   
iPost = MsgBox("Do you want to Post this bill?", vbYesNo)

If iPost = vbYes Then
     
     txtPost.text = "Posted"
     
     cn.Execute "UPDATE PurchaseMaster SET Sname='" & cmbMSupplierName.text & "', " & _
                " VatP=" & Val(txtVatPercentage.text) & ",VatT=" & Val(txtVatAmount.text) & ",Ptime='" & txtTime & "', " & _
                "Pdate = '" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "',SbillNo='" & txtSupplierBill & "', Tamount=" & Val(txtTotalAmt.text) & ", Tpaid = " & Val(txtPaid.text) & ", " & _
                "Discountp = " & Val(txtDiscountPer.text) & ",Discount = " & Val(txtDiscountAmt.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"

        cn.Execute "DELETE FROM PurchaseDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"

        j = 0
            For j = 1 To fgPurchase.Rows - 1
            
         cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"

               Next
        
        rcupdate = True
        
        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
        End If
     End If
Else

'------------------------------Admin group Start------------------

     If cmdNew.Caption = "&Save" Then
        
     strSQL = "INSERT INTO PurchaseMaster (Sname,VatP, VatT, Ptime,Pdate,SbillNo,Tamount,Tpaid,Discountp,Discount,Posted,UName " & _
                ") " & _
                "VALUES ('" & cmbMSupplierName.text & "'," & Val(txtVatPercentage.text) & "," & Val(txtVatAmount.text) & ",'" & txtTime & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "', " & _
                " '" & txtSupplierBill & "'," & Val(txtTotalAmt.text) & "," & Val(txtPaid.text) & "," & Val(txtDiscountPer.text) & "," & Val(txtDiscountAmt.text) & ",'" & txtPost & "','" & txtUName.text & "')"
     cn.Execute strSQL
      rcupdate = True
      
'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),1) as InvNo from PurchaseMaster"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSerialNo = Val(rs!InvNo)
     
      j = 0
            For j = 1 To fgPurchase.Rows - 1
            
     
        cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"

               Next
        
        rcupdate = True
        
'        cn.CommitTrans
        MsgBox "Record added Successfully", vbInformation, "Confirmation"

        ElseIf (cmdEdit.Caption = "&Update") Then
               
                    cn.Execute "UPDATE PurchaseMaster SET Sname='" & cmbMSupplierName.text & "', " & _
                              " VatP=" & Val(txtVatPercentage.text) & ",VatT=" & Val(txtVatAmount.text) & ",Ptime='" & txtTime & "', " & _
                              "Pdate = '" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "',SbillNo='" & txtSupplierBill & "', Tamount=" & Val(txtTotalAmt.text) & ", Tpaid = " & Val(txtPaid.text) & ", " & _
                              "Discountp = " & Val(txtDiscountPer.text) & ",Discount = " & Val(txtDiscountAmt.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM PurchaseDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


         j = 0
            For j = 1 To fgPurchase.Rows - 1
            
     
       cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"

               Next
        
        rcupdate = True
        
        MsgBox "Record updated Successfully", vbInformation, "Confirmation"
        
'  --------------------------------Posting Information-----------------------------------------

   ElseIf cmdPost.Caption = "&Posting" Then
  
iPost = MsgBox("Do you want to Post this bill?", vbYesNo)

If iPost = vbYes Then
     
     txtPost.text = "Posted"
      cn.Execute "UPDATE PurchaseMaster SET Sname='" & cmbMSupplierName.text & "', " & _
                              " VatP=" & Val(txtVatPercentage.text) & ",VatT=" & Val(txtVatAmount.text) & ",Ptime='" & txtTime & "', " & _
                              "Pdate = '" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "',SbillNo='" & txtSupplierBill & "', Tamount=" & Val(txtTotalAmt.text) & ", Tpaid = " & Val(txtPaid.text) & ", " & _
                              "Discountp = " & Val(txtDiscountPer.text) & ",Discount = " & Val(txtDiscountAmt.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"

        cn.Execute "DELETE FROM PurchaseDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"


        j = 0
            For j = 1 To fgPurchase.Rows - 1
            
     
       cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"

               Next
        
        rcupdate = True
        
        MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
        End If

ElseIf cmdUndoPost.Caption = "&Undo Posting" Then
     
     txtPost.text = "Not Posted"

iPost = MsgBox("Do you want to Undo post this bill?", vbYesNo)

If iPost = vbYes Then
     
     txtPost.text = "Not Posted"
     cn.Execute "UPDATE PurchaseMaster SET Sname='" & cmbMSupplierName.text & "', " & _
                              " VatP=" & Val(txtVatPercentage.text) & ",VatT=" & Val(txtVatAmount.text) & ",Ptime='" & txtTime & "', " & _
                              "Pdate = '" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "',SbillNo='" & txtSupplierBill & "', Tamount=" & Val(txtTotalAmt.text) & ", Tpaid = " & Val(txtPaid.text) & ", " & _
                              "Discountp = " & Val(txtDiscountPer.text) & ",Discount = " & Val(txtDiscountAmt.text) & ",Posted='" & txtPost & "', UName='" & txtUName.text & "'  WHERE SerialNo = '" & parseQuotes(txtSerialNo) & "'"


        cn.Execute "DELETE FROM PurchaseDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"

        j = 0
            For j = 1 To fgPurchase.Rows - 1
            
       cn.Execute "INSERT INTO PurchaseDetail (SerialNo,Sname,Pdate,SbillNo,MID,MCatagory,Mname,Qty,PRate, " & _
                           "Amount,SRate,DateExp,DateM,Posted) " & _
                           "Values ('" & parseQuotes(txtSerialNo) & "','" & cmbMSupplierName.text & "','" & Format(DTPPurchaseDate, "dd-mmm-yyyy") & "','" & parseQuotes(txtSupplierBill) & "','" & parseQuotes(fgPurchase.TextMatrix(j, 5)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 6)) & "', " & _
                           "'" & parseQuotes(fgPurchase.TextMatrix(j, 7)) & "', " & _
                           IIf(fgPurchase.TextMatrix(j, 8) = "", "0", fgPurchase.TextMatrix(j, 8)) & "," & _
                           IIf(fgPurchase.TextMatrix(j, 9) = "", "0", fgPurchase.TextMatrix(j, 9)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 10) = "", "0", fgPurchase.TextMatrix(j, 10)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 11) = "", "0", fgPurchase.TextMatrix(j, 11)) & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 12) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 12) & "' ") & ", " & _
                           IIf(fgPurchase.TextMatrix(j, 13) = "", "NUll", "'" & fgPurchase.TextMatrix(j, 13) & "' ") & ", " & _
                           "'" & parseQuotes(txtPost.text) & "')"

               Next
        
        rcupdate = True
        
        MsgBox "Record Undo Posted Successfully", vbInformation, "Confirmation"
        
        End If


'-------------Undo Posting End----
    
    End If

'-------------Admin group end--------------

End If

    Exit Function
    
ErrHandler:

    Select Case Err.Number
        Case -2147217900
            MsgBox "Please select Numeric number in ROL field", vbInformation, "Confirmation"

   End Select

   If Err.Number = -2147217874 Then
    MsgBox "You can't Insert same item from same style multiple times in one BTB LC."
   End If
            MsgBox cn.Errors(0).NativeError & " : " & cn.Errors(0).Description
'    End Select
End Function

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    If Trim(cmbMSupplierName) = "" Then
        MsgBox "Your are missing Supplier Name.", vbInformation
        cmbMSupplierName.SetFocus
        IsValidRecord = False
        Exit Function
 
        
  ElseIf Trim(txtSupplierBill) = "" Then
        MsgBox "Your are missing Supplier Bill Information.", vbInformation
        txtSupplierBill.SetFocus
        IsValidRecord = False
        Exit Function
        
'-----------------------------------------------------------------------
    Else
        

       Exit Function
     End If
    End Function
    
Private Sub FindRecord()

    Dim i As Integer
    Dim strPaymentDetail As String
    Set rsItemDetail = New ADODB.Recordset
    txtSerialNo = rsItemMaster!SerialNo
    cmbMSupplierName = rsItemMaster!SName
    txtVatPercentage = rsItemMaster!VatP
    txtVatAmount = rsItemMaster!VatT
    txtTime = rsItemMaster!Ptime
    DTPPurchaseDate = rsItemMaster!PDate
    txtSupplierBill = rsItemMaster!SbillNo
    txtTotalAmt = rsItemMaster!TAmount
    txtPaid = rsItemMaster!Tpaid
    txtDiscountPer = rsItemMaster!Discountp
    txtDiscountAmt = rsItemMaster!Discount
    txtPost = rsItemMaster!Posted
    txtUName = rsItemMaster!UName
    

    fgPurchase.Rows = 1
    strPaymentDetail = "SELECT  SerialNo, Sname, Pdate, SbillNo, MID,MCatagory, Mname, Qty, PRate, Amount, SRate, DateExp, DateM, Posted FROM PurchaseDetail " & _
                       "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "

'strPaymentDetail = "Exec  Sms_PurchaseDetails_Find '" & parseQuotes(Me.txtSerialNo) & "'"

    rsItemDetail.CursorLocation = adUseClient
    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly


    If rsItemDetail.RecordCount <> 0 Then

        fgPurchase.Rows = rsItemDetail.RecordCount + 1
'                i = 0
        For i = 1 To rsItemDetail.RecordCount
            fgPurchase.TextMatrix(i, 1) = rsItemDetail("SerialNo")
            fgPurchase.TextMatrix(i, 2) = rsItemDetail("Sname")
            fgPurchase.TextMatrix(i, 3) = IIf(IsNull(rsItemDetail("Pdate")), "", Format(rsItemDetail("Pdate"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 4) = rsItemDetail("SbillNo")
            fgPurchase.TextMatrix(i, 5) = rsItemDetail("MID")
            fgPurchase.TextMatrix(i, 6) = rsItemDetail("MCatagory")
            fgPurchase.TextMatrix(i, 7) = rsItemDetail("Mname")
            fgPurchase.TextMatrix(i, 8) = rsItemDetail("Qty")
            fgPurchase.TextMatrix(i, 9) = rsItemDetail("PRate")
            fgPurchase.TextMatrix(i, 10) = rsItemDetail("Amount")
            fgPurchase.TextMatrix(i, 11) = rsItemDetail("SRate")
            fgPurchase.TextMatrix(i, 12) = IIf(IsNull(rsItemDetail("DateExp")), "", Format(rsItemDetail("DateExp"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 13) = IIf(IsNull(rsItemDetail("DateM")), "", Format(rsItemDetail("DateM"), "dd-mmm-yyyy"))
            fgPurchase.TextMatrix(i, 14) = rsItemDetail("Posted")
            
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

    
        strPath = App.Path + "\reports\PurchasePreview.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        


    Set rsDailyRpt = New ADODB.Recordset
If rsDailyRpt.State <> 0 Then rsDailyRpt.Close
      
'            strSQL = "SELECT PurchaseMaster.SerialNo, PurchaseMaster.DTPPurchaseDate,PurchaseMaster.SupplierName,PurchaseMaster.SupplierBill, " & _
'                      "PurchaseMaster.BudgetHead, PurchaseDetail.ProductMCatagory, PurchaseDetail.SubCode, PurchaseDetail.ItemName, " & _
'                      "PurchaseDetail.Quentity,PurchaseDetail.Rate, PurchaseDetail.Amount, PurchaseDetail.ExpDate, " & _
'                      "PurchaseDetail.Posted , PurchaseDetail.Warrenty, PurchaseDetail.Remarks,PurchaseMaster.UName " & _
'                      "FROM PurchaseMaster INNER JOIN " & _
'                      "PurchaseDetail ON PurchaseMaster.SerialNo = PurchaseDetail.SerialNo and PurchaseMaster.SerialNo ='" & Me.txtSerialNo & "' ORDER BY PurchaseDetail.ProductMCatagory "


            strSQL = "exec Sms_Purchase_Privew '" & Me.txtSerialNo & "' "
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
        objReport.Preview "Purchase Privew Report", , , , , 16777216 Or 524288 Or 65536
        Else
        objReport.PrintOut
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
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Item MCatagory Report"
    End Select
End Sub


Private Sub duplicate()
   Dim j As Integer
        
         For j = 1 To fgPurchase.Rows - 2
        
        If Val(fgPurchase.TextMatrix(j, 4)) = Val(fgPurchase.TextMatrix(j + 1, 4)) Then
        MsgBox "Duplicate Item Code Number.", vbInformation
         fgPurchase.TextMatrix(j, 4) = ""
         End If

         Next

End Sub

Public Sub PopulateForm(StrID As String)

    rsItemMaster.Close
    rsItemMaster.Open "select * from PurchaseMaster", cn, adOpenStatic, adLockReadOnly
    rsItemMaster.MoveFirst
    rsItemMaster.Find "SerialNo=" & parseQuotes(StrID)
    If rsItemMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


Private Sub fgPurchase_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Call Calculation

End Sub

Private Sub Timer1_Timer()
    txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub

Private Sub txtDiscountAmt_Change()
Call Calculation
End Sub

Private Sub txtDiscountPer_Change()
Call Calculation
End Sub

Private Sub txtSupplierBill_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub


Private Sub txtSupplierBill_GotFocus()
txtSupplierBill.BackColor = &HFFFFC0
End Sub

Private Sub txtSupplierBill_LostFocus()
    txtSupplierBill.BackColor = vbWhite
End Sub

Private Sub txtTime_Change()
txtTime.text = Time
End Sub

Private Sub txtPaid_Click()
Call Calculation
End Sub

Private Sub txtVatAmount_Change()
Call Calculation
End Sub

Private Sub txtVatPercentage_Change()
Call Calculation
End Sub
