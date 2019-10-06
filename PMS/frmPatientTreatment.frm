VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{CC696B60-4159-11D0-BDCB-0020A90B183A}#3.1#0"; "pvdate2.ocx"
Begin VB.Form frmPatientTreatment 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Patient Treatment Informations [Sun Medical Services]"
   ClientHeight    =   11010
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   15240
   Icon            =   "frmPatientTreatment.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0B4A9&
      Height          =   1215
      Left            =   120
      TabIndex        =   65
      Top             =   8280
      Width           =   3495
      Begin VB.TextBox txtCDiagnosis 
         Height          =   285
         Left            =   1200
         TabIndex        =   69
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txtFollowUp 
         Height          =   285
         Left            =   1200
         TabIndex        =   67
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Clinical Diagnosis "
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
         TabIndex        =   68
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblFollowup 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Follow Up"
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
         TabIndex        =   66
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0B4A9&
      Caption         =   "O/E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   1815
      Left            =   120
      TabIndex        =   54
      Top             =   3840
      Width           =   3495
      Begin VB.TextBox txtSpleen 
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtLiver 
         Height          =   285
         Left            =   720
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtLungs 
         Height          =   285
         Left            =   2400
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtHeart 
         Height          =   285
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtpulse 
         Height          =   285
         Left            =   2400
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtBP 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtothers 
         Height          =   285
         Left            =   720
         TabIndex        =   18
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lblothers 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Others"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   63
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label lblSpleen 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Spleen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   62
         Top             =   960
         Width           =   615
      End
      Begin VB.Label lblLiver 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Liver"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblLungs 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Lungs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   60
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblHeart 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Heart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   59
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblPulse 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Pulse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   58
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblBP 
         BackColor       =   &H00C0B4A9&
         Caption         =   "BP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Test Investigations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2415
      Left            =   120
      TabIndex        =   52
      Top             =   5760
      Width           =   3495
      Begin VB.TextBox txtInvestigation 
         Height          =   2055
         Left            =   120
         TabIndex        =   53
         Top             =   240
         Width           =   3255
      End
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
      Left            =   11520
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   10080
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Medicine Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   7575
      Left            =   3720
      TabIndex        =   21
      Top             =   1920
      Width           =   11415
      Begin VSFlex7LCtl.VSFlexGrid fgSales 
         Height          =   7215
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   11175
         _cx             =   19711
         _cy             =   12726
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatientTreatment.frx":030A
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1815
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   3495
      Begin VB.TextBox txtCC 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtIWeight 
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtBMI 
         Height          =   285
         Left            =   840
         TabIndex        =   8
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox txtWeight 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCC 
         BackColor       =   &H00C0B4A9&
         Caption         =   "C/C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   120
         TabIndex        =   64
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "kg"
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
         Left            =   3120
         TabIndex        =   57
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblIBWeight 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Ideal Body Weight"
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
         TabIndex        =   56
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblBMI 
         BackColor       =   &H00C0B4A9&
         Caption         =   "BMI"
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
         TabIndex        =   51
         Top             =   720
         Width           =   495
      End
      Begin VB.Label lblHeight 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Height"
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
         Left            =   1800
         TabIndex        =   50
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblcm 
         BackColor       =   &H00C0B4A9&
         Caption         =   "cm"
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
         Left            =   3120
         TabIndex        =   49
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblkg 
         BackColor       =   &H00C0B4A9&
         Caption         =   "kg"
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
         Left            =   1440
         TabIndex        =   42
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lblWeight 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Weight"
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
         TabIndex        =   41
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   15015
      Begin VB.ComboBox CmbBloodGroup 
         Height          =   315
         Left            =   10200
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtSerialNo 
         Height          =   375
         Left            =   13560
         TabIndex        =   48
         Text            =   "SerialNo"
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   13560
         Top             =   1200
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
         Left            =   11760
         Locked          =   -1  'True
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   1575
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
         Left            =   11760
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "frmPatientTreatment.frx":03C4
         Top             =   720
         Width           =   1575
      End
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
         Picture         =   "frmPatientTreatment.frx":03C8
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1200
         Width           =   480
      End
      Begin VB.TextBox txtPReg 
         BackColor       =   &H00404080&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   38
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtRefer_Code 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Doctor's ID"
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtDoc_Name 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1200
         Width           =   10335
      End
      Begin VB.ComboBox cmbSex 
         Height          =   315
         Left            =   4200
         TabIndex        =   2
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtAge 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtPatientName 
         Height          =   375
         Left            =   4200
         TabIndex        =   0
         Top             =   240
         Width           =   6855
      End
      Begin PVDATE2Lib.PVDate2 PDate 
         Height          =   375
         Left            =   6240
         TabIndex        =   3
         Top             =   720
         Width           =   2775
         _Version        =   196609
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   16776960
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Border          =   1
         DisplayFormat   =   6
         CalendarFormat  =   0
         DateFormat      =   13
         Separator       =   "-"
         TimeStore       =   -1  'True
         HighlightColor  =   8421376
         BackColor       =   0
         ForeColor       =   16776960
         Value           =   42478.9985416667
      End
      Begin SSDataWidgets_A.SSDBCommand cmdAddItem 
         Height          =   495
         Left            =   13800
         TabIndex        =   26
         Top             =   240
         Width           =   1095
         _Version        =   196611
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
      End
      Begin VB.Label lblBloodGroup 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Blood Group"
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
         Left            =   9120
         TabIndex        =   70
         Top             =   720
         Width           =   1095
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
         Left            =   11160
         TabIndex        =   47
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblAge 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Age"
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
         TabIndex        =   40
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Reg."
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
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Date"
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
         Left            =   5760
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference Code"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   120
         TabIndex        =   25
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label lblSex 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Sex"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblPatientName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Name"
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
         Left            =   3000
         TabIndex        =   22
         Top             =   240
         Width           =   1215
      End
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   615
      Left            =   120
      TabIndex        =   28
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   1080
      TabIndex        =   29
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   7800
      TabIndex        =   30
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   2040
      TabIndex        =   31
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   3000
      TabIndex        =   32
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
   Begin SSDataWidgets_A.SSDBCommand cmdClose 
      Height          =   615
      Left            =   5880
      TabIndex        =   33
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   6840
      TabIndex        =   34
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   8760
      TabIndex        =   35
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
      Left            =   4920
      TabIndex        =   36
      Top             =   10080
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
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
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   615
      Left            =   3960
      TabIndex        =   37
      Top             =   10080
      Width           =   975
      _Version        =   196611
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   6840
      Top             =   10680
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
End
Attribute VB_Name = "frmPatientTreatment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Private rsPrescriptionMaster            As ADODB.Recordset
 Private rsItemDetail                    As ADODB.Recordset
 Private rs                              As ADODB.Recordset
 Private bRecordExists                   As Boolean
 Dim str As String
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
'----Add For Reporting Perpose----------------------------------------------
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

Private Sub cmdAddItem_Click()
frmPrescriptionsearch.Show vbModal
End Sub

Private Sub cmdCancel_Click()
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close
str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
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
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
    cmdPost.Enabled = True
    Call alldisable
'    If Not rsPrescriptionMaster.EOF Then FindRecord
Else
cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    cmdPost.Caption = "&Post"
    cmdDelete.Enabled = True
    cmdPreview.Enabled = True
'    cmdPrint.Enabled = True
    cmdPost.Enabled = True
    cmdUndoPost.Enabled = True
    Call alldisable
'    If Not rsPrescriptionMaster.EOF Then FindRecord
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If txtUsername.text = "Admin" Then
    If idelete = vbYes Then
            cn.Execute "Delete From Prescription Where SerialNo ='" & parseQuotes(txtSerialNo) & "'"
'            cn.Execute "DELETE FROM SalesDetail WHERE SerialNo='" & parseQuotes(txtSerialNo) & "'"
            Call Clear
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
     End If
End Sub
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
' If txtPost.text = "Not Posted" Then
'    If cmdEdit.Caption = "&Edit" Then
'        cmdNew.Enabled = False
'        Call allenable
'        cmdEdit.Caption = "&Update"
'        cmdCancel.Enabled = True
'        cmdClose.Enabled = False
'        cmdOpen.Enabled = False
''        cmdDelete.Enabled = False
'        cmdPrint.Enabled = False
'        cmdPreview.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
'        fgSales.Editable = flexEDKbdMouse
'        txtSerialNo.Enabled = False
'        cmdPost.Enabled = False
'        Call Calculation
'
'    ElseIf cmdEdit.Caption = "&Update" Then
''          Call duplicate
'Call Calculation
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdOpen.Enabled = True
'                cmdPrint.Enabled = True
'                cmdPreview.Enabled = True
''                cmdDelete.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                fgSales.Editable = flexEDNone
'                Call alldisable
'
'                rsPrescriptionMaster.Requery
'                Dim s As String
'                s = txtSerialNo
'                rsPrescriptionMaster.MoveFirst
'                rsPrescriptionMaster.Find "SerialNo='" & parseQuotes(s) & "'"
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
'        cmdPrint.Enabled = False
'        cmdPreview.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
'        fgSales.Editable = flexEDKbdMouse
'        txtSerialNo.Enabled = False
'        cmdPost.Enabled = False
'        cmdUndoPost.Enabled = False
'        Call Calculation
'
'    ElseIf cmdEdit.Caption = "&Update" Then
'
'    Call Calculation
''          Call duplicate
'        If IsValidRecord Then
'            If rcupdate Then
'                cmdEdit.Caption = "&Edit"
'                cmdNew.Enabled = True
'                cmdCancel.Enabled = False
'                cmdClose.Enabled = True
'                cmdOpen.Enabled = True
'                cmdPrint.Enabled = True
'                cmdPreview.Enabled = True
'                cmdDelete.Enabled = True
'                cmdClose.Enabled = True
'                cmdPost.Enabled = True
'                cmdUndoPost.Enabled = True
'                fgSales.Editable = flexEDNone
'                Call alldisable
'
'                rsPrescriptionMaster.Requery
''                Dim s As String
'                s = txtSerialNo
'                rsPrescriptionMaster.MoveFirst
'                rsPrescriptionMaster.Find "SerialNo='" & parseQuotes(s) & "'"
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
Private Sub cmdLDelete_Click()
If fgSales.Rows = 1 Then Exit Sub

     If fgSales.Row >= 1 Then
      If MsgBox("Are you sure to delete the selected record", vbYesNo, "Delete Information") = vbYes Then fgSales.RemoveItem fgSales.Row
     Else
      MsgBox "You have to select a row to delete.", vbInformation, "Information"
    End If

'Call Calculation
End Sub

'Private Sub duplicate()
'   Dim j As Integer
'
'         For j = 1 To fgSales.Rows - 2
'
'        If Val(fgSales.TextMatrix(j, 4)) = Val(fgSales.TextMatrix(j + 1, 4)) Then
'        MsgBox "Duplicate Item Code Number.", vbInformation
'         fgSales.TextMatrix(j, 4) = ""
'         End If
'
'         Next
'
'End Sub

'Private Sub cmdNew_Click()
'
''-----------------Admin Check--------
'Set rs = New ADODB.Recordset
'If rs.State <> 0 Then rs.Close
''str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.TxtUName.text & "'"
'str = "select SerialNo,Passward,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='USER'"
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
'        cmdDelete.Enabled = False
'        cmdOpen.Enabled = False
'        cmdClose.Enabled = False
'        cmdPost.Enabled = False
'        cmdUndoPost.Enabled = False
''        cmdLAdd.Enabled = True
''        cmdLDelete.Enabled = True
'        cmdPrint.Enabled = False
'        cmdPreview.Enabled = False
'        TextClear Me
'        Call Clear
'
'        fgSales.Rows = 1
'        fgSales.Editable = flexEDKbdMouse
'        Call allenable
'        txtPost.text = "Not Posted"
'        txtUserName.text = frmLogin.TxtUName.text
'        txtCustomerName.SetFocus
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
'                cmdPrint.Enabled = True
'                cmdPreview.Enabled = True
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
'        cmdPrint.Enabled = False
'        cmdPreview.Enabled = False
''        cmdLAdd.Enabled = True
'        cmdLDelete.Enabled = True
''        chameleonButton1.Enabled = False
'        TextClear Me
'        Call Clear
'
'        fgSales.Rows = 1
'        fgSales.Editable = flexEDKbdMouse
'        Call allenable
'        txtPost.text = "Not Posted"
'        txtUserName.text = frmLogin.TxtUName.text
'        txtCustomerName.SetFocus
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
'                cmdPrint.Enabled = True
'                cmdPreview.Enabled = True
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
Private Sub Clear()
      txtPReg.text = ""
      PDate.Enabled = True
      PDate.Value = Date
      cmdLDelete.Enabled = True
      fgSales.Enabled = True
      txtPatientName.text = ""
      txtAge.text = ""
      cmbSex.text = ""
      txtBP.text = ""
      txtWeight.text = ""
      txtothers.text = ""
      txtInvestigation.text = ""
'     cmdAddItem.Enabled = True

End Sub

Private Sub allenable()
     txtPReg.Enabled = False
     PDate.Value = Date
     cmdLDelete.Enabled = True
     fgSales.Enabled = True
     txtPatientName.Enabled = True
     txtAge.Enabled = True
     cmbSex.Enabled = True
     txtRefer_Code.Enabled = True
     txtDoc_Name.Enabled = True
     txtBP.Enabled = True
     txtWeight.Enabled = True
     txtInvestigation.Enabled = True
     txtothers.Enabled = True
     txtTime.Enabled = True
     cmdAddItem.Enabled = True
    End Sub
'
Private Sub alldisable()
     txtPReg.Enabled = False
     PDate.Enabled = False
     PDate.Value = Date
     cmdLDelete.Enabled = False
     fgSales.Enabled = False
     txtPatientName.Enabled = False
     txtRefer_Code.Enabled = False
     txtDoc_Name.Enabled = False
     txtAge.Enabled = False
     txtTime.Enabled = False
     cmbSex.Enabled = False
     txtBP.Enabled = False
     txtWeight.Enabled = False
     txtothers.Enabled = False
     txtInvestigation.Enabled = False
     cmdAddItem.Enabled = False


End Sub
'
'
'Private Sub CmdPost_Click()
'Dim s As String
'cmdPost.Caption = "&Posting"
'fgSales.Editable = flexEDKbdMouse
''Call postedCheck
'
'
'If cmdPost.Caption = "&Posting" Then
'     If txtPost.text = "Not Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 cmdNew.Caption = "&New"
'                 cmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 fgSales.Enabled = False
'                 cmdOpen.Enabled = True
''                 chameleonButton1.Enabled = True
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

'Private Sub cmdPreview_Click()
'Tracer = 0
'    Call printReport
'End Sub
'
'Public Sub printReport()
'
'On Error GoTo ErrH
'    Dim strPath    As String
'    Dim strSQL     As String
'    Dim temp       As Double
'    If rsPrescriptionMaster.RecordCount = 0 Then
'        MsgBox "Data not available", vbInformation, "Confarmation"
'        Exit Sub
'    End If
'
'
'        strPath = App.Path + "\reports\Prescription.rpt"
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
'              strSQL = " SELECT Prescription.SerialNo, Prescription.PReg, Prescription.PName, Prescription.Age, Prescription.Sex, " & _
'                      "Prescription.Time, Prescription.PDate,Prescription.Ref_Code, Prescription.RefdName, Prescription.BP, " & _
'                      "Prescription.Weight, Prescription.others, Prescription.Investigation, Prescription.id,Prescription.MCatagory, " & _
'                     "Prescription.MName , Prescription.Qty, Prescription.Cpost, Prescription.UserName, " & _
'                     "FROM Prescription where " & _
'                      "Prescription.txtSerialNo='" & Me.txtSerialNo & "'"



'
'
''            strSQL = "select from Pms_Sales_Privew where Prescription.SerialNo= '" & Me.txtSerialNo & "' "
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
'
'        If Tracer = 0 Then
'         objReport.Preview "Prescription Preview Report", , , , , 16777216 Or 524288 Or 65536
'        Else
'        objReport.PrintOut
'        End If
'
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
'      Case 20545
'            MsgBox "Request cancelled by the user", vbInformation, "Printing Cancel Information"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Item catagory Report"
'    End Select
'End Sub
'
'
'Private Sub cmdUndoPost_Click()
'Dim s As String
'cmdUndoPost.Caption = "&Undo Posting"
'fgSales.Editable = flexEDKbdMouse
''Call postedCheck
'
'Call Calculation
'If cmdUndoPost.Caption = "&Undo Posting" Then
'Call Calculation
'     If txtPost.text = "Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 cmdNew.Caption = "&New"
'                 cmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 fgSales.Enabled = False
'                 cmdOpen.Enabled = True
''                 chameleonButton1.Enabled = True
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
'cmdUndoPost.Caption = "&Modify"
'
'End Sub
'
''---------------------------------------------------------------------------
Private Sub Form_Load()
Call Connect
     ModFunction.StartUpPosition Me
     txtUsername.text = frmLogin.txtUName.text

'       Call alldisable


   Set rsPrescriptionMaster = New ADODB.Recordset

  If rsPrescriptionMaster.State <> 0 Then rsPrescriptionMaster.Close
'     rsPrescriptionMaster.Open "select * FROM Prescription", cn, adOpenStatic, adLockReadOnly
rsPrescriptionMaster.Open "exec Sms_Sales_Master", cn, adOpenStatic, adLockReadOnly
  If rsPrescriptionMaster.RecordCount > 0 Then
      rsPrescriptionMaster.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
    txtpost.text = "Not Posted"




'    If Not rsPrescriptionMaster.EOF Then FindRecord

'-----------------For Record Search----------
Adodc2.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc2.CommandType = adCmdTable
  Adodc2.RecordSource = "Prescription"

  Adodc2.Refresh
  CmbBloodGroup.AddItem "A Positive"
  CmbBloodGroup.AddItem "A Negetive"
  CmbBloodGroup.AddItem "B Positive"
  CmbBloodGroup.AddItem "B Negetive"
  CmbBloodGroup.AddItem "O Positive"
  CmbBloodGroup.AddItem "O Negetive"
  CmbBloodGroup.AddItem "AB Positive"
  CmbBloodGroup.AddItem "AB Negetive"
  
'-------------------End Record Search---------
Call DeleteVisible
Call ModifyVisible
End Sub
'
'Private Sub FindRecord()
'
'    Dim i As Integer
'    Dim strPaymentDetail As String
'    Set rsItemDetail = New ADODB.Recordset
'    txtSerialNo = rsPrescriptionMaster!SerialNo
'    txtPReg = rsPrescriptionMaster!PReg
'    txtPatientName = rsPrescriptionMaster!PName
'    PDate = rsPrescriptionMaster!PDate
''    txtPReg = rsPrescriptionMaster!RegNo
'    txtAge = rsPrescriptionMaster!Age
'    cmbSex = rsPrescriptionMaster!Sex
'    txtBP = rsPrescriptionMaster!BP
'    txtWeight = rsPrescriptionMaster!Weight
'    txtothers = rsPrescriptionMaster!txtothers
'    txtInvestigation = rsPrescriptionMaster!Investigation
'    txtTime = rsPrescriptionMaster!Ttime
'    txtPost = rsPrescriptionMaster!Posted
'    txtUserName = rsPrescriptionMaster!UserName
'
'
'    fgSales.Rows = 1
''    strPaymentDetail = "SELECT  SerialNo, DTPSalesDate, SupplierName,ProductCatagory ,SubCode,ItemName,Quentity, " & _
''                "Rate,Amount,Rol,ExpDate,Posted,Warrenty,Remarks,ConPpsted,Unit FROM PurchaseDetail " & _
''                "WHERE SerialNo='" & parseQuotes(Me.txtSerialNo) & "' order by SerialNo "
'
'strPaymentDetail = "Exec  Sms_SalesDetails_Find1 '" & parseQuotes(Me.txtPReg) & "'"
'
'    rsItemDetail.CursorLocation = adUseClient
'    rsItemDetail.Open strPaymentDetail, cn, adOpenStatic, adLockReadOnly
'
'
'    If rsItemDetail.RecordCount <> 0 Then
'
'        fgSales.Rows = rsItemDetail.RecordCount + 1
''                i = 0
'        For i = 1 To rsItemDetail.RecordCount
'            fgSales.TextMatrix(i, 1) = rsItemDetail("SerialNo")
'            fgSales.TextMatrix(i, 2) = rsItemDetail("MCatagory")
'            fgSales.TextMatrix(i, 3) = rsItemDetail("MName")
'            fgSales.TextMatrix(i, 4) = rsItemDetail("Qty")
'            fgSales.TextMatrix(i, 5) = rsItemDetail("Remarks")
'
'
'
'        rsItemDetail.MoveNext
'        Next
'      End If
'        rsItemDetail.Close
'End Sub
'
'
Private Sub ModifyVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UName,Passward,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
           If rs!Name = "ADMIN" And rs!Passward = "123" Then
              cmdUndoPost.Visible = True

        ElseIf rs!Name = "BORHAN" And rs!Passward = "01920468031" Then

              cmdUndoPost.Visible = True
           Else
               cmdUndoPost.Visible = False

           End If
End Sub
'
Private Sub DeleteVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UName,Passward,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
           If rs!Name = "ADMIN" And rs!Passward = "123" Then
              cmdUndoPost.Visible = True

        ElseIf rs!Name = "BORHAN" And rs!Passward = "01920468031" Then

              cmdUndoPost.Visible = True
           Else
               cmdUndoPost.Visible = False

           End If
End Sub

Private Sub Timer1_Timer()
       txtTime.text = Time$
End Sub


