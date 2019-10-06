VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSupplier 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Informations"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10605
   DrawMode        =   1  'Blackness
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7800
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0B4A9&
      Height          =   4095
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   10455
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
         Height          =   3735
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   10215
         _cx             =   18018
         _cy             =   6588
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSupplier.frx":058A
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   "Supplier Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   -120
      Width           =   10695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   120
      TabIndex        =   9
      Top             =   600
      Width           =   10425
      Begin VB.TextBox txtContactPerson 
         Height          =   375
         Left            =   6720
         TabIndex        =   2
         Top             =   480
         Width           =   3495
      End
      Begin VB.TextBox txtCountry 
         Height          =   345
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   7
         Text            =   " "
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtCity 
         Height          =   345
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   6
         Text            =   " "
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtSupplierID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtSEmail 
         Height          =   345
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   8
         Text            =   " "
         Top             =   2415
         Width           =   8775
      End
      Begin VB.TextBox txtFaxNo 
         Height          =   345
         Left            =   6240
         MaxLength       =   50
         TabIndex        =   5
         Text            =   " "
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtPhoneNo 
         Height          =   345
         Left            =   1440
         MaxLength       =   50
         TabIndex        =   4
         Text            =   " "
         Top             =   1560
         Width           =   3855
      End
      Begin VB.TextBox txtAddress 
         Height          =   555
         Left            =   1440
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   3
         Text            =   "frmSupplier.frx":066F
         Top             =   960
         Width           =   8775
      End
      Begin VB.TextBox txtSName 
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   5115
      End
      Begin VB.Label lblCity 
         BackColor       =   &H00C0B4A9&
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   1515
      End
      Begin VB.Label lblCountry 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         TabIndex        =   19
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label lblContactPerson 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Contact Person"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6720
         TabIndex        =   18
         Top             =   240
         Width           =   2955
      End
      Begin VB.Label lblName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier &Name"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   4995
      End
      Begin VB.Label lbiCompanyID 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier ID"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblEmail 
         BackColor       =   &H00C0B4A9&
         Caption         =   "E-&mail"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   2445
         Width           =   1515
      End
      Begin VB.Label lblFax 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Fax"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5400
         TabIndex        =   13
         Top             =   1560
         Width           =   1155
      End
      Begin VB.Label lblPhone 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Phone"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblAddr 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Address"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
   End
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   615
      Left            =   7440
      TabIndex        =   21
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdEdit 
      Height          =   615
      Left            =   1440
      TabIndex        =   24
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdDelete 
      Height          =   615
      Left            =   8640
      TabIndex        =   25
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdCancel 
      Height          =   615
      Left            =   3840
      TabIndex        =   26
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdFind 
      Height          =   615
      Left            =   2640
      TabIndex        =   27
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   615
      Left            =   5040
      TabIndex        =   28
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPreview 
      Height          =   615
      Left            =   6240
      TabIndex        =   29
      Top             =   8400
      Width           =   1215
      _Version        =   196612
      _ExtentX        =   2143
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
      WordWrap        =   0   'False
      PictureAlignment=   0
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3960
      Top             =   7920
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
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Private rsImage               As ADODB.Recordset
Private rsfactory             As ADODB.Recordset
'Private strStream             As ADODB.Stream
Private strFileName           As String
'Private bEditMode             As USER_MODE
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim str As String
'--------------------------------------------------------------
Private oReportApp                        As CRPEAuto.Application
Private oReport                           As CRPEAuto.Report
Private oReportDatabase                   As CRPEAuto.Database
Private oReportDatabaseTables             As CRPEAuto.DatabaseTables
Private oReportDatabaseTable              As CRPEAuto.DatabaseTable
'Private oReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private oReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                 As CRPEAuto.PrintWindowOptions

Private Sub cmdCancel_AfterClick()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    CmdDelete.Enabled = True
    cmdQuit.Enabled = True
    cmdEdit.Enabled = True
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
    cmdFind.Enabled = True
    txtSupplierID.Enabled = False
       
    Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub

Private Sub allenable()
    txtSName.Enabled = True
    txtAddress.Enabled = True
    txtContactPerson.Enabled = True
    txtCity.Enabled = True
    txtCountry.Enabled = True
    txtPhoneNo.Enabled = True
    txtFaxNo.Enabled = True
    txtSEmail.Enabled = True
End Sub

Private Sub alldisable()
    txtSName.Enabled = False
    txtAddress.Enabled = False
    txtContactPerson.Enabled = False
    txtCity.Enabled = False
    txtCountry.Enabled = False
    txtPhoneNo.Enabled = False
    txtFaxNo.Enabled = False
    txtSEmail.Enabled = False
End Sub

Private Sub allClear()
    ModFunction.TextClear Me
End Sub

Private Sub cmdDelete_AfterClick()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If idelete = vbYes Then
            cn.Execute "Delete From Suppliers Where SupplierID ='" & parseQuotes(txtSupplierID) & "'"
            Call allClear
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdEdit_AfterClick()
 If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdQuit.Enabled = False
        CmdDelete.Enabled = False
        cmdFind.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = True
                cmdQuit.Enabled = True
                CmdDelete.Enabled = True
                cmdFind.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                cmdQuit.Enabled = True
                Call alldisable
                rsfactory.Requery

                Dim s As String
                s = txtSName
                rsfactory.Find "SName='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub cmdFind_AfterClick()
frmSupplierSearch.Show vbModal
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdFirst.Enabled = False
 Else
        cmdFirst.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdPrevious.Enabled = True
        txtSupplierID = Adodc1.Recordset!SupplierID
        txtSName = Adodc1.Recordset!SName
        txtContactPerson = Adodc1.Recordset!ContactName
        txtAddress = Adodc1.Recordset!Address
        txtPhoneNo = Adodc1.Recordset!PhoneNo
        txtFaxNo = Adodc1.Recordset!FaxNo
        txtCity = Adodc1.Recordset!City
        txtCountry = Adodc1.Recordset!Country
        txtSEmail = Adodc1.Recordset!Email

End If
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
        cmdFirst.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdPrevious.Enabled = True
        txtSupplierID = Adodc1.Recordset!SupplierID
        txtSName = Adodc1.Recordset!SName
        txtContactPerson = Adodc1.Recordset!ContactName
        txtAddress = Adodc1.Recordset!Address
        txtPhoneNo = Adodc1.Recordset!PhoneNo
        txtFaxNo = Adodc1.Recordset!FaxNo
        txtCity = Adodc1.Recordset!City
        txtCountry = Adodc1.Recordset!Country
        txtSEmail = Adodc1.Recordset!Email

End If
End Sub

Private Sub cmdNew_AfterClick()
On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdQuit.Enabled = False
        CmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdFind.Enabled = False
        cmdPreview.Enabled = False
        
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SupplierID),0) as SerialNo from Suppliers"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSupplierID.text = Val(rs!SerialNo) + 1
            
        Call allenable
        txtSName.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSupplierID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdQuit.Enabled = False
                CmdDelete.Enabled = True
                cmdPrint.Enabled = True
                cmdFind.Enabled = True
                cmdPreview.Enabled = True
                cmdQuit.Enabled = True
                Call alldisable
                s = txtSName
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "SName='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
'
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True

    If (txtSName.text = "") Then
       MsgBox "Enter Supplier Name"
       txtSName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtAddress.text = "") Then
      MsgBox "Enter Supplier Address"
      txtAddress.SetFocus
      IsValidRecord = False
      Exit Function
    End If
    
If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from Suppliers where upper(SName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtSName))) & "'", cn

             If Not rsfactory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtSName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
End Function

Private Function rcupdate() As Boolean

'    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO Suppliers(SupplierID,SName,ContactName,Address, " & _
                   " PhoneNo,FaxNo,City,Country,Email) " & _
                   " VALUES ('" & parseQuotes(txtSupplierID) & "','" & parseQuotes(txtSName) & "','" & parseQuotes(txtContactPerson) & "', " & _
                   " '" & parseQuotes(txtAddress) & "','" & parseQuotes(txtPhoneNo) & "', " & _
                   " '" & parseQuotes(txtFaxNo) & "','" & parseQuotes(txtCity) & "', " & _
                   " '" & parseQuotes(txtCountry) & "','" & parseQuotes(txtSEmail) & "') "
                   


          rcupdate = True
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update Suppliers Set SName='" & parseQuotes(txtSName) & _
                  "',ContactName='" & parseQuotes(txtContactPerson) & "',Address='" & parseQuotes(txtAddress) & "', " & _
                  " PhoneNo='" & parseQuotes(txtPhoneNo) & _
                  "',FaxNo='" & parseQuotes(txtFaxNo) & "', " & _
                  " City='" & parseQuotes(txtCity) & _
                  "',Country='" & parseQuotes(txtCountry) & "',Email='" & parseQuotes(txtSEmail) & "' " & _
                  " Where SupplierID ='" & parseQuotes(txtSupplierID) & "' "
                  
                 

        rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function



'ErrHandler:
'    cn.RollbackTrans
'   ' rsFactory.Requery
'    Select Case cn.Errors(0).NativeError
'        Case 2627
'            MsgBox "Trying with duplicate CNF Name"
'            txtname = ""
'            txtname.SetFocus
'        Case Else
'            MsgBox Err.Number & " : " & Err.Description
'    End Select

End Function

Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtSupplierID = rsfactory("SupplierID")
        txtSName = rsfactory("SName")
        txtContactPerson = rsfactory("ContactName")
        txtAddress = rsfactory("Address")
        txtPhoneNo = rsfactory("PhoneNo")
        txtFaxNo = rsfactory("FaxNo")
        txtCity = rsfactory("City") & ""
        txtCountry = IIf(IsNull(rsfactory("Country")), "", rsfactory("Country"))
        txtSEmail = IIf(IsNull(rsfactory("Email")), "", rsfactory("Email"))
    End If
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdNext.Enabled = False
 Else
        cmdFirst.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdPrevious.Enabled = True
        txtSupplierID = Adodc1.Recordset!SupplierID
        txtSName = Adodc1.Recordset!SName
        txtContactPerson = Adodc1.Recordset!ContactName
        txtAddress = Adodc1.Recordset!Address
        txtPhoneNo = Adodc1.Recordset!PhoneNo
        txtFaxNo = Adodc1.Recordset!FaxNo
        txtCity = Adodc1.Recordset!City
        txtCountry = Adodc1.Recordset!Country
        txtSEmail = Adodc1.Recordset!Email

End If
End Sub

Private Sub cmdPreview_AfterClick()
    Call printReport
End Sub

Public Sub printReport()
'On Error GoTo ErrorHan
Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\SupplierInformation.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select Suppliers.SupplierID,Suppliers.SName,Suppliers.ContactName,Suppliers.Address,Suppliers.PhoneNo, " & _
             "  " & _
             "Suppliers.FaxNo,Suppliers.City,Suppliers.Country,Suppliers.Email " & _
             "from Suppliers where " & _
             "Suppliers.SupplierID='" & Me.txtSupplierID & "'"

    rsFactProf.Open strSQL, cn, adOpenStatic, adLockReadOnly

    oReportDatabaseTable.SetPrivateData 3, rsFactProf

ObjPrinterSetting.HasPrintSetupButton = True
ObjPrinterSetting.HasRefreshButton = True
ObjPrinterSetting.HasSearchButton = True
ObjPrinterSetting.HasZoomControl = True

'      Set oReportFormulaFieldDefinations = oReport.FormulaFields
'      Set oReportFF = oReportFormulaFieldDefinations.Item(1)
'      oReportFF.text = "'Factory Information'"

oReport.DiscardSavedData
oReport.Preview "Supplier Infromation of '" & txtSName.text & "'", , , , , 16777216 Or 524288 Or 65536

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
'          MsgBox "end of file"
       cmdPrevious.Enabled = False
 Else
        cmdFirst.Enabled = True
        cmdNext.Enabled = True
        cmdLast.Enabled = True
        cmdPrevious.Enabled = True
        txtSupplierID = Adodc1.Recordset!SupplierID
        txtSName = Adodc1.Recordset!SName
        txtContactPerson = Adodc1.Recordset!ContactName
        txtAddress = Adodc1.Recordset!Address
        txtPhoneNo = Adodc1.Recordset!PhoneNo
        txtFaxNo = Adodc1.Recordset!FaxNo
        txtCity = Adodc1.Recordset!City
        txtCountry = Adodc1.Recordset!Country
        txtSEmail = Adodc1.Recordset!Email

End If
End Sub

Private Sub cmdQuit_AfterClick()
Unload Me
End Sub

Public Sub PopulateCnf(StrID As String)


    rsfactory.MoveFirst
    rsfactory.Find "SupplierID=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

Private Sub Form_Load()
Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from Suppliers", cn, adOpenStatic, adLockReadOnly
    Call alldisable
    Call DeleteVisible
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
   
    If Not rsfactory.EOF Then FindRecord
    
    txtSupplierID.Enabled = False
    '-----------------For Record Search----------
Adodc1.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "Suppliers"

  Adodc1.Refresh
'-------------------End Record Search---------

End Sub

Private Sub DeleteVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UName,Password,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
           If rs!Name = "ADMIN" Then
'              cmdUndoPost.Visible = True
            
'        ElseIf rs!Name = "BORHAN" And rs!Password = "0711227051" Then
        
              CmdDelete.Visible = True
'           Else
'               cmdUndoPost.Visible = False
               
           End If
End Sub
