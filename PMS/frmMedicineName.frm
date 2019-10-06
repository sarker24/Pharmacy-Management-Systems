VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmMedicineName 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Name Entry [Doctors Clinic Unit - 2]"
   ClientHeight    =   5955
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   10095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMedicineName.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   10095
   Begin VB.CommandButton cmdStock 
      Caption         =   "Stock View"
      Height          =   495
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame cboCatagory 
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
      TabIndex        =   19
      Top             =   480
      Width           =   9825
      Begin VB.TextBox txtDiscount 
         Height          =   420
         Left            =   8520
         TabIndex        =   7
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cmbGenericName 
         Height          =   420
         Left            =   5880
         TabIndex        =   1
         Top             =   600
         Width           =   3735
      End
      Begin VB.ComboBox cmbMCatagory 
         Height          =   420
         Left            =   240
         TabIndex        =   2
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtOBalance 
         Height          =   420
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   1665
      End
      Begin VB.TextBox txtOBRate 
         Height          =   420
         Left            =   1920
         TabIndex        =   11
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtUnit 
         Height          =   420
         Left            =   7320
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtROL 
         Height          =   420
         Left            =   5880
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtMName 
         Height          =   420
         Left            =   1920
         TabIndex        =   0
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox txtMID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   420
         Left            =   240
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox txtSRate 
         Height          =   420
         Left            =   7320
         TabIndex        =   6
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cmbMSName 
         Height          =   420
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTPOBDate 
         Height          =   420
         Left            =   4200
         TabIndex        =   12
         Top             =   2280
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   741
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   49938435
         CurrentDate     =   40388
      End
      Begin VB.Label lblDiscount 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Discount(%)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   8520
         TabIndex        =   40
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Opening Balance Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4200
         TabIndex        =   31
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblOBRate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Opening Balance Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   30
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Label lblOpeningBalance 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Opening Balance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblUnit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7320
         TabIndex        =   28
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label lblROL 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Reorder Level"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5880
         TabIndex        =   27
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Supplier Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   26
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label lblMedicineName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Medicine Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   25
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label lblMedicineID 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine  ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblMedicineCatagory 
         BackColor       =   &H00C0B4A9&
         Caption         =   "M. Catagory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Generic Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5880
         TabIndex        =   22
         Top             =   360
         Width           =   3735
      End
      Begin VB.Label lblSalesrate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Sales Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   7320
         TabIndex        =   21
         Top             =   2040
         Width           =   1095
      End
   End
   Begin VB.TextBox txtUName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   495
      Left            =   5880
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   " "
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox txtpost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4800
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4800
      Width           =   735
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPost 
      Height          =   615
      Left            =   6840
      TabIndex        =   9
      Top             =   5280
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
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   615
      Left            =   240
      TabIndex        =   8
      Top             =   5280
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
      Left            =   1200
      TabIndex        =   32
      Top             =   5280
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
      Left            =   7800
      TabIndex        =   33
      Top             =   5280
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
      Left            =   5970
      TabIndex        =   34
      Top             =   5280
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
      Left            =   2160
      TabIndex        =   35
      Top             =   5280
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
   Begin SSDataWidgets_A.SSDBCommand cmdClose 
      Height          =   615
      Left            =   3090
      TabIndex        =   36
      Top             =   5280
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
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   615
      Left            =   5040
      TabIndex        =   37
      Top             =   5280
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
      Font3D          =   4
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPreview 
      Height          =   615
      Left            =   4080
      TabIndex        =   38
      Top             =   5280
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4200
      Top             =   4200
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
   Begin VB.Label lblMNameshow 
      Alignment       =   2  'Center
      BackColor       =   &H00004000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   10095
   End
End
Attribute VB_Name = "frmMedicineName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsMedicineMaster      As ADODB.Recordset
Private rsMCatagory           As ADODB.Recordset
Private rsMSName       As ADODB.Recordset


Private rsTemp2               As ADODB.Recordset
Private rs                    As ADODB.Recordset
Private strStream             As ADODB.Stream
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Dim strMood                             As String
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



Private Sub cmbGenericName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbGenericName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbMCatagory_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMCatagory, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub
Private Sub cmbMSName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMSName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

'Private Sub cmdChange_Click()
'Dim s As String
'cmdChange.Caption = "&Modify"
'
'If cmdChange.Caption = "&Modify" Then
'     If txtpost.text = "Posted" Then
'        If IsValidRecord Then
'            If rcupdate Then
'                 cmdNew.Caption = "&New"
'                 cmdEdit.Enabled = True
'                 cmdCancel.Enabled = False
'                 cmdClose.Enabled = True
'                 cmdOpen.Enabled = True
'                 cmdPreview.Enabled = True
'                 cmdDelete.Enabled = True
'                 Call alldisable
'           End If
'        End If
'      End If
'Else
'
' End If
'cmdChange.Caption = "&Change"
'End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    cmdCancel.Enabled = False
    cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdClose.Enabled = True
    cmdEdit.Enabled = True
    cmdOpen.Enabled = True
    cmdDelete.Enabled = True
    cmdPreview.Enabled = True
    cmdPrint.Enabled = True
    cmdPost.Enabled = True
'    cmdChange.Enabled = True
    txtMID.Enabled = False
    Call allClear
    Call alldisable
    If Not rsMedicineMaster.EOF Then FindRecord
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
     idelete = MsgBox("Do you want to delete this record?", vbYesNo)
     If txtUName.text = "Admin" Then
    If idelete = vbYes Then
  
            cn.Execute "Delete From tblMedicineName Where MID ='" & parseQuotes(txtMID) & "'"
            Call allClear
    
    MsgBox "Please Call your System Administrator"
    End If
        
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
    
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
       txtMID = Adodc1.Recordset!Mid
        txtMName = Adodc1.Recordset!MName
        cmbGenericName = Adodc1.Recordset!GenericName
        cmbMCatagory = Adodc1.Recordset!MCatagory
        cmbMSName = Adodc1.Recordset!SName
        txtROL = Adodc1.Recordset!ROL
        txtUnit = Adodc1.Recordset!Unit
        txtOBalance = Adodc1.Recordset!OBalance
        txtOBRate = Adodc1.Recordset!OBRate
        DTPOBDate.Value = Adodc1.Recordset!OBDate
        txtPost.text = Adodc1.Recordset!CPost
        txtSRate.text = Adodc1.Recordset!SRate
End If
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True
       
       txtMID = Adodc1.Recordset!Mid
        txtMName = Adodc1.Recordset!MName
        cmbGenericName = Adodc1.Recordset!GenericName
        cmbMCatagory = Adodc1.Recordset!MCatagory
        cmbMSName = Adodc1.Recordset!SName
        txtROL = Adodc1.Recordset!ROL
        txtUnit = Adodc1.Recordset!Unit
        txtOBalance = Adodc1.Recordset!OBalance
        txtOBRate = Adodc1.Recordset!OBRate
        DTPOBDate.Value = Adodc1.Recordset!OBDate
        txtPost.text = Adodc1.Recordset!CPost
        txtSRate.text = Adodc1.Recordset!SRate
        
End If
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
    Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdClose.Enabled = False
        txtMID.Enabled = False
        cmdOpen.Enabled = False
        DTPOBDate.Value = Date
        cmdDelete.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdPost.Enabled = False
'        cmdChange.Enabled = False
        Call allClear
        txtPost.text = "Not Posted"

       Call allenable
       Call MGenericName
       Call MCatagory
       Call MSName
       txtMName.SetFocus

    ElseIf cmdNew.Caption = "&Save" Then
        Dim P As String
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                cmdNew.Caption = "&New"
                cmdCancel.Enabled = True
                cmdClose.Enabled = True
                txtMID.Enabled = True
                cmdOpen.Enabled = True
                cmdDelete.Enabled = True
                cmdPrint.Enabled = True
                cmdPreview.Enabled = True
                cmdPost.Enabled = True
'                cmdChange.Enabled = True
                
               Call alldisable

'                s = txtMID
'                rsMedicineMaster.Requery
'                rsMedicineMaster.MoveFirst
'                 rsMedicineMaster.Find "MID='" & parseQuotes(s) & "'"
                FindRecord
            End If
        End If
    End If
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

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
        cmdPost.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
        txtMName.SetFocus
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                cmdDelete.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
                cmdPost.Enabled = True
                Call alldisable
                rsMedicineMaster.Requery

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
        cmdPost.Enabled = False
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        cmdOpen.Enabled = False
        cmdDelete.Enabled = False
        txtMName.SetFocus
    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdClose.Enabled = True
                cmdOpen.Enabled = True
                cmdDelete.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
                cmdPost.Enabled = True
                Call alldisable
                rsMedicineMaster.Requery

                End If
            End If
        End If
    End If
'    End If
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
       
       txtMID = Adodc1.Recordset!Mid
        txtMName = Adodc1.Recordset!MName
        cmbGenericName = Adodc1.Recordset!GenericName
        cmbMCatagory = Adodc1.Recordset!MCatagory
        cmbMSName = Adodc1.Recordset!SName
        txtROL = Adodc1.Recordset!ROL
        txtUnit = Adodc1.Recordset!Unit
        txtOBalance = Adodc1.Recordset!OBalance
        txtOBRate = Adodc1.Recordset!OBRate
        DTPOBDate.Value = Adodc1.Recordset!OBDate
        txtPost.text = Adodc1.Recordset!CPost
        txtSRate.text = Adodc1.Recordset!SRate
        txtDiscount.text = Adodc1.Recordset!Discount
End If
End Sub

Private Sub cmdOpen_Click()
frmMFind.Show vbModal
End Sub

Private Sub CmdPost_Click()
Dim s As String
cmdPost.Caption = "&Posting"



If cmdPost.Caption = "&Posting" Then
     If txtPost.text = "Not Posted" Then
        If IsValidRecord Then
            If rcupdate Then
                 cmdNew.Caption = "&New"
                 cmdEdit.Enabled = False
                 cmdCancel.Enabled = False
                 cmdClose.Enabled = True
                 cmdOpen.Enabled = True
                 cmdPreview.Enabled = True
                 cmdDelete.Enabled = True
'                 cmdChange.Enabled = True
'                 txtBillMID.Enabled = False
                 Call alldisable
           End If
        End If
      End If
Else
'    cmdtemSelected.Enabled = False
'    cmdLDelete.Enabled = False
 End If
cmdPost.Caption = "&Post"
End Sub

Private Sub cmdPreview_Click()
' Call printReport
End Sub

Public Sub printReport()

Dim strPath         As String
Dim rsFactProf      As ADODB.Recordset
Dim strSQL          As String


    strPath = App.Path + "\reports\Medicine Barcode.rpt"

    Set oReportApp = CreateObject("Crystal.CRPE.Application")
    Set oReport = oReportApp.OpenReport(strPath)
    Set oReportDatabase = oReport.Database
    Set oReportDatabaseTables = oReportDatabase.Tables
    Set oReportDatabaseTable = oReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = oReport.PrintWindowOptions


    Set rsFactProf = New ADODB.Recordset
If rsFactProf.State <> 0 Then rsFactProf.Close

    strSQL = "select tblMedicineName.MID,tblMedicineName.MName,tblMedicineName.Barcode, " & _
             "from tblMedicineName where " & _
             "tblMedicineName.MID='" & Me.txtMID & "'"

'             tblMedicineName.MID ='" & Me.txtMID & "'"

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
oReport.Preview "Medicine Infromation of '" & txtMName.text & "'", , , , , 16777216 Or 524288 Or 65536

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
       
       txtMID = Adodc1.Recordset!Mid
        txtMName = Adodc1.Recordset!MName
        cmbGenericName = Adodc1.Recordset!GenericName
        cmbMCatagory = Adodc1.Recordset!MCatagory
        cmbMSName = Adodc1.Recordset!SName
        txtROL = Adodc1.Recordset!ROL
        txtUnit = Adodc1.Recordset!Unit
        txtOBalance = Adodc1.Recordset!OBalance
        txtOBRate = Adodc1.Recordset!OBRate
        DTPOBDate.Value = Adodc1.Recordset!OBDate
        txtPost.text = Adodc1.Recordset!CPost
        txtSRate.text = Adodc1.Recordset!SRate
        txtDiscount.text = Adodc1.Recordset!Discount
End If
End Sub

Private Sub cmdStock_Click()
frmMStock.Show vbModal
End Sub

Private Sub Form_Load()
    Call Connect
     ModFunction.StartUpPosition Me
       Call alldisable
       
       DTPOBDate.Value = Date
'       lblMNameshow.= txtMName
         
'       txtPost.text = "Not Posted"
'       txtCActive.text = "Active"
       
   Set rsMedicineMaster = New ADODB.Recordset
If rsMedicineMaster.State <> 0 Then rsMedicineMaster.Close

rsMedicineMaster.Open "select TOP 1 PERCENT MID, MName, GenericName, MCatagory, SName, ROL, Unit, CPost, " & _
                   "UName, OBalance, OBRate, OBDate, SRate ,Discount from tblMedicineName ORDER BY MID DESC", cn, adOpenStatic, adLockReadOnly
   
   
   If rsMedicineMaster.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If
If Not rsMedicineMaster.EOF Then FindRecord
    txtMID.Enabled = False
    txtUName.text = frmLogin.txtUName.text
     
'txtUName.text = "Murad"
'---------------
Adodc1.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "tblMedicineName"

  Adodc1.Refresh
'----------------
'ConStr = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
'Call ModifyVisible
Call DeleteVisible
End Sub

Private Sub DeleteVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UName,Password  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
'           If rs!UName = "ADMIN" And rs!Password = "01920468031" Then
            If rs!UName = "Admin" Then
              cmdDelete.Visible = True
            
        ElseIf rs!UName = "BORHAN" And rs!Password = "01920468031" Then
        
              cmdDelete.Visible = True
           Else
               cmdDelete.Visible = False
               
           End If
End Sub

Private Sub MGenericName()

Dim rsTemp2 As New ADODB.Recordset
     
     rsTemp2.Open ("SELECT DISTINCT GenericName FROM tblGenericName ORDER BY GenericName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbGenericName.AddItem rsTemp2("GenericName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub MCatagory()

Dim rsTemp2 As New ADODB.Recordset
          
     rsTemp2.Open ("SELECT DISTINCT MCatagory FROM MedicineCatagory ORDER BY MCatagory ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbMCatagory.AddItem rsTemp2("MCatagory")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub MSName()

Dim rsTemp2 As New ADODB.Recordset
          
     rsTemp2.Open ("SELECT DISTINCT SName FROM Suppliers ORDER BY SName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbMSName.AddItem rsTemp2("SName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

 Private Sub allenable()
    cmbMCatagory.Enabled = True
    txtMName.Enabled = True
    cmbGenericName.Enabled = True
    cmbMSName.Enabled = True
    cmbGenericName.Enabled = True
    txtOBalance.Enabled = True
    txtOBRate.Enabled = True
    txtROL.Enabled = True
    txtUnit.Enabled = True
    DTPOBDate.Enabled = True
    txtSRate.Enabled = True
    txtDiscount.Enabled = True
'    Record.EOFAction = adDoMoveLast
End Sub

Private Sub alldisable()
    txtMID.Enabled = False
    cmbMCatagory.Enabled = False
    txtMName.Enabled = False
    cmbGenericName.Enabled = False
    cmbMSName.Enabled = False
    cmbGenericName.Enabled = False
    txtOBalance.Enabled = False
    txtOBRate.Enabled = False
    txtSRate.Enabled = False
    txtDiscount.Enabled = False
    txtROL.Enabled = False
    txtUnit.Enabled = False
    DTPOBDate.Enabled = False
    txtUName.Enabled = False

End Sub

Private Sub allClear()
txtMName = ""
txtMID = ""
cmbGenericName = ""
cmbMCatagory = ""
cmbMSName = ""
txtOBalance = ""
txtOBRate = ""
txtROL = ""
txtUnit = ""
txtSRate = ""
txtDiscount = ""
End Sub

Private Function rcupdate() As Boolean
Dim strSQL As String
    Dim iRow As Integer
    Dim blnAlarm As Boolean
    Dim strDeliveryDate As String
    Set rs = New ADODB.Recordset
    Dim str As String
    Dim iPost

str = "select SerialNo,Password,Privilegegroup,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly

 If rs!Privilegegroup = "0" Then

 If cmdNew.Caption = "&Save" Then
    
strSQL = "INSERT INTO tblMedicineName(MName,GenericName,MCatagory,SName,ROL,Unit,CPost,UName,OBalance,OBRate,OBDate,SRate,Discount) " & _
          " VALUES ('" & parseQuotes(txtMName) & "','" & parseQuotes(cmbGenericName) & "','" & parseQuotes(cmbMCatagory) & "', " & _
          " '" & parseQuotes(cmbMSName) & "'," & parseQuotes(txtROL) & ",'" & parseQuotes(txtUnit) & "','" & parseQuotes(txtPost) & "','" & parseQuotes(txtUName) & "', " & _
          " " & CDbl(Val(txtOBalance.text)) & "," & CDbl(Val(txtOBRate.text)) & ",'" & Format(DTPOBDate, "dd-mmm-yyyy") & "'," & CDbl(Val(txtSRate.text)) & "," & CDbl(Val(txtDiscount.text)) & ")"
          
        cn.Execute strSQL
        rcupdate = True

'     -------------For primary key and foreign key relation------------
         If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(MID),1) as InvNo from tblMedicineName"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtMID = Val(rs!InvNo)
                   
           MsgBox "Record Added Successfully.", vbInformation, "Confirmation"
    
    ElseIf (cmdEdit.Caption = "&Update") Then
  
  strSQL = "Update tblMedicineName Set MName='" & parseQuotes(txtMName) & "', " & _
             "GenericName='" & parseQuotes(cmbGenericName) & "',MCatagory='" & parseQuotes(cmbMCatagory) & "', " & _
             "SName='" & parseQuotes(cmbMSName) & "',ROL=" & parseQuotes(txtROL) & ",Unit='" & parseQuotes(txtUnit) & "', " & _
             "CPost='" & parseQuotes(txtPost) & "',UName='" & parseQuotes(txtUName) & "', " & _
             "OBalance= " & Val(txtOBalance.text) & ",OBRate=" & Val(txtOBRate.text) & ",OBDate='" & Format(DTPOBDate, "dd-mmm-yyyy") & "', " & _
             "SRate=" & Val(txtSRate.text) & ",Discount=" & Val(txtDiscount.text) & "  " & _
             "WHERE  MID ='" & parseQuotes(txtMID) & "' "
                    
        cn.Execute strSQL
        rcupdate = True
        
        MsgBox "Record Updated", vbInformation, "Confirmation"
        
        
'--------------------------------Posting Information-----------------------------------------

   ElseIf cmdPost.Caption = "&Posting" Then
    
    txtPost.text = "Posted"
     
     Call allenable
   
    iPost = MsgBox("Do you want to Post this bill?", vbYesNo)

    If iPost = vbYes Then
     
strSQL = "Update tblMedicineName Set MName='" & parseQuotes(txtMName) & "', " & _
             "GenericName='" & parseQuotes(cmbGenericName) & "',MCatagory='" & parseQuotes(cmbMCatagory) & "', " & _
             "SName='" & parseQuotes(cmbMSName) & "',ROL=" & parseQuotes(txtROL) & ",Unit='" & parseQuotes(txtUnit) & "', " & _
             "CPost='" & parseQuotes(txtPost) & "',UName='" & parseQuotes(txtUName) & "', " & _
             "OBalance= " & Val(txtOBalance.text) & ",OBRate=" & Val(txtOBRate.text) & ",OBDate='" & Format(DTPOBDate, "dd-mmm-yyyy") & "', " & _
             "SRate=" & Val(txtSRate.text) & ",Discount=" & Val(txtDiscount.text) & "  " & _
             "WHERE  MID ='" & parseQuotes(txtMID) & "' "

            cn.Execute strSQL
            rcupdate = True
            
            MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
             End If
     End If
 Else
 
'----------------For Admin--------------------

If cmdNew.Caption = "&Save" Then
    
 strSQL = "INSERT INTO tblMedicineName(MName,GenericName,MCatagory,SName,ROL,Unit,CPost,UName,OBalance,OBRate,OBDate,SRate,Discount) " & _
          " VALUES ('" & parseQuotes(txtMName) & "','" & parseQuotes(cmbGenericName) & "','" & parseQuotes(cmbMCatagory) & "', " & _
          " '" & parseQuotes(cmbMSName) & "'," & parseQuotes(txtROL) & ",'" & parseQuotes(txtUnit) & "','" & parseQuotes(txtPost) & "','" & parseQuotes(txtUName) & "', " & _
          " " & CDbl(Val(txtOBalance.text)) & "," & CDbl(Val(txtOBRate.text)) & ",'" & Format(DTPOBDate, "dd-mmm-yyyy") & "'," & CDbl(Val(txtSRate.text)) & "," & CDbl(Val(txtDiscount.text)) & ")"
                            
         cn.Execute strSQL
         rcupdate = True
       
       If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(MID),1) as InvNo from tblMedicineName"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtMID = Val(rs!InvNo)

          MsgBox "Record Added", vbInformation, "Confirmation"
            
    ElseIf (cmdEdit.Caption = "&Update") Then
    
strSQL = "Update tblMedicineName Set MName='" & parseQuotes(txtMName) & "', " & _
             "GenericName='" & parseQuotes(cmbGenericName) & "',MCatagory='" & parseQuotes(cmbMCatagory) & "', " & _
             "SName='" & parseQuotes(cmbMSName) & "',ROL=" & parseQuotes(txtROL) & ",Unit='" & parseQuotes(txtUnit) & "', " & _
             "CPost='" & parseQuotes(txtPost) & "',UName='" & parseQuotes(txtUName) & "', " & _
             "OBalance= " & Val(txtOBalance.text) & ",OBRate=" & Val(txtOBRate.text) & ",OBDate='" & Format(DTPOBDate, "dd-mmm-yyyy") & "', " & _
             "SRate=" & Val(txtSRate.text) & ",Discount=" & Val(txtDiscount.text) & "  " & _
             "WHERE  MID ='" & parseQuotes(txtMID) & "' "
               
              
             cn.Execute strSQL
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
        
        '  --------------------------------Posting Information-----------------------------------------

   ElseIf cmdPost.Caption = "&Posting" Then
   
    txtPost.text = "Posted"
    iPost = MsgBox("Do you want to Post this bill?", vbYesNo)
    
           If iPost = vbYes Then
            
strSQL = "Update tblMedicineName Set MName='" & parseQuotes(txtMName) & "', " & _
             "GenericName='" & parseQuotes(cmbGenericName) & "',MCatagory='" & parseQuotes(cmbMCatagory) & "', " & _
             "SName='" & parseQuotes(cmbMSName) & "',ROL=" & parseQuotes(txtROL) & ",Unit='" & parseQuotes(txtUnit) & "', " & _
             "CPost='" & parseQuotes(txtPost) & "',UName='" & parseQuotes(txtUName) & "', " & _
             "OBalance= " & Val(txtOBalance.text) & ",OBRate=" & Val(txtOBRate.text) & ",OBDate='" & Format(DTPOBDate, "dd-mmm-yyyy") & "', " & _
             "SRate=" & Val(txtSRate.text) & ",Discount=" & Val(txtDiscount.text) & "  " & _
             "WHERE  MID ='" & parseQuotes(txtMID) & "' "
                 
             cn.Execute strSQL
             rcupdate = True
             
             MsgBox "Record Posted Successfully", vbInformation, "Confirmation"
        
             End If

     End If
        
  End If
      
      Exit Function
ErrHandler:
    cn.RollbackTrans
   Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Medicine Name"
            txtMName = ""
            txtMName.SetFocus
'            txtMName = ""
'            txtMName.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select


End Function
Public Sub FindRecord()
If Not rsMedicineMaster.EOF Then
        txtMID = rsMedicineMaster("MID")
        txtMName = rsMedicineMaster("MName")
        cmbGenericName = rsMedicineMaster("GenericName")
        cmbMCatagory = rsMedicineMaster("MCatagory")
        cmbMSName = rsMedicineMaster("SName")
        txtROL = rsMedicineMaster("ROL")
        txtUnit = rsMedicineMaster("Unit")
        txtOBalance = rsMedicineMaster("OBalance")
        txtOBRate = rsMedicineMaster("OBRate")
        DTPOBDate.Value = rsMedicineMaster!OBDate
        txtPost.text = rsMedicineMaster!CPost
        txtSRate.text = rsMedicineMaster!SRate
        txtDiscount.text = rsMedicineMaster!Discount

   End If
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    If (txtMName.text = "") Then
       MsgBox "Enter Valid Medicine Name"
       txtMName.SetFocus
       IsValidRecord = False
       Exit Function
    End If
    If (cmbGenericName.text = "") Then
    MsgBox "Enter Medicine Generic Name"
    cmbGenericName.SetFocus
    IsValidRecord = False
        Exit Function
    End If
    If (cmbMSName.text = "") Then
        MsgBox "Enter Valid Supplier Name"
        cmbMSName.SetFocus
        IsValidRecord = False
        Exit Function
      End If
      If cmdNew.Caption = "&Save" Or cmdNew.Caption = "&Save" Then
'        If rsMedicineMaster.RecordCount > 0 Then
        If rsMedicineMaster.State <> 0 Then rsMedicineMaster.Close
            rsMedicineMaster.Open "select * from tblMedicineName where upper(MName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtMName))) & "'", cn
             If Not rsMedicineMaster.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtMName.SetFocus
          IsValidRecord = False
         Exit Function
            End If
'         End If
    End If
'      If cmdEdit.Caption <> "&Update" Then
'          If rsMedicineMaster.RecordCount <> 0 Then
'                If (UCase(txtProductDescription.text) = UCase(rsMedicineMaster!YarnDescription)) Then
'                      MsgBox "Trying Duplicate Yarn Description"
'                      IsValidRecord = False
'                      Exit Function
'                End If
'          End If
'      End If
End Function

Public Sub PopulateProduct(StrID As String)
    rsMedicineMaster.Close
'    rsMedicineMaster.Open "exec Sms_Medicine_Name_Search 2", cn, adOpenStatic, adLockReadOnly
    rsMedicineMaster.Open "select * from tblMedicineName", cn, adOpenStatic, adLockReadOnly
    rsMedicineMaster.MoveFirst
    rsMedicineMaster.Find "MID=" & parseQuotes(StrID)
    If rsMedicineMaster.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtMName_Change()
'    Barcod1.Caption = txtMName
    lblMNameshow.Caption = txtMName.text
End Sub

Private Sub txtMName_GotFocus()
txtMName.BackColor = &HFFFFC0
End Sub

Private Sub txtMName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtMName_LostFocus()
    txtMName.BackColor = vbWhite
    txtMName.text = StrConv(txtMName.text, vbProperCase)
End Sub

Private Sub cmbGenericName_GotFocus()
cmbGenericName.BackColor = &HFFFFC0
End Sub

Private Sub cmbGenericName_LostFocus()
cmbGenericName.BackColor = vbWhite
cmbGenericName.text = StrConv(cmbGenericName.text, vbProperCase)
End Sub

Private Sub cmbMCatagory_GotFocus()
cmbMCatagory.BackColor = &HFFFFC0
End Sub

Private Sub cmbMCatagory_LostFocus()
cmbMCatagory.BackColor = &HFFFFC0
cmbMCatagory.text = StrConv(cmbMCatagory.text, vbProperCase)
End Sub

Private Sub cmbMSName_LostFocus()
    cmbMSName.BackColor = vbWhite
    cmbMSName.text = StrConv(cmbMSName.text, vbProperCase)
End Sub

Private Sub cmbMSName_GotFocus()
cmbMSName.BackColor = &HFFFFC0
End Sub

Private Sub txtROL_GotFocus()
txtROL.BackColor = vbWhite
End Sub

Private Sub txtROL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtROL_LostFocus()
txtROL.BackColor = vbWhite
End Sub

Private Sub txtSRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtunit_GotFocus()
txtUnit.BackColor = &HFFFFC0
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub txtunit_LostFocus()
    txtUnit.BackColor = vbWhite
End Sub

Private Sub txtoBalance_GotFocus()
txtOBalance.BackColor = &HFFFFC0
End Sub

Private Sub txtoBalance_LostFocus()
txtOBalance.BackColor = vbWhite
End Sub

Private Sub txtoBRate_GotFocus()
txtOBRate.BackColor = &HFFFFC0
End Sub

Private Sub txtoBRate_LostFocus()
txtOBRate.BackColor = &HFFFFC0
End Sub

Private Sub DTPOBDate_GotFocus()
    DTPOBDate.CalendarBackColor = vbWhite
End Sub

Private Sub DTPOBDate_LostFocus()
    DTPOBDate.CalendarBackColor = vbWhite
End Sub


Private Sub txtSRate_GotFocus()
txtSRate.BackColor = &HFFFFC0
End Sub

Private Sub txtSRate_LostFocus()
txtSRate.BackColor = vbWhite
End Sub

