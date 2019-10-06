VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Object = "{8D650141-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3b32.ocx"
Begin VB.Form frmPurchaseSearch 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   6315
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatagorySearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin SSDataWidgets_A.SSDBCommand SSDBCommand1 
      Height          =   375
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6375
      _Version        =   196612
      _ExtentX        =   11245
      _ExtentY        =   661
      _StockProps     =   78
      Caption         =   "Select Medicine Item"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.74
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      BorderStyle     =   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdSupplierName 
      Height          =   615
      Left            =   3720
      TabIndex        =   17
      Top             =   3600
      Width           =   2535
      _Version        =   196612
      _ExtentX        =   4471
      _ExtentY        =   1085
      _StockProps     =   78
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
   Begin SSCalendarWidgets_A.SSDateCombo DTPExpiry 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   3600
      Width           =   3375
      _Version        =   65537
      _ExtentX        =   5953
      _ExtentY        =   873
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "-"
   End
   Begin VB.TextBox txtSalesRate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox txtRate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.TextBox txtAmount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   5
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1695
   End
   Begin SSDataWidgets_B.SSDBCombo cmbMedicineName 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
      _Version        =   196616
      DataMode        =   2
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   32768
      Columns(0).Width=   3200
      _ExtentX        =   10610
      _ExtentY        =   873
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SSCalendarWidgets_A.SSDateCombo DTPManufacture 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   4440
      Width           =   3375
      _Version        =   65537
      _ExtentX        =   5953
      _ExtentY        =   873
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DropDownFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "-"
   End
   Begin SSDataWidgets_A.SSDBCommand cmdKo 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   4200
      TabIndex        =   15
      Top             =   4560
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
   End
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   615
      Left            =   5160
      TabIndex        =   16
      Top             =   4560
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Quit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
   End
   Begin VB.Label lblSupplierName 
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
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   3360
      Width           =   2535
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
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Date of Manufacture"
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
      TabIndex        =   12
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Date of Expiry "
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
      TabIndex        =   10
      Top             =   3360
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Last Purchase Rate"
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
      Left            =   3960
      TabIndex        =   9
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblSalesRate 
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Rate"
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Amount"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblQty 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Quantity"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "frmPurchaseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private rsTemp                      As ADODB.Recordset
'Private rsExport                    As ADODB.Recordset
'Private rsfactory                   As New ADODB.Recordset
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdOK_Click()
'    If fgExport.RowSel < 0 Then
'        MsgBox "Please Select a Catagory From the List."
'        Exit Sub
'    End If
'
'     Call PopulateCompanySearch
'
'     Unload Me
'     Set frmPCatagorySearch = Nothing
'End Sub
'
'Private Sub fgExport_DblClick()
'    cmdOK_Click
'End Sub
'
'Private Sub Form_Load()
'     ModFunction.StartUpPosition Me
'     Set rsTemp = New ADODB.Recordset
'     rsTemp.CursorLocation = adUseClient
'
'        If rsTemp.State <> 0 Then rsTemp.Close
'
'       If strCallingForm = LCase("frmCountry") Then
' '     rsTemp.CursorLocation = adUseClient
'       rsTemp.Open "SELECT DISTINCT PCatagory,CatagoryID FROM USCatagory order by CatagoryID ", cn, adOpenStatic, adLockReadOnly
'
'     End If
'
'         fgExport.Rows = 1
'
'        If strCallingForm = LCase("frmCountry") Then
'             Label58.Caption = "Product Catagory Search"
'              While Not rsTemp.EOF
'            fgExport.AddItem "" & vbTab & rsTemp("CatagoryID") & vbTab & rsTemp("PCatagory")
'            rsTemp.MoveNext
'        Wend
'        GridCount fgExport
'     End If
'
''     If fgExport.Rows = 1 Then fgExport.AddItem ""
'
'End Sub
'
'    Private Sub PopulateCompanySearch()
'    If fgExport.Row > 0 Then
'             frmCatagory.PopulateProduct fgExport.TextMatrix(fgExport.Row, 1)
'          End If
'
'   End Sub
'
'Private Sub cmdFind_Click()
'Set rsTemp = New ADODB.Recordset
'
'    If rsTemp.State <> 0 Then rsTemp.Close
'
'
'     rsTemp.Open "SELECT USCatagory.CatagoryID[CatagoryID],USCatagory.PCatagory[PCatagory]" & _
'                 " FROM USCatagory WHERE USCatagory.PCatagory LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
'                 fgExport.Rows = 1
'
'    While Not rsTemp.EOF
'        fgExport.AddItem "" & vbTab & rsTemp("CatagoryID") & vbTab & rsTemp("PCatagory")
'
'        rsTemp.MoveNext
'    Wend
'     GridCount fgExport
'
'End Sub
'
'
'Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
'
'End Sub
'
'Private Sub txtSearch_Change()
'cmdFind_Click
'End Sub
'

Private Sub cmdQuit_AfterClick()
Unload Me
End Sub
