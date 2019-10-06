VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FrmStockQty 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "FrmStockQty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Total Stock"
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
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStockQty.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStockQty.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmStockQty.frx":173E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc dcBank 
      Height          =   450
      Left            =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   794
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
      BackColor       =   16777215
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
      Caption         =   "dcBank"
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cmbBankName 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1560
      Width           =   2535
      _Version        =   196616
      RowHeight       =   423
      Columns(0).Width=   3200
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.Toolbar tbEO 
      Height          =   600
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Preview"
            Object.ToolTipText     =   "Preview"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Print"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Close"
            Object.ToolTipText     =   "Close"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Product  Name"
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
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label58 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STOCK QUANTITY  "
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
      TabIndex        =   0
      Top             =   0
      Width           =   4545
   End
End
Attribute VB_Name = "FrmStockQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsMaster                            As ADODB.Recordset
Private rsSelect                            As ADODB.Recordset 'sub

Private objReportApp                        As CRPEAuto.Application
Private objReport                           As CRPEAuto.Report
Private objReportDatabase                   As CRPEAuto.Database
Private objReportDatabaseTables             As CRPEAuto.DatabaseTables
Private objReportDatabaseTable              As CRPEAuto.DatabaseTable
Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
Private Tracer                              As Integer
Private strGroupName                        As String

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Call Connect
    ModFunction.StartUpPosition Me
'        DTPTo.Value = Date
        Call LoadcmbBankName
        Check1.Value = 0
    End Sub

Private Sub LoadcmbBankName()
    dcBank.CursorLocation = adUseClient
    dcBank.ConnectionString = cn.ConnectionString
    dcBank.LockType = adLockReadOnly
    dcBank.RecordSource = "SELECT Name FROM USStockDetail order by Name"
    cmbBankName.DataMode = ssDataModeBound
    Set cmbBankName.DataSource = dcBank
    cmbBankName.DataSourceList = dcBank
    cmbBankName.DataFieldList = "Name"
'    cmbBankName.DataField = "Name"
    cmbBankName.ColumnHeaders = True
    cmbBankName.BackColorOdd = &HFFFF00
    cmbBankName.BackColorEven = &HFFC0C0
    cmbBankName.ForeColorEven = &H80000008
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
cmbBankName.Enabled = False
Else
cmbBankName.Enabled = True
End If
End Sub

Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
'            If Validate Then
                Tracer = 0
                Call FetchData
                Call previewReport
'               End If
     Case "Print"
'            If Validate Then
                Tracer = 1
                Call FetchData
                Call previewReport
'               End If
     Case "Close"
               Unload Me
    End Select
End Sub
'Private Function Validate() As Boolean
''           Validate = True
''        If DTPFrom.Value > DTPTo.Value Then
''            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
''            DTPFrom.SetFocus
''            Validate = False
''            Exit Function
''        End If
'    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Public Function FetchData()

    Set rsMaster = New ADODB.Recordset
    
    If Check1.Value = 1 Then
                 
                  
    
     rsMaster.Open "SELECT USStockDetail.SerialNo[SerialNo],USStockMaster.PurchaseDate[PurchaseDate]," & _
                "USStockDetail.Name[Name],USStockDetail.UnitPrice, " & _
                "USStockDetail.Qty-(ISNULL((select sum(USSalesDetail.IssueQty) " & _
                "from USSalesDetail Where USSalesDetail.Did = USStockDetail.Did),0)) as availableQty " & _
                ",(USStockDetail.UnitPrice*(USStockDetail.Qty-(ISNULL((select sum(USSalesDetail.IssueQty) " & _
                "from USSalesDetail Where USSalesDetail.Did = USStockDetail.Did),0))))as Totalprice " & _
                "From USStockMaster, USStockDetail " & _
                "Where USStockMaster.SerialNo = USStockDetail.SerialNo " & _
                "AND USStockDetail.Qty-(ISNULL((select sum(USSalesDetail.IssueQty) " & _
                "from USSalesDetail Where USSalesDetail.Did = USStockDetail.Did),0))!=0", cn, adOpenStatic, adLockReadOnly
                
     Else

    rsMaster.Open "SELECT USStockDetail.Did[DocID],USStockDetail.SerialNo[SerialNo]," & _
                "USStockDetail.Name[Name],USStockDetail.UnitPrice, " & _
                "USStockDetail.Qty-(ISNULL((select sum(USSalesDetail.IssueQty) " & _
                "from USSalesDetail Where USSalesDetail.Did = USStockDetail.Did),0)) as availableQty " & _
                ",(USStockDetail.UnitPrice*(USStockDetail.Qty-(ISNULL((select sum(USSalesDetail.IssueQty) " & _
                "from USSalesDetail Where USSalesDetail.Did = USStockDetail.Did),0))))as Totalprice " & _
                "From USStockMaster, USStockDetail " & _
                "Where USStockMaster.SerialNo = USStockDetail.SerialNo " & _
                "AND USStockDetail.Qty-(ISNULL((select sum(USSalesDetail.IssueQty) " & _
                "from USSalesDetail Where USSalesDetail.Did = USStockDetail.Did),0))!=0 and " & _
                "USStockDetail.Name='" & parseQuotes(cmbBankName.text) & "'", cn, adOpenStatic, adLockReadOnly
                
   End If
   


' rsMaster.Open "SELECT USLedgerMaster.SerialNo, USLedgerMaster.PartyName," & _
'               "USLedgerDetail.Date, USLedgerDetail.Description, " & _
'               "USLedgerDetail.Folio, USLedgerDetail.Debit, " & _
'               "USLedgerDetail.Credit , USLedgerDetail.Total FROM " & _
'               "USLedgerMaster,USLedgerDetail WHERE " & _
'               "USLedgerMaster.SerialNo = USLedgerDetail.SerialNo and " & _
'               "USLedgerMaster.PartyName = '" & parseQuotes(cmbBankName.text) & "' " & _
'               "and USLedgerDetail.Date BETWEEN '" & DTPFrom.Value & "' AND '" & DTPTo.Value & "' " & _
'               "order by USLedgerDetail.LedgerID", cn, adOpenStatic, adLockReadOnly


End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\StockQty.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   
'   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'              objReportFF.text = "'" + Format(DTPFrom, "dd-MMM-yyyy") + "'"
'         Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'             objReportFF.text = "'" + Format(DTPTo, "dd-MMM-yyyy") + "'"


      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Stock Quantity", , , , , 16777216 Or 524288 Or 65536
    
      
     If Tracer = 1 Then
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
            MsgBox "Request cancelled by the user", vbInformation, "Bank Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Bank Information Report"
    End Select
End Sub








