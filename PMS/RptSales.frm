VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptMIStatement 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
   Icon            =   "RptSales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   8
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20250627
      CurrentDate     =   38716
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Date To Date Sale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific Date Sale"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1320
      Top             =   2400
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
            Picture         =   "RptSales.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptSales.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptSales.frx":1A7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbEO 
      Height          =   600
      Left            =   1920
      TabIndex        =   0
      Top             =   2520
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
   Begin MSComCtl2.DTPicker DTPTo 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20250627
      CurrentDate     =   38476
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   375
      Left            =   720
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   20250627
      CurrentDate     =   38278
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Specific Date"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "End"
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
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "From"
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
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   495
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
      TabIndex        =   1
      Top             =   0
      Width           =   6705
   End
End
Attribute VB_Name = "RptMIStatement"
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
'    Call Connect
    ModFunction.StartUpPosition Me
        DTPTo.Value = Date
'        Call LoadcmbBankName
        Option1.Value = True
        lblDate.Visible = False
        Label2.Visible = False
        DTPFrom.Visible = False
        DTPTo.Visible = False
        
        
    End Sub

'Private Sub LoadcmbBankName()
'    dcBank.CursorLocation = adUseClient
'    dcBank.ConnectionString = cn.ConnectionString
'    dcBank.LockType = adLockReadOnly
'
'    dcBank.RecordSource = "SELECT Name FROM USStockDetail order by Name"
'
'    cmbBankName.DataMode = ssDataModeBound
'    Set cmbBankName.DataSource = dcBank
'    cmbBankName.DataSourceList = dcBank
'    cmbBankName.DataFieldList = "Name"
'    cmbBankName.DataField = "Name"
'    cmbBankName.ColumnHeaders = True
'    cmbBankName.BackColorOdd = &HFFFF00
'    cmbBankName.BackColorEven = &HFFC0C0
'    cmbBankName.ForeColorEven = &H80000008
'End Sub

'Private Sub Check1_Click()
'If Check1.Value = 1 Then
'cmbBankName.Enabled = False
'Else
'cmbBankName.Enabled = True
'End If
'End Sub

Private Sub Option1_Click()
        lblDate.Visible = False
        Label2.Visible = False
        DTPFrom.Visible = False
        DTPTo.Visible = False
        DTPicker1.Visible = True
        Label1.Visible = True
End Sub

Private Sub Option2_Click()
        lblDate.Visible = True
        Label2.Visible = True
        DTPFrom.Visible = True
        DTPTo.Visible = True
        DTPicker1.Visible = False
        Label1.Visible = False
End Sub

Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
            If Validate Then
                Tracer = 0
                Call FetchData
                Call previewReport
               End If
     Case "Print"
            If Validate Then
                Tracer = 1
                Call FetchData
                Call previewReport
               End If
     Case "Close"
               Unload Me
    End Select
End Sub
Private Function Validate() As Boolean
           Validate = True
        If DTPFrom.Value > DTPTo.Value Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
            DTPFrom.SetFocus
            Validate = False
            Exit Function
        End If
    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Public Function FetchData()

    Set rsMaster = New ADODB.Recordset
    
    If Option1.Value = True Then
                 
            rsMaster.Open "SELECT USSalesMaster.InvoiceNo, USSalesDetail.Name, " & _
                         "USSalesDetail.IssueQty, USSalesDetail.UnitPrice, " & _
                         "USSalesDetail.Total, USSalesMaster.Date, " & _
                         "USSalesMaster.SoldBy , USSalesDetail.Remarks " & _
                         "FROM USSalesMaster INNER JOIN " & _
                         "USSalesDetail ON " & _
                         "USSalesMaster.InvoiceNo = USSalesDetail.InvoiceNo where " & _
                         "USSalesMaster.Date='" & DTPicker1.Value & "'", cn, adOpenStatic, adLockReadOnly
            
     End If
             
      If Option2.Value = True Then
      
         rsMaster.Open "SELECT USSalesMaster.InvoiceNo, USSalesDetail.Name, " & _
                         "USSalesDetail.IssueQty, USSalesDetail.UnitPrice, " & _
                         "USSalesDetail.Total, USSalesMaster.Date, " & _
                         "USSalesMaster.SoldBy , USSalesDetail.Remarks " & _
                         "FROM USSalesMaster INNER JOIN " & _
                         "USSalesDetail ON " & _
                         "USSalesMaster.InvoiceNo = USSalesDetail.InvoiceNo " & _
                         "Where USSalesMaster.Date BETWEEN '" & DTPFrom.Value & "' AND " & _
                         "'" & DTPTo.Value & "' ", cn, adOpenStatic, adLockReadOnly
              
                                             
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

    
        strPath = App.Path + "\reports\RptSale.rpt"
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
        objReport.Preview "Sales Insformations", , , , , 16777216 Or 524288 Or 65536
    
      
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










