VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RptStockPosition 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medicine Stock Statement "
   ClientHeight    =   3720
   ClientLeft      =   1380
   ClientTop       =   930
   ClientWidth     =   6345
   Icon            =   "RptStockPosition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAllItem 
      BackColor       =   &H00C0B4A9&
      Caption         =   "All Item"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ComboBox cmbSName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5010
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtLPRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Insert your Medicine Sales Rate."
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtSRate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   2250
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtCatagory 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2055
   End
   Begin VB.ComboBox cmbMedicineName 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5985
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2055
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   2880
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
            Picture         =   "RptStockPosition.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptStockPosition.frx":0E64
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptStockPosition.frx":173E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbEO 
      Height          =   600
      Left            =   4200
      TabIndex        =   4
      Top             =   2880
      Width           =   1890
      _ExtentX        =   3334
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
   Begin VB.Label lblBalance 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5010
      TabIndex        =   16
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblMedicineName 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Medicine Name"
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
      TabIndex        =   15
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label lblSalesRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Sales Rate"
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
      Left            =   2250
      TabIndex        =   14
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblLPRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Last P. Rate"
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
      Left            =   3600
      TabIndex        =   13
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblSName 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Supplier Name"
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
      TabIndex        =   12
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblMedicineCatagory 
      BackColor       =   &H00C0B4A9&
      Caption         =   "MedicineCatagory"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Discount (%)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label lblIWSSales 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Medicine Stock Position Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "RptStockPosition"
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
Private bRecordExists              As Boolean


Private Sub chkAllItem_Click()
If chkAllItem.Value = 1 Then
cmbMedicineName.Enabled = False
Else
cmbMedicineName.Enabled = True
End If
End Sub

Private Sub cmbMedicineName_DropDown()
Call MedicineName
End Sub

Private Sub cmbMedicineName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMedicineName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbSName_DropDown()
Call allClear
End Sub

Private Sub cmbSName_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(cmbSName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub cmbMedicineName_LostFocus()
Call Others
Call LPRate
Call Balance
End Sub

Private Sub Form_Load()
    Call Connect
    Call SName
     Call MedicineName
    ModFunction.StartUpPosition Me

End Sub

Private Sub SName()
     
     Dim rsTemp2 As New ADODB.Recordset
     
     
     rsTemp2.Open ("SELECT DISTINCT SName FROM Suppliers ORDER BY SName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbSName.AddItem rsTemp2("SName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
     
End Sub

Private Sub MedicineName()
     
cmbMedicineName.Clear
Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MName, SName FROM tblMedicineName where SName='" & parseQuotes(cmbSName.text) & "'ORDER BY MName ASC"), cn, adOpenStatic
    While Not rsTemp.EOF
    cmbMedicineName.AddItem rsTemp("MName")
    rsTemp.MoveNext
    Wend
    rsTemp.Close
    
'    Call Others
     
End Sub

Private Sub LPRate()

On Error Resume Next
Dim rsTemp3 As New ADODB.Recordset
rsTemp3.Open ("SELECT SerialNo,Mname,PRate FROM PurchaseDetail where Mname='" & parseQuotes(cmbMedicineName.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic

If rsTemp3.RecordCount > 0 Then
txtLPRate = rsTemp3!PRate

End If
rsTemp3.Close
End Sub

Private Sub Others()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MCatagory,OBalance,SRate,Discount FROM tblMedicineName where MName='" & parseQuotes(cmbMedicineName.text) & "'"), cn, adOpenStatic

txtSRate = rsTemp!SRate
txtDiscount = rsTemp!Discount
txtCatagory = rsTemp!MCatagory
    
    rsTemp.Close
End Sub

Private Sub Balance()
Dim rsBalance As New ADODB.Recordset
If rsBalance.State <> 0 Then rsBalance.Close

rsBalance.Open ("select (isnull(sum(PurchaseDetail.Qty),0))-(select isnull(sum(SalesDetail.Qty),0) " & _
                "from SalesDetail where SalesDetail.Posted='Posted' and SalesDetail.Mname='" & parseQuotes(cmbMedicineName.text) & "' ) " & _
                "as Balance from PurchaseDetail  where PurchaseDetail.Mname='" & parseQuotes(cmbMedicineName.text) & "'  and  PurchaseDetail.Posted='Posted'"), cn, adOpenStatic


If rsBalance.RecordCount > 0 Then
'      rsTemp1.MoveFirst
        bRecordExists = True
    Else
        bRecordExists = False
    End If
txtBalance = rsBalance!Balance

If rsBalance!Balance < 0 Then
txtBalance = 0
End If


End Sub

Private Sub allClear()
txtBalance = ""
txtLPRate = ""
cmbMedicineName = ""
txtCatagory = ""
txtSRate = ""
End Sub

Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
     Case "Preview"
If chkAllItem.Value = 1 Then
                Tracer = 0
                Call FetchData
                Call previewReport
               End If
     Case "Print"
If chkAllItem.Value = 1 Then
                Tracer = 1
                Call FetchData
                Call previewReport
               End If
     Case "Close"
               Unload Me
    End Select

End Sub

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function


Public Function FetchData()

    Set rsMaster = New ADODB.Recordset
    
    rsMaster.Open "exec Stock_Status1", cn, adOpenStatic, adLockReadOnly

    
                  
End Function

Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Stock Position.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   

      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Medicine Stock Statement", , , , , 16777216 Or 524288 Or 65536
    
      
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
            MsgBox "Request cancelled by the user", vbInformation, "Medicine Stock Statement Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Sales Summery Information Report"
    End Select
End Sub





