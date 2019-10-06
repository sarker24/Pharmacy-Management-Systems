VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form RptROL 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Reorder Level Statement"
   ClientHeight    =   2670
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   4560
   Icon            =   "RptROL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4560
   Begin VB.Frame FraDateSelect 
      BackColor       =   &H00C0B4A9&
      Caption         =   "View ROL Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      Begin VB.CheckBox chkAll 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Select All Medicine"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   4065
      End
      Begin VB.ComboBox cmbMedicineName 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4065
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   4065
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1800
      Top             =   1800
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
            Picture         =   "RptROL.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptROL.frx":08E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptROL.frx":11C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbEO 
      Height          =   600
      Left            =   2520
      TabIndex        =   4
      Top             =   2040
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
   Begin MSAdodcLib.Adodc dcSupplierName 
      Height          =   360
      Left            =   0
      Top             =   2280
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   635
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
      Caption         =   "dcSupplierName"
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
Attribute VB_Name = "RptROL"
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

Private Sub cmbMedicineName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMedicineName, KeyAscii)
   If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub chkAll_Click()
If chkAll.Value = 1 Then
cmbMedicineName.Enabled = False
Else
cmbMedicineName.Enabled = True
End If
End Sub

Private Sub Form_Load()
    Call Connect
    ModFunction.StartUpPosition Me
    Call MedicineName
'    OpCurrentDate.Visible = True
'    OpCustomDate.Visible = True
'    SSCurrentDate.Date = Date
'    SSTDate.Date = Date
            
End Sub

Private Sub MedicineName()

Dim rsTemp2 As New ADODB.Recordset
     
     
     rsTemp2.Open ("SELECT DISTINCT MedicineName FROM tblMedicineName ORDER BY MedicineName ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbMedicineName.AddItem rsTemp2("MedicineName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

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
  End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Public Function FetchData()

    Set rsMaster = New ADODB.Recordset
     
   If chkAll.Value = 1 Then
    rsMaster.Open " select tunion.Mname,isnull((select sum(tt1.Qty) from PurchaseDetail " & _
                  " tt1 where tt1.Mname=tunion.Mname and tt1.Posted='Posted' ),0) " & _
                  "-isnull((select sum(tt2.Qty) from SalesDetail tt2 where tt2.Mname=tunion.Mname and tt2.Posted='Posted' ),0) " & _
                  "as Qty,Rol=(select distinct ROL from tblMedicineName where tblMedicineName.MedicineName=tunion.Mname)" & _
                  "from (select distinct PurchaseDetail.Mname,PurchaseDetail.Catagory from PurchaseDetail union " & _
                  "all select SalesDetail.Mname,SalesDetail.MCatagory from SalesDetail)  as tunion " & _
                  "Where isnull((select sum(tt1.Qty) from PurchaseDetail tt1 where tt1.Mname=tunion.Mname " & _
                  "and tt1.Posted='Posted' ),0) -isnull((select sum(tt2.Qty) from SalesDetail tt2 where tt2.Mname=tunion.Mname " & _
                  "and tt2.Posted='Posted' ),0)<(select distinct ROL from tblMedicineName where tblMedicineName.MedicineName=tunion.Mname) " & _
                  "group by tunion.Mname", cn, adOpenStatic, adLockReadOnly
                  
    Else
            rsMaster.Open " select tunion.Mname,isnull((select sum(tt1.Qty) from PurchaseDetail  " & _
                  " tt1 where tt1.Mname=tunion.Mname and tt1.Posted='Posted' ),0) " & _
                  "-isnull((select sum(tt2.Qty) from SalesDetail tt2 where tt2.Mname=tunion.Mname and tt2.Posted='Posted' ),0) " & _
                  "as Qty,Rol=(select distinct ROL from tblMedicineName where tblMedicineName.MedicineName=tunion.Mname) " & _
                  "from (select distinct PurchaseDetail.Mname,PurchaseDetail.Catagory from PurchaseDetail union " & _
                  "all select SalesDetail.Mname,SalesDetail.MCatagory from SalesDetail)  as tunion " & _
                  "Where isnull((select sum(tt1.Qty) from PurchaseDetail tt1 where tt1.Mname=tunion.Mname " & _
                  "and tt1.Posted='Posted' ),0) -isnull((select sum(tt2.Qty) from SalesDetail tt2 where tt2.Mname=tunion.Mname " & _
                  "and tt2.Posted='Posted' ),0)<(select distinct ROL from tblMedicineName where tblMedicineName.MedicineName=tunion.Mname) " & _
                  "and tunion.Mname='" & cmbMedicineName.text & "' group by tunion.Mname", cn, adOpenStatic, adLockReadOnly
                  
      End If
                  
End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Medicine Auto ROL.rpt"
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
            MsgBox "Request cancelled by the user", vbInformation, "Sales Summery Information Report"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Sales Summery Information Report"
    End Select
End Sub

