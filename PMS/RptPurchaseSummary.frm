VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form RptPurchaseSummary 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Catagory Wise Purchase Report"
   ClientHeight    =   4200
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   6315
   Icon            =   "RptPurchaseSummary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6315
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   600
      Width           =   6255
      Begin MSForms.ComboBox cmbMCatagory 
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Top             =   240
         Width           =   4575
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "8070;661"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Catagory Name"
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
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame FraDateSelect 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Select Date"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   6255
      Begin VB.CheckBox chkAllSupplier 
         BackColor       =   &H00C0B4A9&
         Caption         =   "All Catagory"
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
         Left            =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OpCurrentDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&rrent Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   120
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton OpCustomDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Cu&stom Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   1920
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
               Picture         =   "RptPurchaseSummary.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptPurchaseSummary.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptPurchaseSummary.frx":173E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   3480
         TabIndex        =   5
         Top             =   1920
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
      Begin MSComCtl2.DTPicker FDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   64684035
         CurrentDate     =   38258
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   64684035
         CurrentDate     =   38258
      End
      Begin MSComCtl2.DTPicker CurrentDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   64684035
         CurrentDate     =   38258
      End
      Begin VB.Label lblTo 
         BackColor       =   &H00C0B4A9&
         Caption         =   "To Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   3720
         TabIndex        =   9
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblFrom 
         BackColor       =   &H00C0B4A9&
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Catagory Wise Purchase Report"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6375
   End
End
Attribute VB_Name = "RptPurchaseSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsMaster                            As ADODB.Recordset
Private rsSelect                            As ADODB.Recordset 'sub

Private rsTemp2                             As ADODB.Recordset
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

Private Sub chkAllSupplier_Click()
If chkAllSupplier.Value = 1 Then
cmbMCatagory.Enabled = False
Else
cmbMCatagory.Enabled = True
End If
End Sub


Private Sub Form_Load()
Call Connect
    Call CatagoryName
    ModFunction.StartUpPosition Me
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    CurrentDate.Value = Date
    TDate.Value = Date
    FDate.Value = Date
End Sub
Private Sub CatagoryName()
     Dim rsTemp2 As New ADODB.Recordset

     rsTemp2.Open ("SELECT DISTINCT MCatagory FROM MedicineCatagory ORDER BY MCatagory ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbMCatagory.AddItem rsTemp2("MCatagory")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
End Sub

Private Sub OpCurrentDate_Click()
    OpCurrentDate.Visible = True
    OpCustomDate.Visible = True
    CurrentDate.Visible = True
    lblFrom.Visible = False
'    Frame1.Visible = False
    FDate.Visible = False
    lblTo.Visible = False
    TDate.Visible = False
End Sub
Private Sub OpCustomDate_Click()
    OpCustomDate.Visible = True
    OpCurrentDate.Visible = True
    lblFrom.Visible = True
    lblTo.Visible = True
'    Frame1.Visible = True
    CurrentDate.Visible = False
    FDate.Visible = True
    TDate.Visible = True
End Sub

Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
     Case "Preview"
            If Validate Then
            If chkAllSupplier.Value = 1 Then
                Tracer = 0
                Call FetchData1
                Call previewReport1
                Else
                Call FetchData
            Call previewReport
                    End If
               End If
     Case "Print"
            If Validate Then
            If chkAllSupplier.Value = 1 Then
                Tracer = 1
                Call FetchData1
                Call previewReport1
                Else
                Call FetchData
            Call previewReport
                    End If
               End If
     Case "Close"
               Unload Me
    End Select
End Sub
Private Function Validate() As Boolean
           Validate = True
        If FDate.Value > TDate.Value Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
            FDate.SetFocus
            Validate = False
            Exit Function
        End If
    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function


Public Function FetchData()

    Set rsMaster = New ADODB.Recordset
    
    If OpCurrentDate.Value = True Then
                 
    
    rsMaster.Open "SELECT     SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, " & _
                    "Qty, PRate, Amount,SRate,Posted " & _
                    "From PurchaseDetail " & _
                    "WHERE    Pdate='" & CurrentDate.Value & "' AND Posted = 'Posted' and PurchaseDetail.MCatagory='" & parseQuotes(cmbMCatagory) & "'", cn, adOpenStatic, adLockReadOnly
     
     
     
     End If
             
      If OpCustomDate.Value = True Then
      
                                             
     rsMaster.Open "SELECT     SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, " & _
                    "Qty, PRate, Amount,SRate, Posted " & _
                    "From PurchaseDetail " & _
                    "WHERE  Pdate BETWEEN '" & FDate.Value & "' AND " & _
                    "'" & TDate.Value & "' AND Posted = 'Posted' and PurchaseDetail.MCatagory='" & parseQuotes(cmbMCatagory) & "'", cn, adOpenStatic, adLockReadOnly
      
      End If

End Function


Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Purchase Catagpry Details.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   
    If OpCurrentDate.Value = True Then
   
   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + Format(CurrentDate, "dd-MMM-yyyy") + "'"
   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbMCatagory + "'"

              
  End If
  
  
  If OpCustomDate.Value = True Then
  
   Set objReportFF = objReportFormulaFieldDefinations.Item(2)

              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"
             
   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
              objReportFF.text = "'" + cmbMCatagory + "'"

             
   End If
      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Purchase Statement Report", , , , , 16777216 Or 524288 Or 65536
    
      
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
            MsgBox "Request cancelled by the user", vbInformation, "Purchase Summery Statement"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Purchase Summery Statement"
    End Select
End Sub

Public Function FetchData1()

    Set rsMaster = New ADODB.Recordset
    
    If OpCurrentDate.Value = True Then
                 
    
    rsMaster.Open "SELECT     SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, " & _
                    "Qty, PRate, Amount,SRate,Posted " & _
                    "From PurchaseDetail " & _
                    "WHERE    Pdate='" & CurrentDate.Value & "' AND Posted = 'Posted' order by SerialNo", cn, adOpenStatic, adLockReadOnly
     
     
     
     End If
             
      If OpCustomDate.Value = True Then
      
                                             
     rsMaster.Open "SELECT     SerialNo, Sname, Pdate, SbillNo, MCatagory, Mname, " & _
                    "Qty, PRate, Amount,SRate, Posted " & _
                    "From PurchaseDetail " & _
                    "WHERE  Pdate BETWEEN '" & FDate.Value & "' AND " & _
                    "'" & TDate.Value & "' AND Posted = 'Posted' order by SerialNo", cn, adOpenStatic, adLockReadOnly
      
      End If

End Function


Public Sub previewReport1()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Purchase Statement Details.rpt"
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
        objReport.Preview "Purchase Statement Report", , , , , 16777216 Or 524288 Or 65536
    
      
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
            MsgBox "Request cancelled by the user", vbInformation, "Purchase Summery Statement"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Purchase Summery Statement"
    End Select
End Sub
