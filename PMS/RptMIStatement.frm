VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form RptMIStatement 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Medicine Income Statement "
   ClientHeight    =   3450
   ClientLeft      =   1380
   ClientTop       =   930
   ClientWidth     =   5970
   Icon            =   "RptMIStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2895
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   5895
      Begin VB.CheckBox chkDetails 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Details Income Statement"
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
         TabIndex        =   10
         Top             =   2040
         Width           =   2775
      End
      Begin VB.OptionButton OpCurrentDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Current Date"
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
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton OpCustomDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Custom Date"
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
         Left            =   120
         TabIndex        =   0
         Top             =   960
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5160
         Top             =   1680
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
               Picture         =   "RptMIStatement.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptMIStatement.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptMIStatement.frx":173E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   3240
         TabIndex        =   3
         Top             =   1920
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
      Begin MSComCtl2.DTPicker CurrentDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67895299
         CurrentDate     =   38258
      End
      Begin MSComCtl2.DTPicker FDate 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67895299
         CurrentDate     =   38258
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   1200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   67895299
         CurrentDate     =   38258
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
         Left            =   1920
         TabIndex        =   9
         Top             =   960
         Width           =   1575
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
         Left            =   3840
         TabIndex        =   8
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medicine Sales Statement          "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0FF&
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5985
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
    CurrentDate.Value = Date
    FDate.Value = Date
    TDate.Value = Date
'        Call LoadcmbBankName
        OpCustomDate.Value = True
        lblFrom.Visible = True
        lblTo.Visible = True
        FDate.Visible = True
        TDate.Visible = True
        
        
    End Sub

Private Sub OpCurrentDate_Click()
        lblFrom.Visible = False
        lblTo.Visible = False
        FDate.Visible = False
        TDate.Visible = False
        CurrentDate.Visible = True
'        Label1.Visible = True
End Sub

Private Sub OpCustomDate_Click()
        lblFrom.Visible = True
        lblTo.Visible = True
        FDate.Visible = True
        TDate.Visible = True
        CurrentDate.Visible = False
'        Label1.Visible = False
End Sub

Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)


Select Case Button.Key
     Case "Preview"
            If Validate Then
            If chkDetails.Value = 1 Then
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
            If chkDetails.Value = 1 Then
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

If OpCustomDate.Value = True Then
              
 Set rsMaster = New ADODB.Recordset
    
    rsMaster.Open "exec Medicine_Income '','','','','','','','','" + Format(FDate, "dd-MMM-yyyy") + "','" + Format(TDate, "dd-MMM-yyyy") + "'", cn, adOpenStatic, adLockReadOnly
                                            
      
      End If
                  
End Function

Public Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Medicine Income Ledger.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
        

  If OpCustomDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(1)

              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"

   End If

   
      
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Medicine Income Insformations", , , , , 16777216 Or 524288 Or 65536
    
      
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
            MsgBox "Request cancelled by the user", vbInformation, "Medicine Income Insformations"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Medicine Income Insformations"
    End Select
End Sub

Public Function FetchData1()

Set rsMaster = New ADODB.Recordset
            
      If OpCustomDate.Value = True Then
              
      rsMaster.Open "SELECT CMID, SDate, CName, CAddress, RegNo,BedNo, Tamount, DiscuntPer, DiscuntTaka, Ttime, Tpaid, Tdue, Advance, Posted, UName, Admitted, Remark " & _
                    "From SalesMaster " & _
                    "WHERE  SDate BETWEEN '" & FDate.Value & "' AND " & _
                    "'" & TDate.Value & "' AND Posted = 'Posted' order by CMID", cn, adOpenStatic, adLockReadOnly
    
      End If

End Function

Public Sub previewReport1()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Medicine Income Statement.rpt"
'        strPath = App.Path + "\reports\Medicine Income Ledger.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
         
         If OpCustomDate.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(1)

              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"

   End If
         
        objReportDatabaseTable.SetPrivateData 3, rsMaster
    
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        
        objReport.DiscardSavedData
        objReport.Preview "Medicine Sales Statement Details", , , , , 16777216 Or 524288 Or 65536
    
      
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
            MsgBox "Request cancelled by the user", vbInformation, "Medicine Sales Statement Details"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Medicine Sales Statement Details"
    End Select
End Sub


