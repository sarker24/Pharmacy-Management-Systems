VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptMSSwithProfit 
   BackColor       =   &H00C0B4A9&
   ClientHeight    =   3375
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Palatino Linotype"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "RptProfit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5490
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
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   5295
      Begin VB.CheckBox chkDetails 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Profit Loss Statement Detail"
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
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   2160
         Top             =   2040
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
               Picture         =   "RptProfit.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptProfit.frx":0E64
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "RptProfit.frx":173E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   3000
         TabIndex        =   2
         Top             =   1800
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
         Left            =   480
         TabIndex        =   6
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   20774915
         CurrentDate     =   40571
      End
      Begin MSComCtl2.DTPicker TDate 
         Height          =   375
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   20774915
         CurrentDate     =   41071
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
         Left            =   3120
         TabIndex        =   4
         Top             =   480
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
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medicine Profit Statement "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5505
   End
End
Attribute VB_Name = "RptMSSwithProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsMaster                            As ADODB.Recordset
'Private rsSelect                            As ADODB.Recordset 'sub

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

Private Sub cmdPrint_Click()
     Tracer = 1
    Call FetchData
    Call previewReport
End Sub

Private Sub cmdPrivew_Click()
    Call FetchData
    Call previewReport
    
End Sub

Private Sub Form_Load()
'    Call Connect
    ModFunction.StartUpPosition Me
'        CurrentDate.Value = Date
        TDate.Value = Date
        FDate.Value = Date

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
            MsgBox "Invalid Date Range", vbInformation, "Party Wise Sample Report"
            FDate.SetFocus
            Validate = False
            Exit Function
        End If
    End Function
    
Private Sub FetchData()

    Set rsMaster = New ADODB.Recordset
                                             
     rsMaster.Open "SELECT SerialNo,CMID, SName,MCatagory, MName, Qty, SRate, Amount,PRate, Discount,(Amount * Discount / 100) AS TDiscount, (Amount - (Amount * Discount / 100)) AS NetAmount, Posted, UName, SDate " & _
                    "From SalesDetail " & _
                    "WHERE  SDate BETWEEN '" & FDate.Value & "' AND " & _
                    "'" & TDate.Value & "' AND Posted = 'Posted' order by SerialNo", cn, adOpenStatic, adLockReadOnly
            
End Sub


Private Sub previewReport()
    
 On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If
    
    strPath = App.Path + "\reports\Profit Loss Statement Summary.rpt"
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(strPath)



    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions
    
    
    
    Set objReportFormulaFieldDefinations = objReport.FormulaFields
'    Set objReportFF = objReportFormulaFieldDefinations.Item(5)
'            objReportFF.text = "'" + str(Profit) + "'"


    Set objReportFF = objReportFormulaFieldDefinations.Item(1)

              objReportFF.text = "'" + Format(FDate, "dd-MM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
             objReportFF.text = "'" + Format(TDate, "dd-MM-yyyy") + "'"
             
             
       objReportDatabaseTable.SetPrivateData 3, rsMaster
          
        
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        objReport.DiscardSavedData

        objReport.Preview "Profit Loss Information", , , , , 16777216 Or 524288 Or 65536
         If Tracer = 1 Then
         objReport.PrintOut
         End If
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub
ErrH:
If Err.Number = 20545 Then

MsgBox "Request cancelled by the user", vbInformation, "Print"
Err.Clear
End If
'MsgBox Err.Description
Set rsMaster = Nothing

End Sub

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Private Sub FetchData1()

    Set rsMaster = New ADODB.Recordset
                                             
     rsMaster.Open "SELECT SerialNo,CMID, SName,MCatagory, MName, Qty, SRate, Amount,PRate, Discount,(Amount * Discount / 100) AS TDiscount, (Amount - (Amount * Discount / 100)) AS NetAmount, Posted, UName, SDate " & _
                    "From SalesDetail " & _
                    "WHERE  SDate BETWEEN '" & FDate.Value & "' AND " & _
                    "'" & TDate.Value & "' AND Posted = 'Posted' order by SerialNo", cn, adOpenStatic, adLockReadOnly
    
     End Sub


Private Sub previewReport1()
    
 On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If
    
    strPath = App.Path + "\reports\Profit Loss Statement Details.rpt"
    Set objReportApp = CreateObject("Crystal.CRPE.Application")
    Set objReport = objReportApp.OpenReport(strPath)



    Set objReportDatabase = objReport.Database
    Set objReportDatabaseTables = objReportDatabase.Tables
    Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
    Set ObjPrinterSetting = objReport.PrintWindowOptions
    
    
    
    Set objReportFormulaFieldDefinations = objReport.FormulaFields
'    Set objReportFF = objReportFormulaFieldDefinations.Item(5)
'            objReportFF.text = "'" + str(Profit) + "'"


    Set objReportFF = objReportFormulaFieldDefinations.Item(1)

              objReportFF.text = "'" + Format(FDate, "dd-MM-yyyy") + "'"

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
             objReportFF.text = "'" + Format(TDate, "dd-MM-yyyy") + "'"
             
             
             
             

       objReportDatabaseTable.SetPrivateData 3, rsMaster
          
        
        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True
        objReport.DiscardSavedData

        objReport.Preview "Profit Loss Information", , , , , 16777216 Or 524288 Or 65536
         If Tracer = 1 Then
         objReport.PrintOut
         End If
        
        Set objReport = Nothing
        Set objReportDatabase = Nothing
        Set objReportDatabaseTables = Nothing
        Set objReportDatabaseTable = Nothing
    Exit Sub
ErrH:
If Err.Number = 20545 Then

MsgBox "Request cancelled by the user", vbInformation, "Print"
Err.Clear
End If
'MsgBox Err.Description
Set rsMaster = Nothing

End Sub

