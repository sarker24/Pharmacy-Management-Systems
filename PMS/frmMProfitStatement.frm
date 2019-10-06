VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Begin VB.Form RptMDueStatement 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Salse Due Statement "
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6435
   Icon            =   "frmMProfitStatement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6435
   StartUpPosition =   2  'CenterScreen
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
      Height          =   3375
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.CheckBox chkDetails 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Details Due Statement"
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
         Width           =   2535
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   960
         Width           =   1695
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5160
         Top             =   480
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
               Picture         =   "frmMProfitStatement.frx":000C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMProfitStatement.frx":08E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMProfitStatement.frx":11C0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbEO 
         Height          =   600
         Left            =   3360
         TabIndex        =   3
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
      Begin SSCalendarWidgets_A.SSDateCombo SSCurrentDate 
         Height          =   390
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   2175
         _Version        =   65537
         _ExtentX        =   3836
         _ExtentY        =   688
         _StockProps     =   93
         DateSeparator   =   "-"
         Format          =   "DD-MMM-YY"
         BevelColorShadow=   8421504
         DividerStyle    =   4
         DropDownFont3D  =   3
         DropDownForeColor=   32768
         Mask            =   2
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSTDate 
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
         _Version        =   65537
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   93
         DateSeparator   =   "-"
         Format          =   "DD-MMM-YY"
         DropDownForeColor=   16384
      End
      Begin SSCalendarWidgets_A.SSDateCombo SSFDate 
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   1320
         Width           =   1935
         _Version        =   65537
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   93
         DateSeparator   =   "-"
         Format          =   "DD-MMM-YY"
         DropDownForeColor=   16384
      End
      Begin MSAdodcLib.Adodc dcCatagory 
         Height          =   330
         Left            =   3600
         Top             =   2760
         Visible         =   0   'False
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   582
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
         Caption         =   "dcCatagory"
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
         Left            =   3960
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Medicine Salse Due Statement"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6465
   End
End
Attribute VB_Name = "RptMDueStatement"
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

Private Sub FetchData()

Set rsMaster = New ADODB.Recordset
    
    If OpCurrentDate.Value = True Then
                  
    rsMaster.Open "SELECT PC_ID,CDate,CTime,ComputerName,Amount," & _
                    "SerialNo,Post,Particulars," & _
                    "UserName,Status " & _
                    "From Medicine_Payment " & _
                    "WHERE    CDate='" & SSCurrentDate.Date & "' AND Post = 'Posted' order by SerialNo", cn, adOpenStatic, adLockReadOnly
       
     End If
             
      If OpCustomDate.Value = True Then
                                           
     rsMaster.Open "SELECT PC_ID,CDate,CTime,ComputerName,Amount," & _
                    "SerialNo,Post,Particulars," & _
                    "UserName,Status " & _
                    "From Medicine_Payment " & _
                    "WHERE  CDate BETWEEN '" & SSFDate.Date & "' AND " & _
                    "'" & SSTDate.Date & "' AND Post = 'Posted' order by SerialNo", cn, adOpenStatic, adLockReadOnly
      
      End If
        
End Sub

Private Sub previewReport()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

    
        strPath = App.Path + "\reports\Medicine Sales Due.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields
   
'   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'              objReportFF.text = "'" + Format(SSFDate, "dd-MMM-yyyy") + "'"
'         Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'             objReportFF.text = "'" + Format(SSTDate, "dd-MMM-yyyy") + "'"


      
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
    SSCurrentDate.Date = Date
        SSTDate.Date = Date
        SSFDate.Date = Date
'        Call LoadcmbBankName
        OpCurrentDate.Value = True
        lblFrom.Visible = False
        lblTo.Visible = False
        SSFDate.Visible = False
        SSTDate.Visible = False


    End Sub

Private Sub OpCurrentDate_Click()
        lblFrom.Visible = False
        lblTo.Visible = False
        SSFDate.Visible = False
        SSTDate.Visible = False
        SSCurrentDate.Visible = True
'        Label1.Visible = True
End Sub

Private Sub OpCustomDate_Click()
        lblFrom.Visible = True
        lblTo.Visible = True
        SSFDate.Visible = True
        SSTDate.Visible = True
        SSCurrentDate.Visible = False
'        Label1.Visible = False
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
        If SSFDate.Date > SSTDate.Date Then
            MsgBox "Invalid Date and select accurate date range", vbInformation, "Party Wise Sample Report"
            SSFDate.SetFocus
            Validate = False
            Exit Function
        End If
    End Function

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

