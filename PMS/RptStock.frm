VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptStock 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Stock Report"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   Icon            =   "RptStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   3840
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   2775
   End
   Begin VB.ComboBox cboPName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   0
      Text            =   "<All>"
      Top             =   2280
      Width           =   2775
   End
   Begin MSComCtl2.DTPicker DTPur 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yy"
      Format          =   51314691
      CurrentDate     =   38388
   End
   Begin MSComctlLib.Toolbar tbEO 
      Height          =   600
      Left            =   840
      TabIndex        =   6
      Top             =   2760
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2880
      Top             =   2760
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
            Picture         =   "RptStock.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptStock.frx":05E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "RptStock.frx":0646
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Product Description"
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
      Left            =   480
      TabIndex        =   4
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Purchase Date"
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
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Product Name"
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
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "RptStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim rs As New ADODB.Recordset
'Dim strSQL As String
''--------------------------------------------------------------------
'Private objReportApp                          As CRPEAuto.Application
'Private objReport                             As CRPEAuto.Report
'Private objReportDatabase                     As CRPEAuto.Database
'Private objReportDatabaseTables               As CRPEAuto.DatabaseTables
'Private objReportDatabaseTable                As CRPEAuto.DatabaseTable
'Private ObjPrinterSetting                   As CRPEAuto.PrintWindowOptions
'Private objReportFormulaFieldDefinations    As CRPEAuto.FormulaFieldDefinitions
'Private objReportFF                         As CRPEAuto.FormulaFieldDefinition
'
'
'Private Sub Form_Load()
'    Call Connect
'    ModFunction.StartUpPosition Me
'DTPur = Date
'End Sub
'
Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
     Case "Preview"
'            If Validate Then
'                Tracer = 0
'                Call FetchData
'                Call previewReport
'               End If
     Case "Print"
'            If Validate Then
'                Tracer = 1
'                Call FetchData
'                Call previewReport
'               End If
     Case "Close"
               Unload Me
    End Select
End Sub
'Private Sub FetchData()
'Dim stDate As Integer
'If cboPName.text <> "<All>" And stDate = 1 Then
'        strn = "SELECT * FROM " & StrTable & " " & _
'                 " Where " & StrTable & ".ProductName='" & cboPName.text & "' " & _
'                 " and  " & StrTable & ".PurchaseDate='" & DTPur & "'"
'        If cmbType.text <> "" Then
'            strn = strn + " and " & StrTable & ".TypeID='" & cmbType & "' "
'        End If
'    ElseIf cboPName.text = "<All>" And stDate = 1 Then
'        strn = "SELECT * FROM " & StrTable & " " & _
'                 " Where  " & StrTable & ".PurchaseDate='" & DTPur & "' "
'        If cmbType.text <> "" Then
'            strn = strn + " and " & StrTable & ".TypeID='" & cmbType & "' "
'        End If
'    ElseIf cboPName.text <> "<All>" And stDate = 0 Then
'        strn = "SELECT * FROM " & StrTable & " " & _
'        " Where  " & StrTable & ".ProductName='" & cboPName.text & "'"
'        If cmbType.text <> "" Then
'            strn = strn + " and " & StrTable & ".TypeID='" & cmbType & "' "
'        End If
'    Else
'        strn = "SELECT * FROM " & StrTable & " "
'        If cmbType.text <> "" Then
'            strn = strn + " Where " & StrTable & ".TypeID='" & cmbType & "'"
'        End If
'    End If
'    strn = strn + " order by " & StrTable & ".ProductName, " & StrTable & ".TypeID , PurchaseDate"
'
'    rsEmp.Open strn, cn, adOpenStatic, adLockReadOnly
'End Sub
'
'

