VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmPrescriptionsearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Prescription Search"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4440
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   4440
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbRemarks 
      BeginProperty Font 
         Name            =   "SutonnyMJ"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   4185
   End
   Begin VB.TextBox txtCatagory 
      BackColor       =   &H80000000&
      Height          =   360
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2175
   End
   Begin VB.ComboBox cmbMedicineName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4185
   End
   Begin VB.TextBox txtQty 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   2640
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Quit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      Font3D          =   4
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdAdd 
      Height          =   615
      Left            =   2400
      TabIndex        =   3
      Top             =   2640
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   8421376
      Font3D          =   4
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand SSDBCommand1 
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      _Version        =   196611
      _ExtentX        =   7858
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Select Medicine Name"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   -2147483647
      CaptionAlignment=   7
      BorderStyle     =   0
   End
   Begin MSAdodcLib.Adodc dcSupplierName 
      Height          =   360
      Left            =   0
      Top             =   2880
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
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
   Begin VB.Label lblMCatagory 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Catagory"
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
      Left            =   2160
      TabIndex        =   10
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label lblRemarks 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Remarks"
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
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label lblQuantity 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Quantity"
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
      TabIndex        =   7
      Top             =   1200
      Width           =   975
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
      TabIndex        =   6
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "frmPrescriptionsearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsTemp1                     As ADODB.Recordset
Private rsTemp2                     As ADODB.Recordset
Private bRecordExists              As Boolean


Private Sub cmbMedicineName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMedicineName, KeyAscii)
End Sub

Private Sub cmbRemarks_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbRemarks, KeyAscii)
End Sub

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
 Dim iRows As Integer
 Dim i As Integer
 If txtQty = "" Then
   txtQty = 0
 End If

 
 frmPatientTreatment.fgSales.AddItem "" & vbTab & vbTab & frmPrescriptionsearch.txtCatagory.text & vbTab & frmPrescriptionsearch.cmbMedicineName.text & _
                    vbTab & frmPrescriptionsearch.txtQty & vbTab & frmPrescriptionsearch.cmbRemarks.text
                    

cmbMedicineName.RemoveItem (cmbMedicineName.ListIndex)
End If
'cmbItemName.Items.Remove (cmbItemName.SelectedItem)

cmbMedicineName.Refresh


ErrHandler:

    Select Case Err.Number
'        Case -2147217900
        Case 13
            MsgBox "Please input numeric number in QTY field", vbInformation, "Confirmation"

   End Select
  allClear
  cmbMedicineName.SetFocus
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    
If Trim(cmbMedicineName) = "" Then
        MsgBox "Please input Medicine Name.", vbInformation
        cmbMedicineName.SetFocus
        IsValidRecord = False
        Exit Function
               
  ElseIf Trim(txtQty) = "" Then
        MsgBox "Please Input Item Quantity.", vbInformation
        txtQty.SetFocus
        IsValidRecord = False
        Exit Function
               
  ElseIf Trim(cmbRemarks) = "" Then
        MsgBox "Please Input Medicine Remarks.", vbInformation
        cmbRemarks.SetFocus
        IsValidRecord = False
        Exit Function
        
   End If
  End Function


Private Function IsValidRecord1() As Boolean
    IsValidRecord1 = True
        
 If Trim(cmbMedicineName) = "" Then
        MsgBox "Input Medicine Name.", vbInformation
        cmbMedicineName.SetFocus
        IsValidRecord1 = False
        Exit Function
        
                
 ElseIf Trim(txtQty) = "" Then
        MsgBox "Input Medicine Quantity.", vbInformation
        txtQty.SetFocus
        IsValidRecord1 = False
        Exit Function
        
        If Trim(cmbRemarks) = "" Then
        MsgBox "Input Medicine Remarks.", vbInformation
        cmbRemarks.SetFocus
        IsValidRecord1 = False
        Exit Function
        
      
    End If
    End If
  End Function

Private Sub cmbMedicineName_LostFocus()
    Call SupplierName1
'    Call Balance
End Sub

Private Sub cmdQuit_AfterClick()
    Unload Me
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Call Connect
     Call MedicineName
     Call Remarks
'     dtExpDate.Date = Date
End Sub

Private Sub cmbMedicineName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

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

Private Sub Remarks()

    Dim rsTemp2 As New ADODB.Recordset
     
     
     rsTemp2.Open ("SELECT DISTINCT Remarks FROM tblRemarks ORDER BY Remarks ASC"), cn, adOpenStatic
    
    While Not rsTemp2.EOF
        cmbRemarks.AddItem rsTemp2("Remarks")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close
     
End Sub


Private Sub SupplierName1()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MCatagory FROM tblMedicineName where MedicineName='" & parseQuotes(cmbMedicineName.text) & "'"), cn, adOpenStatic
    If rsTemp.RecordCount > 0 Then
        txtCatagory = rsTemp!MCatagory
       
 End If
    rsTemp.Close
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub


Private Sub allClear()
    cmbMedicineName = ""
    txtQty = ""
    txtCatagory = ""
    cmbRemarks = ""
End Sub




