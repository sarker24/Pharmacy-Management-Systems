VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSalesSearch 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4125
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSalesSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleMode       =   0  'User
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMID 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtSName 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   4335
   End
   Begin VB.TextBox txtPRate 
      BackColor       =   &H00C0C0C0&
      Height          =   450
      Left            =   2160
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDiscount 
      BackColor       =   &H80000000&
      Height          =   450
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.ComboBox cmbMName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4305
   End
   Begin VB.TextBox txtCatagory 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2055
   End
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   615
      Left            =   3480
      TabIndex        =   3
      Top             =   3480
      Width           =   975
      _Version        =   196612
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
      Left            =   2520
      TabIndex        =   2
      Top             =   3480
      Width           =   975
      _Version        =   196612
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
   Begin VB.TextBox txtBalance 
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtSRate 
      Height          =   450
      Left            =   1200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox txtQty 
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin SSDataWidgets_A.SSDBCommand SSDBCommand1 
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5775
      _Version        =   196612
      _ExtentX        =   10186
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "Select Items"
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
   Begin VB.Label lblMID 
      BackColor       =   &H00C0B4A9&
      Caption         =   "MID"
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
      Left            =   3360
      TabIndex        =   19
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblPRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "P Rate"
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
      TabIndex        =   16
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblDiscount 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Discount"
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
      Left            =   3360
      TabIndex        =   14
      Top             =   1320
      Width           =   975
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
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "S Rate"
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
      Left            =   1200
      TabIndex        =   8
      Top             =   1320
      Width           =   735
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
      Top             =   1320
      Width           =   1575
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
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
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
      TabIndex        =   5
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "frmSalesSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsTemp1                     As ADODB.Recordset
Private rsTemp2                     As ADODB.Recordset
Private rsBalance                   As ADODB.Recordset
Private bRecordExists              As Boolean


Private Sub cmbMName_Change()
'Call MedicineName
End Sub

Private Sub cmbMName_GotFocus()
'Call MedicineName
End Sub

Private Sub cmbMName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMName, KeyAscii)
End Sub

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
 Dim iRows As Integer
 Dim i As Integer
 If txtQty = "" Then
   txtQty = 0
 End If
If txtSRate = "" Then
   txtSRate = 0
 End If
 If txtPRate = "" Then
   txtPRate = 0
 End If
 
 frmSales.fgSales.AddItem "" & vbTab & vbTab & frmSalesSearch.txtSName.text & vbTab & frmSalesSearch.txtMID.text & _
                    vbTab & frmSalesSearch.txtCatagory & vbTab & frmSalesSearch.cmbMName & vbTab & frmSalesSearch.txtQty & _
                    vbTab & frmSalesSearch.txtSRate & vbTab & frmSalesSearch.txtQty * frmSalesSearch.txtSRate & _
                    vbTab & frmSalesSearch.txtPRate & vbTab & (frmSalesSearch.txtDiscount)
'                   vbTab & (frmSalesSearch.txtDiscount * frmSalesSearch.txtQty * frmSalesSearch.txtRate) / 100


cmbMName.RemoveItem (cmbMName.ListIndex)
End If

cmbMName.Refresh


ErrHandler:

    Select Case Err.Number
'        Case -2147217900
        Case 13
            MsgBox "Please select numeric number in QTY/RATE field", vbInformation, "Confirmation"

   End Select
  allClear
  cmbMName.SetFocus
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True
    
    
If Trim(cmbMName) = "" Then
        MsgBox "Please input Medicine Name.", vbInformation
        cmbMName.SetFocus
        IsValidRecord = False
        Exit Function
               
  ElseIf Trim(txtQty) = "" Then
        MsgBox "Please Input Item Quantity.", vbInformation
        txtQty.SetFocus
        IsValidRecord = False
        Exit Function
        
'   ElseIf Trim(txtRate) = "" Then
'        MsgBox "Please input Item Rate.", vbInformation
'        txtRate.SetFocus
'        IsValidRecord = False
'        Exit Function
'
    ElseIf Val(txtQty) > Val(txtBalance) Then
        MsgBox "Medicine Stock are not available.", vbInformation
        cmbMName.SetFocus
        IsValidRecord = False
        Exit Function
        
   End If
  End Function


Private Function IsValidRecord1() As Boolean
    IsValidRecord1 = True
          
 If Trim(cmbMName) = "" Then
        MsgBox "Your are missing Medicine Name.", vbInformation
        cmbMName.SetFocus
        IsValidRecord1 = False
        Exit Function
        
                
 ElseIf Trim(txtQty) = "" Then
        MsgBox "Your are missing Medicine Quantity.", vbInformation
        txtQty.SetFocus
        IsValidRecord1 = False
        Exit Function
        
ElseIf Trim(txtSRate) = "" Then
        MsgBox "Your are missing Sales Rate.", vbInformation
        txtSRate.SetFocus
        IsValidRecord1 = False
        Exit Function
        
    End If
  End Function

Private Sub cmbMName_LostFocus()
    Call SupplierName1
    Call Balance
    Call PRate
End Sub

Private Sub cmdQuit_Click()
    Unload Me
    frmSales.cmAutoPaid.SetFocus
End Sub

Private Sub Form_Load()
     
     ModFunction.StartUpPosition Me
     Call Connect
     Call MedicineName

End Sub

Private Sub cmbMName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys Chr(9)
    End If

End Sub

Private Sub MedicineName()

    Dim rsTemp2 As New ADODB.Recordset


'     rsTemp2.Open ("SELECT DISTINCT MName FROM tblMedicineName ORDER BY MName ASC"), cn, adOpenStatic
'
     rsTemp2.Open ("SELECT DISTINCT top 20 MName FROM tblMedicineName ORDER BY MName ASC"), cn, adOpenStatic

    While Not rsTemp2.EOF
        cmbMName.AddItem rsTemp2("MName")
        rsTemp2.MoveNext
    Wend
    rsTemp2.Close

End Sub

Private Sub SupplierName1()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MID,MCatagory,SRate,Discount,SName FROM tblMedicineName where MName='" & parseQuotes(cmbMName.text) & "'"), cn, adOpenStatic
    
    If rsTemp.RecordCount > 0 Then
        txtMID = rsTemp!Mid
        txtCatagory = rsTemp!MCatagory
        txtSRate = rsTemp!SRate
        txtDiscount = rsTemp!Discount
        txtSName = rsTemp!SName
 End If
    rsTemp.Close
End Sub

Private Sub PRate()
On Error Resume Next
Dim rsTemp1 As New ADODB.Recordset
'rsTemp1.Open ("SELECT SerialNo,MID,MName,PRate FROM PurchaseDetail where MID='" & parseQuotes(txtMID.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic
rsTemp1.Open ("SELECT SerialNo,MID,MName,PRate FROM PurchaseDetail where MName='" & parseQuotes(cmbMName.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic
If rsTemp1.RecordCount > 0 Then
txtPRate = rsTemp1!PRate
End If
rsTemp1.Close
End Sub

Private Sub txtDiscount_Change()
'txtSRate=
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
End Sub

Private Sub Balance()
    Dim rsBalance As New ADODB.Recordset
        If rsBalance.State <> 0 Then rsBalance.Close

rsBalance.Open ("select (isnull(sum(PurchaseDetail.Qty),0))-(select isnull(sum(SalesDetail.Qty),0) " & _
                "from SalesDetail where SalesDetail.Posted='Posted' and SalesDetail.MName='" & parseQuotes(cmbMName.text) & "' ) " & _
                "as Balance from PurchaseDetail where PurchaseDetail.MName='" & parseQuotes(cmbMName.text) & "'  and  PurchaseDetail.Posted='Posted'"), cn, adOpenStatic

'rsBalance.Open ("select (isnull(sum(PurchaseDetail.Qty),0))-(select isnull(sum(SalesDetail.Qty),0) " & _
'                "from SalesDetail where SalesDetail.Posted='Posted' and SalesDetail.MID='" & parseQuotes(txtMID.text) & "' ) " & _
'                "as Balance from PurchaseDetail where PurchaseDetail.MID='" & parseQuotes(txtMID.text) & "'  and  PurchaseDetail.Posted='Posted'"), cn, adOpenStatic


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
    txtDiscount = ""
    txtQty = ""
    txtSRate = ""
    txtPRate = ""
    txtMID = ""
    txtCatagory = ""
    txtBalance = ""
    txtSName = ""
    cmbMName = ""
End Sub
