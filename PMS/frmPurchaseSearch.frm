VERSION 5.00
Object = "{E8671A8B-E5DD-11CD-836C-0000C0C14E92}#1.0#0"; "SSCALA32.OCX"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmPurchaseSearch 
   BackColor       =   &H00C0B4A9&
   Caption         =   "cmbMedicineName"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPurchaseSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6285
   Begin VB.TextBox txtMID 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Insert your Medicine Purchase Rate."
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtVAT 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Insert your Medicine Purchase Rate."
      Top             =   2040
      Width           =   855
   End
   Begin VB.ComboBox cmbSName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6015
   End
   Begin VB.TextBox txtAmount 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      ToolTipText     =   "Insert your Medicine Purchase Rate."
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox txtBalance 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   5130
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtPRate 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """$""#,##0.00;(""$""#,##0.00)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   2
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      ToolTipText     =   "Insert your Medicine Purchase Rate."
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox txtLPRate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Insert your Medicine Sales Rate."
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox txtSRate 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   4890
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtCatagory 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2295
   End
   Begin VB.TextBox txtQty 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Insert your Medicine Quantity."
      Top             =   2040
      Width           =   975
   End
   Begin VB.ComboBox cmbMName 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5985
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtManuf 
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1935
      _Version        =   65537
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "-"
      BevelColorHighlight=   8421376
      BevelType       =   2
   End
   Begin SSCalendarWidgets_A.SSDateCombo dtExpDate 
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
      _Version        =   65537
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   93
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "-"
      BevelColorFace  =   14737632
      BevelColorHighlight=   12632064
      CaptionBevelType=   2
      BevelType       =   2
   End
   Begin SSDataWidgets_A.SSDBCommand SSDBCommand1 
      Height          =   495
      Left            =   0
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   6375
      _Version        =   196612
      _ExtentX        =   11245
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
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   495
      Left            =   5160
      TabIndex        =   14
      Top             =   4080
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
      BevelColorFace  =   12629161
      Font3D          =   4
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdAdd 
      Height          =   495
      Left            =   4200
      TabIndex        =   4
      Top             =   4080
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
      BevelColorFace  =   12629161
      Font3D          =   4
      CaptionAlignment=   7
   End
   Begin SSCalendarWidgets_A.SSDateCombo SSDateCombo1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd-MMM-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   2655
      _Version        =   65537
      _ExtentX        =   4683
      _ExtentY        =   873
      _StockProps     =   93
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DateSeparator   =   "-"
      ForeColorSelected=   16777215
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
      Left            =   4920
      TabIndex        =   31
      Top             =   3240
      Width           =   1215
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
      Height          =   375
      Left            =   4200
      TabIndex        =   29
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label lblVAT 
      BackColor       =   &H00C0B4A9&
      Caption         =   "VAT"
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
      Left            =   3240
      TabIndex        =   27
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Amount"
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
      Left            =   2280
      TabIndex        =   26
      Top             =   1800
      Width           =   855
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
      Left            =   5130
      TabIndex        =   25
      Top             =   1800
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
      TabIndex        =   24
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblRate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Qty Rate"
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
      Left            =   1200
      TabIndex        =   23
      Top             =   1800
      Width           =   975
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
      TabIndex        =   22
      Top             =   1800
      Width           =   975
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
      Left            =   4890
      TabIndex        =   21
      Top             =   2520
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
      Left            =   2640
      TabIndex        =   20
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Date of Expiry"
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
      Left            =   2280
      TabIndex        =   19
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Date of Manufactur"
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
      TabIndex        =   18
      Top             =   3240
      Width           =   2055
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
      TabIndex        =   17
      Top             =   480
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
      TabIndex        =   16
      Top             =   2520
      Width           =   2295
   End
End
Attribute VB_Name = "frmPurchaseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsTemp1                     As ADODB.Recordset
Private rsTemp2                     As ADODB.Recordset
Private rsTemp3                     As ADODB.Recordset
Private rsBalance                   As ADODB.Recordset
Private bRecordExists              As Boolean


Private Sub cmbMName_DropDown()
Call MedicineName
End Sub

Private Sub cmbMName_KeyPress(KeyAscii As Integer)
   KeyAscii = AutoMatchCBBox(cmbMName, KeyAscii)
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

Private Sub cmdAdd_Click()

On Error GoTo ErrHandler

If IsValidRecord Then
 Dim iRows As Integer
 Dim i As Integer
 If txtQty = "" Then
   txtQty = 0
 End If
If txtPRate = "" Then
   txtPRate = 0
 End If

 frmPurchase.fgPurchase.AddItem "" & vbTab & vbTab & frmPurchaseSearch.cmbSName.text & vbTab & vbTab & vbTab & frmPurchaseSearch.txtMID.text & _
                    vbTab & frmPurchaseSearch.txtCatagory.text & vbTab & frmPurchaseSearch.cmbMName.text & vbTab & frmPurchaseSearch.txtQty & _
                    vbTab & frmPurchaseSearch.txtPRate & vbTab & frmPurchaseSearch.txtPRate * frmPurchaseSearch.txtQty & _
                    vbTab & frmPurchaseSearch.txtSRate & vbTab & frmPurchaseSearch.dtExpDate & vbTab & frmPurchaseSearch.dtManuf.text

'
' frmPurchase.fgPurchase.AddItem "" & vbTab & vbTab & frmPurchaseSearch.cmbSName.text & vbTab & vbTab & vbTab & frmPurchaseSearch.txtCatagory.text & _
'                    vbTab & frmPurchaseSearch.cmbMedicineName.text & vbTab & frmPurchaseSearch.txtQty & vbTab & frmPurchaseSearch.txtPRate & _
'                    vbTab & frmPurchaseSearch.txtQty * frmPurchaseSearch.txtPRate & vbTab & frmPurchaseSearch.txtSRate & vbTab & frmPurchaseSearch.dtExpDate & _
'                    vbTab & frmPurchaseSearch.dtManuf.text


cmbMName.RemoveItem (cmbMName.ListIndex)
End If
'cmbItemName.Items.Remove (cmbItemName.SelectedItem)

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
        MsgBox "Please input Item Quantity.", vbInformation
        txtQty.SetFocus
        IsValidRecord = False
        Exit Function

   ElseIf Trim(txtPRate) = "" Then
        MsgBox "Please input Purchase Rate.", vbInformation
        txtPRate.SetFocus
        IsValidRecord = False
        Exit Function

      End If
  End Function


Private Function IsValidRecord1() As Boolean
    IsValidRecord1 = True

    If Trim(txtLPRate) = "" Then
        MsgBox "Please Input Sales Rate.", vbInformation
        txtLPRate.SetFocus
        IsValidRecord1 = False
        Exit Function

 ElseIf Trim(cmbMName) = "" Then
        MsgBox "Please input Medicine Name.", vbInformation
        cmbMName.SetFocus
        IsValidRecord1 = False
        Exit Function


 ElseIf Trim(txtQty) = "" Then
        MsgBox "Your are missing Qty Information.", vbInformation
        txtQty.SetFocus
        IsValidRecord1 = False
        Exit Function

ElseIf Trim(txtPRate) = "" Then
        MsgBox "Your are missing Rate Information.", vbInformation
        txtPRate.SetFocus
        IsValidRecord1 = False
        Exit Function

    End If
  End Function

Private Sub cmbMName_LostFocus()
Call Others
Call LPRate
Call Balance
End Sub

Private Sub cmdQuit_Click()
Unload Me
frmPurchase.txtPaid.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys Chr(9)
    End If
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Call Connect
     Call SName
'     Call MedicineName
     dtExpDate.Date = Date + 365
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
     
cmbMName.Clear
Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MName, SName FROM tblMedicineName where SName='" & parseQuotes(cmbSName.text) & "'ORDER BY MName ASC"), cn, adOpenStatic
    While Not rsTemp.EOF
    cmbMName.AddItem rsTemp("MName")
    rsTemp.MoveNext
    Wend
    rsTemp.Close
    
'    Call Others
     
End Sub

Private Sub txtAmount_GotFocus()
    txtAmount.BackColor = &HFFC0C0
    txtAmount.SelStart = 0
    txtAmount.SelLength = Len(txtAmount)
End Sub

Private Sub txtPRate_Change()
Call Calculation1
End Sub

Private Sub txtPRate_GotFocus()
    txtPRate.BackColor = &HFFC0C0
    txtPRate.SelStart = 0
    txtPRate.SelLength = Len(txtPRate)
End Sub

Private Sub txtPRate_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

Call Calculation1
End Sub

Private Sub Calculation1()
'txtAmount = CDbl(Val(txtPRate)) * CDbl(Val(txtQty))
End Sub

Private Sub txtAmount_Change()
'Call Calculation
End Sub

Private Sub Calculation()
'txtPRate = CDbl(Val(txtAmount)) / CDbl(Val(txtQty))
txtPRate = ((Val(txtAmount) + (Val(txtAmount) * Val(txtVAT)) / 100) - (Val(txtAmount) * (Val(txtDiscount)) / 100)) / Val(txtQty)
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If
Call Calculation

End Sub

Private Sub LPRate()

On Error Resume Next
Dim rsTemp3 As New ADODB.Recordset
'rsTemp3.Open ("SELECT SerialNo,MID,Mname,PRate FROM PurchaseDetail where MID='" & parseQuotes(txtMID.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic
rsTemp3.Open ("SELECT SerialNo,MID,Mname,PRate FROM PurchaseDetail where Mname='" & parseQuotes(cmbMName.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic

If rsTemp3.RecordCount > 0 Then
txtLPRate = rsTemp3!PRate

End If
rsTemp3.Close
End Sub

Private Sub Others()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MID,MCatagory,OBalance,SRate FROM tblMedicineName where MName='" & parseQuotes(cmbMName.text) & "'"), cn, adOpenStatic

txtSRate = rsTemp!SRate
txtMID = rsTemp!Mid
txtCatagory = rsTemp!MCatagory
    
    rsTemp.Close
End Sub

Private Sub txtQty_GotFocus()
    txtQty.BackColor = &HFFC0C0
    txtQty.SelStart = 0
    txtQty.SelLength = Len(txtQty)
End Sub

Private Sub txtQty_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub

Private Sub txtVAT_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    SendKeys Chr(9)
End If

End Sub


Private Sub Balance()
Dim rsBalance As New ADODB.Recordset
If rsBalance.State <> 0 Then rsBalance.Close

rsBalance.Open ("select (isnull(sum(PurchaseDetail.Qty),0))-(select isnull(sum(SalesDetail.Qty),0) " & _
                "from SalesDetail where SalesDetail.Posted='Posted' and SalesDetail.MName='" & parseQuotes(cmbMName.text) & "' ) " & _
                "as Balance from PurchaseDetail  where PurchaseDetail.MName='" & parseQuotes(cmbMName.text) & "'  and  PurchaseDetail.Posted='Posted'"), cn, adOpenStatic

'rsBalance.Open ("select (isnull(sum(PurchaseDetail.Qty),0))-(select isnull(sum(SalesDetail.Qty),0) " & _
'                "from SalesDetail where SalesDetail.Posted='Posted' and SalesDetail.MID='" & parseQuotes(txtMID.text) & "' ) " & _
'                "as Balance from PurchaseDetail  where PurchaseDetail.MID='" & parseQuotes(txtMID.text) & "'  and  PurchaseDetail.Posted='Posted'"), cn, adOpenStatic


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
txtQty = ""
txtPRate = ""
txtAmount = ""
txtBalance = ""
txtLPRate = ""
cmbMName = ""
txtCatagory = ""
txtSRate = ""
txtMID = ""
End Sub
