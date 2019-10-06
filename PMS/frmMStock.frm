VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmMStock 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Stock View"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   Icon            =   "frmMStock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMID 
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ComboBox cmbMName 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   5985
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2055
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1215
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
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Insert your Medicine Sales Rate."
      Top             =   2160
      Width           =   1335
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1095
   End
   Begin VB.ComboBox cmbSName 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   6015
   End
   Begin SSDataWidgets_A.SSDBCommand SSDBCommand1 
      Height          =   495
      Left            =   0
      TabIndex        =   6
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
      Left            =   5040
      TabIndex        =   15
      Top             =   2880
      Width           =   1095
      _Version        =   196612
      _ExtentX        =   1931
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
      Height          =   405
      Left            =   2400
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   14
      Top             =   2640
      Width           =   2055
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
      TabIndex        =   12
      Top             =   1920
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
      TabIndex        =   11
      Top             =   600
      Width           =   1935
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
      TabIndex        =   10
      Top             =   1920
      Width           =   1335
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
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
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
      TabIndex        =   8
      Top             =   1320
      Width           =   2655
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
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
End
Attribute VB_Name = "frmMStock"
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

Private Function IsValidRecord() As Boolean
    IsValidRecord = True


If Trim(cmbMName) = "" Then
        MsgBox "Please input Medicine Name.", vbInformation
        cmbMName.SetFocus
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
        
    End If
  End Function

Private Sub cmbMName_LostFocus()
Call Others
Call LPRate
Call Balance
End Sub

Private Sub cmdQuit_Click()
Unload Me
'frmPurchase.txtPaid.SetFocus
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
     Call MedicineName
'     dtExpDate.Date = Date + 365
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

Private Sub LPRate()

On Error Resume Next
Dim rsTemp3 As New ADODB.Recordset
rsTemp3.Open ("SELECT SerialNo,MID,Mname,PRate FROM PurchaseDetail where Mname='" & parseQuotes(cmbMName.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic
'rsTemp3.Open ("SELECT SerialNo,MID,Mname,PRate FROM PurchaseDetail where MID='" & parseQuotes(txtMID.text) & "'ORDER BY SerialNo DESC"), cn, adOpenStatic

If rsTemp3.RecordCount > 0 Then
txtLPRate = rsTemp3!PRate

End If
rsTemp3.Close
End Sub

Private Sub Others()
On Error Resume Next

     Dim rsTemp As New ADODB.Recordset

rsTemp.Open ("SELECT MID,MCatagory,OBalance,SRate,Discount FROM tblMedicineName where MName='" & parseQuotes(cmbMName.text) & "'"), cn, adOpenStatic

txtMID = rsTemp!Mid
txtSRate = rsTemp!SRate
txtDiscount = rsTemp!Discount
txtCatagory = rsTemp!MCatagory
    
    rsTemp.Close
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
txtBalance = ""
txtLPRate = ""
cmbMName = ""
txtCatagory = ""
txtSRate = ""
txtMID = ""
End Sub


