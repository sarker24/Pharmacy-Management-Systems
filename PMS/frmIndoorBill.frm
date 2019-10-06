VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndoorBill 
   BackColor       =   &H00C0B4A9&
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11430
   Icon            =   "frmIndoorBill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   11430
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optMReciept 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Money Reciept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   16
      Top             =   5400
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker MRDate 
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   65601539
      CurrentDate     =   42611
   End
   Begin VB.TextBox txtCMID 
      Height          =   495
      Left            =   8520
      TabIndex        =   13
      ToolTipText     =   "Cash Memo ID No."
      Top             =   4800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPost 
      Height          =   495
      Left            =   8520
      TabIndex        =   12
      ToolTipText     =   "Cash Memo Posted"
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtAdjust 
      Height          =   495
      Left            =   8520
      TabIndex        =   11
      ToolTipText     =   "Amout Adjust"
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPaid 
      Height          =   495
      Left            =   8520
      TabIndex        =   10
      ToolTipText     =   "Amount Paid"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.OptionButton optDStatement 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Details Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   6480
      Width           =   2415
   End
   Begin VB.OptionButton optSStatement 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Summary Statement"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   6120
      Width           =   2415
   End
   Begin VB.OptionButton optCashMemo 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient Cash Memo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   5760
      Width           =   2415
   End
   Begin VB.TextBox txtPay 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Payable Amount"
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdMPayment 
      Caption         =   "Make Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   5640
      Width           =   1815
   End
   Begin VB.TextBox TxtDue 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   4
      ToolTipText     =   "Due Amount"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox txtRegNo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      ToolTipText     =   "Patient Reg No."
      Top             =   4320
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   6480
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
            Picture         =   "frmIndoorBill.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndoorBill.frx":08E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmIndoorBill.frx":11C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbEO 
      Height          =   600
      Left            =   3840
      TabIndex        =   9
      Top             =   6360
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
   Begin VSFlex7LCtl.VSFlexGrid fgExport 
      Height          =   4005
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11400
      _cx             =   20108
      _cy             =   7064
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12629161
      ForeColor       =   -2147483640
      BackColorFixed  =   12632064
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   0
      BackColorBkg    =   12629161
      BackColorAlternate=   14737632
      GridColor       =   12629161
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmIndoorBill.frx":1A9A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Total Due Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label lblRegNo 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Patient RegNo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4080
      Width           =   1695
   End
End
Attribute VB_Name = "frmIndoorBill"
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
Private rsTemp                              As ADODB.Recordset
Private rs                                  As ADODB.Recordset
Private rs1                                  As ADODB.Recordset
Dim temp1                                   As Double
Dim temp2                                   As Double

 
'-----------Get computer name-------------------------
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'-----------End to get computer name-------------------

Private Sub cmdMPayment_Click()

Dim j As Integer
temp1 = 0
temp2 = 0

'txtPay.text = ""
txtCMID.text = ""
txtAdjust.text = ""
txtPost.text = ""


If txtPay.text = "" Then
            MsgBox "Enter payment amount"
            txtPay.SetFocus
Exit Sub

ElseIf Val(txtPay.text) > Val(TxtDue.text) Then
            MsgBox "Payment is more than due"
            txtPay.SetFocus
Exit Sub
End If


For j = 1 To fgExport.Rows - 1
    If Val(txtPay.text) > fgExport.TextMatrix(j, 9) Then
           temp1 = CDbl(txtPay.text) - CDbl(Val(fgExport.TextMatrix(j, 9)))
           txtPay.text = temp1
           txtCMID.text = Val(fgExport.TextMatrix(j, 1))
           txtAdjust.text = CDbl(Val(fgExport.TextMatrix(j, 9)))
           txtPost.text = fgExport.TextMatrix(j, 10)
           txtPaid.text = CDbl(Val(fgExport.TextMatrix(j, 8))) + CDbl(Val(fgExport.TextMatrix(j, 9)))
           fgExport.TextMatrix(j, 8) = CDbl(Val(txtPaid.text))
           fgExport.TextMatrix(j, 9) = 0
           Duecollection
    cn.Execute "UPDATE SalesMaster SET Tdue =  " & fgExport.TextMatrix(j, 9) & " , Tpaid=" & Val(txtPaid.text) & "  where CMID= " & fgExport.TextMatrix(j, 1) & ""
   
    
     ElseIf Val(txtPay.text) <= fgExport.TextMatrix(j, 9) Then
           temp1 = CDbl(Val(fgExport.TextMatrix(j, 9))) - CDbl(txtPay.text)
           txtAdjust.text = CDbl(Val(txtPay.text))
           txtPaid = CDbl(Val(fgExport.TextMatrix(j, 8))) + CDbl(Val(txtAdjust.text))
'          Val (txtPaid.text = CDbl(Val(fgExport.TextMatrix(j, 8))) + CDbl(Val(txtAdjust.text)))
           fgExport.TextMatrix(j, 8) = CDbl(Val(txtPaid.text))
           txtPay.text = temp1
           fgExport.TextMatrix(j, 9) = temp1
           txtCMID.text = fgExport.TextMatrix(j, 1)
           txtPost.text = fgExport.TextMatrix(j, 10)
           
           Duecollection
           cn.Execute "UPDATE SalesMaster SET Tdue =  " & fgExport.TextMatrix(j, 9) & ", Tpaid=" & Val(txtPaid.text) & " where CMID= " & fgExport.TextMatrix(j, 1) & ""
            
            MsgBox "Patient Due Collection Successfully.", vbInformation, "Confirmation"
            txtPay.text = 0
 
           Exit Sub
    
    End If
   
   
    Next
    
 MsgBox "Record added Successfully", vbInformation, "Confirmation"
 txtPay.text = 0
End Sub

Private Sub Duecollection()

On Error Resume Next
Dim Due As String
Dim str As String
'------------------Define Due Date-------------------------
Dim D1, D2 As Date
D1 = frmSales.SDate
D2 = frmSales.dtpCurrentDate
Dim Particulars As String
Dim Status As String
Dim PartialAmount As Integer

  If D1 = D2 Then
                Particulars = "Current Due Amount"
                Status = "Advance"
          Else
                Particulars = "Previous Due Amount"
                Status = "Due"
  End If
'---------------End Define Due Date-----------------------

'---------Define Computer Name-----------------------------
 Dim dwLen As Long
 Dim strString As String
 dwLen = MAX_COMPUTERNAME_LENGTH + 1
 strString = String(dwLen, "X")
 GetComputerName strString, dwLen
 strString = Left(strString, dwLen)



             cn.Execute "insert into Medicine_Payment (CDate,CTime,ComputerName,Amount,CMID,RegNo,Post,Particulars,UName,Status) " & _
                        " values('" & Format(frmSales.dtpCurrentDate, "dd-mmm-yyyy") & "',    " & _
                        "   '" & frmSales.txtTime & "','" & parseQuotes(strString) & "',         " & _
                        "   " & (txtAdjust.text) & ", " & (txtCMID.text) & " ,                     " & _
                        "   '" & (txtRegNo.text) & "',                             " & _
                        "   '" & (txtPost.text) & "',                             " & _
                        "   '" & (Particulars) & "','" & parseQuotes(frmLogin.txtUName) & "', " & _
                        "   '" & (Status) & "')"



End Sub

Public Function parseQuotes(text As String) As String
    parseQuotes = Replace(text, "'", "''")
End Function

Private Sub fgExport_Click()
If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Row  From the List."
        Exit Sub
    End If

     Call PopulateCompanySearch

Unload Me
Set frmIndoorBill = Nothing

End Sub

Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then

             frmSales.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub Form_Load()
ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     fgExport.Editable = flexEDKbdMouse
     
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
     rsTemp.Open "SELECT TOP 10 CMID,RegNo,BedNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted FROM SalesMaster ORDER BY CMID DESC", cn, adOpenStatic, adLockReadOnly
         
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("CMID") & vbTab & rsTemp("RegNo") & vbTab & rsTemp("BedNo") & vbTab & rsTemp("CName") & _
         vbTab & Format(rsTemp("SDate"), "dd-MM-yyyy") & vbTab & rsTemp("Tamount") & vbTab & rsTemp("DiscuntTaka") & vbTab & rsTemp("Tpaid") & vbTab & rsTemp("Tdue") & _
         vbTab & rsTemp("Posted")
        rsTemp.MoveNext
    Wend
End Sub


Private Sub optMReciept_Click()
If optMReciept.Value = True Then
MRDate.Visible = True
MRDate.Value = Date
End If
If MRDate.Value = False Then
MRDate.Visible = False
End If
'End If
End Sub

Private Sub txtPay_GotFocus()
txtPay.BackColor = &HFFFFFF
txtPay.SelStart = 0
txtPay.SelLength = Len(txtPay)
End Sub

Private Sub txtRegNo_Change()

If rsTemp.State <> 0 Then rsTemp.Close

      rsTemp.Open "SELECT TOP 50 CMID,RegNo,BedNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                  "FROM SalesMaster WHERE SalesMaster.RegNo LIKE '" & txtRegNo.text & "%' And Posted = 'Posted' And Tdue <> 0 ", cn, adOpenStatic, adLockReadOnly

         fgExport.Rows = 1

    While Not rsTemp.EOF
       fgExport.AddItem "" & vbTab & rsTemp("CMID") & vbTab & rsTemp("RegNo") & vbTab & rsTemp("BedNo") & vbTab & rsTemp("CName") & _
         vbTab & Format(rsTemp("SDate"), "dd-MM-yyyy") & vbTab & rsTemp("Tamount") & vbTab & rsTemp("DiscuntTaka") & vbTab & rsTemp("Tpaid") & vbTab & rsTemp("Tdue") & _
         vbTab & rsTemp("Posted")
        rsTemp.MoveNext
        Wend

        Call DueCheck

        TxtDue = Val(rs!Amt)
End Sub

Private Sub txtRegNo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    SendKeys Chr(9)
End If

If rsTemp.State <> 0 Then rsTemp.Close
        
       
      rsTemp.Open "SELECT TOP 50 CMID,RegNo,BedNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                  "FROM SalesMaster WHERE SalesMaster.RegNo LIKE '" & txtRegNo.text & "%' And Posted = 'Posted' And Tdue <> 0 ", cn, adOpenStatic, adLockReadOnly

         fgExport.Rows = 1
    
    While Not rsTemp.EOF
       fgExport.AddItem "" & vbTab & rsTemp("CMID") & vbTab & rsTemp("RegNo") & vbTab & rsTemp("BedNo") & vbTab & rsTemp("CName") & _
         vbTab & Format(rsTemp("SDate"), "dd-MM-yyyy") & vbTab & rsTemp("Tamount") & vbTab & rsTemp("DiscuntTaka") & vbTab & rsTemp("Tpaid") & vbTab & rsTemp("Tdue") & _
         vbTab & rsTemp("Posted")
        rsTemp.MoveNext
        Wend
        
'  End If
        Call DueCheck
        
        TxtDue = Val(rs!Amt)
'  End If
End Sub


Private Sub txtPay_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 13 Then
    SendKeys Chr(9)
End If
'Call Calculation
End Sub

Private Sub DueCheck()
Dim str As String
Set rs = New ADODB.Recordset
If rs.State <> 0 Then rs.Close

str = "Select ISNULL(sum(Tdue),0) as Amt from SalesMaster where RegNo='" & parseQuotes(Me.txtRegNo) & "'"

      rs.Open str, cn, adOpenStatic, adLockReadOnly
       
End Sub



Private Sub tbEO_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
     Case "Preview"
            If optCashMemo.Value = True Then
            If txtRegNo.text = "" Then
            MsgBox "Input Patient RegNo"
            txtRegNo.SetFocus
            Exit Sub
                End If
                Tracer = 0
                Call FetchData
                Call CashMemo
                End If
            If optMReciept.Value = True Then
            If txtRegNo.text = "" Then
            MsgBox "Input Patient RegNo"
            txtRegNo.SetFocus
            Exit Sub
                End If
                Call FetchData1
                Call MoneyReceipt
               End If

     Case "Print"
            If optCashMemo.Value = True Then
            If txtRegNo.text = "" Then
            MsgBox "Input Patient RegNo"
            txtRegNo.SetFocus
            Exit Sub
                End If
                Tracer = 1
                Call FetchData
                Call CashMemo
                End If
            If optMReciept.Value = True Then
            If txtRegNo.text = "" Then
            MsgBox "Input Patient RegNo"
            txtRegNo.SetFocus
            Exit Sub
                End If
                Call FetchData1
                Call MoneyReceipt
               End If
     Case "Close"
               Unload Me
    End Select
End Sub

Public Function FetchData()

    Set rsMaster = New ADODB.Recordset

    If optCashMemo.Value = True Then
    
    rsMaster.Open " SELECT SalesMaster.CMID, SalesMaster.SDate, SalesMaster.CName, SalesMaster.CAddress, " & _
                  " SalesMaster.RegNo,SalesMaster.BedNo, SalesMaster.Tamount,SalesMaster.DiscuntPer, SalesMaster.DiscuntTaka, SalesMaster.Ttime, " & _
                  " SalesMaster.Tpaid, SalesMaster.Tdue, SalesMaster.Posted, SalesMaster.UName, SalesMaster.Admitted, SalesMaster.Remark, " & _
                  " SalesDetail.MCatagory,SalesDetail.MName, SalesDetail.Qty, SalesDetail.SRate, SalesDetail.Amount,SalesDetail.Discount,(SalesDetail.Amount*SalesDetail.Discount)/100 as TDiscount " & _
                  " FROM SalesMaster INNER JOIN " & _
                  " SalesDetail ON SalesMaster.CMID = SalesDetail.CMID " & _
                  " And SalesMaster.RegNo = '" & parseQuotes(txtRegNo) & "'", cn, adOpenStatic, adLockReadOnly

           End If

End Function

Public Sub CashMemo()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Data not available", vbInformation
        Exit Sub
    End If

        strPath = App.Path + "\reports\Indoor Cash Memo.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields


        If optCashMemo.Value = True Then
   
   Set objReportFF = objReportFormulaFieldDefinations.Item(5)
              objReportFF.text = "'" + txtRegNo + "'"


  End If


        objReportDatabaseTable.SetPrivateData 3, rsMaster

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        objReport.Preview "Indoor Cash Memo", , , , , 16777216 Or 524288 Or 65536


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

Public Function FetchData1()

    Set rsMaster = New ADODB.Recordset

    If optMReciept.Value = True Then


    rsMaster.Open "SELECT  PC_ID, CMID, CDate, CTime, RegNo, Amount, ComputerName, Particulars, UName " & _
                    "From Medicine_Payment " & _
                    "WHERE    CDate='" & MRDate.Value & "' AND RegNo = '" & parseQuotes(txtRegNo) & "'", cn, adOpenStatic, adLockReadOnly

     End If

End Function
'
Public Sub MoneyReceipt()
On Error GoTo ErrH
    Dim strPath As String

    If rsMaster.RecordCount = 0 Then
        MsgBox "Payment not available", vbInformation
        Exit Sub
    End If

        strPath = App.Path + "\reports\Money Receipt.rpt"
        Set objReportApp = CreateObject("Crystal.CRPE.Application")
        Set objReport = objReportApp.OpenReport(strPath)
        Set objReportDatabase = objReport.Database
        Set objReportDatabaseTables = objReportDatabase.Tables
        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
        Set ObjPrinterSetting = objReport.PrintWindowOptions
        Set objReportFormulaFieldDefinations = objReport.FormulaFields


        If optCashMemo.Value = True Then

   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
              objReportFF.text = "'" + Format(MRDate, "dd-MM-yyyy") + "'"
   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
              objReportFF.text = "'" + txtRegNo + "'"


  End If


        objReportDatabaseTable.SetPrivateData 3, rsMaster

        ObjPrinterSetting.HasPrintSetupButton = True
        ObjPrinterSetting.HasRefreshButton = True
        ObjPrinterSetting.HasSearchButton = True
        ObjPrinterSetting.HasZoomControl = True

        objReport.DiscardSavedData
        objReport.Preview "Patient Money Receipt", , , , , 16777216 Or 524288 Or 65536


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
            MsgBox "Request cancelled by the user", vbInformation, "Patient Money Receipt"
        Case Else
            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Patient Money Receipt"
    End Select
End Sub
'
'Public Function FetchData2()
'
'    Set rsMaster = New ADODB.Recordset
'
'    If OpCurrentDate.Value = True Then
'
'
'    rsMaster.Open "SELECT  SerialNo,CMID, SName,MCatagory, MName, Qty, SRate, Amount,PRate, Discount,(Amount * Discount / 100) AS TDiscount, (Amount - (Amount * Discount / 100)) AS NetAmount, Posted, UName, SDate " & _
'                    "From SalesDetail " & _
'                    "WHERE    SDate='" & CurrentDate.Value & "' AND Posted = 'Posted' and SalesDetail.SName='" & parseQuotes(cmbSName) & "'", cn, adOpenStatic, adLockReadOnly
'
'     End If
'
'      If OpCustomDate.Value = True Then
'
'
'     rsMaster.Open "SELECT SerialNo,CMID,SName, MCatagory, MName, Qty, SRate, Amount,PRate, Discount,(Amount * Discount / 100) AS TDiscount, (Amount - (Amount * Discount / 100)) AS NetAmount, Posted, UName, SDate " & _
'                    "From SalesDetail " & _
'                    "WHERE  SDate BETWEEN '" & FDate.Value & "' AND " & _
'                    "'" & TDate.Value & "' AND Posted = 'Posted' and SalesDetail.Sname='" & parseQuotes(cmbSName) & "'", cn, adOpenStatic, adLockReadOnly
'
'      End If
'
'End Function
'
'Public Sub previewReport2()
'On Error GoTo ErrH
'    Dim strPath As String
'
'    If rsMaster.RecordCount = 0 Then
'        MsgBox "Data not available", vbInformation
'        Exit Sub
'    End If
'
'
'        strPath = App.Path + "\reports\Supplier Sales Statement.rpt"
'        Set objReportApp = CreateObject("Crystal.CRPE.Application")
'        Set objReport = objReportApp.OpenReport(strPath)
'        Set objReportDatabase = objReport.Database
'        Set objReportDatabaseTables = objReportDatabase.Tables
'        Set objReportDatabaseTable = objReportDatabaseTables.Item(1)
'        Set ObjPrinterSetting = objReport.PrintWindowOptions
'        Set objReportFormulaFieldDefinations = objReport.FormulaFields
'
'
'    If OpCurrentDate.Value = True Then
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(2)
'              objReportFF.text = "'" + Format(CurrentDate, "dd-MMM-yyyy") + "'"
'   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'              objReportFF.text = "'" + cmbSName + "'"
'
'
'  End If
'
'
'  If OpCustomDate.Value = True Then
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(3)
'
'              objReportFF.text = "'" + Format(FDate, "dd-MMM-yyyy") + "'"
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(4)
'             objReportFF.text = "'" + Format(TDate, "dd-MMM-yyyy") + "'"
'
'   Set objReportFF = objReportFormulaFieldDefinations.Item(1)
'              objReportFF.text = "'" + cmbSName + "'"
'
'
'   End If
'
'
'        objReportDatabaseTable.SetPrivateData 3, rsMaster
'
'        ObjPrinterSetting.HasPrintSetupButton = True
'        ObjPrinterSetting.HasRefreshButton = True
'        ObjPrinterSetting.HasSearchButton = True
'        ObjPrinterSetting.HasZoomControl = True
'
'        objReport.DiscardSavedData
'        objReport.Preview "Sales Insformations", , , , , 16777216 Or 524288 Or 65536
'
'
'     If Tracer = 1 Then
'    objReport.PrintOut
'    End If
'
'        Set objReport = Nothing
'        Set objReportDatabase = Nothing
'        Set objReportDatabaseTables = Nothing
'        Set objReportDatabaseTable = Nothing
'    Exit Sub
'
'ErrH:
'
'    Select Case Err.Number
'        Case 20545
'            MsgBox "Request cancelled by the user", vbInformation, "Bank Information Report"
'        Case Else
'            MsgBox "Error " & Err.Number & " - " & Err.Description, vbCritical, "Bank Information Report"
'    End Select
'End Sub
