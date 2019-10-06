VERSION 5.00
Begin VB.Form frmProductSearch 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Product"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmProductSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Private rsTemp                      As ADODB.Recordset
'Private rsExport                    As ADODB.Recordset
'Private rsfactory                   As New ADODB.Recordset
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdFind_Click()
'Set rsTemp = New ADODB.Recordset
'
'    If rsTemp.State <> 0 Then rsTemp.Close
'
'
'     rsTemp.Open "SELECT USProduct.ProductID[ProductID],USProduct.PCatagory[PCatagory],USProduct.ProductDescription[ProductDescription]" & _
'                 " FROM USProduct WHERE USProduct.PCatagory LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
'                 fgExport.Rows = 1
'
'    While Not rsTemp.EOF
'        fgExport.AddItem "" & vbTab & rsTemp("ProductID") & vbTab & rsTemp("PCatagory") & vbTab & rsTemp("ProductDescription")
'
'        rsTemp.MoveNext
'    Wend
'     GridCount fgExport
'
'End Sub
'
'Private Sub cmdOK_Click()
'    If fgExport.RowSel < 0 Then
'        MsgBox "Please Select a Product Name From the List."
'        Exit Sub
'    End If
'
'     Call PopulateCompanySearch
'
'     Unload Me
'     Set frmProductSearch = Nothing
'End Sub
'
'Private Sub fgExport_DblClick()
'    cmdOK_Click
'End Sub
'
'Private Sub Form_Load()
'     ModFunction.StartUpPosition Me
'     Set rsTemp = New ADODB.Recordset
'     rsTemp.CursorLocation = adUseClient
'
'        If rsTemp.State <> 0 Then rsTemp.Close
'
'             rsTemp.Open "SELECT ProductID,PCatagory,ProductDescription FROM USProduct order by PCatagory", cn, adOpenStatic, adLockReadOnly
'
''              End If
'
'
'         fgExport.Rows = 1
'             While Not rsTemp.EOF
'             fgExport.AddItem "" & vbTab & rsTemp("ProductID") & vbTab & rsTemp("PCatagory") & vbTab & rsTemp("ProductDescription")
'             rsTemp.MoveNext
'        Wend
'
'        GridCount fgExport
'    '    End If
''        End If
'
'
'End Sub
'
'    Private Sub PopulateCompanySearch()
'
'          If fgExport.Row > 0 Then
'             frmProductName.PopulateProduct fgExport.TextMatrix(fgExport.Row, 1)
'          End If
'   End Sub
'
'
'Private Sub txtSearch_Change()
'cmdFind_Click
'End Sub
