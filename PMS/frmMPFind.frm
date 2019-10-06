VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMPFind 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Purchase Search."
   ClientHeight    =   8775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "frmMPFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8775
   ScaleWidth      =   10695
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Text            =   " "
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   8040
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   750
      Left            =   9600
      Picture         =   "frmMPFind.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   8640
      Picture         =   "frmMPFind.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7920
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   7560
      Picture         =   "frmMPFind.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7920
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   7575
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   10455
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   7245
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10200
         _cx             =   17992
         _cy             =   12779
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMPFind.frx":1A6A
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
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   7920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MM-yyyyy"
      Format          =   65798147
      CurrentDate     =   41840
   End
   Begin VB.Label lblIteamGroup 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Field Name"
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
      Left            =   120
      TabIndex        =   8
      Top             =   7800
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Available Value"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   7800
      Width           =   1815
   End
End
Attribute VB_Name = "frmMPFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
  
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
If cboMode.text = "Bill Number" Then

      rsTemp.Open "SELECT TOP 50 SerialNo,Pdate,Sname,SbillNo,Posted " & _
                 "FROM PurchaseMaster WHERE PurchaseMaster.SerialNo LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
ElseIf cboMode.text = "Date Search" Then
      
      rsTemp.Open "SELECT TOP 50 SerialNo,Pdate,Sname,SbillNo,Posted " & _
                 "FROM PurchaseMaster WHERE PurchaseMaster.Pdate = '" & dtDate.Value & "'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "SBill No" Then

      rsTemp.Open "SELECT TOP 50 SerialNo,Pdate,Sname,SbillNo,Posted " & _
                 "FROM PurchaseMaster WHERE PurchaseMaster.SbillNo LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Supplier Name" Then

      rsTemp.Open "SELECT TOP 50 SerialNo,Pdate,Sname,SbillNo,Posted " & _
                 "FROM PurchaseMaster WHERE PurchaseMaster.Sname LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
Else

      rsTemp.Open "SELECT TOP 50 SerialNo,Pdate,Sname,SbillNo,Posted " & _
                 "FROM PurchaseMaster WHERE PurchaseMaster.Posted LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

End If

'   rsTemp.Open
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
       fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("Sname") & vbTab & Format(rsTemp("Pdate"), "dd-mmm-yyyy") & _
         vbTab & rsTemp("SbillNo") & vbTab & rsTemp("Posted")
         
        rsTemp.MoveNext
        Wend

End Sub

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Bill No From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me

Set frmPurchaseSearch = Nothing
End Sub

Private Sub dtDate_click()
cmdFind_Click
End Sub

Private Sub fgExport_Click()
cmdOk_Click
End Sub

Private Sub fgExport_KeyPress(KeyAscii As Integer)
fgExport_Click
End Sub

Private Sub Form_Load()
     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     fgExport.Editable = flexEDKbdMouse
     
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
     rsTemp.Open "SELECT TOP 5 SerialNo,Pdate,Sname,SbillNo,Posted FROM PurchaseMaster ORDER BY SerialNo DESC", cn, adOpenStatic, adLockReadOnly
         
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SerialNo") & vbTab & rsTemp("Sname") & vbTab & Format(rsTemp("Pdate"), "dd-mmm-yyyy") & _
         vbTab & rsTemp("SbillNo") & vbTab & rsTemp("Posted")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
'     If fgExport.Rows = 1 Then fgExport.AddItem ""

       cboMode.AddItem "Bill Number"
       cboMode.AddItem "Date Search"
       cboMode.AddItem "Supplier Name"
       cboMode.AddItem "SBill No"
       cboMode.AddItem "Posted"
       cboMode.text = "Bill Number"

       dtDate.Value = Date
       

End Sub
  
    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmPurchase.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub
    
Private Sub txtSearch_Change()
cmdFind_Click
End Sub


Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys Chr(9)

End If
End Sub
