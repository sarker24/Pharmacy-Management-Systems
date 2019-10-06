VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesMasterSearch 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Master Search"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "frmSalesMasterSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   7080
      Width           =   735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   6255
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   11175
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   5805
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   10920
         _cx             =   19262
         _cy             =   10239
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSalesMasterSearch.frx":058A
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
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      Height          =   750
      Left            =   9480
      Picture         =   "frmSalesMasterSearch.frx":06B6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Ok"
      Height          =   750
      Left            =   8520
      Picture         =   "frmSalesMasterSearch.frx":0F80
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Refresh"
      Height          =   750
      Left            =   7440
      Picture         =   "frmSalesMasterSearch.frx":184A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6840
      Width           =   1095
   End
   Begin VB.ComboBox cboMode 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Text            =   " "
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   7080
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   7080
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
      Format          =   49938435
      CurrentDate     =   41840
   End
   Begin VB.Label Label58 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SALES MASTER  SEARCH  "
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11265
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
      Left            =   240
      TabIndex        =   9
      Top             =   6840
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
      TabIndex        =   8
      Top             =   6840
      Width           =   1815
   End
End
Attribute VB_Name = "frmSalesMasterSearch"
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

Private Sub cmdOk_Click()
    If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Row  From the List."
        Exit Sub
    End If

     Call PopulateCompanySearch

Unload Me
Set frmSalesMasterSearch = Nothing
End Sub

Private Sub fgExport_Click()
cmdOk_Click
End Sub

Private Sub cmdFind_Click()

If rsTemp.State <> 0 Then rsTemp.Close
        
       
If cboMode.text = "Bill Number" Then

      rsTemp.Open "SELECT TOP 50 CMID,RegNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                 "FROM SalesMaster WHERE SalesMaster.CMID LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
        
ElseIf cboMode.text = "Customer Name" Then

      rsTemp.Open "SELECT TOP 50 CMID,RegNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                 "FROM SalesMaster WHERE SalesMaster.CName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Reg No" Then

      rsTemp.Open "SELECT TOP 50 CMID,RegNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                 "FROM SalesMaster WHERE SalesMaster.RegNo LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly

ElseIf cboMode.text = "Date Search" Then
      
      rsTemp.Open "SELECT TOP 50 CMID,RegNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                 "FROM SalesMaster WHERE SalesMaster.SDate = '" & dtDate.Value & "'", cn, adOpenStatic, adLockReadOnly

Else
 rsTemp.Open "SELECT TOP 50 CMID,RegNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted " & _
                 "FROM SalesMaster WHERE SalesMaster.Posted LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
End If

'   rsTemp.Open
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
       fgExport.AddItem "" & vbTab & rsTemp("CMID") & vbTab & rsTemp("RegNo") & vbTab & rsTemp("CName") & _
         vbTab & Format(rsTemp("SDate"), "dd-MM-yyyy") & vbTab & rsTemp("Tamount") & vbTab & rsTemp("DiscuntTaka") & vbTab & rsTemp("Tpaid") & vbTab & rsTemp("Tdue") & _
         vbTab & rsTemp("Posted")
        rsTemp.MoveNext
        Wend


End Sub

Private Sub Form_Load()

     ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     fgExport.Editable = flexEDKbdMouse
     
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
       
     rsTemp.Open "SELECT TOP 5 CMID,RegNo,CName,SDate,Tamount,DiscuntTaka,Tpaid,Tdue,Posted FROM SalesMaster ORDER BY CMID DESC", cn, adOpenStatic, adLockReadOnly
         
         fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("CMID") & vbTab & rsTemp("RegNo") & vbTab & rsTemp("CName") & _
         vbTab & Format(rsTemp("SDate"), "dd-MM-yyyy") & vbTab & rsTemp("Tamount") & vbTab & rsTemp("DiscuntTaka") & vbTab & rsTemp("Tpaid") & vbTab & rsTemp("Tdue") & _
         vbTab & rsTemp("Posted")
        rsTemp.MoveNext
    Wend
     GridCount fgExport
       
       cboMode.AddItem "Bill Number"
       cboMode.AddItem "Reg No"
       cboMode.AddItem "Customer Name"
       cboMode.AddItem "Date Search"
       cboMode.AddItem "Not Posted"
       cboMode.text = "Bill Number"

       dtDate.Value = Date

End Sub

    Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then

             frmSales.PopulateForm fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys Chr(9)
End If
End Sub

Private Sub fgExport_KeyPress(KeyAscii As Integer)
fgExport_Click
End Sub

Private Sub txtSearch_Change()
 cmdFind_Click
Text1.text = txtSearch.text
End Sub

Private Sub cmdFind1_Click()

 Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     If rsTemp.State <> 0 Then rsTemp.Close
'       rsTemp.CursorLocation = adUseClient
      rsTemp.Open "SELECT CMID,RegNo,CName,SDate,Posted " & _
                 " " & _
                 " FROM SalesMaster where SDate LIKE '" & txtSearch.text & "%'order by SDate", cn, adOpenStatic, adLockReadOnly

'   End If

         fgExport.Rows = 1

     While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("CMID") & vbTab & rsTemp("RegNo") & vbTab & rsTemp("CName") & vbTab & Format(rsTemp("SDate"), "dd-MM-yyyy") & _
         vbTab & rsTemp("Posted")

        rsTemp.MoveNext
    Wend
        GridCount fgExport

End Sub


