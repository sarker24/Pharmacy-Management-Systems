VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmMFind 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Find"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9720
   Icon            =   "frmMFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6540
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   5880
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Medicine Details Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9615
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   5130
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9465
         _cx             =   16695
         _cy             =   9049
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
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
         BackColorFixed  =   14737632
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
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMFind.frx":08CA
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
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   5880
      Width           =   1095
      _Version        =   196611
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
      BevelColorFace  =   12632064
      Font3D          =   4
   End
   Begin SSDataWidgets_A.SSDBCommand cmdOk 
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   5880
      Width           =   1095
      _Version        =   196611
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Ok"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12632064
      Font3D          =   4
   End
   Begin SSDataWidgets_A.SSDBCommand cmdRefresh 
      Height          =   495
      Left            =   6360
      TabIndex        =   5
      Top             =   5880
      Width           =   1095
      _Version        =   196611
      _ExtentX        =   1931
      _ExtentY        =   873
      _StockProps     =   78
      Caption         =   "&Refresh"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelColorFace  =   12632064
      Font3D          =   4
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   " Enter Medicince Name"
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
      Top             =   5880
      Width           =   2655
   End
End
Attribute VB_Name = "frmMFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset

Private Sub cmdOk_Click()
 If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Medicine Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmMFind = Nothing
End Sub

Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmMedicineName.PopulateProduct fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub cmdFind_Click()
If rsTemp.State <> 0 Then rsTemp.Close
     rsTemp.Open "SELECT tblMedicineName.MID,tblMedicineName.MName,tblMedicineName.SName" & _
                 " FROM tblMedicineName WHERE tblMedicineName.MName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
                 fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("MID") & vbTab & rsTemp("MName") & vbTab & rsTemp("SName")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

Private Sub cmdQuit_Click()
Unload Me
End Sub

Private Sub fgExport_DblClick()
    cmdOk_Click
End Sub

Private Sub Form_Load()
 ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
'     rsTemp.Open "SELECT SupplierID,SName,ContactName,Address,PhoneNo,FaxNo,City,Country, " & _
'                 "Email FROM Suppliers", cn, adOpenStatic, adLockReadOnly
      rsTemp.Open "exec Sms_Medicine_Name_Search 1", cn, adOpenStatic, adLockReadOnly
        
            fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("MID") & vbTab & rsTemp("MName") & _
        vbTab & rsTemp("SName") & vbTab & rsTemp("CPost")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub
Private Sub txtSearch_Change()
 cmdFind_Click
End Sub




