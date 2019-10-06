VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmSupplierSearch 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Supplier Search"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10395
   ControlBox      =   0   'False
   Icon            =   "frmSupplierSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   735
      Left            =   9000
      TabIndex        =   5
      Top             =   7440
      Width           =   1095
      _Version        =   196612
      _ExtentX        =   1931
      _ExtentY        =   1296
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
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   7560
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Supplier Details Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   10335
      Begin VSFlex7LCtl.VSFlexGrid fgExport 
         Height          =   6210
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10065
         _cx             =   17754
         _cy             =   10954
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSupplierSearch.frx":058A
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
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000001&
      Caption         =   "Supplier Details Search"
      Default         =   -1  'True
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      MousePointer    =   1  'Arrow
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   10575
   End
   Begin SSDataWidgets_A.SSDBCommand cmdOk 
      Height          =   735
      Left            =   7920
      TabIndex        =   6
      Top             =   7440
      Width           =   1095
      _Version        =   196612
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      Height          =   735
      Left            =   6840
      TabIndex        =   7
      Top             =   7440
      Width           =   1095
      _Version        =   196612
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      Caption         =   " Enter Supplier Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   7560
      Width           =   2655
   End
End
Attribute VB_Name = "frmSupplierSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsTemp                      As ADODB.Recordset
Private rsExport                    As ADODB.Recordset
Private rsfactory                   As New ADODB.Recordset

Private Sub cmdFind_AfterClick()
If rsTemp.State <> 0 Then rsTemp.Close
     rsTemp.Open "SELECT Suppliers.SupplierID[SupplierID],Suppliers.SName[SName]" & _
                 " FROM Suppliers WHERE Suppliers.SName LIKE '" & txtSearch.text & "%'", cn, adOpenStatic, adLockReadOnly
                 fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SupplierID") & vbTab & rsTemp("SName")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub

Private Sub cmdOk_AfterClick()
 If fgExport.RowSel < 0 Then
        MsgBox "Please Select a Supplier Name From the List."
        Exit Sub
    End If
     
     Call PopulateCompanySearch
   
Unload Me
Set frmSupplierSearch = Nothing
End Sub

Private Sub PopulateCompanySearch()
        If fgExport.Row > 0 Then
        
             frmSupplier.PopulateCnf fgExport.TextMatrix(fgExport.Row, 1)
        End If
    End Sub

Private Sub cmdQuit_AfterClick()
Unload Me
End Sub


Private Sub Form_Load()
 ModFunction.StartUpPosition Me
     Set rsTemp = New ADODB.Recordset
     rsTemp.CursorLocation = adUseClient
     
        If rsTemp.State <> 0 Then rsTemp.Close
        
     rsTemp.Open "SELECT SupplierID,SName,ContactName,Address,PhoneNo,FaxNo,City,Country, " & _
                 "Email FROM Suppliers", cn, adOpenStatic, adLockReadOnly
        
            fgExport.Rows = 1
    
    While Not rsTemp.EOF
        fgExport.AddItem "" & vbTab & rsTemp("SupplierID") & vbTab & rsTemp("SName") & _
        vbTab & rsTemp("ContactName") & vbTab & rsTemp("Address") & vbTab & rsTemp("PhoneNo") & _
        vbTab & rsTemp("FaxNo") & vbTab & rsTemp("City") & _
        vbTab & rsTemp("Country") & vbTab & rsTemp("Email")
         
        rsTemp.MoveNext
    Wend
     GridCount fgExport
End Sub


Private Sub txtSearch_Change()
 cmdFind_AfterClick
End Sub
