VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmMedicineRemarks 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Remarks "
   ClientHeight    =   2745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      Begin VB.TextBox txtRID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtRName 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "SutonnyMJ"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label lblGenericName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Medicine Remarks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label lblRID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Remarks ID"
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
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPreview 
      Height          =   615
      Left            =   4800
      TabIndex        =   5
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "Pre&view"
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
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdPrint 
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Print"
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
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdEdit 
      Height          =   615
      Left            =   960
      TabIndex        =   7
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Edit"
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
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdDelete 
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Delete"
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
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdFind 
      Height          =   615
      Left            =   1920
      TabIndex        =   9
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Find"
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
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   615
      Left            =   5760
      TabIndex        =   10
      Top             =   2040
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
      BevelColorFace  =   12632064
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&New"
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
      Font3D          =   3
      WordWrap        =   0   'False
      CaptionAlignment=   7
      PictureAlignment=   0
   End
   Begin SSDataWidgets_A.SSDBCommand cmdCancel 
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   2040
      Width           =   975
      _Version        =   196611
      _ExtentX        =   1720
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "&Cancel"
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
      Font3D          =   3
      CaptionAlignment=   7
   End
End
Attribute VB_Name = "frmMedicineRemarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs                     As ADODB.Recordset
Private rsRName               As ADODB.Recordset
'Private strStream             As ADODB.Stream
Private strFileName            As String
Private bRecordExists          As Boolean
Dim str                        As String
'Private rm                    As New ADODB.Recordset
'Private rc                    As New ADODB.Recordset


Private Sub cmdCancel_Click()

   cmdCancel.Enabled = False
   cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdQuit.Enabled = True
    cmdEdit.Enabled = True
    cmdFind.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
'    cmdPreview.Enabled = True
    txtRName.Enabled = False
    Call allClear
'    txtCompanyID.Enabled = False
    Call alldisable
End Sub

Private Sub cmdDelete_Click()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If frmLogin.txtUName.text = "Admin" Then
    If idelete = vbYes Then
            cn.Execute "Delete From tblRemarks Where RID ='" & parseQuotes(txtRID) & "'"
            Call allClear
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please call your system Administrator.", vbInformation, "Confirmation"
     End Select
     End If
End Sub

Private Sub cmdFind_Click()
 strCallingForm = LCase("frmRemarks")
    frmGNameSearch.Show vbModal
    cmdFind.Enabled = True
    cmdCancel.Enabled = True
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
       Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        cmdQuit.Enabled = False
        cmdFind.Enabled = False
        cmdDelete.Enabled = False
'        cmdPreview.Enabled = False
        cmdPrint.Enabled = False
'        txtRName.Enabled = True
        Call allClear
'       ModFunction.TextEnable Me, True
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(RID),0) as SerialNo from tblRemarks"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtRID.text = Val(rs!SerialNo) + 1
           Call allenable
'           Call alldisable
        txtRName.SetFocus

    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                txtRName.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdQuit.Enabled = True
                cmdFind.Enabled = True
                cmdDelete.Enabled = True
'                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
'                ModFunction.TextEnable Me, False
                Call alldisable
                s = txtRName
                rsRName.Requery
                rsRName.MoveFirst
                rsRName.Find "Remarks='" & parseQuotes(s) & "'"
               
                FindRecord
            End If
        End If
    End If

    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Sub cmdEdit_Click()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtRName.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdQuit.Enabled = False
        cmdFind.Enabled = False
'        cmdPreview.Enabled = False
        cmdDelete.Enabled = False
        cmdPrint.Enabled = False

    ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                cmdQuit.Enabled = True
                cmdFind.Enabled = True
'                cmdPreview.Enabled = True
                cmdDelete.Enabled = True
                cmdPrint.Enabled = True
                rsRName.Requery

                Dim s As String
                s = txtRName
                rsRName.Find "Remarks='" & parseQuotes(s) & "'"
                
                FindRecord
            End If
        End If
    End If
End Sub


Private Sub cmdQuit_AfterClick()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
   If (KeyCode = 13 And Me.ActiveControl.Name <> "txtAddress") Then SendKeys "{TAB}", True
End Sub

Private Sub Form_Load()

    Call Connect
    ModFunction.StartUpPosition Me
    Set rsRName = New ADODB.Recordset
'    Set rsImage = New ADODB.Recordset
    rsRName.Open "select  DISTINCT * from tblRemarks", cn, adOpenStatic, adLockReadOnly
    
ModFunction.TextEnable Me, False
    
    Call alldisable

   If rsRName.RecordCount > 0 Then

        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
   If Not rsRName.EOF Then FindRecord
    
    txtRName.Enabled = False
    
End Sub

Private Sub allClear()
'    ModFunction.TextClear Me
txtRName.text = ""
End Sub

Private Function rcupdate() As Boolean

     On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO tblRemarks(RID,Remarks) " & _
                   " VALUES ('" & parseQuotes(txtRID) & "','" & parseQuotes(txtRName) & "')"
                   

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

        cn.Execute "Update tblRemarks Set Remarks='" & parseQuotes(txtRName) & _
                  "'WHERE  RID ='" & parseQuotes(txtRID) & "' "
                      
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rstblRemarks.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Medicine Catagory Name"
            txtRName = ""
            txtRName.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsRName.EOF Then
        txtRID = rsRName("RID")
        txtRName = rsRName("Remarks")
        
   End If
End Sub


Private Sub allenable()
    txtRName.Enabled = True
    
End Sub

Private Sub alldisable()
    txtRName.Enabled = False
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtRName.text = "") Then
       MsgBox "Enter Medicine Catagory Name"
       txtRName.SetFocus
       IsValidRecord = False
       Exit Function
       
       End If

    If (txtRID.text = "") Then
     MsgBox "Enter RID"
     txtRID.SetFocus
     IsValidRecord = False
     Exit Function
      
    End If
    
    If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
        If rsRName.RecordCount > 0 Then
        If rsRName.State <> 0 Then rsRName.Close
            rsRName.Open "select * from tblRemarks where upper(Remarks)='" & Strings.UCase(Strings.Trim(parseQuotes(txtRName))) & "'", cn

             If Not rsRName.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtRName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
    
End Function

Public Sub PopulatetblRemarks(StrID As String)

    rsRName.MoveFirst
    rsRName.Find "RID=" & parseQuotes(StrID)
    If rsRName.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub






