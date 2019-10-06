VERSION 5.00
Object = "{8D650146-6025-11D1-BC40-0000C042AEC0}#3.0#0"; "ssdw3a32.ocx"
Begin VB.Form frmMedicineCatagory 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Medicine Catagory"
   ClientHeight    =   3480
   ClientLeft      =   1395
   ClientTop       =   945
   ClientWidth     =   5670
   Icon            =   "frmMedicineCatagory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5670
   Begin SSDataWidgets_A.SSDBCommand cmdPreview 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   2160
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5415
      Begin VB.TextBox txtMCID 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox txtMCatagory 
         Appearance      =   0  'Flat
         Height          =   465
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   11
         Text            =   " "
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Medicine Catagory"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Catagory ID"
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
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin SSDataWidgets_A.SSDBCommand cmdEdit 
      Height          =   495
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
      Height          =   495
      Left            =   2040
      TabIndex        =   5
      Top             =   2760
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2160
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
   Begin SSDataWidgets_A.SSDBCommand cmdCancel 
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
   Begin SSDataWidgets_A.SSDBCommand cmdQuit 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   2760
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
      BevelColorFace  =   12632064
      Font3D          =   3
      CaptionAlignment=   7
   End
   Begin SSDataWidgets_A.SSDBCommand cmdNew 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   975
      _Version        =   196612
      _ExtentX        =   1720
      _ExtentY        =   873
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
End
Attribute VB_Name = "frmMedicineCatagory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rs                     As ADODB.Recordset
Private rsMCatagory            As ADODB.Recordset
'Private strStream             As ADODB.Stream
Private strFileName            As String
Private bRecordExists          As Boolean
Dim str                        As String
'Private rm                    As New ADODB.Recordset
'Private rc                    As New ADODB.Recordset


Private Sub cmdCancel_AfterClick()

   cmdCancel.Enabled = False
   cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    cmdQuit.Enabled = True
    cmdEdit.Enabled = True
    cmdFind.Enabled = True
    cmdDelete.Enabled = True
    cmdPrint.Enabled = True
    cmdPreview.Enabled = True
    txtMCatagory.Enabled = False
    Call allClear
'    txtCompanyID.Enabled = False
    Call alldisable
End Sub

Private Sub cmdDelete_AfterClick()
On Error GoTo ErrHandler
     Dim idelete As Integer
    idelete = MsgBox("Do you want to delete this record?", vbYesNo)
    If idelete = vbYes Then
            cn.Execute "Delete From MedicineCatagory Where MCID ='" & parseQuotes(txtMCID) & "'"
            Call allClear
    End If
ErrHandler:
    Select Case Err.Number
        Case -2147217913
            MsgBox "Please select record first for delete", vbInformation, "Confirmation"
     End Select
End Sub

Private Sub cmdFind_AfterClick()
 strCallingForm = LCase("frmMedicineCatagory")
    frmMCatagorySearch.Show vbModal
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
        cmdPreview.Enabled = False
        cmdPrint.Enabled = False
'        txtMCatagory.Enabled = True
        Call allClear
'       ModFunction.TextEnable Me, True
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(MCID),0) as SerialNo from MedicineCatagory"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtMCID.text = Val(rs!SerialNo) + 1
           Call allenable
'           Call alldisable
        txtMCatagory.SetFocus

    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then

                txtMCatagory.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                cmdQuit.Enabled = True
                cmdFind.Enabled = True
                cmdDelete.Enabled = True
                cmdPreview.Enabled = True
                cmdPrint.Enabled = True
'                ModFunction.TextEnable Me, False
                Call alldisable
                s = txtMCatagory
                rsMCatagory.Requery
                rsMCatagory.MoveFirst
                rsMCatagory.Find "MCatagory='" & parseQuotes(s) & "'"
               
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

Private Sub cmdEdit_AfterClick()
    If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtMCatagory.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        cmdQuit.Enabled = False
        cmdFind.Enabled = False
        cmdPreview.Enabled = False
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
                cmdPreview.Enabled = True
                cmdDelete.Enabled = True
                cmdPrint.Enabled = True
                rsMCatagory.Requery

                Dim s As String
                s = txtMCatagory
                rsMCatagory.Find "MCatagory='" & parseQuotes(s) & "'"
                
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
    Set rsMCatagory = New ADODB.Recordset
'    Set rsImage = New ADODB.Recordset
    rsMCatagory.Open "select  DISTINCT * from MedicineCatagory", cn, adOpenStatic, adLockReadOnly
    
ModFunction.TextEnable Me, False
    
    Call alldisable
    Call DeleteVisible

   If rsMCatagory.RecordCount > 0 Then

        bRecordExists = True
    Else
        bRecordExists = False
    End If
    
   If Not rsMCatagory.EOF Then FindRecord
    
    txtMCatagory.Enabled = False
    
End Sub

Private Sub DeleteVisible()
Dim str As String
Set rs = New ADODB.Recordset
str = "select UName,Password,Upper(UName)as Name  from PMSUser where UName ='" & frmLogin.txtUName.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
           If rs!Name = "ADMIN" And rs!Password = "123" Then
              cmdDelete.Visible = True
            
        ElseIf rs!Name = "BORHAN" And rs!Password = "01920468031" Then
        
              cmdDelete.Visible = True
           Else
               cmdDelete.Visible = False
               
           End If
End Sub

Private Sub allClear()
'    ModFunction.TextClear Me
txtMCatagory.text = ""
End Sub

Private Function rcupdate() As Boolean

     On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then


        
        cn.Execute "INSERT INTO MedicineCatagory(MCID,MCatagory) " & _
                   " VALUES ('" & parseQuotes(txtMCID) & "','" & parseQuotes(txtMCatagory) & "')"
                   

          rcupdate = True
          MsgBox "Record Added", vbInformation, "Confirmation"
    Else

        cn.Execute "Update MedicineCatagory Set MCatagory='" & parseQuotes(txtMCatagory) & _
                  "'WHERE  MCID ='" & parseQuotes(txtMCID) & "' "
                      
        rcupdate = True
        MsgBox "Record Updated", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsMedicineCatagory.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate Medicine Catagory Name"
            txtMCatagory = ""
            txtMCatagory.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function
Public Sub FindRecord()
If Not rsMCatagory.EOF Then
        txtMCID = rsMCatagory("MCID")
        txtMCatagory = rsMCatagory("MCatagory")
        
   End If
End Sub


Private Sub allenable()
    txtMCatagory.Enabled = True
    
End Sub

Private Sub alldisable()
    txtMCatagory.Enabled = False
End Sub


Private Function IsValidRecord() As Boolean
    IsValidRecord = True


    If (txtMCatagory.text = "") Then
       MsgBox "Enter Medicine Catagory Name"
       txtMCatagory.SetFocus
       IsValidRecord = False
       Exit Function
       
       End If

    If (txtMCID.text = "") Then
     MsgBox "Enter MCID"
     txtMCID.SetFocus
     IsValidRecord = False
     Exit Function
      
    End If
    
    If cmdEdit.Caption <> "&Update" Or cmdEdit.Caption = "&Update" Then
        If rsMCatagory.RecordCount > 0 Then
        If rsMCatagory.State <> 0 Then rsMCatagory.Close
            rsMCatagory.Open "select * from MedicineCatagory where upper(MCatagory)='" & Strings.UCase(Strings.Trim(parseQuotes(txtMCatagory))) & "'", cn

             If Not rsMCatagory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          txtMCatagory.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
    
End Function

Public Sub PopulateMedicineCatagory(StrID As String)

    rsMCatagory.MoveFirst
    rsMCatagory.Find "MCID=" & parseQuotes(StrID)
    If rsMCatagory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub


