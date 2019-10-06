VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPReg 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Indoor Patient Registration"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7950
   Icon            =   "frmPReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H00C0C000&
      Caption         =   ">"
      Height          =   495
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Find Next"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdLast 
      BackColor       =   &H00C0C000&
      Caption         =   ">>|"
      Height          =   495
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Find Last"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H00C0C000&
      Caption         =   "<"
      Height          =   495
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Find Previous"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton cmdFirst 
      BackColor       =   &H00C0C000&
      Caption         =   "|<<"
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Find First"
      Top             =   3120
      Width           =   735
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Q&uit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   945
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3600
      Width           =   945
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3600
      Width           =   945
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Width           =   945
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3600
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3600
      Width           =   945
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   7695
      Begin VB.TextBox txtPName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         MaxLength       =   50
         TabIndex        =   0
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txtRegNo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtPhoneNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   4095
      End
      Begin VB.TextBox txtPID 
         Height          =   375
         Left            =   3120
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "Member Serial No"
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblMName 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label lblMRID 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Patient Registration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblMPhone 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Contact Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   4095
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3000
      Top             =   3240
      Visible         =   0   'False
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   2
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=SMS;Data Source=NOTEBOOK"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DCSearch"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmPReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsfactory             As ADODB.Recordset
Private strFileName           As String
Private bRecordExists         As Boolean
Private rm                    As New ADODB.Recordset
Private rs                    As New ADODB.Recordset
Dim str As String


Private Sub alldisable()
    txtPID.Enabled = False
    txtRegNo.Enabled = False
    txtPName.Enabled = False
    txtPhoneNo.Enabled = False
End Sub

Private Sub allenable()
    txtRegNo.Enabled = True
    txtPName.Enabled = True
    txtPID.Enabled = True
    txtPhoneNo.Enabled = True
End Sub
'
Public Sub allClear()
    txtRegNo.text = ""
    txtPName.text = ""
    txtPID.text = ""
    txtPhoneNo.text = ""
End Sub

Private Sub cmdCancel_Click()
cmdCancel.Enabled = False
   cmdNew.Enabled = True
    cmdEdit.Caption = "&Edit"
    cmdNew.Caption = "&New"
    CmdExit.Enabled = True
    cmdEdit.Enabled = True
    txtPID.Enabled = False
    cmdOpen.Enabled = True
   Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub

Private Sub cmdEdit_Click()
 If cmdEdit.Caption = "&Edit" Then
        cmdNew.Enabled = False
        Call allenable
        txtRegNo.SetFocus
        cmdEdit.Caption = "&Update"
        cmdCancel.Enabled = True
        CmdExit.Enabled = False
        cmdOpen.Enabled = False

 ElseIf cmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                cmdEdit.Caption = "&Edit"
                cmdNew.Enabled = True
                cmdCancel.Enabled = False
                CmdExit.Enabled = True
                cmdOpen.Enabled = True
            Call alldisable
                rsfactory.Requery

'                Dim s As String
'                s = txtPID
'                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
'                FindRecord

            End If
        End If
    End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdFirst_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.EOF = True Then
       cmdFirst.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtPID = Adodc1.Recordset!PID
       txtRegNo = Adodc1.Recordset!RegNo
       txtPName = Adodc1.Recordset!PName
       txtPhoneNo = Adodc1.Recordset!PhoneNo
       
End If
End Sub

Private Sub cmdLast_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
'          MsgBox "end of file"
       cmdLast.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtPID = Adodc1.Recordset!PID
       txtRegNo = Adodc1.Recordset!RegNo
       txtPName = Adodc1.Recordset!PName
       txtPhoneNo = Adodc1.Recordset!PhoneNo
       
End If
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If cmdNew.Caption = "&New" Then
        cmdNew.Caption = "&Save"
        cmdEdit.Enabled = False
        cmdCancel.Enabled = True
        CmdExit.Enabled = False
        cmdOpen.Enabled = False
'        chameleonButton1.Enabled = False
        Call allClear
        
If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(PID),0) as PID from PRegtration"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtPID.text = Val(rs!PID) + 1
            
        Call allenable
        txtPName.SetFocus
    ElseIf cmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtPID.Enabled = False
                cmdNew.Caption = "&New"
                cmdEdit.Enabled = True
                cmdCancel.Enabled = False
                CmdExit.Enabled = True
                cmdOpen.Enabled = True
                Call alldisable
                s = txtPID
                rsfactory.Requery
'                rsfactory.MoveFirst
'                rsfactory.Find "PID='" & parseQuotes(s) & "'"
'                FindRecord

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

Private Sub RegNo()
  Dim strSQL As String
  
  strSQL = "Update PRegtration set RegNo= (select CONVERT(varchar, YEAR(getdate()))+''" & _
 "+''+REPLICATE('0', 6-LEN(PID)-2)+ + CONVERT(varchar,(select count(*)from PRegtration))" & _
 "from PRegtration where PID=(select max(PID)" & _
 "from PRegtration)) " & _
 "where PID=(select max(PID) " & _
 "from PRegtration)"

cn.Execute strSQL
End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True

'    If (txtRegNo.text = "") Then
'       MsgBox "Enter User Name"
'       txtRegNo.SetFocus
'       IsValidRecord = False
'       Exit Function
'    End If

    If (txtPName.text = "") Then
      MsgBox "Enter Passward"
      txtPName.SetFocus
      IsValidRecord = False
      Exit Function
    End If

    If (txtPhoneNo.text = "") Then
      MsgBox "Enter Patient Mobile No"
      txtPID.SetFocus
      IsValidRecord = False
      Exit Function
    End If

If cmdNew.Caption = "&Save" Then
        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from PRegtration where upper(PName)='" & Strings.UCase(Strings.Trim(parseQuotes(txtRegNo))) & "'", cn

'             If Not rsfactory.EOF Then
'        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
'          txtRegNo.SetFocus
'          IsValidRecord = False
'         Exit Function
'            End If

         End If
    End If
End Function

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If cmdNew.Caption = "&Save" Then
    
        cn.Execute "INSERT INTO PRegtration(PID,RegNo,PName,PhoneNo) " & _
                   " VALUES ('" & parseQuotes(txtPID) & "','" & parseQuotes(txtRegNo) & "', " & _
                   " '" & parseQuotes(txtPName) & "','" & parseQuotes(txtPhoneNo) & "')"
                   
                   
         rcupdate = True
         Call RegNo
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update PRegtration Set RegNo='" & parseQuotes(txtRegNo) & "',PName='" & parseQuotes(txtPName) & _
                  "',PhoneNo='" & parseQuotes(txtPhoneNo) & "' WHERE PID = '" & parseQuotes(txtPID) & "'"
          
     rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans

    Exit Function



ErrHandler:
    cn.RollbackTrans
    rsfactory.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate PRegtration Name"
            txtRegNo = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function


Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtPID = rsfactory("PID")
        txtRegNo = rsfactory("RegNo")
        txtPName = rsfactory("PName")
        txtPhoneNo = rsfactory("PhoneNo")

    End If
End Sub

Private Sub cmdNext_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
       cmdNext.Enabled = False
 Else
       cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtPID = Adodc1.Recordset!PID
       txtRegNo = Adodc1.Recordset!RegNo
       txtPName = Adodc1.Recordset!PName
       txtPhoneNo = Adodc1.Recordset!PhoneNo
       

End If
End Sub

Private Sub cmdOpen_Click()
frmPRegSearch.Show vbModal
    cmdOpen.Enabled = True
    cmdCancel.Enabled = True
End Sub

Public Sub PopulateCnf(StrID As String)

    rsfactory.MoveFirst
    rsfactory.Find "PID=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

Private Sub cmdPrevious_Click()
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF = True Then
       cmdPrevious.Enabled = False
 Else
      cmdFirst.Enabled = True
       cmdNext.Enabled = True
       cmdLast.Enabled = True
       cmdPrevious.Enabled = True

       txtPID = Adodc1.Recordset!PID
       txtRegNo = Adodc1.Recordset!RegNo
       txtPName = Adodc1.Recordset!PName
       txtPhoneNo = Adodc1.Recordset!PhoneNo
       
End If
End Sub

Private Sub Form_Load()
Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from PRegtration order by RegNo", cn, adOpenStatic, adLockReadOnly
    
    Call alldisable
'    Call Department
    
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsfactory.EOF Then FindRecord

    txtPID.Enabled = False

'    txtPName.AddItem "Expenditure"
'    txtPName.AddItem "Income"
'    txtPName.AddItem "Asset"
'    txtPName.AddItem "Liabilities"

    Adodc1.ConnectionString = "Driver={SQL Server};" & _
           "Server=" & sServerName & ";" & _
           "Database=" & SDatabaseName & ";" & _
           "Trusted_Connection=yes"

  Adodc1.CommandType = adCmdTable
  Adodc1.RecordSource = "PRegtration"

  Adodc1.Refresh

End Sub

Private Sub txtPName_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
       SendKeys Chr(9)
txtPName.text = StrConv(txtPName.text, vbProperCase)
    End If
End Sub

Private Sub txtPName_LostFocus()
txtPName.text = StrConv(txtPName.text, vbProperCase)
End Sub

Private Sub txtPhoneNo_LostFocus()
If Len(txtPhoneNo) < 11 Or Len(txtPhoneNo) > 11 Then
MsgBox "Input valid Phone no."
txtPhoneNo.SetFocus
End If
'    End If
txtPhoneNo.BackColor = vbWhite
End Sub
