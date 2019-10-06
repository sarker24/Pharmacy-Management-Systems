VERSION 5.00
Begin VB.Form frmUser 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " User Information"
   ClientHeight    =   5070
   ClientLeft      =   1380
   ClientTop       =   930
   ClientWidth     =   6000
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   6000
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "User Information Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.TextBox txtUID 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox TxtUName 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   11
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
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
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   10
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox txtSNumber 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
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
         Left            =   1920
         TabIndex        =   9
         Top             =   360
         Width           =   3135
      End
      Begin VB.ComboBox cboEn 
         BackColor       =   &H00D0B5A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   6
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "E&xit"
         Height          =   735
         Left            =   3360
         Picture         =   "frmUser.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdNew 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&New"
         Height          =   735
         Left            =   480
         Picture         =   "frmUser.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdEdit 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Edit"
         Height          =   735
         Left            =   1440
         Picture         =   "frmUser.frx":171E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdOpen 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Open"
         Height          =   735
         Left            =   4320
         Picture         =   "frmUser.frx":1FE8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   990
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00C0B4A9&
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   2400
         Picture         =   "frmUser.frx":28B2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   990
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00C0B4A9&
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0B4A9&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label lblSerial 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblPgroup 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Privilege Group"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label MSG 
         Alignment       =   2  'Center
         BackColor       =   &H00D0B5A8&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   600
         TabIndex        =   7
         Top             =   3960
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmUser"
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
    txtUID.Enabled = False
    TxtUName.Enabled = False
    txtPassword.Enabled = False
'    txtSNumber.Enabled = False
    cboEn.Enabled = False
End Sub

Private Sub allenable()
    TxtUName.Enabled = True
    txtPassword.Enabled = True
    txtUID.Enabled = True
    txtSNumber.Enabled = True
    cboEn.Enabled = True
End Sub
'
Public Sub allClear()
    TxtUName.text = ""
    txtPassword.text = ""
    txtUID.text = ""
    cboEn.text = ""
'    txtSerialNumber.text = ""
End Sub

Private Sub cmdCancel_Click()
CmdCancel.Enabled = False
   CmdNew.Enabled = True
    CmdEdit.Caption = "&Edit"
    CmdNew.Caption = "&New"
    CmdExit.Enabled = True
    CmdEdit.Enabled = True
    txtSNumber.Enabled = False
    CmdOpen.Enabled = True
   Call allClear
    Call alldisable
    If Not rsfactory.EOF Then FindRecord
End Sub

Private Sub cmdEdit_Click()
 If CmdEdit.Caption = "&Edit" Then
        CmdNew.Enabled = False
        Call allenable
        TxtUName.SetFocus
        CmdEdit.Caption = "&Update"
        CmdCancel.Enabled = True
        CmdExit.Enabled = False
        CmdOpen.Enabled = False

 ElseIf CmdEdit.Caption = "&Update" Then
        If IsValidRecord Then
            If rcupdate Then
                CmdEdit.Caption = "&Edit"
                CmdNew.Enabled = True
                CmdCancel.Enabled = False
                CmdExit.Enabled = True
                CmdOpen.Enabled = True
            Call alldisable
                rsfactory.Requery

                Dim s As String
                s = txtSNumber
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
'                Call search
'                Call countrysearch
                FindRecord

            End If
        End If
    End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
    On Error GoTo ProcError
      Set rs = New ADODB.Recordset
    If CmdNew.Caption = "&New" Then
        CmdNew.Caption = "&Save"
        CmdEdit.Enabled = False
        CmdCancel.Enabled = True
        CmdExit.Enabled = False
        CmdOpen.Enabled = False
        Call allClear

If rs.State <> 0 Then rs.Close
           str = "Select ISNULL(max(SerialNo),0) as SerialNo from PMSUser"
           rs.Open str, cn, adOpenStatic, adLockReadOnly
           txtSNumber.text = Val(rs!SerialNo) + 1

        Call allenable
        TxtUName.SetFocus
    ElseIf CmdNew.Caption = "&Save" Then
        Dim s As String
        If IsValidRecord Then
            If rcupdate Then
                txtSNumber.Enabled = False
                CmdNew.Caption = "&New"
                CmdEdit.Enabled = True
                CmdCancel.Enabled = False
                CmdExit.Enabled = True
                CmdOpen.Enabled = True
            Call alldisable
                s = txtSNumber
                rsfactory.Requery
                rsfactory.MoveFirst
                rsfactory.Find "SerialNo='" & parseQuotes(s) & "'"
                FindRecord

            End If
        End If
    End If
'
    Exit Sub

ProcError:
    Select Case Err.Number
    Case 0:
    Case Else
        MsgBox Err.Description
    End Select

End Sub

Private Function IsValidRecord() As Boolean
    IsValidRecord = True

    If (TxtUName.text = "") Then
       MsgBox "Enter User Name"
       TxtUName.SetFocus
       IsValidRecord = False
       Exit Function
    End If

    If (txtPassword.text = "") Then
      MsgBox "Enter Passward"
      txtPassword.SetFocus
      IsValidRecord = False
      Exit Function
    End If

    If (txtUID.text = "") Then
      MsgBox "Enter Confirm Passward"
      txtUID.SetFocus
      IsValidRecord = False
      Exit Function
    End If

If CmdNew.Caption = "&Save" Then
        If rsfactory.RecordCount > 0 Then
        If rsfactory.State <> 0 Then rsfactory.Close
            rsfactory.Open "select * from PMSUser where upper(UName)='" & Strings.UCase(Strings.Trim(parseQuotes(TxtUName))) & "'", cn

             If Not rsfactory.EOF Then
        MsgBox "This Record already exists Please Enter Another Record.", vbInformation, Me.Caption & " - " & App.Title
          TxtUName.SetFocus
          IsValidRecord = False
         Exit Function
            End If

         End If
    End If
End Function

Private Function rcupdate() As Boolean

    On Error GoTo ErrHandler

    cn.BeginTrans
    If CmdNew.Caption = "&Save" Then

        cn.Execute "INSERT INTO PMSUser(SerialNo,UID,UName, " & _
                   " Password,Privilegegroup) " & _
                   " VALUES ('" & parseQuotes(txtSNumber) & "','" & parseQuotes(txtUID) & "', " & _
                   " '" & parseQuotes(TxtUName) & "', " & _
                   " '" & parseQuotes(txtPassword) & "', " & _
                   " '" & parseQuotes(cboEn) & "') "



          rcupdate = True
          MsgBox "Record Added Successfully", vbInformation, "Confirmation"
    Else

        cn.Execute "Update PMSUser Set UID='" & parseQuotes(txtUID) & _
                  "',UName='" & parseQuotes(TxtUName) & "', " & _
                  " Password='" & parseQuotes(txtPassword) & _
                  "',Privilegegroup='" & parseQuotes(cboEn) & "' " & _
                  " Where SerialNo ='" & parseQuotes(txtSNumber) & "' "


        rcupdate = True
        MsgBox "Record Updated Successfully", vbInformation, "Confirmation"
    End If

    cn.CommitTrans
'    Exit Sub
    Exit Function

ErrHandler:
    cn.RollbackTrans
   ' rsFactory.Requery
    Select Case cn.Errors(0).NativeError
        Case 2627
            MsgBox "Trying with duplicate CNF Name"
            TxtUName = ""
            TxtUName.SetFocus
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Function


Public Sub FindRecord()
If Not rsfactory.EOF Then
        txtSNumber = rsfactory("SerialNo")
        txtUID = rsfactory("UID")
        TxtUName = rsfactory("UName")
        txtPassword = rsfactory("Password")
        cboEn = rsfactory("Privilegegroup")

    End If
End Sub

Private Sub cmdOpen_Click()
frmUserSearch.Show vbModal
    CmdOpen.Enabled = True
    CmdCancel.Enabled = True
End Sub

Public Sub PopulateCnf(StrID As String)

    rsfactory.MoveFirst
    rsfactory.Find "SerialNo=" & parseQuotes(StrID)
    If rsfactory.EOF Then MsgBox "No Such Record Exists.", vbOKOnly, "Find" Else FindRecord

End Sub

Private Sub Form_Load()
Call Connect
       ModFunction.StartUpPosition Me
    Set rsfactory = New ADODB.Recordset
    rsfactory.Open "select * from PMSUser", cn, adOpenStatic, adLockReadOnly
    Call alldisable
   If rsfactory.RecordCount > 0 Then
        bRecordExists = True
    Else
        bRecordExists = False
    End If

    If Not rsfactory.EOF Then FindRecord

    cboEn.AddItem "1"
    cboEn.AddItem "0"
'    cboEn.AddItem "POWER USER"
'    txtSNumber.Enabled = False
End Sub




