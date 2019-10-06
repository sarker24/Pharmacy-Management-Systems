VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "frmChangePassword.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "Change"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3075
      TabIndex        =   4
      Top             =   2805
      Width           =   1410
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4530
      TabIndex        =   5
      Top             =   2805
      Width           =   1000
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1275
      Width           =   3500
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1635
      Width           =   3500
   End
   Begin VB.TextBox txtNewPassword 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1995
      Width           =   3500
   End
   Begin VB.TextBox txtRetype 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2355
      Width           =   3500
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   11
      Top             =   1275
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   10
      Top             =   1635
      Width           =   975
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   9
      Top             =   1995
      Width           =   1050
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Re-type Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   165
      TabIndex        =   8
      Top             =   2355
      Width           =   1320
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Note:Please fill all required parameters."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   645
      Width           =   2835
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE PASSWORD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   210
      TabIndex        =   6
      Top             =   360
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   -6525
      Picture         =   "frmChangePassword.frx":000C
      Top             =   -30
      Width           =   14325
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Public State                        As FORM_STATE
'
'Dim RSPass                          As Recordset
'Dim sSQL                            As String
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdSave_Click()
'Dim obj As Control
'
'For Each obj In Me
'If TypeOf obj Is TextBox Or TypeOf obj Is ComboBox Then
'    If obj.text = "" Then
'        MsgBox obj.Name & " could not be left blank. Please complete the field.", vbExclamation, Me.Caption
'        obj.SetFocus
'        Exit Sub
'    End If
'End If
'Next obj
'
'If txtNewPassword.text <> txtRetype.text Then
'    MsgBox "Password(s) did not match.Please check it!", vbExclamation
'    Exit Sub
'End If
'
'Set RSPass = New ADODB.Recordset
'If RSPass.State = adStateOpen Then RSPass.Close
'RSPass.Open "SELECT * FROM PMSUser WHERE UID='" & PMSUser.UID & "'", cn, adOpenStatic, adLockReadOnly
'
'If RSPass.Fields("Password") <> Encode(txtPassword.text) Then
'    MsgBox "Old password mismatch.Please check it!", vbExclamation
'    Exit Sub
'End If
'
'If State = EditStateMode Then
'    cn.Execute "UPDATE PMSUser SET PMSUser.Password='" & Encode(txtNewPassword.text) & "' WHERE PMSUser.Username='" & PMSUser.UID & "'"
'
'    MsgBox "Your existing password has been successfully changed!", vbInformation
'    Unload Me
'End If
'End Sub
'
'Private Sub Form_Activate()
'On Error Resume Next
'    txtPassword.SetFocus
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo ErrHandler
'CenterForm frmChangePassword
'
'State = EditStateMode
'txtUsername.text = PMSUser.UID
'
'Exit Sub
'ErrHandler:
'    MsgBox "Error Number: " & Err.Number & vbCrLf & "Description: " & Err.Description, vbExclamation, Me.Caption
'End Sub
'
'Private Sub Form_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then
'    Unload Me
'ElseIf KeyAscii = 13 Then
'    SendKeys "{TAB}"
'End If
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    Set frmChangePassword = Nothing
'End Sub
'
'Private Sub txtNewPassword_GotFocus()
'HLText txtNewPassword
'End Sub
'
'Private Sub txtPassword_GotFocus()
'HLText txtPassword
'End Sub
'
'
'Private Sub txtRetype_GotFocus()
'HLText txtRetype
'End Sub
'
'Private Sub txtUsername_GotFocus()
'    HLText txtUsername
'End Sub
'
'
