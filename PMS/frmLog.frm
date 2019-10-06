VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H00C0B4A9&
   BorderStyle     =   0  'None
   Caption         =   "Login To application ..."
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLog.frx":0E42
   ScaleHeight     =   2820
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   -120
      TabIndex        =   9
      Text            =   "Doctors Clinic Unit - 2"
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox txtTime 
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDate 
      Enabled         =   0   'False
      Height          =   405
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtCTime 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   120
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   2880
      Picture         =   "frmLog.frx":3BD8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton CmdEnter 
      BackColor       =   &H00C0B4A9&
      Height          =   495
      Left            =   1920
      Picture         =   "frmLog.frx":44A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
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
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "01735414303"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtUname 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   1800
      TabIndex        =   0
      Text            =   "admin"
      Top             =   1200
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker ExpiryDate 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   49676289
      CurrentDate     =   41037
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   600
      TabIndex        =   3
      Tag             =   "&User Name:"
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   2
      Tag             =   "&Password:"
      Top             =   1680
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
                                                                                                                                               
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub CmdEnter_Click()
    Call Connect
    Dim str As String
    Dim cm As New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set cm = New ADODB.Connection '
    Set cn = New ADODB.Connection
'   str = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;Initial Catalog=" & SDatabaseName & ";Data Source=" & sServerName
    str = "Provider=SQLOLEDB;Trusted_Connection=Yes;User ID=sa;Database=" & SDatabaseName & ";Server=" & sServerName
    cn.Open str
    txtDate.text = Date        'Format((cm.Execute("Select GetDate()")), "dd-MM-yyyy")
    txtTime.text = Time        'Format((cm.Execute("Select GetDate()")), "hh:mm:ss")
    str = "select UName,Password from PMSUser where UName ='" & txtUName.text & "'"
 '    str = "select UserID,Password,Upper(EmployeeNick) EmployeeNick from USUser where UserID='" & txtUname.text & "'"
         If rs.State <> 0 Then rs.Close
            rs.Open str, cn, adOpenStatic, adLockReadOnly
           If rs.RecordCount = 0 Then Exit Sub
    If rs!Password = Trim(CStr(txtPassword.text)) Then
    frmScanLogOn.Show vbModal
        frmMain.Show
'        frmROL.Show vbModal
        
'        If rs!Name = "MD" And rs!Password = "01920468031" Then
        If rs!UName = "Admin" Then
        
        frmLogin.Hide
'            cm.BeginTrans
'            frmMain.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuComunication.Enabled = True
             frmMain.mnuSetUp.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Enabled = True
             frmMain.mnuMSInfo.Visible = True
             frmMain.mnuPLStatement.Visible = True
             frmMain.mnuMSales.Enabled = True
             frmMain.Calculator.Enabled = True
             frmMain.mnuMPurchase.Enabled = True
             frmMain.mnuSupplierName.Enabled = True
             frmMain.mnuTools.Enabled = True
             

cn.Execute " insert into PMSLogin Values ('" & txtUName.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtTime.text, "hh:mm:ss:tt") & "')"
'str = " insert into PMSLogin Values ('" & txtUname.text & "','" & txtDate.text & "','" & txtTime.text & "')"

'            cm.Execute str
'            cm.CommitTrans
ElseIf rs!UName = "Borhan" Then

            frmLogin.Hide

             frmMain.Enabled = True
             frmMain.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuComunication.Enabled = True
             frmMain.mnuMSales.Enabled = True
             frmMain.mnuSetUp.Enabled = True
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuReport.Enabled = True
             frmMain.Calculator.Enabled = True
             frmMain.mnuSupplierName.Enabled = True
             frmMain.mnuMPurchase.Enabled = True
             frmMain.mnuTools.Enabled = True
'            Unload Me
cn.Execute " insert into PMSLogin Values ('" & txtUName.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtTime.text, "hh:mm:ss:tt") & "')"

        Else
            frmMain.Enabled = True
             frmMain.Enabled = True
             frmMain.mnuBackUp.Enabled = True
             frmMain.mnuComunication.Enabled = True
             frmMain.mnuMPurchase.Enabled = False
             frmMain.mnuSetUp.Enabled = True
             frmMain.mnuUser.Visible = False
             frmMain.mnuMSInfo.Visible = False
             frmMain.mnuPLStatement.Visible = False
             frmMain.mnuHelp.Enabled = True
             frmMain.mnuMIStatement.Enabled = True
             frmMain.Calculator.Enabled = True
             frmMain.mnuSupplierName.Enabled = True
             frmMain.mnuMSales.Enabled = True
             frmMain.mnuTools.Enabled = True

cn.Execute " insert into PMSLogin Values ('" & txtUName.text & "','" & Format(txtDate.text, "yyyy-mm-dd") & "','" & Format(txtCTime.text, "hh:mm:ss") & "')"

            frmLogin.Hide
        End If
    Else
            MsgBox "Invalid Password. Please try again.", vbInformation, "Confarmation"
            txtPassword.text = ""
            txtPassword.SetFocus
    End If

End Sub

Private Sub CmdEnter_GotFocus()
    CmdEnter.FontBold = True
End Sub

Private Sub CmdEnter_LostFocus()
    CmdEnter.FontBold = False
End Sub

Private Sub CmdCancel_GotFocus()
    cmdCancel.FontBold = True
End Sub

Private Sub CmdCancel_LostFocus()
    cmdCancel.FontBold = False
End Sub

Private Sub txtUname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    txtPassword.SetFocus
    End If
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
    Call CmdEnter_Click
    End If
End Sub

Private Sub Form_Load()
'   ModConnection.StartUpPosition Me
'   Text1.Enabled = False
Call TimeExpired
End Sub

Private Sub TimeExpired()
Dim CurrentDate As Date
Dim ExpiryDate
CurrentDate = Now
ExpiryDate = "1-1-2019"
If CurrentDate > ExpiryDate Then
'MsgBox ("Your system has Expired")
'frmLogIn.Hide
Unload Me
End If

End Sub

Private Sub Timer1_Timer()
     txtTime.text = Format(Time$, "hh:mm:ss AM/PM")
     txtCTime.text = Format(Time$, "hh:mm:ss AM/PM")
End Sub

Private Sub txtUname_GotFocus()
    txtUName.BackColor = &HFFC0C0
    txtUName.SelStart = 0
    txtUName.SelLength = Len(txtUName)
End Sub

Private Sub txtUname_LostFocus()
    txtUName.BackColor = &HFFFFFF
    txtUName.text = StrConv(txtUName.text, vbProperCase)
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.BackColor = &HFFC0C0
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = &HFFFFFF
End Sub
