VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmROL 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Auto Reorder Level Informations"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "frmROL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   9600
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LstRol 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   14420
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Catagory"
         Object.Width           =   1870
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Medicine Name"
         Object.Width           =   9172
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "ROL Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Available Qty"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmROL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private rsMaster                As ADODB.Recordset
'Private rsSelect                As ADODB.Recordset 'sub
'Dim Str                         As String
'Dim rs                          As New ADODB.Recordset
'
'
'Private Sub cmdExit_Click()
'Unload Me
'End Sub
'
'
'Private Sub Form_Load()
'Call Connect
'    ModFunction.StartUpPosition Me
'    Call MedicineName
'End Sub
'
'Private Sub MedicineName()
'
'Set rsMaster = New ADODB.Recordset
'rsMaster.Open " select tunion.Mname,isnull((select sum(tt1.Qty) from PurchaseDetail " & _
'                  " tt1 where tt1.Mname=tunion.Mname and tt1.Posted='Posted' ),0) " & _
'                  "-isnull((select sum(tt2.Qty) from SalesDetail tt2 where tt2.Mname=tunion.Mname and tt2.Posted='Posted' ),0) " & _
'                  "as Qty,Rol=(select distinct ROL from tblMedicineName where tblMedicineName.MedicineName=tunion.Mname)" & _
'                  "from (select distinct PurchaseDetail.Mname,PurchaseDetail.Catagory from PurchaseDetail union " & _
'                  "all select SalesDetail.Mname,SalesDetail.MCatagory from SalesDetail)  as tunion " & _
'                  "Where isnull((select sum(tt1.Qty) from PurchaseDetail tt1 where tt1.Mname=tunion.Mname " & _
'                  "and tt1.Posted='Posted' ),0) -isnull((select sum(tt2.Qty) from SalesDetail tt2 where tt2.Mname=tunion.Mname " & _
'                  "and tt2.Posted='Posted' ),0)<(select distinct ROL from tblMedicineName where tblMedicineName.MedicineName=tunion.Mname) " & _
'                  "group by tunion.Mname", cn, adOpenStatic, adLockReadOnly
'
''    If rs.State <> 0 Then rs.Close
''    rs.Open Str, cn, adOpenStatic, adLockReadOnly
''    While Not rs.EOF
'
''        With LstRol.ListItems.Add
'
''            .text = rs!Catagory
''            .SubItems(1) = rs!Mname
''            .SubItems(2) = rs!Qty
''            .SubItems(3) = rs!ROL
''
''
''        End With
''        rs.MoveNext
''    Wend
'
'End Sub
