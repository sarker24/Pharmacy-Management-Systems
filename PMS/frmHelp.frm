VERSION 5.00
Begin VB.Form frmHelp 
   Appearance      =   0  'Flat
   BackColor       =   &H80000017&
   Caption         =   "IMS Help"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fram1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   " Inventory Management Systems  "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00C0B4A9&
         Caption         =   "Ok"
         Height          =   495
         Left            =   7680
         Picture         =   "frmHelp.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label LblCaption 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5535
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   8655
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Fram1_DragDrop(Source As Control, X As Single, Y As Single)
'text.txtMsg = "Inventory Management Systems" & _
'              "Developed By" & _
'              "Murad Hossain Sarker" & _
'              "MD. Abdus Samad Borhan" & _
'              "By this Software you can easily handle your sales and Stock system from a point" & _
'              "For any Information Contact with Us" _
'              "Mobile No: 011059279, 0171227051"
End Sub

Private Sub Label1_Click()
'Label1.Caption = "Inventory Management Systems" & vbCrLf & _
'                 "Developed By" & vbCrLf & vbCrLf & _
'                 "Murad Hossain Sarker" & vbCrLf & _
'                 "MD. Abdus Samad Borhan" & vbCrLf & vbCrLf& _
'                 "By this Software you can easily handle your sales and Stock system from a point" & _
'                 "For any Information Contact with Us" & _
'                 "Mobile No: 011059279, 0171227051"
''"5900-T Hollis Street" & vbCrLf & _
'                                     "Emeryville, CA 94608" & vbCrLf & vbCrLf & _
'                                     "Voice: 510-595-2400" & vbCrLf & _
'                                     "Fax:   510-595-2424" & vbCrLf & vbCrLf & _
'                                     "www.videosoft.com"

End Sub

