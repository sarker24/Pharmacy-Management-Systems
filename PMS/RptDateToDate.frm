VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form RptDateToDate 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Date To Date Sales Report"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3075
   Icon            =   "RptDateToDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3075
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2280
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   57933827
      CurrentDate     =   38476
   End
   Begin VB.CommandButton cmdPrivew 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Privew"
      Height          =   735
      Left            =   240
      Picture         =   "RptDateToDate.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Print"
      Height          =   735
      Left            =   1080
      Picture         =   "RptDateToDate.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   855
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Exit"
      Height          =   735
      Left            =   1920
      Picture         =   "RptDateToDate.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPFrom 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   57933827
      CurrentDate     =   38278
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00C0B4A9&
      Caption         =   "From"
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
      Left            =   480
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0B4A9&
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Product Description"
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
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "RptDateToDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    Unload Me
End Sub

