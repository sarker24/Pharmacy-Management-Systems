VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBT 
   BackColor       =   &H00C0B4A9&
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
   Icon            =   "frmBF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   240
      TabIndex        =   11
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12629161
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Debt"
      TabPicture(0)   =   "frmBF.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Credit"
      TabPicture(1)   =   "frmBF.frx":0B24
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0B4A9&
         Height          =   3135
         Left            =   0
         TabIndex        =   13
         Top             =   480
         Width           =   5775
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1920
            TabIndex        =   25
            Top             =   2160
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   1920
            TabIndex        =   23
            Top             =   1200
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            Format          =   22806529
            CurrentDate     =   38476
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "frmBF.frx":10BE
            Left            =   1920
            List            =   "frmBF.frx":10CE
            TabIndex        =   22
            Top             =   1680
            Width           =   3135
         End
         Begin VB.TextBox txtSerialNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Height          =   375
            Index           =   0
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Width           =   3135
         End
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
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
            TabIndex        =   15
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox txtBname 
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
            TabIndex        =   14
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Check Number"
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
            Left            =   240
            TabIndex        =   24
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblSlNo 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Serial No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   21
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label lblDate 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   20
            Top             =   1320
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Transaction Type"
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
            Index           =   0
            Left            =   360
            TabIndex        =   19
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   960
            TabIndex        =   18
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Bank Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0B4A9&
         Height          =   3255
         Left            =   -75000
         TabIndex        =   12
         Top             =   360
         Width           =   5775
         Begin VB.ComboBox Combo2 
            Height          =   315
            ItemData        =   "frmBF.frx":10F7
            Left            =   1800
            List            =   "frmBF.frx":1101
            TabIndex        =   3
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox Text5 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   2
            Top             =   1320
            Width           =   3135
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   375
            Left            =   1800
            TabIndex        =   33
            Top             =   960
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            CustomFormat    =   "dd-MMM-yyyy"
            Format          =   22806529
            CurrentDate     =   38476
         End
         Begin VB.TextBox Text4 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1800
            TabIndex        =   4
            Top             =   2160
            Width           =   3135
         End
         Begin VB.TextBox txtSerialNo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Height          =   375
            Index           =   1
            Left            =   1800
            TabIndex        =   26
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   5
            Top             =   2640
            Width           =   3135
         End
         Begin VB.TextBox Text2 
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
            Left            =   1800
            TabIndex        =   1
            Top             =   600
            Width           =   3135
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Transaction Type"
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
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Check Number"
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
            Left            =   120
            TabIndex        =   32
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label lblSlNo 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Serial No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   600
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblDate 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   1080
            TabIndex        =   30
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Payment To"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   29
            Top             =   1320
            Width           =   1335
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   28
            Top             =   2640
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0B4A9&
            Caption         =   "Bank Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   27
            Top             =   720
            Width           =   1215
         End
      End
   End
   Begin VB.CommandButton cmdOpen 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Open"
      Height          =   795
      Left            =   4920
      Picture         =   "frmBF.frx":1112
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4080
      Width           =   945
   End
   Begin VB.CommandButton chameleonButton1 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Pre&view"
      Height          =   795
      Left            =   4065
      Picture         =   "frmBF.frx":19DC
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4080
      Width           =   825
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0B4A9&
      Caption         =   "C&lose"
      Height          =   795
      Left            =   3240
      Picture         =   "frmBF.frx":22A6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   825
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00C0B4A9&
      Caption         =   "&Edit"
      Height          =   795
      Left            =   1560
      Picture         =   "frmBF.frx":2B70
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   825
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0B4A9&
      Caption         =   "&New"
      Height          =   795
      Left            =   720
      Picture         =   "frmBF.frx":343A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4080
      Width           =   825
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0B4A9&
      Caption         =   "Cancel"
      Height          =   795
      Left            =   2400
      Picture         =   "frmBF.frx":3D04
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
      Width           =   825
   End
End
Attribute VB_Name = "frmBT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

