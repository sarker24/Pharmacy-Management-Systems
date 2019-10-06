VERSION 5.00
Object = "{BDDD132C-614B-11D3-B85E-85ADB7D07209}#1.0#0"; "dXSBar.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H00C0B4A9&
   Caption         =   "Pharmacy  Management System [Doctors Clinic Unit - 2]"
   ClientHeight    =   9045
   ClientLeft      =   1065
   ClientTop       =   450
   ClientWidth     =   14970
   Icon            =   "frmMain.frx":0000
   Moveable        =   0   'False
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   8670
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1235
            MinWidth        =   1235
            Text            =   "User :"
            TextSave        =   "User :"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   1940
            MinWidth        =   1940
            TextSave        =   "2017-08-05"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "10:22"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   15743
            Text            =   "Software Developed by ""MAS IT SOLUTIONS"". Hot Line : 02-9031260, 01915682291"
            TextSave        =   "Software Developed by ""MAS IT SOLUTIONS"". Hot Line : 02-9031260, 01915682291"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Palatino Linotype"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   10425
      ScaleWidth      =   14940
      TabIndex        =   0
      Top             =   0
      Width           =   14970
      Begin DXSIDEBARLibCtl.dxSideBar dxSideBar1 
         Height          =   10455
         Left            =   0
         OleObjectBlob   =   "frmMain.frx":2292B
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Menu First 
      Caption         =   ".........."
   End
   Begin VB.Menu mnuSetUp 
      Caption         =   "Set&Up"
      Begin VB.Menu mnuIPRegistration 
         Caption         =   "Indoor Patient Registraton"
      End
      Begin VB.Menu mnuMRestriction 
         Caption         =   "Medicine Restriction"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuUser 
         Caption         =   "&User Information"
      End
      Begin VB.Menu mnuCPassword 
         Caption         =   "Change Pasword"
      End
      Begin VB.Menu mnuMedicineCagory 
         Caption         =   "Medicine &Catagory"
      End
      Begin VB.Menu mnuGName 
         Caption         =   "&Generic Name"
      End
      Begin VB.Menu mnuRemarks 
         Caption         =   "Remarks"
      End
      Begin VB.Menu mnuMedicineName 
         Caption         =   "&Medicine Name"
      End
      Begin VB.Menu mnuSupplierName 
         Caption         =   "Supplier Name"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "&Transaction"
      Begin VB.Menu mnuPrescription 
         Caption         =   "Prescription"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMSales 
         Caption         =   "Medicine &Sales "
      End
      Begin VB.Menu mnuMPurchase 
         Caption         =   "Medicine &Purchase "
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
      Begin VB.Menu mnuMList 
         Caption         =   "Medicine List"
      End
      Begin VB.Menu mnuSalesStatement 
         Caption         =   "Sales Statement"
         Begin VB.Menu mnuMIStatement 
            Caption         =   "Medicine Income Statement"
         End
         Begin VB.Menu mnuMSDetails 
            Caption         =   "Supplier and Item Sales Report"
         End
         Begin VB.Menu mnuCWSStatement 
            Caption         =   "Catagory wise Sales Statement"
         End
         Begin VB.Menu mnuMELedger 
            Caption         =   "Medicine Employee Ledger"
         End
         Begin VB.Menu mnuDCStatement 
            Caption         =   "Due Collection Statement"
         End
         Begin VB.Menu mnuMSInOut 
            Caption         =   "Medicine Sales Summery Indoor Outdoor"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSDueInformantion 
            Caption         =   "&Indoor Patient Due Statement"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuPStatement 
         Caption         =   "Purchase Statement"
         Begin VB.Menu mnuMPDetails 
            Caption         =   "Supplier wise Purchase Statement"
         End
         Begin VB.Menu mnuMPSummary 
            Caption         =   "Catagory Wise Purchase Statement"
         End
         Begin VB.Menu mnuPSBMedicine 
            Caption         =   "Medicine Purchase Statement"
         End
      End
      Begin VB.Menu mnuSInformation 
         Caption         =   "Supplier Information"
      End
      Begin VB.Menu mnuMSInfo 
         Caption         =   "Medicine Stock Information"
         Begin VB.Menu mnuMSSDW 
            Caption         =   "Medicine Stock Statement Date wise"
         End
         Begin VB.Menu mnuMSPositions 
            Caption         =   "Medicine Item Stock Positions"
         End
         Begin VB.Menu mnuMCSPosition 
            Caption         =   "Medicine Catagory Stock Positions"
         End
         Begin VB.Menu mnuMSWSPosition 
            Caption         =   "Medicine Supplier Wise Stock Positions"
         End
      End
      Begin VB.Menu mnuMEStatement 
         Caption         =   "Medicine Expiry Statement"
      End
      Begin VB.Menu mnuSRStatement 
         Caption         =   "Sales Return Statement"
      End
      Begin VB.Menu mnuOPDDueStatement 
         Caption         =   "Medicine  Due Statement"
      End
      Begin VB.Menu mnuMARRol 
         Caption         =   "Medicine Auto Requisitions ROL"
      End
      Begin VB.Menu mnuProfit 
         Caption         =   "Medicine Sales Summery with Profit"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPSCSummery 
         Caption         =   "Purchase Statement Credit Summery"
      End
      Begin VB.Menu mnuPLStatement 
         Caption         =   "Profit & Loss Statement"
      End
   End
   Begin VB.Menu mnuAcounts 
      Caption         =   "&Acounts"
      Visible         =   0   'False
      Begin VB.Menu mnuVoucher 
         Caption         =   "&Voucher"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "T&ools"
      Begin VB.Menu Calculator 
         Caption         =   "&Calculator"
      End
      Begin VB.Menu mnuComunication 
         Caption         =   "Communication"
      End
      Begin VB.Menu mnuSmessage 
         Caption         =   "Send Message"
      End
   End
   Begin VB.Menu mnuBackUp 
      Caption         =   "Back&Up"
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "&Window"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "H&elp"
      Begin VB.Menu mnuHelpSupport 
         Caption         =   "Help & Support"
      End
      Begin VB.Menu mnuAboutPharmacy 
         Caption         =   "About Pharmachy Software"
      End
   End
   Begin VB.Menu mnuLogoff 
      Caption         =   "&Log off"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calculator_Click()
On Error Resume Next
   Shell "calc.exe"
End Sub

Private Sub dxSideBar1_OnClickItemLink(ByVal pGroup As DXSIDEBARLibCtl.IdxGroup, ByVal pLink As DXSIDEBARLibCtl.IdxItemLink, ByVal GroupIndex As Integer, ByVal ItemLinkIndex As Integer)
  
  Select Case GroupIndex
    
    
    Case 0 'Inventory
    
        Select Case ItemLinkIndex
            Case 0
                frmPurchase.Show vbModal
            Case 1
               frmSales.Show vbModal
            Case 2
               frmSupplier.Show vbModal
            Case 3
                End
'                frmBackUp.Show vbModal
 '           Case 4

        End Select
   End Select

End Sub

Private Sub mnuCalculator_Click()
On Error Resume Next
   Shell "calc.exe"
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Call mnuLogoff_Click
End Sub

Private Sub mnuAboutPharmacy_Click()
FrmAbout.Show vbModal
End Sub

Private Sub mnuACStatement_Click()
RptAdvanceCollection.Show vbModal
End Sub

Private Sub mnuBackUp_Click()
frmBackUp.Show vbModal
End Sub

Private Sub mnuComunication_Click()
Call Shell("C:\Program Files\Internet Explorer\IEXPLORE.EXE", vbMaximizedFocus)
End Sub

Private Sub mnuCPassword_Click()
frmChangePassword.Show vbModal
End Sub

Private Sub mnuCWSStatement_Click()
RptSalesSummery.Show vbModal
End Sub

Private Sub mnuDCStatement_Click()
RptDueCSummary.Show vbModal
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuGName_Click()
frmGenericName.Show vbModal
End Sub

Private Sub MDIForm_Load()
Me.StatusBar1.Panels(2) = frmLogin.txtUName
'Me.StatusBar1.Panels(3) = Date
'Me.StatusBar1.Panels(4) = Time
End Sub

Private Sub mnuIPRegistration_Click()
frmPReg.Show vbModal
End Sub

Private Sub mnuLogoff_Click()
Dim res As VbMsgBoxResult
    res = MsgBox("Are you sure you want to log off?", vbYesNo + vbQuestion)
    If res = vbYes Then
    Unload Me
    'frmLogin.Show
    frmLogin.Show
    frmLogin.txtUName = ""
    frmLogin.txtPassword = ""
    Else
    End If
End Sub

Private Sub mnuMARRol_Click()
RptROL.Show vbModal
End Sub

Private Sub mnuMCSPosition_Click()
RptStockCatagory.Show vbModal
End Sub

Private Sub mnuMedicineCagory_Click()
frmMedicineCatagory.Show vbModal
End Sub

Private Sub mnuMedicineName_Click()
frmMedicineName.Show vbModal
End Sub

Private Sub mnuMELedger_Click()
RptELedger.Show vbModal
End Sub

Private Sub mnuMEStatement_Click()
RptMedicineExpiry.Show vbModal
End Sub

Private Sub mnuMIStatement_Click()
RptMIStatement.Show vbModal
End Sub

Private Sub mnuMList_Click()
RptMedicineList.Show vbModal
End Sub

Private Sub mnuMPDetails_Click()
RptPurchaseDetails.Show vbModal
End Sub

Private Sub mnuMPSummary_Click()
RptPurchaseSummary.Show vbModal
End Sub

Private Sub mnuMPurchase_Click()
frmPurchase.Show vbModal
End Sub

Private Sub mnuMRestriction_Click()
frmMedicineRestrictions.Show vbModal
End Sub

Private Sub mnuMSales_Click()
frmSales.Show vbModal
End Sub

Private Sub mnuMSDetails_Click()
RptSalesDetails.Show vbModal
End Sub

Private Sub mnuMSInOut_Click()
RptIOSStatement.Show vbModal
End Sub

Private Sub mnuMSPositions_Click()
RptStockPosition.Show vbModal
End Sub

Private Sub mnuMSSDW_Click()
frmStockMasterSearch.Show vbModal
End Sub

Private Sub mnuMSSummery_Click()
RptSalesSummery.Show vbModal
End Sub

Private Sub mnuMSWSPosition_Click()
RptSSStatement.Show vbModal
End Sub

Private Sub mnuOPDDueStatement_Click()
RptOPDPatientDue.Show vbModal
End Sub

Private Sub mnuPLStatement_Click()
RptMSSwithProfit.Show vbModal
End Sub

Private Sub mnuPrescription_Click()
frmPatientTreatment.Show vbModal
End Sub

Private Sub mnuProfit_Click()
RptMSSwithProfit.Show vbModal
End Sub

Private Sub mnuPSBSupplier_Click()

End Sub

Private Sub mnuPSBMedicine_Click()
RptIStatement.Show vbModal
End Sub

Private Sub mnuPStatement_Click()
'RptProfitStatement.Show vbModal
End Sub

Private Sub mnuRemarks_Click()
frmMedicineRemarks.Show vbModal
End Sub

Private Sub mnuSDueInformantion_Click()
'RptSalesDueStatement.Show vbModal
End Sub

Private Sub mnuShutdown_Click()
Unload Me
End Sub

Private Sub mnuSInformation_Click()
RptSupplierLedger.Show vbModal
End Sub

Private Sub mnuSmessage_Click()
frmSendMessage.Show vbModal
End Sub

Private Sub mnuSRStatement_Click()
RptSalesReturn.Show vbModal
End Sub

Private Sub mnuUser_Click()
frmUser.Show vbModal
End Sub

Private Sub mnuSupplierName_Click()
frmSupplier.Show vbModal
End Sub

