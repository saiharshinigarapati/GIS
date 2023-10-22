VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.4#0"; "CODEJO~1.OCX"
Begin VB.MDIForm mmain 
   BackColor       =   &H8000000F&
   Caption         =   "Geographic Information System in Flood-Prone and Disease Affected Areas in Tandag City"
   ClientHeight    =   7725
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   14025
   Icon            =   "mmain.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "mmain.frx":0CCA
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar MDIStatus 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   7380
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   12
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   442
            MinWidth        =   442
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "LOGIN USER :"
            TextSave        =   "LOGIN USER :"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   8149
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   529
            MinWidth        =   529
            Picture         =   "mmain.frx":55A07
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Today:"
            TextSave        =   "Today:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "2/20/2015"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "11:22 PM"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   882
            MinWidth        =   882
         EndProperty
         BeginProperty Panel11 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Width           =   970
            MinWidth        =   970
         EndProperty
         BeginProperty Panel12 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   935
            MinWidth        =   935
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList i24x24 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":55DA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":58F8D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":5C0F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":5F3BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":6271E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":65910
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":68BB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":6BAFE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":6EE6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":6FB46
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":70820
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":714FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":719E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mmain.frx":71ED6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14025
      _ExtentX        =   24739
      _ExtentY        =   1270
      ButtonWidth     =   1482
      ButtonHeight    =   1217
      Appearance      =   1
      Style           =   1
      ImageList       =   "i24x24"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Log-out"
            Key             =   "Lock"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Map"
            Key             =   "Map"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Barangay"
            Key             =   "Bara"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "User"
            Key             =   "User"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   240
      Top             =   2160
      _Version        =   851972
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu Logout 
         Caption         =   "Log-out"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Trans 
      Caption         =   "&Transaction"
      Begin VB.Menu Geographical 
         Caption         =   "Map View"
      End
   End
   Begin VB.Menu Monitoring 
      Caption         =   "&Monitoring"
      Begin VB.Menu FlodA 
         Caption         =   "Flood Area"
      End
      Begin VB.Menu De 
         Caption         =   "Dengue Affected Area"
      End
      Begin VB.Menu Mal 
         Caption         =   "Malaria Affected Area"
      End
      Begin VB.Menu Meas 
         Caption         =   "Measles Affected Area"
      End
   End
   Begin VB.Menu Maintenance 
      Caption         =   "&Record Master"
      Begin VB.Menu Barangay 
         Caption         =   "Barangay Information"
      End
      Begin VB.Menu ca 
         Caption         =   "-"
      End
      Begin VB.Menu Flood 
         Caption         =   "Flood Level"
      End
   End
   Begin VB.Menu Report 
      Caption         =   "&Report"
      Begin VB.Menu BI 
         Caption         =   "Barangay Information"
      End
      Begin VB.Menu Brgy 
         Caption         =   "Barangay List"
      End
      Begin VB.Menu asa 
         Caption         =   "-"
      End
      Begin VB.Menu M 
         Caption         =   "Tandag Map"
      End
      Begin VB.Menu FM 
         Caption         =   "Flood Map"
      End
      Begin VB.Menu FlArea 
         Caption         =   "High-Flood Prone Area"
      End
      Begin VB.Menu Dengue 
         Caption         =   "Dengue Outbreak Area"
      End
      Begin VB.Menu Malaria 
         Caption         =   "Malaria Outbreak Area"
      End
      Begin VB.Menu Measles 
         Caption         =   "Measles Outbreak Area"
      End
   End
   Begin VB.Menu Utilities 
      Caption         =   "&Utilities"
      Begin VB.Menu User 
         Caption         =   "Manage User"
      End
      Begin VB.Menu About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "mmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Barangay_Click()
frmBarangay.Show
End Sub

Private Sub BI_Click()
Set rs = New ADODB.Recordset
      Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Barangay '", cn, 1, 2
          Set rptBrgyInfo.DataSource = rs
        
     rptBrgyInfo.Show
End Sub

Private Sub Brgy_Click()
Set rs = New ADODB.Recordset
      Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Barangay '", cn, 1, 2
          Set rptBarangay.DataSource = rs
        
     rptBarangay.Show
End Sub

Private Sub De_Click()
frmDengueBrgy.Show
End Sub

Private Sub Dengue_Click()
Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Barangay '", cn, 1, 2
'Set rs = New ADODB.Recordset
    ' rs.Open "Select * from tbl_Barangay where Dengue like '%" & "High" & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
         Set rptDengueArea.DataSource = rs

         rptDengueArea.Show
End Sub

Private Sub DengueA_Click()
frmDengue.Show vbModal
End Sub

Private Sub Exit_Click()
If MsgBox("Are you sure you want to close the program?", vbYesNo, "") = vbYes Then
Unload Me
End If
End Sub

Private Sub FlArea_Click()

Set rs = New ADODB.Recordset

Set rs = New ADODB.Recordset
     rs.Open "Select * from tbl_Barangay where Flood like '%" & "High" & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
         Set rptFloodArea.DataSource = rs

         rptFloodArea.Show
End Sub

Private Sub FlodA_Click()
frmFloodBrgy.Show
End Sub

Private Sub Flood_Click()
frmFlood.Show 1
End Sub

Private Sub FM_Click()
Set rs = New ADODB.Recordset
      Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Barangay '", cn, 1, 2
Set rptFloodMap.DataSource = rs
rptFloodMap.Show
End Sub

Private Sub Geographical_Click()
frmMapVIew.Show vbModal
End Sub

Private Sub Logout_Click()
Unload Me
frmLogin.Show 1
End Sub



Private Sub M_Click()

'Set rptMapTandag.Sections(1).Controls("Image1").Picture = LoadPicture("" & App.Path & "\Map\Normal.jpg")

Set rs = New ADODB.Recordset
      Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Barangay '", cn, 1, 2
Set rptMapTandag.DataSource = rs
rptMapTandag.Show
End Sub

Private Sub Mal_Click()
frmMalariaBrgy.Show
End Sub

Private Sub Malaria_Click()
Set rs = New ADODB.Recordset
        rs.Open "Select * from tbl_Barangay '", cn, 1, 2
'Set rs = New ADODB.Recordset
    ' rs.Open "Select * from tbl_Barangay where Malaria like '%" & "High" & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
         Set rptMalariaArea.DataSource = rs

         rptMalariaArea.Show
End Sub

Private Sub MalariaA_Click()
frmMalaria.Show 1
End Sub

Private Sub MDIForm_Activate()
'SkinFramework.LoadSkin App.Path & "\Styles\Vista.cjstyles", "NormalBlue.ini"
   ' SkinFramework.ApplyWindow Me.hWnd
   ' SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End Sub

Private Sub MDIForm_Load()
SkinFramework.LoadSkin App.Path & "\Styles\Vista.cjstyles", "NormalBlue.ini"
   SkinFramework.ApplyWindow Me.hWnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End Sub

Private Sub Meas_Click()
frmMeaslesBrgy.Show
End Sub

Private Sub Measles_Click()

'Set rs = New ADODB.Recordset
    ' rs.Open "Select * from tbl_Barangay where Measles like '%" & "High" & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
         Set rptMeaslesArea.DataSource = rs

         rptMeaslesArea.Show
End Sub

Private Sub MeaslesAC_Click()
frmMeasles.Show 1
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Key
    Case "Lock"
            Unload Me
            frmLogin.Show 1
     Case "User"
        frmUser.Show 1
    Case "Map"
        frmMapVIew.Show vbModal
    Case "Bara"
        frmBarangay.Show
End Select
End Sub

Private Sub User_Click()
frmUser.Show 1
End Sub

