VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBarangay 
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19155
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   19155
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Caption         =   "Commands"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4575
      Left            =   5640
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      Begin VB.CommandButton cmdNew 
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         Picture         =   "frmBarangay.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   """ Barangay - Information """
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtMeasles 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   3360
         Width           =   1575
      End
      Begin VB.TextBox txtMalaria 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   2880
         Width           =   1575
      End
      Begin VB.TextBox txtDengue 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtPopulation 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtBrgy 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboFlood 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   3135
      End
      Begin VB.Label Label13 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "People"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   25
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label10 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "People"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   24
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label9 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "People"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   3600
         TabIndex        =   23
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Measles Cases  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   19
         Top             =   3360
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Malaria Cases :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Population :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label11 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Brgy ID :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Flood Level :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Dengue Cases :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lstBrgy 
      Height          =   4455
      Left            =   0
      TabIndex        =   16
      Top             =   5520
      Width           =   19215
      _ExtentX        =   33893
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   8040
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   25
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":0CCA
            Key             =   "bar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":15A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":227E
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":2B58
            Key             =   "guy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":3432
            Key             =   "trolley"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":3D0C
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":45E6
            Key             =   "app"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":4EC0
            Key             =   "right"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":579A
            Key             =   "line"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":6074
            Key             =   "exclaimation"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":694E
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":7228
            Key             =   "db"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":7B02
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":83DC
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":8CB6
            Key             =   "gng"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":9590
            Key             =   "key"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":9E6A
            Key             =   "arrows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":A744
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":B01E
            Key             =   "magnifier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":B8F8
            Key             =   "synon"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":C1D2
            Key             =   "people"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":CAAC
            Key             =   "silverlock"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":D386
            Key             =   "server"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":DC60
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBarangay.frx":E53A
            Key             =   "plus"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtCidNo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      TabIndex        =   17
      Top             =   4320
      Width           =   1335
   End
End
Attribute VB_Name = "frmBarangay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim search As String
Public Cid As Integer
Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub cmdClose_Click()


If Me.cmdDelete.Enabled = True And Me.cmdUpdate.Enabled = True Then
Me.cmdNew.Enabled = False
'Me.cmdOk.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
clear
btnfalse
End If

If cmdNew.Enabled = False Then
Me.cmdNew.Enabled = True
'Me.cmdOk.Enabled = False
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
clear
btnfalse
Exit Sub
ElseIf cmdNew.Enabled = True Then
    If MsgBox("Are you sure you want to close?", vbYesNo, "") = vbYes Then
        Unload Me
    End If
End If
End Sub

Private Sub cmdDelete_Click()
 If MsgBox("Are you sure you want to delete?", vbYesNo, "Message!") = vbYes Then
 Set rs = New ADODB.Recordset
 rs.Open "Select * from tbl_Barangay where SNo like '" & Me.txtID.Text & "'", cn, 1, 2
 rs!Flag = 0
 rs.Update
Me.cmdNew.Enabled = True
'Me.cmdOk.Enabled = False
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
MsgBox "Record delete.", vbInformation, ""
rec ""
clear
 End If
End Sub

Private Sub cmdNew_Click()
If Me.cmdDelete.Enabled = True And Me.cmdUpdate = True Then
Me.cmdNew.Enabled = False
'Me.cmdOk.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
clear
Exit Sub
Else



Me.cmdNew.Enabled = False
'Me.cmdOk.Enabled = True
Me.cmdUpdate.Enabled = False
Me.cmdDelete.Enabled = False
clear
btntrue
End If
End Sub

Private Sub cmdOk_Click()


    Set rs = New ADODB.Recordset
    rs.Open "Select * from tbl_Barangay where BrgyID like '" & Me.txtID.Text & "'", cn, 1, 2
    If rs.RecordCount = 1 Then
    MsgBox "Please check Barangay ID#. may cause dupplicate of record", vbExclamation, ""
    Exit Sub
    Else

    If Me.txtID.Text = "" Then
    MsgBox "Please input Barangay ID#.", vbExclamation, ""
    Exit Sub
    End If
    If Me.txtBrgy.Text = "" Then
    MsgBox "Please input Barangay.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtPopulation.Text = "" Then
    MsgBox "Please enter number of population in area.", vbExclamation, ""
    Exit Sub
    ElseIf Me.cboFlood.Text = "" Then
    MsgBox "Please select Flood-level.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtDengue.Text = "" Then
    MsgBox "Please select Dengue-level.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtMalaria.Text = "" Then
    MsgBox "Please select Malaria-level.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtMeasles.Text = "" Then
    MsgBox "Please select Measles-level.", vbExclamation, ""
    Exit Sub
    End If
    
    





Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay", cn, 1, 2
With rs
        .AddNew
        !BrgyID = Me.txtID.Text
        !Barangay = Me.txtBrgy.Text
        !Population = Me.txtPopulation.Text
        !Flood = Me.cboFlood.Text
        !Dengue = 0
        !Malaria = 0
        !Measles = 0
        !Flag = 1
        .Update
End With
        Me.cmdNew.Enabled = True
        Me.cmdOk.Enabled = False
        MsgBox "Record save.", vbInformation, ""
        rec ""
        clear

End If
End Sub

Private Sub cmdSearch_Click()
search = InputBox("Enter barangay to search")
             rec search
End Sub

Private Sub cmdUpdate_Click()
    
    

    If Me.txtID.Text = "" Then
    MsgBox "Please input Barangay ID#.", vbExclamation, ""
    Exit Sub
    End If
    
    If Me.txtBrgy.Text = "" Then
    MsgBox "Please input Barangay.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtPopulation.Text = "" Then
    MsgBox "Please enter number of population in area.", vbExclamation, ""
    Exit Sub
    ElseIf Me.cboFlood.Text = "" Then
    MsgBox "Please select Flood-level.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtDengue.Text = "" Then
    MsgBox "Please Input Dengue-Cases.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtMalaria.Text = "" Then
    MsgBox "Please Input Malaria-Cases.", vbExclamation, ""
    Exit Sub
    ElseIf Me.txtMeasles.Text = "" Then
    MsgBox "Please Input Measles-Cases.", vbExclamation, ""
    Exit Sub
    End If
    


Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where BrgyID like '" & Me.txtID.Text & "'", cn, 1, 2
With rs
     !BrgyID = Me.txtID.Text
        !Barangay = Me.txtBrgy.Text
        !Population = Me.txtPopulation.Text
        !Flood = Me.cboFlood.Text
        !Dengue = Me.txtDengue.Text
        !Malaria = Me.txtMalaria.Text
        !Measles = Me.txtMeasles.Text
        .Update
End With
'        Me.cmdNew.Enabled = True
'        Me.cmdOK.Enabled = False
'        Me.cmdUpdate.Enabled = False
'        Me.cmdDelete.Enabled = False
        MsgBox "Record update.", vbInformation, ""
        rec ""
      
        
End Sub

Private Sub Command2_Click()
Me.txtCompany.Text = lstSupplier.SelectedItem.ListSubItems(1).Text
sNo = lstSupplier.SelectedItem.Text
End Sub

Private Sub Command3_Click()
Me.Frame3.Visible = False
End Sub

Private Sub Command4_Click()
Frame4.Visible = False
End Sub

Private Sub Command5_Click()
rec Me.Text1.Text
End Sub

Private Sub Command1_Click()
frmLookupInfoCompany.Show 1
End Sub

Private Sub Form_Load()
rec search


'Flood
Set rs = New ADODB.Recordset
rs.Open " Select * from tbl_FloodLevel", cn, 1, 2
Do Until rs.EOF
Me.cboFlood.AddItem rs!Flood
rs.MoveNext
Loop

'Dengue
'Set rs = New ADODB.Recordset
'rs.Open " Select * from tbl_DengueLevel", cn, 1, 2
'Do Until rs.EOF
'Me.txtDengue.AddItem rs!Dengue
'rs.MoveNext
'Loop

'Malaria
'Set rs = New ADODB.Recordset
'rs.Open " Select * from tbl_MalariaLevel", cn, 1, 2
'Do Until rs.EOF
'Me.txtMalaria.AddItem rs!Malaria
'rs.MoveNext
'Loop

'Measles
'Set rs = New ADODB.Recordset
'rs.Open " Select * from tbl_MeaslesLevel", cn, 1, 2
'Do Until rs.EOF
'Me.txtMeasles.AddItem rs!Measles
'rs.MoveNext
'Loop



btnfalse
End Sub
Sub btnfalse()
Me.txtID.Enabled = False
Me.txtBrgy.Enabled = False
Me.txtPopulation.Enabled = False
Me.cboFlood.Enabled = False
Me.txtDengue.Enabled = False
Me.txtMalaria.Enabled = False
Me.txtMeasles.Enabled = False
End Sub
Sub btntrue()
Me.txtID.Enabled = True
Me.txtBrgy.Enabled = True
Me.txtPopulation.Enabled = True
Me.cboFlood.Enabled = True
Me.txtDengue.Enabled = True
Me.txtMalaria.Enabled = True
Me.txtMeasles.Enabled = True

End Sub

Private Sub txtLname_Change()
recSupplier Me.txtLname.Text
End Sub




Private Sub lstStock_DblClick()
If lstStock.ListItems.Count = 0 Then
MsgBox "Empty List.", vbExclamation, ""
Exit Sub
End If


Me.txtID.Text = lstStock.SelectedItem.Text
Me.txtBrgy.Text = lstStock.SelectedItem.ListSubItems(1).Text
Me.txtCompany.Text = lstStock.SelectedItem.ListSubItems(2).Text
Me.cboCategory.Text = lstStock.SelectedItem.ListSubItems(3).Text
Me.txtReorderpoint.Text = lstStock.SelectedItem.ListSubItems(4).Text
Me.txtMaximum.Text = lstStock.SelectedItem.ListSubItems(5).Text
Me.txtPrice.Text = lstStock.SelectedItem.ListSubItems(6).Text
Me.txtRPrice.Text = lstStock.SelectedItem.ListSubItems(8).Text
Me.txtRPlus.Text = lstStock.SelectedItem.ListSubItems(7).Text
Me.txtWPrice.Text = lstStock.SelectedItem.ListSubItems(10).Text
Me.txtWPlus.Text = lstStock.SelectedItem.ListSubItems(9).Text
'Me.cmdOk.Enabled = False
Me.cmdUpdate.Enabled = True
Me.cmdDelete.Enabled = True

Me.txtID.Enabled = False
Me.txtBrgy.Enabled = True
Me.txtCompany.Enabled = True
Me.cboCategory.Enabled = True
Me.txtReorderpoint.Enabled = True
Me.txtMaximum.Enabled = True
Me.txtPrice.Enabled = True
Me.txtRPrice.Enabled = True
Me.txtRPlus.Enabled = True
Me.txtWPrice.Enabled = True
Me.txtWPlus.Enabled = True
End Sub

Private Sub txtCompany_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub txtPrice_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57
    Case 46
    Case vbKeyBack
    Case Else
        KeyAscii = 0
End Select
End Sub


Private Sub txtRPlus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtRPrice.Text = Val(Me.txtRPlus.Text) + Val(Me.txtPrice.Text)
End If
Select Case KeyAscii
    Case 48 To 57
    Case 46
    Case vbKeyBack
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtRPrice_Change()
On Error Resume Next
'txtRPlus.Text = ((Val(Me.txtRPrice.Text) / Val(Me.txtPrice.Text)) * 100) - 100
txtRPlus.Text = Format(txtRPlus.Text, "###.00")
End Sub

Private Sub txtWPlus_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtWPrice.Text = Val(Me.txtWPlus.Text) + Val(Me.txtPrice.Text)
End If
Select Case KeyAscii
    Case 48 To 57
    Case 46
    Case vbKeyBack
    Case Elses
        KeyAscii = 0
End Select
End Sub

Private Sub txtWPrice_Change()
On Error Resume Next
'txtWPlus.Text = ((Val(Me.txtWPrice.Text) / Val(Me.txtPrice.Text)) * 100) - 100
txtWPlus.Text = Format(txtWPlus.Text, "###.00")
End Sub

Sub clear()
Me.txtID.Text = ""
Me.txtBrgy.Text = ""
Me.txtPopulation.Text = ""
Me.cboFlood.Text = ""
Me.txtDengue.Text = ""
Me.txtMalaria.Text = ""
Me.txtMeasles.Text = ""
    
End Sub
Sub rec(ByVal lookup As String)

On Error Resume Next
With lstBrgy
Set lstBrgy.SmallIcons = img32
Set lstBrgy.Icons = img32
    .ListItems.clear
    .ColumnHeaders.clear
    .ColumnHeaders.Add , , "ID No.", 2000
    .ColumnHeaders.Add , , "Barangay", 3000
     .ColumnHeaders.Add , , "Population", 2700
     .ColumnHeaders.Add , , "Flood-Level", 2300
     .ColumnHeaders.Add , , "Dengue-Level", 2500
     .ColumnHeaders.Add , , "Malaria-Level", 2500
     .ColumnHeaders.Add , , "Measles-Level", 2500
     
End With

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '%" & search & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
Do Until rs.EOF
With rs
    Set lst = lstBrgy.ListItems.Add(, , rs!BrgyID, 2, 2)
        lst.ListSubItems.Add , , !Barangay
        lst.ListSubItems.Add , , !Population
        lst.ListSubItems.Add , , !Flood
        lst.ListSubItems.Add , , !Dengue
        lst.ListSubItems.Add , , !Malaria
        lst.ListSubItems.Add , , !Measles
       
End With
        rs.MoveNext
    Loop
End Sub

Private Sub lstBrgy_DblClick()
On Error Resume Next
If lstBrgy.ListItems.Count = 0 Then
MsgBox "Empty List.", vbExclamation, ""
Exit Sub
End If


Me.txtID.Text = lstBrgy.SelectedItem.Text
Me.txtBrgy.Text = lstBrgy.SelectedItem.ListSubItems(1).Text
Me.txtPopulation.Text = lstBrgy.SelectedItem.ListSubItems(2).Text
Me.cboFlood.Text = lstBrgy.SelectedItem.ListSubItems(3).Text
Me.txtDengue.Text = lstBrgy.SelectedItem.ListSubItems(4).Text
Me.txtMalaria.Text = lstBrgy.SelectedItem.ListSubItems(5).Text
Me.txtMeasles.Text = lstBrgy.SelectedItem.ListSubItems(6).Text

'Me.cmdOk.Enabled = False
Me.cmdUpdate.Enabled = True
Me.cmdDelete.Enabled = True

Me.txtID.Enabled = False
Me.txtBrgy.Enabled = True
Me.txtPopulation.Enabled = True
Me.cboFlood.Enabled = True
Me.txtDengue.Enabled = True
Me.txtMalaria.Enabled = True
Me.txtMeasles.Enabled = True

End Sub

Private Sub txtDengue_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57
    Case 46
    Case vbKeyBack
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtMalaria_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57
    Case 46
    Case vbKeyBack
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtMeasles_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 48 To 57
    Case 46
    Case vbKeyBack
    Case Else
        KeyAscii = 0
End Select
End Sub
