VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFlood 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                              Flood Level"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   -120
      Width           =   4215
      Begin VB.TextBox txtDescription 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         Left            =   1320
         TabIndex        =   0
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtCategory 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
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
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label12 
         Caption         =   "Flood Level"
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
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "Level No :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lstCategory 
      Height          =   3735
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   6588
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
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   5160
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFlood.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFlood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

If Me.cmdClose.Caption = "Cancel" Then
Me.txtCategory.Text = ""
Me.txtDescription.Text = ""
Me.cmdOk.Visible = True
Me.cmdClose.Caption = "Close"
Exit Sub
Else
Unload Me
End If

End Sub

Private Sub cmdOk_Click()

If Me.txtDescription.Text = "" Then
MsgBox "Please enter category description", vbExclamation, ""
Exit Sub
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_FloodLevel ", cn, 1, 2
With rs
        .AddNew
        !Flood = Me.txtDescription.Text
        .Update
End With
        MsgBox "Record Saved!", vbInformation, ""
        rec
End Sub
Sub rec()
Set lstCategory.SmallIcons = ImageList1
Set lstCategory.Icons = ImageList1
With lstCategory
    .ListItems.clear
    .ColumnHeaders.clear
    .ColumnHeaders.Add , , "Category No.", 0
    .ColumnHeaders.Add , , "Flood-Level", 4250
End With
Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_FloodLevel ", cn, 1, 2

Do Until rs.EOF
With rs
    Set lst = lstCategory.ListItems.Add(, , !FloodID)
        lst.ListSubItems.Add , , !Flood, 1, 1
      
End With
        rs.MoveNext
    Loop
End Sub

Private Sub cmdUpdate_Click()
Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_FloodLevel where FloodID like '" & Me.txtCategory.Text & "'", cn, 1, 2
With rs
    
        !Flood = Me.txtDescription.Text
        .Update
End With
        MsgBox "Record update!", vbInformation, ""
        rec
End Sub

Private Sub Form_Load()
rec
End Sub



Private Sub lstCategory_DblClick()
Me.txtCategory.Text = lstCategory.SelectedItem.Text
Me.txtDescription.Text = lstCategory.SelectedItem.ListSubItems(1).Text
Me.cmdOk.Visible = False
Me.cmdClose.Caption = "Cancel"
End Sub
