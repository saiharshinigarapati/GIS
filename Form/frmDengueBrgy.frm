VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDengueBrgy 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dengue"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   19425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      Picture         =   "frmDengueBrgy.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Picture         =   "frmDengueBrgy.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   8040
      Width           =   3375
      Begin VB.TextBox txtCategory 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Barangay Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
   Begin MSComctlLib.ListView lstBrgy 
      Height          =   7215
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   19335
      _ExtentX        =   34105
      _ExtentY        =   12726
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483634
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
      Left            =   120
      Top             =   10440
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
            Picture         =   "frmDengueBrgy.frx":1360
            Key             =   "bar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":1C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":2914
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":31EE
            Key             =   "guy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":3AC8
            Key             =   "trolley"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":43A2
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":4C7C
            Key             =   "app"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":5556
            Key             =   "right"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":5E30
            Key             =   "line"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":670A
            Key             =   "exclaimation"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":6FE4
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":78BE
            Key             =   "db"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":8198
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":8A72
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":934C
            Key             =   "gng"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":9C26
            Key             =   "key"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":A500
            Key             =   "arrows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":ADDA
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":B6B4
            Key             =   "magnifier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":BF8E
            Key             =   "synon"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":C868
            Key             =   "people"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":D142
            Key             =   "silverlock"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":DA1C
            Key             =   "server"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":E2F6
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDengueBrgy.frx":EBD0
            Key             =   "plus"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmDengueBrgy.frx":F4AA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      Caption         =   "Dengue Outbreak Area"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmDengueBrgy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub rec(ByVal lookup As String)


With lstBrgy
Set lstBrgy.SmallIcons = img32
Set lstBrgy.Icons = img32
    .ListItems.clear
    .ColumnHeaders.clear
    .ColumnHeaders.Add , , "ID No.", 2000
    .ColumnHeaders.Add , , "Barangay", 3000
     .ColumnHeaders.Add , , "Population", 2700
     .ColumnHeaders.Add , , "No. of Dengue Cases", 2300
     
End With

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '%" & Me.txtCategory & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
Do Until rs.EOF
With rs
    Set lst = lstBrgy.ListItems.Add(, , rs!BrgyID, 2, 2)
        lst.ListSubItems.Add , , !Barangay
        lst.ListSubItems.Add , , !Population
        lst.ListSubItems.Add , , !Dengue

       
End With
        rs.MoveNext
    Loop
End Sub

Private Sub txtCategory_Click()
rec Me.txtCategory.Text
End Sub

Private Sub txtCategory_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rec ""
End If
End Sub

Private Sub cmdClose_Click()

    If MsgBox("Are you sure you want to close?", vbYesNo, "") = vbYes Then
        Unload Me
    End If

End Sub

Private Sub cmdPrint_Click()
Set rs = New ADODB.Recordset
     rs.Open "Select * from tbl_Barangay where Barangay like '%" & Me.txtCategory & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
         Set rptDengueArea.DataSource = rs

         rptDengueArea.Show
End Sub

Private Sub Form_Load()
rec ""

'Flood
'Set rs = New ADODB.Recordset
'rs.Open " Select * from tbl_DengueLevel", cn, 1, 2
'Do Until rs.EOF
'Me.txtCategory.AddItem rs!Dengue
'rs.MoveNext
'Loop
End Sub

