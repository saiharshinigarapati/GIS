VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6720
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
      Height          =   825
      Left            =   5640
      Picture         =   "frmUser.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete User"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5640
      Picture         =   "frmUser.frx":0696
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Change Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5640
      Picture         =   "frmUser.frx":09A0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&New User"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5640
      Picture         =   "frmUser.frx":1562
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin MSComctlLib.ListView lstUser 
      Height          =   3855
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   6800
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
      Left            =   4080
      Top             =   -120
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
            Picture         =   "frmUser.frx":1BDF
            Key             =   "bar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":24B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":3193
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":3A6D
            Key             =   "guy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":4347
            Key             =   "trolley"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":4C21
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":54FB
            Key             =   "app"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":5DD5
            Key             =   "right"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":66AF
            Key             =   "line"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":6F89
            Key             =   "exclaimation"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":7863
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":813D
            Key             =   "db"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":8A17
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":92F1
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":9BCB
            Key             =   "gng"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":A4A5
            Key             =   "key"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":AD7F
            Key             =   "arrows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":B659
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":BF33
            Key             =   "magnifier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":C80D
            Key             =   "synon"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":D0E7
            Key             =   "people"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":D9C1
            Key             =   "silverlock"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":E29B
            Key             =   "server"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":EB75
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUser.frx":F44F
            Key             =   "plus"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "User List"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   240
      X2              =   4680
      Y1              =   360
      Y2              =   360
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()

Unload Me

End Sub

Private Sub cmdDelete_Click()
On Error GoTo err
If MsgBox("Are you sure you want to delete?", vbYesNo, "Message!") = vbYes Then
Set rs = New ADODB.Recordset
rs.Open " Select * from tbl_User where Username like '" & lstUser.SelectedItem & "'", cn, 1, 2
rs.Delete
MsgBox " Account successfully was deleted ", vbInformation, "Message"
rec
End If
err:

End Sub

Private Sub cmdSave_Click()


Unload Me
frmUserCreate.Show 1
End Sub

Private Sub cmdUpdate_Click()
frmUserChangePass.txtUser.Text = lstUser.SelectedItem.Text
Unload Me
frmUserChangePass.Show 1
End Sub

Sub rec()
With lstUser
Set lstUser.SmallIcons = img32
Set lstUser.Icons = img32
    .ListItems.clear
    .ColumnHeaders.clear
    .ColumnHeaders.Add , , "Username", 1800
    .ColumnHeaders.Add , , "Complete-Name", 1800
    .ColumnHeaders.Add , , "Position", 1800
    .ColumnHeaders.Add , , "Password", 0
End With
Set rs = New ADODB.Recordset
rs.Open " Select * from tbl_User", cn, 1, 2
Do Until rs.EOF
Set lst = lstUser.ListItems.Add(, , rs!UserName, 15, 15)
    lst.ListSubItems.Add , , rs!Name
    lst.ListSubItems.Add , , rs!Position
    lst.ListSubItems.Add , , rs!Password
    rs.MoveNext
    Loop
End Sub

Private Sub Form_Load()
rec
End Sub

Private Sub lstUser_DblClick()
If MsgBox("You are sure you want to view", vbYesNo, "") = vbYes Then
frmUserPass.Text1.Text = lstUser.SelectedItem.ListSubItems(3).Text

'Set rs = New ADODB.Recordset
'    rs.Open "Select * from tbl_Audit", cn, 1, 2
'    With rs
'    .AddNew
'    !Name = frmMain.lblUser.Caption
'    !Action = "View a password"
'    !DateIn = frmMain.Label5
'    !TimeIn = frmMain.Label6
'    .Update
'    End With
Unload Me
frmUserPass.Show
'frmUserPass.Show 1
End If

End Sub


