VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BD0C1912-66C3-49CC-8B12-7B347BF6C846}#13.4#0"; "CODEJO~1.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5550
   ControlBox      =   0   'False
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   135
      Left            =   0
      ScaleHeight     =   135
      ScaleWidth      =   5655
      TabIndex        =   8
      Top             =   960
      Width           =   5655
   End
   Begin VB.TextBox txtUser 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtPass 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   1
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Login"
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   5655
      TabIndex        =   4
      Top             =   0
      Width           =   5655
      Begin VB.Image Image1 
         Height          =   765
         Left            =   720
         Picture         =   "frmLogin.frx":0000
         Top             =   120
         Width           =   750
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Please enter correct  user name and password to access the system ..."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   435
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   3855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   3720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmLogin.frx":0782
      OLEDBString     =   $"frmLogin.frx":081D
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin XtremeSkinFramework.SkinFramework SkinFramework 
      Left            =   600
      Top             =   3360
      _Version        =   851972
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
Dim ctr As Integer

If txtUser = "" And txtPass = "" Then
MsgBox "Please Input Username and Password.", vbExclamation, ""
Exit Sub

ElseIf txtPass = "" Then
MsgBox "Please Input Password.", vbExclamation, ""
Exit Sub
End If


Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_User where Username like '" & txtUser.Text & "'", cn, 1, 2
If rs!Password <> Me.txtPass.Text Then
        
        MsgBox "Wrong Password", vbExclamation, ""
        txtUser.Text = ""
        txtPass.Text = ""
        Exit Sub
Else
    With rs
        User = !UserName
        pass = !Password
        strname = !Name
    End With
'        If User = txtUser.Text And pass = txtPass.Text Then
'            Set rs = New ADODB.Recordset
'            rs.Open "Select * from tbl_LogHistory", cn, 1, 2
'            With rs
'                .AddNew
'                !UserName = User
'                !DateLog = Date
'                !TimeLog = Time
'                .Update
'            End With
           mmain.MDIStatus.Panels(3) = strname
            Unload Me
            mmain.Show
           
'        Else
'            MsgBox "Wrong Username / Password", vbCritical, "Error"
'            txtUser.Text = ""
'            txtPass.Text = ""
'        End If
SkinFramework.LoadSkin App.Path & "\Styles\Vista.cjstyles", "NormalBlue.ini"
SkinFramework.ApplyWindow Me.hWnd
SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End If

End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub Form_Load()
txtPass.PasswordChar = Chr(149)


SkinFramework.LoadSkin App.Path & "\Styles\Vista.cjstyles", "NormalBlue.ini"
    SkinFramework.ApplyWindow Me.hWnd
    SkinFramework.ApplyOptions = SkinFramework.ApplyOptions Or xtpSkinApplyMetrics
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
KeyAscii = Capitalize(KeyAscii)
End Sub


