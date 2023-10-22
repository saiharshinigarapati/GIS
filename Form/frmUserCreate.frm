VERSION 5.00
Begin VB.Form frmUserCreate 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtName 
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
      Left            =   1680
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtRetype 
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.TextBox txtPass 
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
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
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
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
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
      Left            =   3120
      TabIndex        =   6
      Top             =   3120
      Width           =   1335
   End
   Begin VB.ComboBox cboPos 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      TabIndex        =   4
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Complete Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-type Password"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "frmUserCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = Capitalize(KeyAscii)
End Sub


Private Sub cboPos_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub cmdCancel_Click()
Unload Me
frmUser.Show

End Sub

Private Sub cmdOk_Click()

If Me.cboPos.Text = "" Then
MsgBox "Please select a position.", vbExclamation, "Message"
Exit Sub
End If

If txtPass.Text <> txtRetype.Text Then
MsgBox "Password Mismatch", vbCritical, "Message"
txtPass.Text = ""
txtRetype.Text = ""
Exit Sub
Else
On Error GoTo err
Set rs = New ADODB.Recordset
rs.Open "Select *from tbl_User", cn, 1, 2
With rs
    .AddNew
    !UserName = txtUser.Text
    !Password = txtRetype.Text
    !Name = txtName.Text
    !Position = cboPos.Text
    .Update
End With
MsgBox "Account was successfully create!", vbInformation, "Message"
txtUser.Text = ""
txtPass.Text = ""
txtRetype.Text = ""
txtName.Text = ""
cboPos.Text = ""
Unload Me
frmUser.Show

'Set rs = New ADODB.Recordset
'    rs.Open "Select * from tbl_Audit", cn, 1, 2
'    With rs
'    .AddNew
'    !Name = frmMain.lblUser.Caption
'    !Action = "Create a user account"
'    !DateIn = frmMain.Label5
'    !TimeIn = frmMain.Label6
'    .Update
'    End With

Exit Sub
End If
err:
MsgBox "Username already register, Please choose other username!", vbCritical, "Message"
txtUser.Text = ""
txtPass.Text = ""
txtRetype.Text = ""
txtName.Text = ""
cboPos.Text = ""
End Sub

Private Sub Form_Load()
txtPass.PasswordChar = Chr(149)
txtRetype.PasswordChar = Chr(149)

cboPos.AddItem "Client"
cboPos.AddItem "Administrator"

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = Capitalize(KeyAscii)
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
KeyAscii = Capitalize(KeyAscii)
End Sub


