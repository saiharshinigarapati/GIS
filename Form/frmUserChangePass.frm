VERSION 5.00
Begin VB.Form frmUserChangePass 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton txtOk 
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
      Left            =   2040
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton txtCancel 
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
      Left            =   3480
      TabIndex        =   5
      Top             =   2400
      Width           =   1335
   End
   Begin VB.TextBox txtRe 
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
      Left            =   2040
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox txtUser 
      Enabled         =   0   'False
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
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.TextBox txtOld 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox txtNew 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Old-Password"
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
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1215
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
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
      Top             =   1320
      Width           =   3135
   End
End
Attribute VB_Name = "frmUserChangePass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Form_Load()
txtOld.PasswordChar = Chr(149)
txtNew.PasswordChar = Chr(149)
txtRe.PasswordChar = Chr(149)
End Sub

Private Sub txtCancel_Click()
Unload Me
frmUser.Visible = True

End Sub

Private Sub txtOk_Click()

If Me.txtUser.Text = "" Then
MsgBox "Please enter usename", vbInformation, ""
Exit Sub
End If

If Me.txtOld.Text = "" Then
MsgBox "Please enter the old password", vbInformation, ""
Exit Sub
End If

If Me.txtRe.Text = "" And Me.txtNew.Text = "" Then
MsgBox "Please enter the new password or retype new password", vbInformation, ""
Exit Sub
End If

If txtNew.Text <> txtRe.Text Then
MsgBox "Password Mismatch", vbCritical, "Message"
txtNew.Text = ""
txtRe.Text = ""
Exit Sub
Else

'On Error GoTo err:
Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_User where Username like '" & txtUser.Text & "' and Password like '" & txtOld.Text & "' ", cn, 1, 2
With rs
        !Password = txtNew.Text
        .Update
End With
    MsgBox "Your password was successfully change!", vbInformation, "Message"
   txtUser.Text = ""
    txtOld.Text = ""
    txtNew.Text = ""
    txtRe.Text = ""
    Unload Me
    frmUser.Visible = True
    
'
'    Set rs = New ADODB.Recordset
'    rs.Open "Select * from tbl_Audit", cn, 1, 2
'    With rs
'    .AddNew
'    !UserName = frmMain.lblUser.Caption
'    !Action = "Change a password"
'    !DateIn = frmMain.Label5
'    !Time = frmMain.Label6
'    .Update
'    End With

Exit Sub
End If
'err:
'    MsgBox "Please enter correct Username/Old password", vbCritical, "Message"
'    txtUser.Text = ""
'    txtOld.Text = ""
'    txtNew.Text = ""
'    txtRe.Text = ""

End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
KeyAscii = Capitalize(KeyAscii)
End Sub


