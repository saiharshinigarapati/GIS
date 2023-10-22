VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{859DE455-7B89-457A-9743-C1081A14D235}#7.0#0"; "SuperPicture.ocx"
Begin VB.Form frmMapVIew 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MAP VIEWING"
   ClientHeight    =   10770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18630
   ForeColor       =   &H80000011&
   Icon            =   "frmMapView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10770
   ScaleWidth      =   18630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "v"
   Begin VB.Frame Frame6 
      Caption         =   " Barangay "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   14520
      TabIndex        =   32
      Top             =   7320
      Width           =   3975
      Begin VB.TextBox txtSearch 
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
         TabIndex        =   33
         Top             =   2400
         Width           =   3735
      End
      Begin MSComctlLib.ListView lstBrgy 
         Height          =   2055
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648447
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   14520
      TabIndex        =   19
      Top             =   3480
      Width           =   3855
      Begin VB.PictureBox Picture9 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   22
         Top             =   1560
         Width           =   735
      End
      Begin VB.PictureBox Picture8 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.PictureBox Picture7 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   20
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000B&
         Caption         =   "Map Legend"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   26
         Top             =   120
         Width           =   2010
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000B&
         Caption         =   "River"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   25
         Top             =   1680
         Width           =   7890
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000B&
         Caption         =   "Main Road"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   24
         Top             =   1200
         Width           =   7890
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000B&
         Caption         =   "Road"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   23
         Top             =   720
         Width           =   7890
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   14520
      TabIndex        =   10
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   12
         Top             =   1200
         Width           =   735
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   11
         Top             =   1800
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         ScaleHeight     =   495
         ScaleWidth      =   735
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblBarangay 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   31
         Top             =   120
         Width           =   3090
      End
      Begin VB.Label lblPopulation 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   36
         Top             =   2520
         Width           =   7890
      End
      Begin VB.Image Image4 
         Height          =   450
         Left            =   360
         Picture         =   "frmMapView.frx":0CCA
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Population       :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   35
         Top             =   2520
         Width           =   7890
      End
      Begin VB.Label lblMeasles 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   30
         Top             =   1920
         Width           =   7890
      End
      Begin VB.Label lblMalaria 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   29
         Top             =   1320
         Width           =   7890
      End
      Begin VB.Label lblDengue 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2400
         TabIndex        =   28
         Top             =   720
         Width           =   7890
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   360
         Picture         =   "frmMapView.frx":122B
         Top             =   1800
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "frmMapView.frx":1EF5
         Top             =   1200
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmMapView.frx":2BBF
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "High Flood Susceptibility"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   7890
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Moderate Flood Susceptibility"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   16
         Top             =   1320
         Width           =   7890
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Low Flood Susceptibility"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   15
         Top             =   1920
         Width           =   7890
      End
      Begin VB.Label txtCaption 
         BackColor       =   &H8000000E&
         Caption         =   "Flood Hazard"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   2970
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " View Option "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   14520
      TabIndex        =   7
      Top             =   5880
      Width           =   3975
      Begin VB.OptionButton Option3 
         Caption         =   "Urban Area Map"
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
         TabIndex        =   37
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Flood Map"
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
         Left            =   1800
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Normal Map"
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
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CheckBox chkPanMode 
      Caption         =   "Panning Mode"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1785
      TabIndex        =   6
      Top             =   195
      Width           =   1665
   End
   Begin VB.CheckBox chkUseQuickBar 
      Caption         =   "Enable QuickBar"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3510
      TabIndex        =   5
      Top             =   195
      Value           =   1  'Checked
      Width           =   2010
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "-"
      Height          =   285
      Left            =   1290
      TabIndex        =   4
      Top             =   135
      Width           =   300
   End
   Begin VB.CommandButton cmdIn 
      Caption         =   "+"
      Height          =   285
      Left            =   945
      TabIndex        =   3
      Top             =   135
      Width           =   300
   End
   Begin VB.TextBox txtZoomLevel 
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   840
   End
   Begin VB.ListBox lstEvents 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   9705
      Width           =   14295
   End
   Begin SuperPicture.SuperPicCtl SuperPicCtl1 
      Height          =   9030
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   14310
      _ExtentX        =   25241
      _ExtentY        =   15928
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   14640
      TabIndex        =   18
      Top             =   240
      Width           =   3855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   14640
      TabIndex        =   27
      Top             =   3600
      Width           =   3855
   End
   Begin MSComctlLib.ImageList img32 
      Left            =   13200
      Top             =   0
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
            Picture         =   "frmMapView.frx":3889
            Key             =   "bar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":4163
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":4E3D
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":5717
            Key             =   "guy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":5FF1
            Key             =   "trolley"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":68CB
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":71A5
            Key             =   "app"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":7A7F
            Key             =   "right"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":8359
            Key             =   "line"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":8C33
            Key             =   "exclaimation"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":950D
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":9DE7
            Key             =   "db"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":A6C1
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":AF9B
            Key             =   "earth"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":B875
            Key             =   "gng"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":C14F
            Key             =   "key"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":CA29
            Key             =   "arrows"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":D303
            Key             =   "lock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":DBDD
            Key             =   "magnifier"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":E4B7
            Key             =   "synon"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":ED91
            Key             =   "people"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":F66B
            Key             =   "silverlock"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":FF45
            Key             =   "server"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":1081F
            Key             =   "minus"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMapView.frx":110F9
            Key             =   "plus"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "Note: Double Click Listview to do some function.."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   9360
      TabIndex        =   38
      Top             =   240
      Width           =   5010
   End
End
Attribute VB_Name = "frmMapVIew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*' Comments are sparse in here... Should be self-explanatory.  More detailed comments are in the control.
'*'
Option Explicit


Private Sub chkPanMode_Click()

    '*' Toggle pan mode based upon the check box.
    '*'
    If chkPanMode.Value = 1 Then
        SuperPicCtl1.PanActive = True
    Else
        SuperPicCtl1.PanActive = False
    End If
    
End Sub

Private Sub chkUseQuickBar_Click()

    '*' Toggle the use of the quickbar
    '*'
    If chkUseQuickBar.Value = 1 Then
        SuperPicCtl1.UseQuickBar = True
    Else
        SuperPicCtl1.UseQuickBar = False
    End If
    
End Sub

Private Sub cmdIn_Click()

    '*' Zoom in by an increment of 10% if the percentage is less than 100%
    '*'
    If SuperPicCtl1.Zoom < 1000 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom + 10
    End If
    
End Sub

Private Sub cmdOut_Click()

    '*' Zoom out by an incrment of 10% if the percentage is more than 10%
    '*'
    If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 10
    End If
    
End Sub

Private Sub Form_DblClick()

    '*' Unload the image.
    '*'
    SuperPicCtl1.UnloadImage
    
End Sub

Private Sub Form_Load()

    '*' By default, use the quickbar.
    '*'
    SuperPicCtl1.LoadImage "" & App.Path & "\Map\Normal.jpg"
    SuperPicCtl1.UseQuickBar = True
    Option1.Value = True
    
    If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 80
    End If
    
    txtCaption.Caption = "Brgy."
    Label1.Caption = "Dengue Cases:"
    Label2.Caption = "Malaria Cases:"
    Label3.Caption = "Measles Cases:"
    Label4.Visible = True
    Image4.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Image1.Visible = True
    Image2.Visible = True
    Image3.Visible = True
    Image4.Visible = True
    rec ""

End Sub

Private Sub Form_Resize()

On Error Resume Next

    '*' Store the position if the window state is Max'd or Min'd.  Do it before resizing, since a restore will return
    '*' it to this size.
    '*'
'    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then
'        SuperPicCtl1.StorePosition
'    End If
    
'    SuperPicCtl1.Move 0, 600, Me.ScaleWidth, Me.ScaleHeight - 660 - lstEvents.Height
'    lstEvents.Move 0, Me.ScaleHeight - lstEvents.Height, Me.ScaleWidth, lstEvents.Height
    
'    lstEvents.AddItem "Resize()"
'    CleanList
    
    '*' Recall the position.  Will only work on Restore.
    '*'
 '   SuperPicCtl1.RecallPosition
        
End Sub



Private Sub lblBarangay_Change()
Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & lblBarangay.Caption & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     'lblBarangay.Caption = "Awasian"
     lblPopulation.Caption = !Population
     
End With
End Sub

Private Sub lstBrgy_DblClick()

If Me.Option1.Value = True Then
lblBarangay.Visible = True
'Me.lblBarangay.Caption = lstBrgy.SelectedItem.Text
Me.lblBarangay.Caption = lstBrgy.SelectedItem.ListSubItems(1).Text
Me.lblPopulation.Caption = lstBrgy.SelectedItem.ListSubItems(2).Text
'Me.cboFlood.Text = lstBrgy.SelectedItem.ListSubItems(3).Text
Me.lblDengue.Caption = lstBrgy.SelectedItem.ListSubItems(4).Text
Me.lblMalaria.Caption = lstBrgy.SelectedItem.ListSubItems(5).Text
Me.lblMeasles.Caption = lstBrgy.SelectedItem.ListSubItems(6).Text
End If

If Me.Option3.Value = True Then
lblBarangay.Visible = True


If lstBrgy.SelectedItem.ListSubItems(1).Text = "Urban Area - 1" Then
        
        SuperPicCtl1.LoadImage "" & App.Path & "\Map\Urban1.jpg"
       
        If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 80
        End If
        
        Me.lblBarangay.Caption = lstBrgy.SelectedItem.ListSubItems(1).Text
        
    Exit Sub
ElseIf lstBrgy.SelectedItem.ListSubItems(1).Text = "Urban Area - 2" Then

        SuperPicCtl1.LoadImage "" & App.Path & "\Map\Urban2.jpg"
       
        If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 80
        End If
        
         Me.lblBarangay.Caption = lstBrgy.SelectedItem.ListSubItems(1).Text

End If
End If

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then

        Frame6.Caption = " Barangay "
        rec ""
         SuperPicCtl1.LoadImage "" & App.Path & "\Map\Normal.jpg"
       
        If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 80
        End If
       
      '  Frame2.Visible = False
       ' Frame3.Visible = False
    txtCaption.Caption = "Brgy."
    Label1.Caption = "Dengue Cases:"
    Label2.Caption = "Malaria Cases:"
    Label3.Caption = "Measles Cases:"
    lblPopulation.Visible = False
    Picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Image1.Visible = True
    Image2.Visible = True
    Image3.Visible = True
    lblDengue.Visible = True
    lblMalaria.Visible = True
    lblMeasles.Visible = True
    
    lblPopulation.Visible = True
     
    Label4.Visible = True
    Image4.Visible = True
    lstBrgy.Enabled = True
End If
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then

        Frame6.Caption = " Barangay "
        rec ""
        txtCaption.Caption = "Flood Hazard"
       lblBarangay.Visible = False
        
        SuperPicCtl1.LoadImage "" & App.Path & "\Map\Flood.jpg"
        
       If SuperPicCtl1.Zoom > 10 Then
       SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 80
       End If
        
       ' Frame2.Visible = True
         'Frame3.Visible = True
    Label1.Caption = "High Flood Susceptibility"
    Label2.Caption = "Moderate Flood Susceptibility"
    Label3.Caption = "Low Flood Susceptibility"
    Picture1.Visible = True
    Picture2.Visible = True
    Picture3.Visible = True
    lblPopulation.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    lblDengue.Visible = False
    lblMalaria.Visible = False
    lblMeasles.Visible = False
    Label4.Visible = False
    Image4.Visible = False
    lblPopulation.Visible = False
    
    lblDengue.Caption = ""
    lblMalaria.Caption = ""
    lblMeasles.Caption = ""
    lblPopulation.Caption = ""
    lstBrgy.Enabled = False
End If
End Sub








Private Sub Option3_Click()
If Me.Option3.Value = True Then
Frame6.Caption = " Urban Area "
urbanArea ""
lstBrgy.Enabled = True


SuperPicCtl1.LoadImage "" & App.Path & "\Map\Urban1.jpg"
       
        If SuperPicCtl1.Zoom > 10 Then
        SuperPicCtl1.Zoom = SuperPicCtl1.Zoom - 80
        End If


Label1.Caption = "High Flood Susceptibility"
    Label2.Caption = "Moderate Flood Susceptibility"
    Label3.Caption = "Low Flood Susceptibility"
    Picture1.Visible = True
    Picture2.Visible = True
    Picture3.Visible = True
    lblPopulation.Visible = False
    Image1.Visible = False
    Image2.Visible = False
    Image3.Visible = False
    lblDengue.Visible = False
    lblMalaria.Visible = False
    lblMeasles.Visible = False
    Label4.Visible = False
    Image4.Visible = False
    lblPopulation.Visible = False
    
    lblDengue.Caption = ""
    lblMalaria.Caption = ""
    lblMeasles.Caption = ""
    lblPopulation.Caption = ""
End If
End Sub

Private Sub SuperPicCtl1_Click()



   lstEvents.AddItem "Click()"
    CleanList
    
    
   
    
End Sub

Private Sub SuperPicCtl1_DblClick()

    'lstEvents.AddItem "DblClick()"
    'CleanList
    
    'SuperPicCtl1.LoadImage
        
End Sub

Private Sub SuperPicCtl1_GotFocus()

    lstEvents.AddItem "GotFocus()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_KeyDown(KeyCode As Integer, Shift As Integer)

    lstEvents.AddItem "KeyDown(" & KeyCode & ", " & Shift & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_KeyPress(KeyAscii As Integer)

    lstEvents.AddItem "KeyPress(" & KeyAscii & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_KeyUp(KeyCode As Integer, Shift As Integer)

    lstEvents.AddItem "KeyUp(" & KeyCode & ", " & Shift & ")"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_LostFocus()

    lstEvents.AddItem "LostFocus()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lstEvents.AddItem "MouseDown(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
    CleanList

End Sub

Private Sub SuperPicCtl1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Awasian.......................................
If Option1.Value = True And Option2.Value = False Then
Dim A As Integer
Dim B As Integer
A = 993
B = 579
If X = A And Y = B Then
'txtCaption.Caption = "Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Awasian" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Awasian"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Bagong Lungsod.......................................

Dim C As Integer
Dim D As Integer
C = 1684
D = 682
If X = C And Y = D Then
'txtCaption.Caption = "Bagong Lungsod"
If Option1.Value = True And Option2.Value = False Then
Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Bagong Lungsod (Pob.)" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Bagong Lungsod (Pob.)"
     lblPopulation.Caption = !Population
End With
End If
End If

'.........................................

'Bioto.......................................
If Option1.Value = True And Option2.Value = False Then
Dim E As Integer
Dim F As Integer
E = 1480
F = 771
If X = E And Y = F Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Bioto" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Bioto"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Buenavista.......................................
If Option1.Value = True And Option2.Value = False Then
Dim G As Integer
Dim H As Integer
G = 596
H = 190
If X = G And Y = H Then
'txtCaption.Caption = "Brgy. Buenavista"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Buenavista" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Buenavista"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Bongtod Pob. (East West).......................................
If Option1.Value = True And Option2.Value = False Then
Dim I As Integer
Dim J As Integer
I = 1662
J = 614
If X = I And Y = J Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Bongtod Pob. (East West)" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Bongtod Pob. (East West)"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Dagocdoc (Pob.).......................................
If Option1.Value = True And Option2.Value = False Then
Dim K As Integer
Dim L As Integer
K = 1662
L = 649
If X = K And Y = L Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Dagocdoc (Pob.)" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Dagocdoc (Pob.)"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Mabua.......................................
If Option1.Value = True And Option2.Value = False Then
Dim M As Integer
Dim N As Integer
M = 1705
N = 772
If X = M And Y = N Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Mabua" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Mabua"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Mabuhay.......................................
If Option1.Value = True And Option2.Value = False Then
Dim O As Integer
Dim P As Integer
O = 663
P = 1113
If X = O And Y = P Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Mabuhay" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Mabuhay"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Maitum.......................................
If Option1.Value = True And Option2.Value = False Then
Dim Q As Integer
Dim R As Integer
Q = 549
R = 659
If X = Q And Y = R Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Maitum" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Maitum"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Maticdum.......................................
If Option1.Value = True And Option2.Value = False Then
Dim Hz As Integer
Dim SR As Integer
Hz = 1051
SR = 1182
If X = Hz And Y = SR Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Maticdum" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Maticdum"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Pandanon.......................................
If Option1.Value = True And Option2.Value = False Then
Dim T As Integer
Dim U As Integer
T = 868
U = 1053
If X = T And Y = U Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Pandanon" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Pandanon"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Pangi.......................................
If Option1.Value = True And Option2.Value = False Then
Dim V As Integer
Dim W As Integer
V = 954
W = 318
If X = V And Y = W Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Pangi" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Pangi"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Quezon.......................................
If Option1.Value = True And Option2.Value = False Then
Dim Ww As Integer
Dim XW As Integer
Ww = 1255
XW = 786
If X = Ww And Y = XW Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Quezon" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Quezon"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Rosario.......................................
If Option1.Value = True And Option2.Value = False Then
Dim Ya As Integer
Dim Zb As Integer
Ya = 1649
Zb = 887
If X = Ya And Y = Zb Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Rosario" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Rosario"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Salvacion.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AB As Integer
Dim BB As Integer
AB = 1044
BB = 242
If X = AB And Y = BB Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Salvacion" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Salvacion"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'San Agustin Norte.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AC As Integer
Dim BD As Integer
AC = 1170
BD = 506
If X = AC And Y = BD Then
'txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "San Agustin Norte" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "San Agustin Norte"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'San Agustin Sur.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AE As Integer
Dim BF As Integer
AE = 1449
BF = 662
If X = AE And Y = BF Then
txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "San Agustin Sur" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "San Agustin Sur"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'San Antonio.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AG As Integer
Dim BH As Integer
AG = 1306
BH = 174
If X = AG And Y = BH Then
txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "San Antonio" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "San Antonio"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'San Isidro.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AI As Integer
Dim BJ As Integer
AI = 1196
BJ = 1007
If X = AI And Y = BJ Then
txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "San Antonio" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "San Antonio"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'San Jose.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AK As Integer
Dim BL As Integer
AK = 1434
BL = 942
If X = AK And Y = BL Then
txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "San Jose" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "San Jose"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................

'Telaje.......................................
If Option1.Value = True And Option2.Value = False Then
Dim AM As Integer
Dim BN As Integer
AM = 1623
BN = 751
If X = AM And Y = BN Then
txtCaption.Caption = "Brgy. Awasian"

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '" & "Telaje" & "'", cn, 1, 2
With rs
     lblDengue.Caption = !Dengue & " " & "People"
     lblMalaria.Caption = !Malaria & " " & "People"
     lblMeasles.Caption = !Measles & " " & "People"
     lblBarangay.Caption = "Telaje"
     lblPopulation.Caption = !Population
End With
End If
End If
'.........................................


lstEvents.AddItem "MouseMove(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
CleanList
    

End Sub

Private Sub SuperPicCtl1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


   lstEvents.AddItem "MouseUp(" & Button & ", " & Shift & ", " & X & ", " & Y & ")"
   CleanList
    
End Sub

Private Sub SuperPicCtl1_Paint()

    lstEvents.AddItem "Paint()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_Scroll()

    lstEvents.AddItem "Scroll()"
    CleanList
    
End Sub

Private Sub SuperPicCtl1_ZoomChanged(ByVal ZoomPercent As Long)

    lstEvents.AddItem "ZoomChanged(" & ZoomPercent & ")"
    CleanList
    
    '*' Toggle quickbar buttons based upon current percentage.
    '*'
    SuperPicCtl1.AllowZoomIn = (ZoomPercent < 1000)
    SuperPicCtl1.AllowZoomOut = (ZoomPercent > 10)
    

    
    txtZoomLevel.Text = ZoomPercent & "%"
    

    
End Sub

Private Sub SuperPicCtl1_ZoomInClick()

    lstEvents.AddItem "ZoomIn()"
    CleanList
    
    cmdIn_Click
    
End Sub

Private Sub SuperPicCtl1_ZoomOutClick()

    lstEvents.AddItem "ZoomOut()"
    CleanList
    
    cmdOut_Click
    
End Sub

Private Sub CleanList()

    If lstEvents.ListCount > 10 Then
        Do Until lstEvents.ListCount = 10
            Call lstEvents.RemoveItem(0)
        Loop
    End If
    
    lstEvents.ListIndex = lstEvents.ListCount - 1
    
End Sub

Sub rec(ByVal lookup As String)

'On Error Resume Next
With lstBrgy
Set lstBrgy.SmallIcons = img32
Set lstBrgy.Icons = img32
    .ListItems.clear
    .ColumnHeaders.clear
     .ColumnHeaders.Add , , "ID No.", 0
    .ColumnHeaders.Add , , "Barangay", 3700
     .ColumnHeaders.Add , , "Population", 0
     .ColumnHeaders.Add , , "Flood-Level", 0
     .ColumnHeaders.Add , , "Dengue-Level", 0
     .ColumnHeaders.Add , , "Malaria-Level", 0
     .ColumnHeaders.Add , , "Measles-Level", 0

     
End With

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_Barangay where Barangay like '%" & txtSearch.Text & "%'and Flag like '" & 1 & "' ORDER By BrgyID", cn, 1, 2
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

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.Option1.Value = True Then
    rec ""
    
    ElseIf Me.Option2.Value = True Then
    rec ""
    
    ElseIf Me.Option3.Value = True Then
    urbanArea ""

    End If
End If
End Sub

Sub urbanArea(ByVal lookup As String)

'On Error Resume Next
With lstBrgy
Set lstBrgy.SmallIcons = img32
Set lstBrgy.Icons = img32
    .ListItems.clear
    .ColumnHeaders.clear
     .ColumnHeaders.Add , , "ID No.", 0
    .ColumnHeaders.Add , , "Urban Area", 3700
     .ColumnHeaders.Add , , "Population", 0
    .ColumnHeaders.Add , , "Flood", 0

     
End With

Set rs = New ADODB.Recordset
rs.Open "Select * from tbl_UrbanArea where UrbanArea like '%" & txtSearch.Text & "%'and Flag like '" & 1 & "' ORDER By UrbanID", cn, 1, 2
Do Until rs.EOF
With rs
     Set lst = lstBrgy.ListItems.Add(, , rs!UrbanID, 2, 2)
        lst.ListSubItems.Add , , !urbanArea
        lst.ListSubItems.Add , , !Population
        lst.ListSubItems.Add , , !Flood
    
       
End With
        rs.MoveNext
    Loop
End Sub
