VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   11490
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   15330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   11490
   ScaleWidth      =   15330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   5273
      ScaleHeight     =   3105
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   4178
      Width           =   4785
      Begin VB.TextBox txtUserName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2685
         TabIndex        =   5
         Text            =   "Admin"
         Top             =   1080
         Width           =   1485
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2685
         TabIndex        =   4
         Text            =   "Admin"
         Top             =   1515
         Width           =   1485
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00F0F0F0&
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   0
         ScaleHeight     =   525
         ScaleWidth      =   4725
         TabIndex        =   1
         Top             =   2550
         Width           =   4755
         Begin VB.CommandButton cmdCancel 
            BackColor       =   &H00F9F9F9&
            Caption         =   "&Cancel"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   135
            Width           =   1470
         End
         Begin VB.CommandButton cmdNew 
            BackColor       =   &H00F9F9F9&
            Caption         =   "Login"
            BeginProperty Font 
               Name            =   "Garamond"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   735
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   135
            Width           =   1470
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2250
         TabIndex        =   8
         Top             =   -15
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1695
         TabIndex        =   7
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   195
         Left            =   1695
         TabIndex        =   6
         Top             =   1560
         Width           =   825
      End
      Begin VB.Image imgLogin 
         Height          =   1755
         Left            =   0
         Picture         =   "frmLogin.frx":10E33
         Stretch         =   -1  'True
         Top             =   450
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -15
         Picture         =   "frmLogin.frx":1740A
         Stretch         =   -1  'True
         Top             =   -15
         Width           =   4785
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    If Not txtUserName.Text = "Admin" Or Not txtPassword.Text = "Admin" Then
        MsgBox "Invalid user or password", vbCritical, "Invald user/password"
        txtUserName.SetFocus
        Exit Sub
    End If
    
    If FrmMain.Enabled = False Then
        FrmMain.Enabled = True
    Else
        FrmMain.Show
        Unload frmLogin
    End If
End Sub

