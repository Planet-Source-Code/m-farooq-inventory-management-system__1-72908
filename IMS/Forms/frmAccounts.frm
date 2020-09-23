VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmAccounts 
   BackColor       =   &H00F0F0F0&
   BorderStyle     =   0  'None
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   5145
      Left            =   0
      ScaleHeight     =   5115
      ScaleWidth      =   6090
      TabIndex        =   8
      Top             =   3975
      Width           =   6120
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgAccounts 
         Height          =   5085
         Left            =   15
         TabIndex        =   9
         Top             =   15
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   8969
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   13408563
         ForeColorFixed  =   16777215
         BackColorBkg    =   15790320
         FocusRect       =   2
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
   End
   Begin VB.PictureBox picProdInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC9933&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   1140
      ScaleHeight     =   2745
      ScaleWidth      =   3675
      TabIndex        =   31
      Top             =   900
      Visible         =   0   'False
      Width           =   3705
      Begin VB.TextBox txtOpAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   36
         Top             =   1965
         Width           =   1380
      End
      Begin VB.TextBox txtOpStock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   315
         TabIndex        =   35
         Top             =   1950
         Width           =   1380
      End
      Begin VB.CommandButton cmdProdCancel 
         BackColor       =   &H00F9F9F9&
         Caption         =   "CANCEL"
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
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2400
         Width           =   945
      End
      Begin VB.TextBox txtSalePrice 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   315
         TabIndex        =   33
         Top             =   1320
         Width           =   1380
      End
      Begin VB.TextBox txtReOrder 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   34
         Top             =   1335
         Width           =   1380
      End
      Begin VB.TextBox txtProdName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   315
         TabIndex        =   32
         Top             =   735
         Width           =   3015
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00F9F9F9&
         Caption         =   "OK"
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
         Left            =   810
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1125
         TabIndex        =   44
         Top             =   -15
         Width           =   1305
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1935
         TabIndex        =   43
         Top             =   1725
         Width           =   1410
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   42
         Top             =   1710
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   41
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Order Point"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1935
         TabIndex        =   40
         Top             =   1095
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   315
         TabIndex        =   39
         Top             =   450
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product Info"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1110
         TabIndex        =   45
         Top             =   -15
         Width           =   1305
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   0
         Left            =   0
         Picture         =   "frmAccounts.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3735
      End
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BackColor       =   &H00CC9933&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   780
      ScaleHeight     =   2505
      ScaleWidth      =   4455
      TabIndex        =   22
      Top             =   1155
      Visible         =   0   'False
      Width           =   4485
      Begin VB.OptionButton optTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00CC9933&
         Caption         =   "Find By Exact Account Title"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   225
         TabIndex        =   28
         Top             =   495
         Width           =   2340
      End
      Begin VB.OptionButton optId 
         Appearance      =   0  'Flat
         BackColor       =   &H00CC9933&
         Caption         =   "Find By Account ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   225
         TabIndex        =   27
         Top             =   1230
         Width           =   1710
      End
      Begin VB.TextBox txtFindTitle 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   225
         TabIndex        =   26
         Top             =   765
         Width           =   3945
      End
      Begin VB.TextBox txtFindId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   25
         Top             =   1515
         Width           =   3945
      End
      Begin VB.CommandButton cmdFindCancel 
         BackColor       =   &H00F9F9F9&
         Caption         =   "Cancel"
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
         Left            =   2190
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1995
         Width           =   1470
      End
      Begin VB.CommandButton cmdFind 
         BackColor       =   &H00F9F9F9&
         Caption         =   "Find Record"
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
         Left            =   675
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1995
         Width           =   1470
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   1500
         TabIndex        =   29
         Top             =   -15
         Width           =   1380
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find Account"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1515
         TabIndex        =   30
         Top             =   0
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   270
         Index           =   1
         Left            =   0
         Picture         =   "frmAccounts.frx":0E2D
         Stretch         =   -1  'True
         Top             =   15
         Width           =   4455
      End
   End
   Begin VB.PictureBox picAccounts 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   4005
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   6090
      TabIndex        =   3
      Top             =   -15
      Width           =   6120
      Begin VB.CommandButton cmdProdInfo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Product Info"
         Enabled         =   0   'False
         Height          =   285
         Left            =   4935
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2535
         Width           =   1005
      End
      Begin VB.ComboBox cmbType 
         Height          =   315
         ItemData        =   "frmAccounts.frx":1C5A
         Left            =   3395
         List            =   "frmAccounts.frx":1C82
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   1875
         Width           =   2500
      End
      Begin VB.TextBox txtId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   195
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1875
         Width           =   1200
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   195
         TabIndex        =   1
         Top             =   2535
         Width           =   4725
      End
      Begin VB.ComboBox cmbHead 
         Height          =   315
         ItemData        =   "frmAccounts.frx":1CF6
         Left            =   180
         List            =   "frmAccounts.frx":1D09
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3195
         Width           =   5730
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A C C O U N T S"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   1425
         TabIndex        =   46
         Top             =   825
         Width           =   3105
      End
      Begin VB.Label lblnav 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         Left            =   1215
         TabIndex        =   21
         Top             =   375
         Width           =   225
      End
      Begin VB.Label lblnav 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         Left            =   225
         TabIndex        =   20
         Top             =   375
         Width           =   225
      End
      Begin VB.Label lblnav 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         Left            =   885
         TabIndex        =   19
         Top             =   375
         Width           =   225
      End
      Begin VB.Label lblnav 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         Left            =   555
         TabIndex        =   18
         Top             =   375
         Width           =   225
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   3276
         TabIndex        =   17
         Top             =   375
         Width           =   615
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   5
         Left            =   5520
         TabIndex        =   16
         Top             =   375
         Width           =   330
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   3
         Left            =   4119
         TabIndex        =   15
         Top             =   375
         Width           =   570
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   4
         Left            =   4917
         TabIndex        =   14
         Top             =   375
         Width           =   375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   2598
         TabIndex        =   13
         Top             =   375
         Width           =   450
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   1980
         TabIndex        =   12
         Top             =   375
         Width           =   390
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   0
         Picture         =   "frmAccounts.frx":1D3D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   6090
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AccountType"
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
         Left            =   3390
         TabIndex        =   10
         Top             =   1635
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Id"
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
         Left            =   195
         TabIndex        =   7
         Top             =   1635
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Title"
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
         Left            =   210
         TabIndex        =   6
         Top             =   2310
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account Head"
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
         Left            =   195
         TabIndex        =   5
         Top             =   2970
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A C C O U N T S"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Left            =   1485
         TabIndex        =   47
         Top             =   855
         Width           =   3105
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   1  'Opaque
         BorderWidth     =   2
         Height          =   240
         Left            =   -30
         Top             =   15
         Width           =   6150
      End
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00CC9933&
      BackStyle       =   1  'Opaque
      Height          =   3015
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6090
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ProdSalePrice As Single
Dim ProdReOrderPoint As Single
Dim SQLQry As String
Dim AvgPrice As Single

Private Sub cmbType_LostFocus()
    If cmbType.Text = "Product" Then
        cmdProdInfo.Enabled = True
    Else
        cmdProdInfo.Enabled = False
    End If

    
End Sub

Private Sub cmdCancel_Click()
    Clear Me
'Calling form load event
    Call Form_Load

'Changing the mode of button
    Modes False, True, Me

'Setting focus on TxtName
    cmbType.SetFocus

  
'Highlighting TxtName
    High TxtName

'Unlock Navigation
    UnLockNav Me

End Sub

Private Sub cmdDelete_Click()
    MsgBox "In process"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    If optTitle.Value = True Then
        If Trim(txtFindTitle.Text) = "" Then
            MsgBox "Enter Account Title to find", vbCritical, "Message ..."
            txtFindTitle.SetFocus
            Exit Sub
        Else
            FindRecord "Title"
        End If
    ElseIf optId.Value = True Then
        If Trim(txtFindId.Text) = "" Or Val(txtFindId) = 0 Then
            MsgBox "Enter Account ID to find", vbCritical, "Message ..."
            txtFindId.SetFocus
            Exit Sub
        Else
            FindRecord "Id"
        End If
    End If
    
    
    
End Sub

Private Sub cmdFindCancel_Click()
    picFind.Visible = False
    cmbType.SetFocus
End Sub

Private Sub cmdFindRecord_Click()
    picFind.Visible = True
        optTitle.Value = True
            txtFindTitle.SetFocus
            
End Sub

Private Sub cmdOk_Click()
    
    If txtSalePrice.Text = "" Or Val(txtSalePrice) = 0 Then
        MsgBox "Enter Sale Price for this Product", vbCritical, "Message"
        txtSalePrice.SetFocus
        Exit Sub
    End If
    
    
    picAccounts.Enabled = True
    picProdInfo.Visible = False
    
    TxtName.Text = txtProdName.Text
    cmbHead.Text = "Assets"
    TxtName.SetFocus
End Sub

Private Sub cmdProdCancel_Click()
    picAccounts.Enabled = True
    picProdInfo.Visible = False
    
End Sub

Private Sub cmdProdInfo_Click()
    picProdInfo.Visible = True
    txtProdName.SetFocus
    
    picAccounts.Enabled = False
    
End Sub

Private Sub cmdSave_Click()
'If Name field is blank
    If TxtName.Text = "" Then
        MsgBox "Enter Account Title", vbCritical, "Message"
        TxtName.SetFocus
        Exit Sub
    End If

'If account head not selected then
    If cmbHead.ListIndex = -1 Then
        MsgBox "Select Account Head", vbCritical, "Message"
        cmbHead.SetFocus
        Exit Sub
    End If
    
'If account Type not selected then
    If cmbType.ListIndex = -1 Then
        MsgBox "Select Account Type", vbCritical, "Message"
        cmbType.SetFocus
        Exit Sub
    End If

'If Account Type = Product (as we have other option to enter product detail)
    If cmbType.Text = "Product" Then
        If txtProdName.Text = "" Or txtSalePrice.Text = "" Or txtReOrder.Text = "" Then
            MsgBox "Please enter product informations", vbCritical, "Message"
            cmdProdInfo_Click
            Exit Sub
        End If
    End If

'Finding Head Account ID through CMBHead.Text
    If cmbHead.Text = "Assets" Then
        mHeadId = 1
    ElseIf cmbHead.Text = "Liabilities" Then
        mHeadId = 2
    ElseIf cmbHead.Text = "Capital" Then
        mHeadId = 3
    ElseIf cmbHead.Text = "Revenue" Then
        mHeadId = 4
    ElseIf cmbHead.Text = "Expense" Then
        mHeadId = 5
    End If

'Check if Account Type is Product or anyother
    If Not cmbType.Text = "Product" Then
        txtSalePrice.Text = "0"
        txtReOrder.Text = "0"
    End If

'If New record
    If cmdNew.Enabled = False Then
        
        Con.Execute "Insert into Accounts(AcId,HeadId,AcTitle,AcType,SalePrice,ReOrderPoint) Values(" & Val(txtId) & ", " & Val(mHeadId) & ", '" & TxtName.Text & "', '" & cmbType.Text & "',  " & Val(txtSalePrice) & ", " & Val(txtReOrder) & " )"
        
        If cmbType.Text = "Product" Then
            Con.Execute "Insert into Product_Openings(ProdId, Qty, Amount) values(" & Val(txtId) & ", " & Val(txtOpStock) & ", " & Val(txtOpAmount) & ")"
        End If
        
        MsgBox "New Account Created", vbInformation, "Done"
           
'Updating Maximum Number
        UpdateMaxNumber "AcId", Val(txtId)
    
'If Existing Record
    ElseIf cmdNew.Enabled = True Then
        If UpdateRecord.State = 1 Then UpdateRecord.Close
        
        Set UpdateRecord = New ADODB.Recordset
        UpdateRecord.Open "Update Accounts set HeadId = " & Val(mHeadId) & ", AcTitle = '" & TxtName.Text & "', AcType = '" & cmbType & "' , SalePrice = " & Val(txtSalePrice) & ", ReOrderPoint = " & Val(txtReOrder) & " where AcId = " & Val(txtId) & " ", Con, adOpenDynamic, adLockOptimistic
        
        If cmbType.Text = "Product" Then
            Con.Execute "Update Product_Openings set Qty = " & Val(txtOpStock) & ", Amount = " & Val(txtOpAmount) & " Where ProdId = " & Val(txtId) & ""
        End If
        
        MsgBox "Existing Record Updated", vbInformation, "Done"
    End If
    
    cmdCancel_Click
End Sub




Private Sub Command1_Click()

End Sub


Private Sub fgAccounts_DblClick()
Dim mRow As Integer
    mRow = fgAccounts.RowSel
        
    txtId.Text = fgAccounts.TextMatrix(mRow, 0)
    TxtName.Text = fgAccounts.TextMatrix(mRow, 1)
    cmbType.Text = fgAccounts.TextMatrix(mRow, 2)
    cmbHead.Text = fgAccounts.TextMatrix(mRow, 3)
    
    cmbType.SetFocus
   
    If cmbType.Text = "Product" Then
        Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open "Select AcTitle,SalePrice,ReOrderPoint from Accounts where AcId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
                    txtProdName.Text = RS(0)
                    txtSalePrice.Text = Val(RS(1))
                    txtReOrder.Text = Val(RS(2))
            
        Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open "Select Qty,Amount from Product_Openings where ProdId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
                    txtOpStock.Text = Val(RS(0))
                    txtOpAmount.Text = Val(RS(1))
                    
        cmdProdInfo.Enabled = True
    Else
        cmdProdInfo.Enabled = False
        
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'Changing Control focus on Enter
    ChangeFocusOnEnter KeyAscii, Me
End Sub

Private Sub Form_Load()
'Setting up flexgrid data
    Call fgAccountsData
    Call GridSetting
    
'Setting Navigational Recordset
    If RsNAV.State = 1 Then RsNAV.Close
        
        Set RsNAV = New ADODB.Recordset
        RsNAV.Open "Select AcId, HeadId, AcTitle, AcType, SalePrice, ReOrderPoint from Accounts", Con, adOpenStatic, adLockOptimistic
            If RsNAV.RecordCount > 0 Then
                Call BoundData  'Showing Data in textboxes
            Else
                Exit Sub
            End If
            
    cmdProdInfo.Enabled = False
    
End Sub

Public Sub fgAccountsData()
    If RS.State = 1 Then RS.Close
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from ViewHeadWise order by AcId", Con, adOpenStatic, adLockOptimistic

    Set fgAccounts.DataSource = RS
    
End Sub

Public Sub GridSetting()
    With fgAccounts
        .ColWidth(0) = 950
        .ColWidth(1) = 2000
        .ColWidth(2) = 1350
        .ColWidth(3) = 1350
        
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Account Title"
        .TextMatrix(0, 2) = "Account Type"
        .TextMatrix(0, 3) = "Head Account"
    
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        
        .ColAlignment(0) = 5
        
        .RowHeight(0) = 400
    End With
End Sub

Public Sub BoundData()
    txtId.Text = Val(RsNAV(0))
    If Val(RsNAV(1)) = 1 Then
        cmbHead.Text = "Assets"
    ElseIf Val(RsNAV(1)) = 2 Then
        cmbHead.Text = "Liabilities"
    ElseIf Val(RsNAV(1)) = 3 Then
        cmbHead.Text = "Capital"
    ElseIf Val(RsNAV(1)) = 4 Then
        cmbHead.Text = "Revenue"
    ElseIf Val(RsNAV(1)) = 5 Then
        cmbHead.Text = "Expense"
    End If
    
    TxtName.Text = RsNAV(2)
    cmbType.Text = RsNAV(3)
    
    If cmbType.Text = "Product" Then
        txtProdName.Text = RsNAV(2)
        txtSalePrice.Text = Val(RsNAV(4))
        txtReOrder.Text = Val(RsNAV(5))
    End If
    
End Sub

Private Sub Form_Resize()
    Me.Left = Me.Left + 1200
    Me.Top = Me.Top + 200
End Sub

Private Sub Label19_Click()
    Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseNormalOnLbl
    
End Sub

Private Sub lbl_Click(Index As Integer)
    Select Case Index
        Case 0      'NEW
            'Clearing controls of the form
                Clear Me
            
            'Calling Max Number
                Call AutoId
            
            'Changing the mode of button
                Modes True, False, Me
            
            'Setting focus on TxtName
                cmbType.SetFocus
            
            'Locking Navigation
                LockNav Me
    
        Case 1      'Save
        
            'If Name field is blank
                If TxtName.Text = "" Then
                    MsgBox "Enter Account Title", vbCritical, "Message"
                    TxtName.SetFocus
                    Exit Sub
                End If
            
            'If account head not selected then
                If cmbHead.ListIndex = -1 Then
                    MsgBox "Select Account Head", vbCritical, "Message"
                    cmbHead.SetFocus
                    Exit Sub
                End If
                
            'If account Type not selected then
                If cmbType.ListIndex = -1 Then
                    MsgBox "Select Account Type", vbCritical, "Message"
                    cmbType.SetFocus
                    Exit Sub
                End If
            
            'If Account Type = Product (as we have other option to enter product detail)
                If cmbType.Text = "Product" Then
                    If txtProdName.Text = "" Or txtSalePrice.Text = "" Or txtReOrder.Text = "" Then
                        MsgBox "Please enter product informations", vbCritical, "Message"
                        cmdProdInfo_Click
                        Exit Sub
                    End If
                End If
            
            'Finding Head Account ID through CMBHead.Text
                If cmbHead.Text = "Assets" Then
                    mHeadId = 1
                ElseIf cmbHead.Text = "Liabilities" Then
                    mHeadId = 2
                ElseIf cmbHead.Text = "Capital" Then
                    mHeadId = 3
                ElseIf cmbHead.Text = "Revenue" Then
                    mHeadId = 4
                ElseIf cmbHead.Text = "Expense" Then
                    mHeadId = 5
                End If
            
            'Check if Account Type is Product or anyother
                If Not cmbType.Text = "Product" Then
                    txtSalePrice.Text = "0"
                    txtReOrder.Text = "0"
                End If
            
            'If New record
                If lbl(0).Enabled = False Then
                    
                    Con.Execute "Insert into Accounts(AcId,HeadId,AcTitle,AcType,SalePrice,ReOrderPoint) Values(" & Val(txtId) & ", " & Val(mHeadId) & ", '" & TxtName.Text & "', '" & cmbType.Text & "',  " & Val(txtSalePrice) & ", " & Val(txtReOrder) & " )"
                    
                    If cmbType.Text = "Product" Then
                        Con.Execute "Insert into Product_Openings(ProdId, Qty, Amount) values(" & Val(txtId) & ", " & Val(txtOpStock) & ", " & Val(txtOpAmount) & ")"
                                            
                    'Average Cost Price / Sale Price
                        AvgPrice = Val(txtOpAmount) / Val(txtOpStock)
                        Con.Execute "Insert into Avg_Cost(ProdId, PurAvg, SalAvg) Values(" & Val(txtId) & ", " & Val(AvgPrice) & ", " & Val(txtSalePrice) & ")"
                            
                    End If
                        
                    
                    MsgBox "New Account Created", vbInformation, "Done"
                       
            'Updating Maximum Number
                    UpdateMaxNumber "AcId", Val(txtId)
                
            'If Existing Record
                ElseIf lbl(0).Enabled = True Then
                    If UpdateRecord.State = 1 Then UpdateRecord.Close
                    
                    Set UpdateRecord = New ADODB.Recordset
                    UpdateRecord.Open "Update Accounts set HeadId = " & Val(mHeadId) & ", AcTitle = '" & TxtName.Text & "', AcType = '" & cmbType & "' , SalePrice = " & Val(txtSalePrice) & ", ReOrderPoint = " & Val(txtReOrder) & " where AcId = " & Val(txtId) & " ", Con, adOpenDynamic, adLockOptimistic
                    
                    If cmbType.Text = "Product" Then
                        Con.Execute "Update Product_Openings set Qty = " & Val(txtOpStock) & ", Amount = " & Val(txtOpAmount) & " Where ProdId = " & Val(txtId) & ""
                    'Average Cost Price
                        AvgPrice = Val(txtOpAmount) / Val(txtOpStock)
                        
                        Con.Execute "Update Avg_Cost SET PurAvg = " & Val(AvgPrice) & " where ProdID = " & Val(txtId) & ""
                    
                    End If
                    
                    MsgBox "Existing Record Updated", vbInformation, "Done"
                End If
                
                cmdCancel_Click
        Case 2                  'Cancel
                Clear Me
            'Calling form load event
                Call Form_Load
            
            'Changing the mode of button
                Modes False, True, Me
            
            'Setting focus on TxtName
                cmbType.SetFocus
              
            'Highlighting TxtName
                High TxtName
            
            'Unlock Navigation
                UnLockNav Me
            
        Case 3
                MsgBox "In process"

        Case 4
                picFind.Visible = True
                    optTitle.Value = True
                        txtFindTitle.SetFocus
        
        Case 5
                Unload Me
    End Select

End Sub

Private Sub lbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0, 1, 2, 3, 4, 5
            MouseOver lbl(Index)
            
    End Select

End Sub

Private Sub lblnav_Click(Index As Integer)
    Select Case Index
        Case 0              'Move First
            RsNAV.MoveFirst
            
            If RsNAV.BOF = True Then
                MsgBox "First Record", vbInformation, "Message"
                RsNAV.MoveFirst
            Else
                Call BoundData
            End If
        
        Case 1              'Move previous
            RsNAV.MovePrevious
            
            If RsNAV.BOF = True Then
                MsgBox "First Record", vbInformation, "Message"
                RsNAV.MoveFirst
            Else
                Call BoundData
            End If
    
        Case 2              'Move Next
            RsNAV.MoveNext
        
         
            If RsNAV.EOF = True Then
                MsgBox "Last Record", vbInformation, "Message"
                RsNAV.MoveLast
            Else
                Call BoundData
            End If
            
        Case 3
            RsNAV.MoveLast
            
            If RsNAV.EOF = True Then
                MsgBox "Last Record", vbInformation, "Message"
                RsNAV.MoveLast
            Else
                Call BoundData
            End If
        End Select
        
End Sub

Private Sub lblnav_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0, 1, 2, 3
            MouseOver lblnav(Index)
    End Select
    
End Sub

Private Sub picAccounts_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   MouseNormalOnLbl
End Sub

Public Sub AutoId()
'Calling MaxNumber function to get Auto Id for the record
    MaxNumber "AcId", "Max_Codes"
    txtId.Text = Val(MaxNmbr)
End Sub

Private Sub Picture4_Click()

End Sub

Private Sub txtFindId_KeyPress(KeyAscii As Integer)
    ONU KeyAscii, txtFindId
End Sub

Private Sub txtId_GotFocus()
    TxtName.SetFocus
End Sub

Private Sub txtOpAmount_KeyPress(KeyAscii As Integer)
    ONU KeyAscii, txtOpAmount
End Sub

Private Sub txtOpAmount_Validate(Cancel As Boolean)
    If txtOpAmount.Text = "" Then
        txtOpAmount.Text = "0"
    End If
End Sub

Private Sub txtOpStock_KeyPress(KeyAscii As Integer)
    ONU KeyAscii, txtOpStock
End Sub

Private Sub txtOpStock_Validate(Cancel As Boolean)
    If txtOpStock.Text = "" Then
        txtOpStock.Text = "0"
    End If
End Sub

Private Sub txtReOrder_KeyPress(KeyAscii As Integer)
    ONU KeyAscii, txtReOrder
End Sub

Private Sub txtReOrder_Validate(Cancel As Boolean)
    If txtReOrder.Text = "" Then
        txtReOrder.Text = "0"
    End If
End Sub

Private Sub txtSalePrice_KeyPress(KeyAscii As Integer)
    ONU KeyAscii, txtSalePrice
End Sub

Private Sub txtSalePrice_Validate(Cancel As Boolean)
    If txtSalePrice.Text = "" Or Val(txtSalePrice) = 0 Then
        MsgBox "Enter Sale Price for this Product", vbCritical, "Message"
        txtSalePrice.SetFocus
    End If
End Sub

Public Sub FindRecord(MatchWith As String)
    Select Case MatchWith
        Case "Title"
            SQLQry = "SELECT Accounts.AcId,Accounts.HeadId, Accounts.AcTitle, Accounts.AcType,  Accounts.SalePrice, Accounts.ReOrderPoint " & _
            "FROM Account_Heads INNER JOIN Accounts ON Account_Heads.Id = Accounts.HeadId where AcTitle = '" & txtFindTitle.Text & "'"
            
            Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open SQLQry, Con, adOpenStatic, adLockOptimistic
                    If RS.RecordCount <= 0 Then
                        MsgBox "No account found with this Title", vbCritical, "Message.."
                        txtFindTitle.SetFocus
                        Exit Sub
                    Else
                        ShowFoundData
                    End If
                    
                    
        Case "Id"
            SQLQry = "SELECT Accounts.AcId,Accounts.HeadId, Accounts.AcTitle, Accounts.AcType,  Accounts.SalePrice, Accounts.ReOrderPoint " & _
            "FROM Account_Heads INNER JOIN Accounts ON Account_Heads.Id = Accounts.HeadId where AcId = " & Val(txtFindId) & ""
            
            Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open SQLQry, Con, adOpenStatic, adLockOptimistic
                    If RS.RecordCount <= 0 Then
                        MsgBox "No account found with this ID", vbCritical, "Message.."
                        txtFindId.SetFocus
                        Exit Sub
                    Else
                        ShowFoundData
                    End If
    
    End Select
End Sub

Public Sub ShowFoundData()
    txtId.Text = Val(RS(0))
    If Val(RS(1)) = 1 Then
        cmbHead.Text = "Assets"
    ElseIf Val(RS(1)) = 2 Then
        cmbHead.Text = "Liabilities"
    ElseIf Val(RS(1)) = 3 Then
        cmbHead.Text = "Capital"
    ElseIf Val(RS(1)) = 4 Then
        cmbHead.Text = "Revenue"
    ElseIf Val(RS(1)) = 5 Then
        cmbHead.Text = "Expense"
    End If
    
    TxtName.Text = RS(2)
    cmbType.Text = RS(3)
    
    If cmbType.Text = "Product" Then
        txtProdName.Text = RS(2)
        txtSalePrice.Text = Val(RS(4))
        txtReOrder.Text = Val(RS(5))
    End If

    If cmbType.Text = "Product" Then
        Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open "Select AcTitle,SalePrice,ReOrderPoint from Accounts where AcId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
                    txtProdName.Text = RS(0)
                    txtSalePrice.Text = Val(RS(1))
                    txtReOrder.Text = Val(RS(2))
            
        Set RS = New ADODB.Recordset
            If RS.State = 1 Then RS.Close
                RS.Open "Select Qty,Amount from Product_Openings where ProdId = " & Val(txtId) & "", Con, adOpenStatic, adLockOptimistic
                    txtOpStock.Text = Val(RS(0))
                    txtOpAmount.Text = Val(RS(1))
                    
        cmdProdInfo.Enabled = True
    Else
        cmdProdInfo.Enabled = False
        
    End If
    
    
    picFind.Visible = False
    cmbType.SetFocus
End Sub

Public Sub MouseNormalOnLbl()
'Commands Lable
    lbl(0).ForeColor = vbWhite
    lbl(0).Font.Underline = False
    
    lbl(1).ForeColor = vbWhite
    lbl(1).Font.Underline = False
    
    lbl(2).ForeColor = vbWhite
    lbl(2).Font.Underline = False
    
    lbl(3).ForeColor = vbWhite
    lbl(3).Font.Underline = False
    
    lbl(4).ForeColor = vbWhite
    lbl(4).Font.Underline = False
    
    lbl(5).ForeColor = vbWhite
    lbl(5).Font.Underline = False
    
'Navigation Label
    lblnav(0).ForeColor = vbWhite
    lblnav(0).Font.Underline = False

    lblnav(1).ForeColor = vbWhite
    lblnav(1).Font.Underline = False
    
    lblnav(2).ForeColor = vbWhite
    lblnav(2).Font.Underline = False
    
    lblnav(3).ForeColor = vbWhite
    lblnav(3).Font.Underline = False

End Sub


