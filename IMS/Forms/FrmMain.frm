VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00F0F0F0&
   Caption         =   "Inventory Management System"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "FrmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmMain.frx":000C
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0F0F0&
      ForeColor       =   &H80000008&
      Height          =   11700
      Left            =   0
      ScaleHeight     =   11670
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   4125
         Index           =   4
         Left            =   60
         ScaleHeight     =   4095
         ScaleWidth      =   2535
         TabIndex        =   20
         Top             =   6570
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   10
            Left            =   450
            Picture         =   "FrmMain.frx":10E3F
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   34
            Top             =   3645
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Exit to windows"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   7
               Left            =   465
               TabIndex        =   35
               Top             =   60
               Visible         =   0   'False
               Width           =   1305
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            FillStyle       =   0  'Solid
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   4
            Left            =   450
            Picture         =   "FrmMain.frx":123C0
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   32
            Top             =   375
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Backup"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   10
               Left            =   500
               TabIndex        =   33
               Top             =   45
               Visible         =   0   'False
               Width           =   630
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   5
            Left            =   450
            Picture         =   "FrmMain.frx":13941
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   30
            Top             =   900
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Restore"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   11
               Left            =   500
               TabIndex        =   31
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   6
            Left            =   450
            Picture         =   "FrmMain.frx":14EC2
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   28
            Top             =   1425
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Logoff"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   2
               Left            =   500
               TabIndex        =   29
               Top             =   60
               Visible         =   0   'False
               Width           =   540
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   9
            Left            =   450
            Picture         =   "FrmMain.frx":16443
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   26
            Top             =   2985
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "MS Paint"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   6
               Left            =   495
               TabIndex        =   27
               Top             =   60
               Visible         =   0   'False
               Width           =   705
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   7
            Left            =   450
            Picture         =   "FrmMain.frx":179C4
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   23
            Top             =   1935
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Calculator"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   3
               Left            =   495
               TabIndex        =   25
               Top             =   285
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Calculator"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   12
               Left            =   465
               TabIndex        =   24
               Top             =   60
               Visible         =   0   'False
               Width           =   825
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   8
            Left            =   450
            Picture         =   "FrmMain.frx":18F45
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   21
            Top             =   2460
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Notepad"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   4
               Left            =   495
               TabIndex        =   22
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Accessories"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   465
            TabIndex        =   36
            Top             =   0
            Width           =   1620
         End
         Begin VB.Image ImgRestore 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   900
            Width           =   450
         End
         Begin VB.Image imgLogoff 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   1425
            Width           =   450
         End
         Begin VB.Image imgCalc 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   1935
            Width           =   450
         End
         Begin VB.Image imgNotepad 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   2460
            Width           =   450
         End
         Begin VB.Image imgPaint 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   2985
            Width           =   450
         End
         Begin VB.Image imgExit 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   3645
            Width           =   450
         End
         Begin VB.Image imgBackUp 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   375
            Width           =   450
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   0
            X2              =   2520
            Y1              =   3495
            Y2              =   3510
         End
         Begin VB.Image Image1 
            Height          =   315
            Index           =   5
            Left            =   -15
            Picture         =   "FrmMain.frx":1A4C6
            Stretch         =   -1  'True
            Top             =   -75
            Width           =   2565
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   3
         Left            =   60
         ScaleHeight     =   1290
         ScaleWidth      =   2535
         TabIndex        =   16
         Top             =   5055
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   3
            Left            =   450
            Picture         =   "FrmMain.frx":1B1F3
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   17
            Top             =   570
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Reports"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   8
               Left            =   500
               TabIndex        =   18
               Top             =   45
               Visible         =   0   'False
               Width           =   675
            End
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reports"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   930
            TabIndex        =   19
            Top             =   0
            Width           =   690
         End
         Begin VB.Image imgReports 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   570
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   315
            Index           =   4
            Left            =   -15
            Picture         =   "FrmMain.frx":1C774
            Stretch         =   -1  'True
            Top             =   -75
            Width           =   2565
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   2
         Left            =   60
         ScaleHeight     =   1290
         ScaleWidth      =   2535
         TabIndex        =   12
         Top             =   3540
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   2
            Left            =   450
            Picture         =   "FrmMain.frx":1D4A1
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   13
            Top             =   600
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Journal Voucher"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   5
               Left            =   500
               TabIndex        =   14
               Top             =   45
               Visible         =   0   'False
               Width           =   1425
            End
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Transaction Book"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   525
            TabIndex        =   15
            Top             =   0
            Width           =   1500
         End
         Begin VB.Image imgVoucher 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   600
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   315
            Index           =   3
            Left            =   -15
            Picture         =   "FrmMain.frx":1EA22
            Stretch         =   -1  'True
            Top             =   -75
            Width           =   2565
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   1
         Left            =   60
         ScaleHeight     =   1290
         ScaleWidth      =   2535
         TabIndex        =   8
         Top             =   2010
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   11
            Left            =   450
            Picture         =   "FrmMain.frx":1F74F
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   9
            Top             =   555
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Accounts"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   9
               Left            =   240
               TabIndex        =   10
               Top             =   45
               Visible         =   0   'False
               Width           =   780
            End
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account Books"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   630
            TabIndex        =   11
            Top             =   0
            Width           =   1290
         End
         Begin VB.Image imgAccount 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   555
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   315
            Index           =   2
            Left            =   -15
            Picture         =   "FrmMain.frx":20CD0
            Stretch         =   -1  'True
            Top             =   -75
            Width           =   2565
         End
      End
      Begin VB.PictureBox PicMnu 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   1320
         Index           =   0
         Left            =   60
         ScaleHeight     =   1290
         ScaleWidth      =   2535
         TabIndex        =   2
         Top             =   510
         Width           =   2565
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   450
            Picture         =   "FrmMain.frx":219FD
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   5
            Top             =   855
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Sales"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   210
               Index           =   1
               Left            =   500
               TabIndex        =   6
               Top             =   45
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin VB.PictureBox PicBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   450
            Picture         =   "FrmMain.frx":22F7E
            ScaleHeight     =   315
            ScaleWidth      =   2025
            TabIndex        =   3
            Top             =   360
            Width           =   2055
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00ECE7E3&
               BackStyle       =   0  'Transparent
               Caption         =   "Purchases"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   210
               Index           =   0
               Left            =   500
               TabIndex        =   4
               Top             =   45
               Visible         =   0   'False
               Width           =   915
            End
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase / Sale Books"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   315
            TabIndex        =   7
            Top             =   -15
            Width           =   1920
         End
         Begin VB.Image imgSale 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   855
            Width           =   450
         End
         Begin VB.Image imgPurchase 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   345
            Left            =   30
            Stretch         =   -1  'True
            Top             =   360
            Width           =   450
         End
         Begin VB.Image Image1 
            Height          =   315
            Index           =   1
            Left            =   -15
            Picture         =   "FrmMain.frx":8AEA6
            Stretch         =   -1  'True
            Top             =   -75
            Width           =   2565
         End
      End
      Begin VB.Image Image1 
         Height          =   360
         Index           =   7
         Left            =   -15
         Picture         =   "FrmMain.frx":8BBD3
         Stretch         =   -1  'True
         Top             =   11175
         Width           =   2730
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "APPLICATION MENU"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F0F0F0&
         Height          =   270
         Left            =   240
         TabIndex        =   1
         Top             =   0
         Width           =   2220
      End
      Begin VB.Image Image1 
         Height          =   330
         Index           =   6
         Left            =   -1215
         Picture         =   "FrmMain.frx":8ED0A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3945
      End
   End
   Begin MSComctlLib.ImageList img2 
      Left            =   5265
      Top             =   8505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":91E41
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":92853
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":93265
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":935FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":93999
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":93D33
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":940CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":94ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":954F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":95F03
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":96915
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":97327
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":97D39
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9874B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":98CE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":99283
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img1 
      Left            =   4545
      Top             =   8505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   50
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9ADD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9C767
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9D443
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":9EDD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A0767
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A20F9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A3A8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A4765
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A543F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A5D1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A69F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A76D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A7FB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A8C93
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A956F
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AA24B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":ABBDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AD573
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":ADE4F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AE729
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AF003
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":AF8DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B01B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B0751
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B0A6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B0D85
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B165F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B1F39
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B2813
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B34ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B3807
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B3C59
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B4AAB
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B4EFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B57D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":B60B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BB8A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BC17D
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BC497
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BD171
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BDA4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BE325
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BEBFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BF4D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":BFDB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C068D
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C0F67
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":CF84D
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":D1A11
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":DC5D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label LblHelpLine 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This is your Help Line"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   2715
      TabIndex        =   37
      Top             =   11235
      Width           =   12600
   End
   Begin VB.Image Image1 
      Height          =   360
      Index           =   0
      Left            =   2475
      Picture         =   "FrmMain.frx":DCDC7
      Stretch         =   -1  'True
      Top             =   11190
      Width           =   13080
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   11280
      Left            =   15
      Top             =   90
      Width           =   2760
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VIndex As Integer
Dim SideMenuVisible As Boolean 'On click if frame(Side menu) is visible the invisibole it or if invisible then visible it



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
                
            PicBar_Click (10)
    End If


    If KeyCode = vbKeyA And Shift = 2 Then
        PicBar_Click (11)                           'Accounts
    End If
    

End Sub

Private Sub Form_Load()
    Dim i%
        For i% = 0 To PicBar.Count - 1
            PicBar(i%).Picture = LoadResPicture("Light", 0)
        Next

        LblHelpLine.Caption = ""
    SetPicCaption
    
    imgPurchase.Picture = img2.ListImages(11).Picture
    imgSale.Picture = img1.ListImages(9).Picture
    imgAccount.Picture = img1.ListImages(14).Picture
    imgVoucher.Picture = img1.ListImages(11).Picture
    imgReports.Picture = img1.ListImages(35).Picture
    imgBackUp.Picture = img1.ListImages(41).Picture
    ImgRestore.Picture = img1.ListImages(25).Picture
    imgLogoff.Picture = img1.ListImages(18).Picture
    imgCalc.Picture = img1.ListImages(47).Picture
    imgNotepad.Picture = img1.ListImages(48).Picture
    imgPaint.Picture = img1.ListImages(49).Picture
    imgExit.Picture = img1.ListImages(50).Picture


End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetColorToDefault
End Sub

Private Sub PicBar_Click(Index As Integer)

    
'Getting CLICK Action
    Select Case Index
        Case 0                                                          'Purchases
            frmPurchase.Show 1
            
        Case 1                                                          'Sales
            frmSale.Show 1
        
        Case 2                                                          'Journal Vouchers / Transaction
            frmTransaction.Show 1
            
        Case 3                                                          'Reports
            FrmRep.Show 1
        
        Case 4                                                          'BackUp
                                                    
        Case 5                                                          'Restore
        
        Case 6                                                          'Log off
            FrmMain.Enabled = False
            frmLogin.Show
        
        Case 7                                                          'Calculator
            Shell "C:\windows\system32\calc.exe", vbNormalFocus
        
        Case 8                                                          'Notepad
            Shell "C:\windows\system32\notepad.exe", vbNormalFocus
        
        Case 9                                                          'MS Paint
            Shell "C:\windows\system32\mspaint.exe", vbNormalFocus
        
        Case 10                                                         'Exit
            If MsgBox("Do you want to exit?", vbQuestion + vbYesNo, "Exit?") = vbYes Then
                Unload Me
            End If
            
       Case 11
            frmAccounts.Show 1
        
    End Select
    

    
End Sub

Private Sub PicBar_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    VIndex = Index
   
    Select Case Index
'---------------------------- MAIN MENU
        Case 0
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Purchases
            LblHelpLine.Caption = "Add your Purchases"
        Case 1
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Sales
            LblHelpLine.Caption = "Add your Sales"
        Case 2
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Journal Vouchers
            LblHelpLine.Caption = "Add your Transactions"
        Case 3
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Reports
            LblHelpLine.Caption = "View Reports"
        Case 4
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Back Up
            LblHelpLine.Caption = "Create Database Backup"
        Case 5
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Restore
            LblHelpLine.Caption = "Restore Database"
        Case 6
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Logoff
            LblHelpLine.Caption = "Logoff the System"
        Case 7
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Calculator
            LblHelpLine.Caption = "Calculator"
        Case 8
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Notepad
            LblHelpLine.Caption = "Note pad"
        Case 9
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'MS Paint
            LblHelpLine.Caption = "MS Paint"
        Case 10
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Exit
            LblHelpLine.Caption = "Exit to Windows"
   
        Case 11
            PicBar(Index).Picture = LoadResPicture("Down", 0)       'Accounts
            LblHelpLine.Caption = "Add Transactional Accounts"
    
    End Select

    SetPicCaption

End Sub

Public Sub SetColorToDefault()
    PicBar(VIndex).Picture = LoadResPicture("Light", 0)
    LblHelpLine.Caption = ""
    SetPicCaption
End Sub

Private Sub PicMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    LblHelpLine.Caption = ""

    SetPicCaption

End Sub

Private Sub PicMnu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim i%
        For i% = 0 To PicBar.Count - 1
            PicBar(i%).Picture = LoadResPicture("Light", 0)
        Next

        LblHelpLine.Caption = ""
    SetPicCaption
End Sub


Public Sub SetPicCaption()
    Call ReplaceCap
End Sub



Public Sub ReplaceCap()
    Dim a As Integer
        For a = 0 To 11
            Select Case a
                Case 0
                    PrintToCenter "Purchases", PicBar(a), False
                Case 1
                    PrintToCenter "Sales", PicBar(a), False
                Case 2
                    PrintToCenter "Journal Voucher", PicBar(a), False
                Case 3
                    PrintToCenter "Reports", PicBar(a), False
                Case 4
                    PrintToCenter "Backup", PicBar(a), False
                Case 5
                    PrintToCenter "Restore", PicBar(a), False
                Case 6
                    PrintToCenter "Log Off", PicBar(a), False
                Case 7
                    PrintToCenter "Calculator", PicBar(a), False
                Case 8
                    PrintToCenter "Notepad", PicBar(a), False
                Case 9
                    PrintToCenter "MS Paint", PicBar(a), False
                Case 10
                    PrintToCenter "Exit to Windows", PicBar(a), False
                Case 11
                    PrintToCenter "Accounts", PicBar(a), False
                    
            End Select
        Next
End Sub

