VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   9270
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   9270
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuActivities 
      Caption         =   "Activities"
      Begin VB.Menu mnuNewAccounts 
         Caption         =   "New Accounts"
         Shortcut        =   ^A
      End
      Begin VB.Menu MnuPurchases 
         Caption         =   "Purchases"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuSales 
         Caption         =   "Sales"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuTransaction 
         Caption         =   "Transactions"
         Shortcut        =   ^T
      End
      Begin VB.Menu hylast 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuExit_Click()
    If MsgBox("Exit to Windows ?", vbQuestion + vbYesNo, "Exit") = vbYes Then
        Unload Me
    End If
    

End Sub

Private Sub mnuNewAccounts_Click()
    frmAccounts.Show vbModal
End Sub

Private Sub mnuOpeningBalances_Click()
    frmAcOpening.Show vbModal
End Sub

Private Sub MnuPurchases_Click()
    frmPurchase.Show vbModal
End Sub

Private Sub mnuSales_Click()
    frmSale.Show vbModal
End Sub

Private Sub mnuTransaction_Click()
    frmTransaction.Show vbModal
End Sub
