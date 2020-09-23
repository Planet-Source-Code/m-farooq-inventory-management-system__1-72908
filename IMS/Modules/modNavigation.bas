Attribute VB_Name = "modNavigation"
'Locking the Navigation Button on New Record
Public Function LockNav(Frm As Object)
        Frm.lblnav(0).Enabled = False
        Frm.lblnav(1).Enabled = False
        Frm.lblnav(2).Enabled = False
        Frm.lblnav(3).Enabled = False
End Function

'UnLocking the Navigation Button on Cancel
Public Function UnLockNav(Frm As Object)
        Frm.lblnav(0).Enabled = True
        Frm.lblnav(1).Enabled = True
        Frm.lblnav(2).Enabled = True
        Frm.lblnav(3).Enabled = True
End Function


