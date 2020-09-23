Attribute VB_Name = "modMenu"
'********************** MENU COLOR CHANGE **********************

Public Sub OnColorChange(lblBack As Label, lblFore As Label) 'lblback = back color lable / lblfore = caption
    lblBack.BackColor = vbBlack
    lblFore.ForeColor = vbWhite
End Sub

Public Sub OutColorChange(lblBack As Label, lblFore As Label) 'lblback = back color lable / lblfore = caption
    lblBack.BackColor = &HF9F9F9
    lblFore.ForeColor = vbBlack
End Sub

