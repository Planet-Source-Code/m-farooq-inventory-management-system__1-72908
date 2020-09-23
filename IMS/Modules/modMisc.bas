Attribute VB_Name = "modMisc"
Public ctl As Control
Public RsMax As New ADODB.Recordset
Public MaxNmbr As Integer
Public mHeadId As Integer
Public TrialDr As Double
Public TrialCr As Double
Public TransactionType As String
Public TransactionRef As String
Public StrQry As String            'Variable Reporting
'Declaring SLEEP API for using as wait between work
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Switching On / Off buttons according to the mode
Public Function Modes(pNew As Boolean, pOther As Boolean, Frm As Object)
'If new record then
    If pNew = True Then
        Frm.lbl(0).Enabled = False
        Frm.lbl(3).Enabled = False
        Frm.lbl(4).Enabled = False
        Frm.lbl(5).Enabled = False
        
'If Not new record then
    ElseIf pOther = True Then
        Frm.lbl(0).Enabled = True
        Frm.lbl(3).Enabled = True
        Frm.lbl(4).Enabled = True
        Frm.lbl(5).Enabled = True
    End If
End Function

'On GotFocus Highlight the Text
Public Function High(txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt)
End Function

'Set focus to the next control usong Enter
Public Function Cng(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        KeyAscii = 0
    End If
    
End Function

'Clear all text Boxes and Combo Boxes
Public Function Clear(Frm As Form)
    For Each ctl In Frm.Controls
        
        If TypeOf ctl Is TextBox Then
            ctl.Text = ""
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.ListIndex = -1
        End If
    
    Next
        
End Function

'Only numbers in textbox
Public Function ONU(KeyAscii As Integer, txt As TextBox)
    If KeyAscii > 47 And KeyAscii < 58 Or KeyAscii = 8 Or KeyAscii = 46 Then
    
    
        If KeyAscii = 46 Then
    
    
            If InStr(txt.Text, ".") Then
                KeyAscii = 0
                Exit Function
            Else
                txt.Text = txt.Text
            End If
        Else
        End If
    Else
        KeyAscii = 0
    End If
End Function

'Getting Maximum number for Code field
Public Function MaxNumber(FieldName As String, TableName As String)
    Set RsMax = New ADODB.Recordset
    If RsMax.State = 1 Then RsMax.Close
    RsMax.Open "SELECT Max(" & FieldName & ") + 1 from " & TableName & " where MaxCode =1", Con, adOpenDynamic, adLockOptimistic
    
    If IsNull(RsMax(0)) Then
        MaxNumber = 1
    Else
        MaxNumber = RsMax(0)
    End If

    MaxNmbr = Val(MaxNumber)

End Function

'Updating Maximum Number
Public Function UpdateMaxNumber(FieldName As String, TBox As Integer)
    Con.Execute "Update Max_Codes Set " & FieldName & " = " & Val(TBox) & ""
End Function
Public Sub ChangeFocusOnEnter(KeyAscii As Integer, Frm As Object)
    If KeyAscii = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Public Function MouseOver(mLabel As Label)
    mLabel.ForeColor = vbYellow
    mLabel.Font.Underline = True
End Function

