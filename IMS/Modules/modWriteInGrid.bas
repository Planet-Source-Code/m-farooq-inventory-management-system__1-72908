Attribute VB_Name = "modWriteInGrid"

Public Function EditGrid(Msh As MSHFlexGrid, KeyAscii As Integer)
        
'==================================== Editing Grid Code For Transaction ====================================
        
'If First Row selected then
    Select Case Msh.Row
        Case 0
            KeyAscii = 0
            Exit Function
    End Select
    
'Block Every Key Except Enter in Column 1 (Account Name)

        Select Case Msh.Col

            Case 1
            
'If ENTER is pressed on Account Name Column then move to next column
            If KeyAscii = 13 Then
                SendKeys "{right}"
                Exit Function
            End If

                KeyAscii = 0
                Exit Function
        End Select

'---------------------------------------------------------
        
        Select Case KeyAscii
            
            Case 8: 'IF KEY IS BACKSPACE THEN
                If Msh.Text <> "" Then Msh.Text = Left$(Msh.Text, (Len(Msh.Text) - 1))
            
            Case 13: 'IF KEY IS ENTER THEN
                
                Select Case Msh.Col
                    
                    Case Is < 4
                        '----------Move Curssor to Right Side untill col >= 3------
                        Msh.SetFocus
                        SendKeys "{right}"
                    
                    Case 4
                        Msh.SetFocus
                        If (Msh.Row + 1) = Msh.Rows Then
                            ''-------- Null Value Chk in last col---------------
                            
'If values are missing in Debit and Credit Side
                            If Len(Msh.TextMatrix(Msh.Row, 1)) = 0 Or Msh.TextMatrix(Msh.Row, 1) = "" Then
                                 MsgBox "Select Account Name", vbInformation, "Message"
                                 Msh.Col = 1
                                 Exit Function
                            End If
                            
'If values are missing in Debit and Credit Side
                            If Val(Msh.TextMatrix(Msh.Row, 3)) = 0 And Val(Msh.TextMatrix(Msh.Row, 4)) = 0 Then
                                 MsgBox "Enter Debit / Credit Amount", vbInformation, "Message"
                                 Msh.Col = 3
                                 Exit Function
                            End If
                            
'If values are in both Debit and Credit Side
                            If Val(Msh.TextMatrix(Msh.Row, 3)) > 0 And Val(Msh.TextMatrix(Msh.Row, 4)) > 0 Then
                                 MsgBox "Enter Only Debit OR Credit Amount", vbInformation, "Message"
                                 Msh.Col = 3
                                 Exit Function
                            End If
                            
                            ''-------------New Row Create----------
                            Msh.Rows = Msh.Rows + 1
                            Msh.Col = 1
                            ''-------------------------------------
                            
'''                            ''-------------Adding Amount to Debit / credit Textbox----------
'''                            frmTransaction.txtCredit = Val(frmTransaction.txtCredit) + Val(Msh.TextMatrix(Msh.Row, 4))
'''                            frmTransaction.txtDebit = Val(frmTransaction.txtDebit) + Val(Msh.TextMatrix(Msh.Row, 3))
'''
'''                            ''-------------------------------------
                            
                            
                        End If
                        
                        SendKeys "{home}" + "{down}"   '' + "{right}"
                
                End Select
            
            Case Else 'KeyAscii Select
                ''-------write any code for Any Validation-------------
                Select Case Msh.Col
                    Case 3, 4
                        ''-------Allow Number Validation-------------
                        ONUGrid KeyAscii, Msh
                End Select
                ''-------------Write Data in Cells----------
                Msh.Text = Msh.Text + Chr$(KeyAscii)
                ''-------------------------------------
            End Select 'KeyAscii Select
End Function
'Only numbers in Grid
Public Function ONUGrid(KeyAscii As Integer, txt As MSHFlexGrid)
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

Public Function SetPSGRID(Msh As MSHFlexGrid)          'Purchase / Sale Grid
    
'Column Setting
    With Msh
        .Rows = 2
        .Cols = 7
        
               
        .ColWidth(0) = 400
        .ColWidth(1) = 5200         'Product description
        .ColWidth(2) = 1000         'Qty
        .ColWidth(3) = 1000         'Rate
        .ColWidth(4) = 1500         'Amount
        .ColWidth(5) = 0            'Product Code
        .ColWidth(6) = 0            'Avg Cost

'Column Captions
        
        .TextMatrix(0, 1) = "Product Description"
        .TextMatrix(0, 2) = "Quantity"
        .TextMatrix(0, 3) = "Price"
        .TextMatrix(0, 4) = "Amount"
        
    End With
End Function

Public Function EditGridPS(Msh As MSHFlexGrid, KeyAscii As Integer) ' Purchase / Sale Grid Edit
        
'==================================== Editing Grid Code For Purchase / Sale ====================================
        
'If First Row selected then
    Select Case Msh.Row
        Case 0
            KeyAscii = 0
            Exit Function
    End Select
    
'Block Every Key Except Enter in Column 1 (Product Name)

        Select Case Msh.Col

            Case 1

'If ENTER is pressed on Product Name Column then move to next column
            If KeyAscii = 13 Then
                SendKeys "{right}"
                Exit Function
            End If

                KeyAscii = 0
                Exit Function
        End Select

'---------------------------------------------------------
        
        Select Case KeyAscii
            
            Case 8: 'IF KEY IS BACKSPACE THEN
                If Msh.Text <> "" Then Msh.Text = Left$(Msh.Text, (Len(Msh.Text) - 1))
            
            Case 13: 'IF KEY IS ENTER THEN
                
                Select Case Msh.Col
                    
                    Case Is < 3
                        '----------Move Curssor to Right Side untill col >= 3------
                        Msh.SetFocus
                        SendKeys "{right}"
                    
                    Case 3
                        Msh.SetFocus
                        If (Msh.Row + 1) = Msh.Rows Then
                            ''-------- Null Value Chk in last col---------------
                            
'If Product Name is missing
                            If Len(Msh.TextMatrix(Msh.Row, 1)) = 0 Or Msh.TextMatrix(Msh.Row, 1) = "" Then
                                 MsgBox "Select Product Name", vbInformation, "Message"
                                 Msh.Col = 1
                                 Exit Function
                            End If
                            
'If values are missing of Qty or Price
                            If Val(Msh.TextMatrix(Msh.Row, 2)) = 0 Or Val(Msh.TextMatrix(Msh.Row, 3)) = 0 Then
                                 MsgBox "Enter Quantity / Price", vbInformation, "Message"
                                 Msh.Col = 2
                                 Exit Function
                            End If
                            
'Total of Row i.e Qty * Price
                            Msh.TextMatrix(Msh.Row, 4) = Val(Msh.TextMatrix(Msh.Row, 2)) * Val(Msh.TextMatrix(Msh.Row, 3))

                            ''-------------New Row Create----------
                            Msh.Rows = Msh.Rows + 1
                            Msh.Col = 1
                            ''-------------------------------------
                            
                        End If
                        
                        SendKeys "{home}" + "{down}"   '' + "{right}"
                
                End Select
            
            Case Else 'KeyAscii Select
                ''-------write any code for Any Validation-------------
                Select Case Msh.Col
                    Case 2, 3, 4
                        ''-------Allow Number Validation-------------
                        ONUGrid KeyAscii, Msh
                End Select
                ''-------------Write Data in Cells----------
                Msh.Text = Msh.Text + Chr$(KeyAscii)
                ''-------------------------------------
            End Select 'KeyAscii Select
End Function

Public Function EditGridAcOp(Msh As MSHFlexGrid, KeyAscii As Integer)   'For Account Opening
        
'==================================== Editing Grid Code For A/c Opening ====================================
        
'If First Row selected then
    Select Case Msh.Row
        Case 0
            KeyAscii = 0
            Exit Function
    End Select
    
'Block Every Key Except Enter in Column 1 (Account Name)

        Select Case Msh.Col

            Case 1
            
'If ENTER is pressed on Account Name Column then move to next column
            If KeyAscii = 13 Then
                SendKeys "{right}"
                Exit Function
            End If

                KeyAscii = 0
                Exit Function
        End Select

'---------------------------------------------------------
        
        Select Case KeyAscii
            
            Case 8: 'IF KEY IS BACKSPACE THEN
                If Msh.Text <> "" Then Msh.Text = Left$(Msh.Text, (Len(Msh.Text) - 1))
            
            Case 13: 'IF KEY IS ENTER THEN
                
                Select Case Msh.Col
                    
                    Case Is < 5
                        '----------Move Curssor to Right Side untill col >= 3------
                        Msh.SetFocus
                        SendKeys "{right}"
                    
                    Case 5
                        Msh.SetFocus
                        If (Msh.Row + 1) = Msh.Rows Then
                            ''-------- Null Value Chk in last col---------------
                            
'If values are missing in Debit and Credit Side
                            If Len(Msh.TextMatrix(Msh.Row, 1)) = 0 Or Msh.TextMatrix(Msh.Row, 1) = "" Then
                                 MsgBox "Select Account Name", vbInformation, "Message"
                                 Msh.Col = 1
                                 Exit Function
                            End If
                            
'If values are missing in Debit and Credit Side
                            If Val(Msh.TextMatrix(Msh.Row, 4)) = 0 And Val(Msh.TextMatrix(Msh.Row, 5)) = 0 Then
                                 MsgBox "Enter Debit / Credit Amount", vbInformation, "Message"
                                 Msh.Col = 4
                                 Exit Function
                            End If
                            
'If values are in both Debit and Credit Side
                            If Val(Msh.TextMatrix(Msh.Row, 4)) > 0 And Val(Msh.TextMatrix(Msh.Row, 5)) > 0 Then
                                 MsgBox "Enter Only Debit OR Credit Amount", vbInformation, "Message"
                                 Msh.Col = 3
                                 Exit Function
                            End If
                            
                            ''-------------New Row Create----------
                            Msh.Rows = Msh.Rows + 1
                            Msh.Col = 1
                            ''-------------------------------------
                            
                        End If
                        
                        SendKeys "{home}" + "{down}"   '' + "{right}"
                
                End Select
            
            Case Else 'KeyAscii Select
                ''-------write any code for Any Validation-------------
                Select Case Msh.Col
                    Case 3, 4, 5
                        ''-------Allow Number Validation-------------
                        ONUGrid KeyAscii, Msh
                End Select
                ''-------------Write Data in Cells----------
                Msh.Text = Msh.Text + Chr$(KeyAscii)
                ''-------------------------------------
            End Select 'KeyAscii Select
End Function

