Attribute VB_Name = "modWriteOnPictureBox"
Public Sub PrintToCenter(Msg As String, PicBox As PictureBox, FormTop As Boolean)

    If FormTop = False Then                     'Its mean menu
        With PicBox
           .AutoRedraw = -1
           .Font = "Courier New"
           .FontSize = 8
           .FontBold = False
           .ForeColor = vbBlack
           
           HalfWidth = .TextWidth(Msg) / 2     ' Calculate one-half width.
           HalfHeight = .TextHeight(Msg) / 2   ' Calculate one-half height.
           .CurrentX = .ScaleWidth / 2 - HalfWidth   ' Set X.
           .CurrentY = .ScaleHeight / 2 - HalfHeight ' Set Y.
        End With
        
    Else                                        'its mean Form Title Bar
    
        With PicBox
           .AutoRedraw = -1
           .Font = "Courier New"
           .FontSize = 14
           .FontBold = True
           .ForeColor = vbWhite
           
           HalfWidth = .TextWidth(Msg) / 2     ' Calculate one-half width.
           HalfHeight = .TextHeight(Msg) / 2   ' Calculate one-half height.
           .CurrentX = .ScaleWidth / 2 - HalfWidth   ' Set X.
           .CurrentY = .ScaleHeight / 2 - HalfHeight ' Set Y.
        End With
    End If
    
    PicBox.Print Msg   ' Print message.

End Sub


