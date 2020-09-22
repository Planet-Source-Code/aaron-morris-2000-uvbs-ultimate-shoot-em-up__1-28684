Attribute VB_Name = "modfonts"
Private Type struct_font
    fASC As String * 1
    fRect As RECT
    fDestX As Integer
    fDesty As Integer
End Type

Private font(35) As struct_font

Public Function font_to_screen(x As Integer, y As Integer, text As String)

    Dim LenofText As Integer
    Dim drawtext As Integer
    Dim cvo As String * 1
    Dim posof As Integer
    Dim offset As Integer
    Dim kerning As Integer
    
    offset = 55
    kerning = 0
    LenofText = Len(text)
    For drawtext = 1 To LenofText
            
        cvo = Trim$(Mid$(UCase(text), drawtext, 1))
        
        posof = Asc(cvo)
        'If space
        
        
        If posof > 64 Then
            offset = 55
        Else
            If posof = 32 Then
                kerning = kerning + 20
                GoTo SPACE
            Else
                offset = 48
            End If
        End If
        
        posof = posof - offset
        BitBlt frmShoot.hdc, x + (kerning), y, 20, 20, frmShoot.fontset.hdc, font(posof).fRect.Left, font(posof).fRect.top + 80, SRCAND
        BitBlt frmShoot.hdc, x + (kerning), y, 20, 20, frmShoot.fontset.hdc, font(posof).fRect.Left, font(posof).fRect.top, SRCPAINT
        kerning = kerning + 20
    
SPACE:
    'if space don't blit
    Next drawtext

End Function

Public Function load_font_file()
   
    Open App.Path & "\data\fontdat.dat" For Input As #1

    Dim temp1 As String
    Dim temp2 As String
    
    Dim row As Integer
    row = 0
    
    Do Until EOF(1)
       Line Input #1, temp1
       font(row).fASC = Mid(temp1, 1, 1) 'ASCII
       font(row).fRect.Left = Mid(temp1, 3, 3) 'ASCII
       font(row).fRect.top = Mid(temp1, 7, 3) 'ASCII
       font(row).fRect.Right = Mid(temp1, 11, 3) 'ASCII + font(temp2).fRect.Left 'right
       font(row).fRect.Bottom = Mid(temp1, 15, 3) 'ASCII + font(temp2).fRect.top 'bottom
       row = row + 1
    Loop
    
    Close #1

End Function
