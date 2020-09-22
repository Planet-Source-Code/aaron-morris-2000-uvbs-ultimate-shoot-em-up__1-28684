Attribute VB_Name = "modSelections"
Option Explicit

Public nSelection As Integer
Private Const lColor1 As Long = &H4411FF
Private Const lColor2 As Long = &H11FFFF

Public Sub Options_Screen(nCase As Integer)
    
    Dim lPlayColor As Long
    Dim lHighColor As Long
    Dim lOptionsColor As Long
    Dim lQuitColor As Long
   
    '----------------------------
    '---Menu Selections
    '----------------------------
    
    Select Case nSelection
        Case 0:
            lPlayColor = lColor1
            lHighColor = lColor2
            lOptionsColor = lColor2
            lQuitColor = lColor2
            
        Case 1:
            lPlayColor = lColor2
            lHighColor = lColor1
            lOptionsColor = lColor2
            lQuitColor = lColor2
                
        Case 2:
            lPlayColor = lColor2
            lHighColor = lColor2
            lOptionsColor = lColor1
            lQuitColor = lColor2
            
        Case 3:
            lPlayColor = lColor2
            lHighColor = lColor2
            lOptionsColor = lColor2
            lQuitColor = lColor1
            
    End Select

    SetTextColor frmShoot.hdc, lPlayColor
    If g_Snd = False Then
      TextOut frmShoot.hdc, 200, 100, "Sound Off", 9
    Else
       TextOut frmShoot.hdc, 200, 100, "Sound On", 8
    End If
    
    
    
    SetTextColor frmShoot.hdc, lHighColor
    TextOut frmShoot.hdc, 200, 150, "Game Play - Easy", 16
    
    SetTextColor frmShoot.hdc, lOptionsColor
    If g_Muz = False Then
        TextOut frmShoot.hdc, 200, 200, "Music Off", 9
    Else
        TextOut frmShoot.hdc, 200, 200, "Music On", 8
    End If
    
    SetTextColor frmShoot.hdc, lQuitColor
    TextOut frmShoot.hdc, 200, 250, "Back", 4
    
    'Show the UVBS Title
    DrawTitle
    
End Sub
Private Sub DrawTitle()
    BitBlt frmShoot.hdc, 100, 20, 150, 40, frmShoot.Ship.hdc, 90, 160, SRCAND
    BitBlt frmShoot.hdc, 100, 20, 150, 40, frmShoot.Ship.hdc, 90, 120, SRCPAINT
End Sub

Public Sub Title_Screen(nCase As Integer)
   'frmShoot.Cls 'clear the screen
       
    Dim lPlayColor As Long
    Dim lHighColor As Long
    Dim lOptionsColor As Long
    Dim lQuitColor As Long
    

    Select Case nSelection
        Case 0:
            lPlayColor = lColor1
            lHighColor = lColor2
            lOptionsColor = lColor2
            lQuitColor = lColor2
        
        Case 1:
            lPlayColor = lColor2
            lHighColor = lColor1
            lOptionsColor = lColor2
            lQuitColor = lColor2
                
        Case 2:
            lPlayColor = lColor2
            lHighColor = lColor2
            lOptionsColor = lColor1
            lQuitColor = lColor2
       
            
        Case 3:
            lPlayColor = lColor2
            lHighColor = lColor2
            lOptionsColor = lColor2
            lQuitColor = lColor1
        
    End Select

    
    SetTextColor frmShoot.hdc, lPlayColor
    TextOut frmShoot.hdc, 200, 100, "Play Game", 9
    SetTextColor frmShoot.hdc, lHighColor
    TextOut frmShoot.hdc, 200, 150, "Controls", 8
    SetTextColor frmShoot.hdc, lOptionsColor
    TextOut frmShoot.hdc, 200, 200, "Options", 7
    SetTextColor frmShoot.hdc, lQuitColor
    TextOut frmShoot.hdc, 200, 250, "Quit", 4
    
    'Show the UVBS Title
    DrawTitle

End Sub

Public Sub controls()
   
    SetTextColor frmShoot.hdc, lColor2
    TextOut frmShoot.hdc, 50, 100, "[ESCAPE] = Back To Menu", 23
    TextOut frmShoot.hdc, 50, 150, "[Z] = Shield Up/Down", 20
    TextOut frmShoot.hdc, 50, 200, "[UP] = Forwards", 15
    TextOut frmShoot.hdc, 50, 250, "[DOWN] = Backwards", 17
    TextOut frmShoot.hdc, 50, 300, "[LEFT] = Left", 13
    TextOut frmShoot.hdc, 50, 350, "[RIGHT] = Right", 15
    TextOut frmShoot.hdc, 50, 400, "[SPACE] = Fire Weapon", 21
    'Show the UVBS Title
    DrawTitle

End Sub

Public Sub scroll_menu()
'Performance counter
    '***As VB does not support 64 bit longs we can trick it by
    'using the Currency variable
    
    Dim frequency As Currency
    Dim currentTime As Currency
    Dim startTime As Currency
    Dim time_elapsed As Double
    Dim result As Double
    Dim next_time As Double
    Dim last_time As Double
    Dim perf_flag As Boolean
    Dim time_scale As Double
    Dim time_count As Double
    time_scale = 0.001
    perf_flag = True
   
    ' get the frequency counter
    ' return zero if hardware doesn't support high-res performance counters
    ' In this case we use the timeGettimer
    
    If QueryPerformanceFrequencyAny(frequency) = 0 Then
         perf_flag = False
    Else
        time_count = Val(frequency) / 40
        time_scale = (1 / Val(frequency))
    End If
    
        
    '--------main loop-----
    Do While g_GameState = 0
        

        If (perf_flag = True) Then
            
            QueryPerformanceCounterAny currentTime
 
            
        Else
            currentTime = timeGetTime()
        End If
        
           If (currentTime > next_time) Then
                'Set our counter
                time_elapsed = (currentTime - last_time) * time_scale
                last_time = currentTime
               DoEvents
               frmShoot.Cls
               Draw_Tiles
               '-----
                nSelection = nOption
                If g_OptionsFlag = 0 Then
                    Title_Screen 0
                ElseIf g_OptionsFlag = 1 Then
                    Options_Screen 0
                ElseIf g_OptionsFlag = 2 Then
                    controls
                End If
                
               'Increase the timer 'All done this time
               next_time = currentTime + time_count
        End If
      
    Loop

End Sub
