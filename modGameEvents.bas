Attribute VB_Name = "modGameEvents"
'The Game state...So I can tell if in game loop of in menu
Public g_GameState As Integer
'Turns on the End of level badies
Public bEndOfLevelEnemy As Boolean

'Constants...They pretty much make sense
Public Const NUM_LASERS = 19
Public NUM_ENEMIES As Integer

Public Enum IGameObjs
    IPlayer1 = 0
    IEnemies = 1
    ILevel_End_Enemies = 2
    IPlayer1_Laser = 3
    IEnemies_Laser = 4
    ILevel_End_Enemies_Laser = 5
    IPower_Ups = 6
    IInfomation = 7
End Enum


'You ---- The player
Public Type Ship
    x As Integer
    y As Integer
    DirectionX As Integer
    DirectionY As Integer
    shRect As RECT
    spd As Integer
    Width As Integer
    Height As Integer
    Life As Integer
    nScore As Integer
    bShield As Boolean
    nShield As Integer
    active_laser As Integer
End Type

Public Type Enemy
    x As Integer
    y As Integer
    shRect As RECT
    spd As Integer
    SpdX As Integer
    Width As Integer
    Height As Integer
    active As Boolean
    Life As Integer
End Type

Public Type EOLEnemy_Lasers
    posX As Integer
    posY As Integer
    bActive As Boolean
    offsetX As Integer ' Where to start the lasers
    offsetY As Integer
End Type

Public Type EOLEnemy 'End of Level Enemy
    x As Integer
    y As Integer
    spd As Integer
    SpdX As Integer
    FirePower() As EOLEnemy_Lasers
    Width As Integer
    Height As Integer
    active As Boolean
    Life As Integer
End Type

Public e_active_laser As Integer   'Enemies laser flag

'Your lasers
Public Type Laser
    LaserX As Integer
    LaserY As Integer
    LaserRect As RECT
    LasarCounter As Integer
    laseracitve As Boolean
End Type

Public Type Power_Ups
    x As Integer
    y As Integer
    dx As Integer
    dy As Integer
    Height As Integer
    Width As Integer
    bActive As Boolean
    AnimFrame As Integer
    AnimMax As Integer
    AnimRate As Integer
    AnimCounter As Integer
End Type



'Objects to the structures

Public Las(NUM_LASERS) As Laser
Public en() As Enemy
Public en_fire() As Laser
Public Player As Ship '1 play --- you -- only 1 object :)

Private BIG_BADIE(0) As EOLEnemy
Private PlayerPU As Power_Ups

'Initialize stuff
Public Sub Init_Game()
    frmShoot.WindowState = 0 'See above
    frmShoot.ScaleHeight = SCREENHEIGHT
    frmShoot.ScaleWidth = SCREENWIDTH

    'Start The game
    GameLoop
End Sub

Public Sub GameLoop()
    
    Load_Map
    create_paterns
    frmShoot.Show ' Force the form to show....
    
    '-----Set up player..Default positions etc...
    Player.shRect.Left = 0
    Player.shRect.top = 0
    Player.Height = 50
    Player.Width = 40
    Player.nShield = 100
    Player.shRect.Right = Player.shRect.Left + Player.Width
    Player.shRect.Bottom = Player.shRect.top + Player.Height
    Player.spd = 3
    Player.x = Player.x + Player.spd
    Player.y = Player.y + Player.spd
    Player.Life = 100 '100%
    Player.active_laser = 0
    Player.nScore = 0
    


    Reset_Tiles
        
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
    
        
    '--------main game loop-----
    Do While g_GameState = 1
        

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
                'Must do events so the keyboard pressess will be detected
                
                'Urrrrggggg For the time being I'm using CLS to clear the screen
                'Very ugly but simple
               
                frmShoot.Cls
                Draw_Tiles
            If Player.Life > 0 Then
                'Draw Ship
                Draw_Ship
                'fire
                Player_Fire
                
                'check hit
                collision_det
                collision_det_ply
                
                If bEndOfLevelEnemy = True Then
                    End_Level_Enemy_go
                Else
                    Draw_Enemies
                    'enemies fire
                    Enemy_Fire
                End If
                
                If PlayerPU.bActive = True Then
                    Power_Ups
                End If
                
                Display_Information
                
                'check Events based on tile positions
                Events
            Else
                GameOver
                g_GOFlag = True
            End If
               'Increase the timer 'All done this time
               next_time = currentTime + time_count
        End If
      
    Loop
    'Loop broken call menu

End Sub

Private Sub Display_Information()
    
    Dim nrg_pc As Integer
    
    
    'Update player score with my custom font functions
    font_to_screen 10, 10, "Score"
    font_to_screen 120, 10, CStr(Player.nScore)
    font_to_screen 180, 10, "Shield"
    font_to_screen 310, 10, CStr(Player.nShield)
                
    
    
    nrg_pc = Int(Player.Life)
    BitBlt frmShoot.hdc, 510, 10, nrg_pc, 14, frmShoot.Ship.hdc, 190, 0, SRCCOPY

End Sub

'SUB - Detect the hit on the player
Private Sub collision_det_ply()
    If NUM_ENEMIES > 0 Then
        
        Dim e As Integer
        
        For e = 0 To NUM_ENEMIES
            
            'Detect a hit by enemies laser
            If en_fire(e).laseracitve = True Then
                If en_fire(e).LaserX >= Player.shRect.Left Then
                    If en_fire(e).LaserX <= Player.shRect.Left + Player.Width Then
                        If en_fire(e).LaserY >= Player.shRect.top Then
                            If en_fire(i).LaserY <= Player.shRect.top + Player.Height Then
                                Destory_Player e
                            End If
                        End If
                    End If
                End If
            
            'See if enemy has hit player
            ElseIf en(e).active = True Then
                If en(e).x >= Player.shRect.Left Then
                    If (en(e).x + en(e).Width) <= Player.shRect.Left + Player.Width Then
                        If en(e).y >= Player.shRect.top Then
                            If (en(e).y + en(e).Height) <= Player.shRect.top + Player.Height Then
                                'hit player
                                Destory_Player e
                                'hit enemy
                                Destory_Enemy e, -1
                            End If
                        End If
                    End If
                End If
            End If
         
         Next e
    
    End If

End Sub



'SUB - Deteced hit kill the player
Private Sub Destory_Player(e As Integer)

'Is shield up ?
 If Player.nShield > 0 Then 'check to see if we have any shields left
    If Player.bShield = True Then ' Are the shields up ?
        If g_Snd = True Then play_sound 3 'Hit shield Sound
        Exit Sub
    End If
 End If
    
    'we did not exit so the player has been hit!
    If g_Snd = True Then play_sound 1
    Player.Life = Player.Life - 10
    'Changed to include hit detection for big badie
    
    If bEndOfLevelEnemy = False Then 'Not EOL Badie
        en_fire(e).laseracitve = False
    Else
        'Is end of level enemy
        BIG_BADIE(0).FirePower(e).bActive = False
    End If
    
    
End Sub

'SUB - Detect the hit on the enemies
Private Sub collision_det()

    If NUM_ENEMIES > 0 Then
    
        Dim i As Integer
        Dim e As Integer
        
        For e = 0 To NUM_ENEMIES
            If en(e).active Then
                For i = 0 To NUM_LASERS
                    If Las(i).laseracitve = True Then
                        If Las(i).LaserX + 25 >= en(e).x Then
                            If Las(i).LaserX + 25 <= (en(e).x + en(e).Width) Then
                                If Las(i).LaserY >= en(e).y Then
                                    If Las(i).LaserY <= (en(e).y + en(e).Height) Then
                                            'Kill enemy
                                            'We hit do remove enemy
                                            Destory_Enemy e, i
                                        
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next i
            End If
        Next e
        
    End If
    
End Sub

Private Sub Create_Explosions(rLoc As RECT, ExlodeID As Integer)
        
        Select Case explodeid
        'enemy ship 1
            Case 0
                rLoc.Right = 39
                rLoc.Bottom = 39
            End Select

        BitBlt frmShoot.hdc, rLoc.Left, rLoc.top, rLoc.Right, rLoc.Bottom, frmShoot.Ship.hdc, 40, 140, SRCAND
        BitBlt frmShoot.hdc, rLoc.Left, rLoc.top, rLoc.Right, rLoc.Bottom, frmShoot.Ship.hdc, 0, 140, SRCPAINT
        
        'Play the explosions sound
        If g_Snd = True Then play_sound 1
End Sub


'SUB - Detected hit on enemy --- Kill it
Private Sub Destory_Enemy(e As Integer, Optional i As Integer)
    
    If i <> -1 Then     'Check that we destroyed the enemy with laser and
                        'Not by a collision
        Las(i).laseracitve = False 'Delete collition laser (Not working...laser could kill more enemies)
        Player.nScore = Player.nScore + 1
    End If
    
    en(e).Life = en(e).Life - 1
    
    If en(e).Life <= 0 Then
        en(e).active = False
        
        Dim r As RECT
        
        r.Left = en(e).x
        r.top = en(e).y
        Create_Explosions r, 0
        
    End If
    
End Sub

'SUB - Create a new enemy
Private Sub Set_Enemy()
    WaveFinder
End Sub

'SUB - Draw the enemies to screen
Private Sub Draw_Enemies()
   If NUM_ENEMIES > 0 Then
    Dim i As Integer
    
    For i = 0 To NUM_ENEMIES
           If en(i).active = True Then
                
                en(i).y = en(i).y + en(i).spd
                en(i).x = en(i).x + en(i).SpdX
                BitBlt frmShoot.hdc, en(i).x, en(i).y, 39, 30, frmShoot.Ship.hdc, 120, 0, SRCAND
                BitBlt frmShoot.hdc, en(i).x, en(i).y, 39, 30, frmShoot.Ship.hdc, 80, 0, SRCPAINT
  
                If en(i).y > patterns(i).y2 Then
                    en(i).active = False
                End If
            End If
  
    Next i
    End If
End Sub

'SUB - Draw YOU the player to screen
Private Sub Draw_Ship()

    If Player.shRect.Left >= SCREENWIDTH Then
        Player.DirectionX = -Player.spd
    Else
        If Player.shRect.Left < 0 Then
            Player.DirectionX = Player.spd
        End If
    End If
       
    
    If Player.shRect.top >= SCREENHEIGHT Then
        Player.DirectionY = -Player.spd
    Else
        If Player.shRect.top <= 0 Then
            Player.DirectionY = Player.spd
        End If
    End If
    
    Player.shRect.top = Player.shRect.top + Player.DirectionY
    Player.shRect.Left = Player.shRect.Left + Player.DirectionX
 
    If Player.bShield = False Then
        BitBlt frmShoot.hdc, Player.shRect.Left, Player.shRect.top, Player.shRect.Right, Player.shRect.Bottom, frmShoot.Ship.hdc, 40, 0, SRCAND
        BitBlt frmShoot.hdc, Player.shRect.Left, Player.shRect.top, Player.shRect.Right, Player.shRect.Bottom, frmShoot.Ship.hdc, 0, 0, SRCPAINT
    Else
        If Player.nShield > 0 Then
            Player.nShield = Player.nShield - 1
            BitBlt frmShoot.hdc, Player.shRect.Left, Player.shRect.top, Player.shRect.Right, Player.shRect.Bottom, frmShoot.Ship.hdc, 40, 80, SRCAND
            BitBlt frmShoot.hdc, Player.shRect.Left, Player.shRect.top, Player.shRect.Right, Player.shRect.Bottom, frmShoot.Ship.hdc, 0, 80, SRCPAINT
        Else
            Player.bShield = False
        End If
    End If
    
End Sub

'SUB - The player fires a laser
Private Sub Player_Fire()
    
    Dim i As Integer
    
    For i = 0 To NUM_LASERS
        
        If Las(i).laseracitve = True Then
            
            BitBlt frmShoot.hdc, Las(i).LaserX, Las(i).LaserY, 10, 10, frmShoot.Ship.hdc, 30, 50, SRCAND
            BitBlt frmShoot.hdc, Las(i).LaserX, Las(i).LaserY, 10, 10, frmShoot.Ship.hdc, 20, 50, SRCPAINT
            'if power-up
            BitBlt frmShoot.hdc, Las(i).LaserX - 10, Las(i).LaserY, 10, 10, frmShoot.Ship.hdc, 30, 50, SRCAND
            BitBlt frmShoot.hdc, Las(i).LaserX - 10, Las(i).LaserY, 10, 10, frmShoot.Ship.hdc, 20, 50, SRCPAINT
            
            Las(i).LaserY = Las(i).LaserY - 8
            If Las(i).LaserY < -20 Then
                Las(i).laseracitve = False
            End If
        Else
            Las(i).LaserX = -10
            Las(i).LaserY = -10
            Player.active_laser = i
            Las(i).laseracitve = True
        End If
    
    Next i
End Sub

'SUB - The Enemy fires a laser

Private Sub Enemy_Fire()
If NUM_ENEMIES > 0 Then
    Dim i As Integer

    For i = 0 To NUM_ENEMIES
        
            If en_fire(i).laseracitve = True Then
            
                BitBlt frmShoot.hdc, en_fire(i).LaserX, en_fire(i).LaserY, 10, 10, frmShoot.Ship.hdc, 60, 50, SRCAND
                BitBlt frmShoot.hdc, en_fire(i).LaserX, en_fire(i).LaserY, 10, 10, frmShoot.Ship.hdc, 50, 50, SRCPAINT
                
                en_fire(i).LaserY = en_fire(i).LaserY + 8
                If en_fire(i).LaserY > 800 Then
                    en_fire(i).laseracitve = False
                End If
            Else
                e_active_laser = i
               
                If rnd_fire(i) = True Then
                    en_fire(i).laseracitve = True
                   
                End If
            End If
       
    Next i
End If
End Sub

'SUB - Randomize the enemy fire
Private Function rnd_fire(CurrentEn As Integer) As Boolean
    Randomize Timer
    Dim rand As Integer
   
    rand = Rnd(10) * 10
    If en(CurrentEn).active = True Then 'Enemy is NOT Killed
    If en(CurrentEn).y > 0 Then
    
    If rand < 9 Then
        'set fire start position
        en_fire(CurrentEn).LaserX = en(CurrentEn).x + 10
        en_fire(CurrentEn).LaserY = en(CurrentEn).y + 15
     
        rnd_fire = True
        If g_Snd = True Then play_sound 4
    Else
        rnd_fire = False
    End If
    
    End If
    End If

End Function


Public Sub Kill_Game()
    'Quit Game :(
    End
End Sub

'Displays -- "GET READY!" for a few seconds at the start of level!!!
Private Sub Blit_GR()
    BitBlt frmShoot.hdc, 250, 300, 170, 39, frmShoot.Ship.hdc, 320, 40, SRCAND
    BitBlt frmShoot.hdc, 250, 300, 170, 39, frmShoot.Ship.hdc, 320, 0, SRCPAINT
End Sub
'Displays -- "LEVEL UP!" for a few seconds at the start of level!!!
Private Sub Level_UP()
    BitBlt frmShoot.hdc, 200, 300, 270, 38, frmShoot.Ship.hdc, 320, 120, SRCAND
    BitBlt frmShoot.hdc, 200, 300, 270, 38, frmShoot.Ship.hdc, 320, 80, SRCPAINT
End Sub
'Displays -- "GAME OVER!" for a few seconds at the start of level!!!
Private Sub GameOver()
    BitBlt frmShoot.hdc, 250, 300, 155, 24, frmShoot.Ship.hdc, 150, 90, SRCAND
    BitBlt frmShoot.hdc, 250, 300, 155, 24, frmShoot.Ship.hdc, 150, 60, SRCPAINT
End Sub

Private Sub Set_Power_Ups()
    PlayerPU.x = patterns(0).x1
    PlayerPU.y = patterns(0).y1
    PlayerPU.bActive = True
    PlayerPU.dx = 1
    PlayerPU.dy = 2
    PlayerPU.AnimMax = 1
    PlayerPU.AnimRate = 4
    PlayerPU.Height = 25
    PlayerPU.Width = 25
End Sub

Private Sub Power_Ups()
    
    Dim x1 As Integer
    Dim x2 As Integer
    Dim y1 As Integer
    Dim y2 As Integer
    Dim h As Integer
    Dim w As Integer
    
    PlayerPU.x = PlayerPU.x + PlayerPU.dx
    PlayerPU.y = PlayerPU.y + PlayerPU.dy
    'Do some nice animation here !
    
    If PlayerPU.AnimFrame = 0 Then
        h = 20
        w = 20
        x1 = 130
        x2 = 110
        y1 = 50
        y2 = 50
    ElseIf PlayerPU.AnimFrame = 1 Then
        h = 20
        w = 20
        x1 = 130
        x2 = 110
        y1 = 70
        y2 = 70
    End If
    
    BitBlt frmShoot.hdc, PlayerPU.x, PlayerPU.y, w, h, frmShoot.Ship.hdc, x1, y1, SRCAND
    BitBlt frmShoot.hdc, PlayerPU.x, PlayerPU.y, w, h, frmShoot.Ship.hdc, x2, y2, SRCPAINT
    
    PlayerPU.AnimCounter = PlayerPU.AnimCounter + 1
    
    If PlayerPU.AnimCounter >= PlayerPU.AnimRate Then
        PlayerPU.AnimFrame = PlayerPU.AnimFrame + 1
        PlayerPU.AnimCounter = 0
    End If
    
    If PlayerPU.AnimFrame > PlayerPU.AnimMax Then PlayerPU.AnimFrame = 0
    
    Check_Col_Power_Up 'Check collition with power up
    If PlayerPU.y > SCREENHEIGHT Then PlayerPU.bActive = False
    
End Sub

Private Sub Check_Col_Power_Up()
    If PlayerPU.x >= Player.shRect.Left Then
        If (PlayerPU.x + PlayerPU.Width) <= (Player.shRect.Left + Player.Width) Then
            If PlayerPU.y >= Player.shRect.top Then
                If (PlayerPU.y + PlayerPU.Height) <= (Player.shRect.top + Player.Height) Then
                    'We hit the power-up
                    If Player.Life < 90 Then
                        Player.Life = Player.Life + 10 'add ten to life !
                    Else
                        'Round life to 100
                        Player.Life = 100
                    End If
                    If g_Snd = True Then play_sound 2 'Make a nice "Power-Up" sound
                    PlayerPU.bActive = False
                End If
            End If
        End If
    End If
End Sub

Private Sub Events()
    Select Case nCurrentTileID
    
        Case Is < 3
            Blit_GR
      
        Case 4
           CWave = 1
           Set_Enemy
            End_Level_Enemy_Setup
        
        Case 10
            CWave = 2
            Set_Enemy
        
        Case 11
            Set_Power_Ups
                    
        Case 17
            CWave = 3
            Set_Enemy
        Case 23
            CWave = 1
            Set_Enemy
        
        Case 35
            CWave = 3
            Set_Enemy
            
        
        Case 36
            Set_Power_Ups
            
        Case 37
            Blit_GR
               
        Case 38
            CWave = 2
            Set_Enemy
                
        Case 45
            CWave = 1
            Set_Enemy
        
        Case 54
            CWave = 2
            Set_Enemy
        Case 59
            CWave = 4
            Set_Enemy
        Case 64
            CWave = 4
            Set_Enemy
        Case 68
            Set_Power_Ups
            Level_UP
        Case 74
            End_Level_Enemy_Setup

    End Select
End Sub

Private Sub End_Level_Enemy_Setup()

    frmShoot.elb.Picture = LoadPicture(App.Path + "\bits\elb.bmp")
    'End of Level Enemy
    BIG_BADIE(0).active = True
    BIG_BADIE(0).y = -200
    BIG_BADIE(0).x = 250
    BIG_BADIE(0).Height = 149
    BIG_BADIE(0).Width = 149
    BIG_BADIE(0).spd = 2
    BIG_BADIE(0).SpdX = 1
    BIG_BADIE(0).Life = 100
    ReDim BIG_BADIE(0).FirePower(4)
    BIG_BADIE(0).FirePower(0).offsetX = 0
    BIG_BADIE(0).FirePower(1).offsetX = 30
    BIG_BADIE(0).FirePower(2).offsetX = 100
    BIG_BADIE(0).FirePower(3).offsetX = 140
    BIG_BADIE(0).FirePower(4).offsetX = 75
    BIG_BADIE(0).FirePower(0).offsetY = 150
    BIG_BADIE(0).FirePower(1).offsetY = 130
    BIG_BADIE(0).FirePower(2).offsetY = 130
    BIG_BADIE(0).FirePower(3).offsetY = 150
    BIG_BADIE(0).FirePower(4).offsetY = 140
    
    BIG_BADIE(0).FirePower(0).posX = 100
    BIG_BADIE(0).FirePower(1).posX = 110
    BIG_BADIE(0).FirePower(2).posX = 120
    BIG_BADIE(0).FirePower(3).posX = 130
    
    bEndOfLevelEnemy = True
    
    
End Sub

Private Sub End_Level_Enemy_go()

    BIG_BADIE(0).x = BIG_BADIE(0).x + BIG_BADIE(0).SpdX
    BIG_BADIE(0).y = BIG_BADIE(0).y + BIG_BADIE(0).spd
    
    BitBlt frmShoot.hdc, BIG_BADIE(0).x, BIG_BADIE(0).y, 149, 149, frmShoot.elb.hdc, 150, 0, SRCAND
    BitBlt frmShoot.hdc, BIG_BADIE(0).x, BIG_BADIE(0).y, 149, 149, frmShoot.elb.hdc, 0, 0, SRCPAINT
    
    If BIG_BADIE(0).y > 200 Then
        BIG_BADIE(0).spd = -2
    End If
    
    If BIG_BADIE(0).y < -50 Then
        BIG_BADIE(0).spd = 2
    End If
    
    If BIG_BADIE(0).x > (SCREENWIDTH - 100) Then
        BIG_BADIE(0).SpdX = -2
    End If
    If BIG_BADIE(0).x < 0 Then
        BIG_BADIE(0).SpdX = 2
    End If
    
    'Firing
    End_Level_Enemy_FIRE
    collision_det_BIDBADIE
    collision_det_ply_enemy
    'Display Energy
    font_to_screen 10, 400, "Enemy"
    Dim nrg_pc As Integer
    nrg_pc = Int(BIG_BADIE(0).Life)
    BitBlt frmShoot.hdc, 10, 430, nrg_pc, 14, frmShoot.Ship.hdc, 190, 0, SRCCOPY
    If nrg_pc <= 0 Then Kill_End_Level_Enemy
End Sub
Private Sub End_Level_Enemy_FIRE()
    Dim i As Integer
    If BIG_BADIE(0).y + BIG_BADIE(0).Height > 1 Then
    
    For i = 0 To UBound(BIG_BADIE(0).FirePower())
        
            If BIG_BADIE(0).FirePower(i).bActive = True Then
            
                BitBlt frmShoot.hdc, BIG_BADIE(0).FirePower(i).posX, BIG_BADIE(0).FirePower(i).posY, 9, 29, frmShoot.Ship.hdc, 330, 160, SRCAND
                BitBlt frmShoot.hdc, BIG_BADIE(0).FirePower(i).posX, BIG_BADIE(0).FirePower(i).posY, 9, 29, frmShoot.Ship.hdc, 320, 160, SRCPAINT
                
                BIG_BADIE(0).FirePower(i).posY = BIG_BADIE(0).FirePower(i).posY + 8
                
                If BIG_BADIE(0).FirePower(i).posY > 800 Then
                    BIG_BADIE(0).FirePower(i).bActive = False
                End If
            Else
                   If g_Snd = True Then play_sound 1
                'If rnd_fire(i) = True Then
                    BIG_BADIE(0).FirePower(i).bActive = True
                    BIG_BADIE(0).FirePower(i).posX = BIG_BADIE(0).x + BIG_BADIE(0).FirePower(i).offsetX
                    BIG_BADIE(0).FirePower(i).posY = BIG_BADIE(0).y + BIG_BADIE(0).FirePower(i).offsetY
                'End If
            End If
       
    Next i
    End If
End Sub
 
'SUB - Detect the hit on the enemies
Private Sub collision_det_BIDBADIE()
    Dim i As Integer
        
            For i = 0 To NUM_LASERS
                If Las(i).laseracitve = True Then
                    If Las(i).LaserX + 25 >= BIG_BADIE(0).x Then
                        If Las(i).LaserX + 25 <= (BIG_BADIE(0).x + BIG_BADIE(0).Width) Then
                            If Las(i).LaserY >= BIG_BADIE(0).y Then
                                If Las(i).LaserY <= (BIG_BADIE(0).y + BIG_BADIE(0).Height) Then
                                    'Get the pixel value from the source picture box
                                    'We cannot check on the visible screen as there are tiles
                                    'and other things that contain colors!!!
                                    If GetPixel(frmShoot.elb.hdc, Las(i).LaserX - BIG_BADIE(0).x, Las(i).LaserY - BIG_BADIE(0).y) <> &H0 Then '
                                        Player.nScore = Player.nScore + 1
                                        If g_Snd = True Then play_sound 1
                                        BitBlt frmShoot.hdc, Las(i).LaserX, Las(i).LaserY, 39, 40, frmShoot.Ship.hdc, 40, 140, SRCAND
                                        BitBlt frmShoot.hdc, Las(i).LaserX, Las(i).LaserY, 39, 40, frmShoot.Ship.hdc, 0, 140, SRCPAINT
                                        Las(i).laseracitve = False
                                        BIG_BADIE(0).Life = BIG_BADIE(0).Life - 1
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next i
End Sub
Private Sub collision_det_ply_enemy()
        Dim e As Integer
        
        For e = 0 To UBound(BIG_BADIE(0).FirePower())
            If BIG_BADIE(0).FirePower(e).bActive = True Then
                If BIG_BADIE(0).FirePower(e).posX >= Player.shRect.Left Then
                    If BIG_BADIE(0).FirePower(e).posX <= Player.shRect.Left + Player.Width Then
                        If BIG_BADIE(0).FirePower(e).posY >= Player.shRect.top Then
                            If BIG_BADIE(0).FirePower(e).posY <= Player.shRect.top + Player.Height Then
                                Destory_Player e
                            End If
                        End If
                    End If
                End If
            End If
         Next e
    

End Sub
Private Sub Kill_End_Level_Enemy()
    Dim i As Integer
    Dim ii As Integer
    For i = 0 To BIG_BADIE(0).Width Step 50
        For ii = 0 To BIG_BADIE(0).Height Step 50
                If g_Snd = True Then play_sound 1
                BitBlt frmShoot.hdc, BIG_BADIE(0).x + i, BIG_BADIE(0).y + ii, 39, 40, frmShoot.Ship.hdc, 40, 140, SRCAND
                BitBlt frmShoot.hdc, BIG_BADIE(0).x + i, BIG_BADIE(0).y + ii, 39, 40, frmShoot.Ship.hdc, 0, 140, SRCPAINT
        Next ii
    Next i
    bEndOfLevelEnemy = False
End Sub
