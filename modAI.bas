Attribute VB_Name = "modAI"
Public Type Pat1
    x1 As Integer
    y1 As Integer
    x2 As Integer
    y2 As Integer
    SpdX As Integer
    spdY As Integer
End Type
Public CWave As Integer
Public patterns(10) As Pat1


Sub create_paterns()
    'Speed of movement for the pattern
    'Pattern 1
'----------------
'  s
'
'       e
'----------------
    
    patterns(0).SpdX = 1
    patterns(0).spdY = 2
    patterns(0).x1 = 50
    patterns(0).x2 = 50
    patterns(0).y1 = -50
    patterns(0).y2 = SCREENHEIGHT + 50
    
'----------------
'       s
'
'       e
'----------------

    'Pattern 2
    patterns(1).SpdX = 0
    patterns(1).spdY = 3
    patterns(1).x1 = 300
    patterns(1).x2 = 300
    patterns(1).y1 = -50
    patterns(1).y2 = SCREENHEIGHT + 50
    
'----------------
'               s
'
'       e
'----------------

    'Pattern 3
    patterns(2).SpdX = -1
    patterns(2).spdY = 3
    patterns(2).x1 = 600
    patterns(2).x2 = 600
    patterns(2).y1 = -50
    patterns(2).y2 = SCREENHEIGHT + 50

    
'----------------
'           s
'
'           e
'----------------

    'Pattern 2
    patterns(3).SpdX = 0
    patterns(3).spdY = 2
    patterns(3).x1 = 550
    patterns(3).x2 = 550
    patterns(3).y1 = -50
    patterns(3).y2 = SCREENHEIGHT + 50
    
    'Pattern 4
    patterns(4).SpdX = 0
    patterns(4).spdY = 2
    patterns(4).x1 = 600
    patterns(4).x2 = 600
    patterns(4).y1 = -50
    patterns(4).y2 = SCREENHEIGHT + 50
    'pat5
    patterns(5).SpdX = 1
    patterns(5).spdY = 2
    patterns(5).x1 = 10
    patterns(5).x2 = 10
    patterns(5).y1 = -50
    patterns(5).y2 = SCREENHEIGHT + 50
    'pat6
    patterns(6).SpdX = 1
    patterns(6).spdY = 2
    patterns(6).x1 = 0
    patterns(6).x2 = 0
    patterns(6).y1 = -150
    patterns(6).y2 = SCREENHEIGHT + 50
    'pat7
    patterns(7).SpdX = 1
    patterns(7).spdY = 2
    patterns(7).x1 = 0
    patterns(7).x2 = 0
    patterns(7).y1 = -100
    patterns(7).y2 = SCREENHEIGHT + 50
     'pat8
    patterns(8).SpdX = 1
    patterns(8).spdY = 4
    patterns(8).x1 = 150
    patterns(8).x2 = 150
    patterns(8).y1 = -100
    patterns(8).y2 = SCREENHEIGHT + 50
    'pat9
    patterns(9).SpdX = 1
    patterns(9).spdY = 3
    patterns(9).x1 = 220
    patterns(9).x2 = 220
    patterns(9).y1 = -50
    patterns(9).y2 = SCREENHEIGHT + 50
    'pat10 ---- "------------ Big Enemy 1"
    patterns(10).SpdX = 0
    patterns(10).spdY = 2
    patterns(10).x1 = 220
    patterns(10).x2 = 220
    patterns(10).y1 = -250
    patterns(10).y2 = SCREENHEIGHT + 250

End Sub

Function WaveFinder()
    Dim u As Integer
    'Use the patters within the wave attack
If CWave = 1 Then
        NUM_ENEMIES = 4 'Ships in this attack wave
        ReDim en(NUM_ENEMIES) As Enemy
        ReDim en_fire(NUM_ENEMIES) As Laser
  
End If
If CWave = 2 Then
        NUM_ENEMIES = 6 'Ships in this attack wave
        ReDim en(NUM_ENEMIES) As Enemy
        ReDim en_fire(NUM_ENEMIES) As Laser
End If
If CWave = 3 Then
        NUM_ENEMIES = 8 'Ships in this attack wave
        ReDim en(NUM_ENEMIES) As Enemy
        ReDim en_fire(NUM_ENEMIES) As Laser
       
End If
If CWave = 4 Then
        NUM_ENEMIES = 10 'Ships in this attack wave
        ReDim en(NUM_ENEMIES) As Enemy
        ReDim en_fire(NUM_ENEMIES) As Laser
End If
    
        '-- Assign the movement pattern to each ship
        For u = 0 To NUM_ENEMIES - 1
            en(u).active = True
            en(u).y = patterns(u).y1
            en(u).x = patterns(u).x1
            en(u).spd = patterns(u).spdY
            en(u).SpdX = patterns(u).SpdX
            en(u).Width = 40
            en(u).Height = 30
            en(u).Life = 1
        Next u
End Function
