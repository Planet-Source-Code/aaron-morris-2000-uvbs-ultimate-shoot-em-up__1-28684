Attribute VB_Name = "modKeyHandler"
Option Explicit

'Flags used for detecting the current menu being viewed
Public nOption As Integer
Public g_OptionsFlag As Integer
Private g_SelectedOption As Integer
Public g_GOFlag As Boolean

Public Function KeyEvents(KeyCode As Integer)

'Here I'm using the g_GameState flag to figure out if I'm in a menu (if so which) or
'I'm detection the keys for the player movements

    If g_GameState = 0 Then
        If g_OptionsFlag = 0 Then
            Select Case KeyCode
            'g_OptionsFlag = 0 - Main Title Screen
                
                Case vbKeyUp
                
                    If nOption > 0 Then
                        nOption = nOption - 1
                    Else
                        nOption = 3
                    End If
              
                
                Case vbKeyDown
                    If nOption < 3 Then
                        nOption = nOption + 1
                    Else
                        nOption = 0
                    End If
               
                    
                Case vbKeySpace
                    
                    
                    Select Case nOption
                        Case 0:
                            g_GameState = 1
                            Init_Game
                            Exit Function
                         
                        Case 1:
                            g_OptionsFlag = 2
                            Exit Function
                        Case 2:
                            g_OptionsFlag = 1
                                                        
                            Exit Function
                        Case 3:
                            Kill_Game 'end game
                    
                    End Select
                
            End Select
        
            'Title_Screen
        Else
            If g_OptionsFlag = 1 Then
            'g_OptionsFlag = 1 - Options
            
                Select Case KeyCode
                    Case vbKeyUp
                        If nOption > 0 Then
                            nOption = nOption - 1
                        Else
                            nOption = 3
                        End If
                        
                    
                    Case vbKeyDown
                        If nOption < 3 Then
                            nOption = nOption + 1
                        Else
                            nOption = 0
                        End If
                    
                        
                    Case vbKeySpace, vbEnter
                     
                        
                        Select Case nOption
                            Case 0:
                                If g_Snd = False Then
                                    g_Snd = True
                                Else
                                    g_Snd = False
                                End If
                            Case 1:
                                '
                         
                            Case 2:
                                'If g_Muz = False Then
                                '    g_Muz = True
                                'Else
                                '    g_Muz = False
                                'End If
                            Case 3:
                            
                                g_OptionsFlag = 0
                                                        
                            Exit Function
                        End Select
                
                End Select
            
               ' Options_Screen
            End If
            If g_OptionsFlag = 2 Then
            'g_OptionsFlag = 2 - controls
                g_OptionsFlag = 0
            End If
        End If
    End If
         
         'In Game !!!
         If g_GameState = 1 Then
            If g_GOFlag = False Then
             Select Case KeyCode
                      
                Case vbKeyLeft
                    Player.DirectionX = -Player.spd
                    
                Case vbKeyRight
                    Player.DirectionX = Player.spd
                    
                Case vbKeyUp
                    Player.DirectionY = -Player.spd
                
                Case vbKeyDown
                    Player.DirectionY = Player.spd
                    
                Case vbKeyZ
                    If Player.nShield > 0 Then
                        Player.bShield = True
                        
                    Else
                        Player.bShield = False
                    End If
                Case vbKeyEscape
                    'return to menu
                   g_GameState = 0
                   g_OptionsFlag = 0
                   
                   
                   
                    
                Case vbKeySpace
                If bShoot = False Then
                    bShoot = True
                If Player.active_laser <= NUM_LASERS Then
                    If g_Snd = True Then play_sound 0
                    Las(Player.active_laser).laseracitve = True
                    Las(Player.active_laser).LaserX = Player.shRect.Left + Int((Player.Width / 2))
                    Las(Player.active_laser).LaserY = Player.shRect.top
                    Player.active_laser = Player.active_laser + 1
               
                End If
                End If
            
            End Select
            Else
                If KeyCode = vbKeyEscape Then
                    'return to menu
                    g_GameState = 0
                    g_OptionsFlag = 0
                End If
            End If
     
        End If

End Function
