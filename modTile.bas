Attribute VB_Name = "modTile"
Option Explicit
'----Screen Size
Public Const SCREENWIDTH = 640
Public Const SCREENHEIGHT = 480
Private Const TILE_SIZE_SQ As Integer = 64
Private Const TILES_ACROSS As Integer = SCREENWIDTH / TILE_SIZE_SQ
Private Const TILES_UP As Integer = SCREENHEIGHT / TILE_SIZE_SQ
Private Const TOTAL_TILESX As Integer = TILES_ACROSS
Private Const TOTAL_TILESY As Integer = (TILES_UP * 99)
Private Const MAP_SPEED = -1 'Speed of the scrolling map

Public nCurrentTileID As Integer
Public Map(TOTAL_TILESX, TOTAL_TILESY) As Integer



Private CurrentMapY As Integer
Private CurrentYLevel As Integer

Private Type tilez
    FileX As Integer
    FileY As Integer
    ScreenX As Integer
    ScreenY As Integer
End Type

Private bitmap_buffer As Object
Private tls(30) As tilez 'Tile bmp pointer

Public Function Load_Map()
    On Error GoTo UVBS_Error    'Handle Errors

    Open App.Path & "\data\level1.dat" For Input As #1

    Dim temp1 As String
    Dim temp2 As String
    Dim results As String
    Dim spos As Integer
    Dim lnt As Integer
    Dim XMap As Integer
    Dim YMap As Integer
    Dim C As Integer
    Dim x As Integer
    Dim y As Integer
    
    
    XMap = 0
    YMap = 0
    
    Do Until EOF(1)
        Input #1, temp1
        spos = 1
        lnt = 1
        XMap = 0
               
        For C = 1 To Len(temp1) - 1
                  temp2 = Mid(Trim(temp1), C, 1)
                                                      
                  If temp2 = "." Then
                      lnt = (C - spos)
                      results = Mid$(temp1, spos, lnt)
                      Map(XMap, YMap) = Val(results)
                      spos = C + 1
                      XMap = XMap + 1
                  End If
                  
        Next C
        
        YMap = YMap + 1
    Loop
    
    Close #1

    'Load the array with x,y positions of where the bitmaps are in the file!!!
    
' TO DO ----- Opt'tions
'    C = 0
    
'    For y = 0 To 200 Step TILE_SIZE_SQ
'        For x = 0 To 399 Step TILE_SIZE_SQ
'            tls(C).FileX = x
'            tls(C).FileY = y
'            C = C + 1
'        Next x
'   Next y
    
    tls(0).FileX = 0
    tls(0).FileY = 0
    tls(1).FileX = 64
    tls(1).FileY = 0
    tls(2).FileX = 128
    tls(2).FileY = 0
    tls(3).FileX = 192
    tls(3).FileY = 0
    tls(4).FileX = 256
    tls(4).FileY = 0
    tls(5).FileX = 320
    tls(5).FileY = 0
    tls(6).FileX = 128
    tls(6).FileY = 128
    tls(7).FileX = 0
    tls(7).FileY = 65
    tls(8).FileX = 65
    tls(8).FileY = 65
    tls(9).FileX = 128
    tls(9).FileY = 65
    tls(10).FileX = 192
    tls(10).FileY = 65
    tls(11).FileX = 256
    tls(11).FileY = 65
    tls(12).FileX = 320
    tls(12).FileY = 65
    tls(13).FileX = 0
    tls(13).FileY = 129
    tls(14).FileX = 65
    tls(14).FileY = 129
  
    CurrentMapY = -64


    'Load bitmap to memory
    Set bitmap_buffer = frmShoot.Buffer
    bitmap_buffer.Picture = LoadPicture(App.Path & "\bits\map.bmp")
    
    Exit Function
    
UVBS_Error:
    MsgBox "An error occurred while loading the map file!", 0, "Error"
    End
    
End Function

Public Function Reset_Tiles()
    nCurrentTileID = 0
End Function
    
Public Sub Draw_Tiles()
   prepare_tiles
  'Call BitBlt(frmShoot.hdc, 0, 0, 640, 480, BK_Surface.hdc, 0, 0, vbSrcPaint)
End Sub

Private Sub prepare_tiles()
    
    Dim XGraph As Integer
    Dim YGraph As Integer
    Dim y As Integer
    Dim x As Integer
    
    XGraph = 0
    YGraph = 0
    
    For y = TILES_UP + nCurrentTileID To nCurrentTileID Step MAP_SPEED
        For x = 0 To TILES_ACROSS
            
     
            Call BitBlt(frmShoot.hdc, XGraph, YGraph + CurrentMapY, TILE_SIZE_SQ, TILE_SIZE_SQ, bitmap_buffer.hdc, tls(Map(x, y)).FileX, tls(Map(x, y)).FileY, vbSrcPaint)
            XGraph = XGraph + TILE_SIZE_SQ
            'font_to_screen 435, 10, CStr(nCurrentTileID)
            
        Next x
        
            YGraph = YGraph + TILE_SIZE_SQ
            XGraph = 0
    Next y

    'Set current map position to next block
        
    CurrentMapY = CurrentMapY + 1
    
    If (CurrentMapY = 0) Then
        CurrentMapY = -64
        nCurrentTileID = nCurrentTileID + 1
        If nCurrentTileID = TOTAL_TILESY Then nCurrentTileID = 0
    End If

End Sub

    
