VERSION 5.00
Begin VB.Form frmShoot 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UVBS By A.Morris - V2.1"
   ClientHeight    =   7200
   ClientLeft      =   7830
   ClientTop       =   5865
   ClientWidth     =   9600
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Impact"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   Begin VB.PictureBox elb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3480
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox fontset 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1560
      ScaleHeight     =   47
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   63
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Buffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   600
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox Ship 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2400
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   40
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
End
Attribute VB_Name = "frmShoot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    'Set the tiles scrolling behind the menu
    scroll_menu
End Sub

'VB Standard key event handler
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Call the keyhandler function
    KeyEvents KeyCode
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
        'set key flag ... So the player does not hold the space bar down to fire
        If g_GameState = 1 Then
        Select Case KeyCode
            Case vbKeySpace
                bShoot = False
            Case vbKeyZ
                Player.bShield = False
        End Select
        End If
End Sub
'Set up some flags on loading
Private Sub Form_Load()
    'My custom bitmaping fonts
    load_font_file
    Load_Sounds
    Load_Map
    g_Snd = True
    g_GameState = 0     'Start
    Ship.Picture = LoadPicture(App.Path + "\bits\bit.bmp")
    fontset.Picture = LoadPicture(App.Path + "\bits\fonts.bmp")
   
End Sub
'SUB - Shut down all
Private Sub Form_Unload(Cancel As Integer)
    Kill_Game
End Sub

