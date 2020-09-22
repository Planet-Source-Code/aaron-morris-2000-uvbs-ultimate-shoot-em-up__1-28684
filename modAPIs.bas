Attribute VB_Name = "modAPIs"
Public bShoot As Boolean
Public g_Snd As Boolean ' Turn sounds on or off
'Public g_Muz As Boolean 'Flag 4 music

'--------------------C API's---------------------

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020            ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086           ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6             ' (DWORD) dest = source AND dest


Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function QueryPerformanceFrequencyAny Lib "kernel32" Alias "QueryPerformanceFrequency" (lpFrequency As Any) As Long
Public Declare Function QueryPerformanceCounterAny Lib "kernel32" Alias "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Public Type RECT
        Left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
