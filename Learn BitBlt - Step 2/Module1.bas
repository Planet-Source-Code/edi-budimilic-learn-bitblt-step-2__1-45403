Attribute VB_Name = "Module1"
'These are the API functions for use of BitBlt and other cool things
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCERASE = &H4400328
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086

'For easier keys detection
'Currently I use VB for detection but there are better ways for this, you can use API for detection or you can use DirectX...
Type KeysPT
    Up As Boolean 'Boolean gives back True or False for easier definition...
    Down As Boolean
    Left As Boolean
    Right As Boolean
End Type
Public Keys As KeysPT

'Player position
Type Player1PT
    XX As Integer 'left position
    YY As Integer 'top position
End Type
Public Player1 As Player1PT

Public Sub MainLoop()
'Starting player position
Player1.XX = 140
Player1.YY = 230
'Font color on the Form
Form1.ForeColor = vbWhite

Do 'Starting/doing loop
Form1.Cls

'First draw background picture / no mask picture needed because there is nothing to be transparent
'In that case on the end of the code we use SRCCOPY (copy from memory and drops it on the screen)
BitBlt Form1.hdc, 0, 0, 300, 500, DcPic.Teren, 0, 0, SRCCOPY

'Sub that moves our Player
MovePlayer

'Draw text on the screen using API-TextOut
TextOut Form1.hdc, 4, 4, "Use cursors for movement, try...", 32  'this no.35 is number of characters drawn on the screen
TextOut Form1.hdc, 170, 4, "and no blinking, cool ha! :)", 28       'you can get that no. with this 'Len(TextVariable)' that you don't have to count yourself every time you write something

DoEvents
Sleep 1  'Sleep/block whole process for 1 milisecond (1sec. = 1000 milliseconds )
Loop 'Continue loop
End Sub


Public Sub MovePlayer()
'Ako je forma detektirala i oznacila u varijabli Stisak neke tipke...
'onda ovo prati njezine vrijednosti i pomice igraca...

'If Form detects some key press and mark it in variable...
'then this track it's variables and moves a Player
If Keys.Up = True Then Player1.YY = Player1.YY - 1
If Keys.Down = True Then Player1.YY = Player1.YY + 1 'You can add Variable for Player Speed and change the last number with it..
If Keys.Left = True Then Player1.XX = Player1.XX - 1
If Keys.Right = True Then Player1.XX = Player1.XX + 1

'Now after we moved him lets draw his Mask with Picture
BitBlt Form1.hdc, Player1.XX, Player1.YY, 12, 12, DcPic.PlayerMask, 0, 0, SRCAND     'Maska
BitBlt Form1.hdc, Player1.XX, Player1.YY, 12, 12, DcPic.Player, 0, 0, SRCPAINT  'Slika preko
End Sub
