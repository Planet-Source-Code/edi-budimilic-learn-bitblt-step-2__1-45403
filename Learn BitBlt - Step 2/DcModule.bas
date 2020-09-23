Attribute VB_Name = "DcModule"
Option Explicit
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


'Dc Pictures
'Here I create variables that are needed to put pictures in them
Public Type DcPicPT
    Player As Long      'Player
    PlayerMask As Long  'Player Mask
    Teren As Long       'Soccer field as background
End Type
Public DcPic As DcPicPT


Public Sub LoadDC()
'Load Pictures
'Set in which variable will go which picture
' It's recommended for picture to be a .bmp uncompressed format which makes
' things faster to process. You can use .jpg pictures if you want (due small file size)
' but it's not good if you planing to make a big game
PicToDc DcPic.Player, App.Path, "Pictures\Player.bmp"
PicToDc DcPic.PlayerMask, App.Path, "Pictures\PlayerMask.bmp"
PicToDc DcPic.Teren, App.Path, "Pictures\Teren.bmp"
End Sub




Public Sub DestroyDC()
'Empty memory (DC)
'Everything you create it must be deleted or your RAM will fill up...
DeleteDC DcPic.Player: DeleteObject DcPic.Player
DeleteDC DcPic.PlayerMask: DeleteObject DcPic.PlayerMask
DeleteDC DcPic.Teren: DeleteObject DcPic.Teren
End Sub
'This Sub is to easier add a pictures in memory - It's used in LoadDC()
Public Sub PicToDc(ByRef SrcDC, hDir, hFile)
Screen.ActiveForm.Picture = LoadPicture(hDir & "\" & hFile)  'Load Pic from File
SrcDC = CreateCompatibleDC(Screen.ActiveForm.hdc)           'Create DC to hold stage
SelectObject SrcDC, Screen.ActiveForm.Picture               'Select bitmap in DC
End Sub
