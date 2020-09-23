VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn BitBlt Step 2             <- by EdiFreak ->"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4485
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   499
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   299
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
LoadDC   'Load graphics - put it into RAM
MainLoop 'It starts mail loop / engine
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If you want keys detection to work you must set this -> Form1.KeyPreview=True (in properties)
If KeyCode = vbKeyUp Then Keys.Up = True
If KeyCode = vbKeyDown Then Keys.Down = True
If KeyCode = vbKeyLeft Then Keys.Left = True
If KeyCode = vbKeyRight Then Keys.Right = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If you want keys detection to work you must set this -> Form1.KeyPreview=True (in properties)
If KeyCode = vbKeyUp Then Keys.Up = False
If KeyCode = vbKeyDown Then Keys.Down = False
If KeyCode = vbKeyLeft Then Keys.Left = False
If KeyCode = vbKeyRight Then Keys.Right = False

If KeyCode = vbKeyEscape Then Unload Form1 'To exit program/game
End Sub

Private Sub Form_Unload(Cancel As Integer)
DestroyDC 'Delete pictures from Memory
Unload Form1
End
End Sub
