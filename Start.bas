Attribute VB_Name = "Start"
Option Explicit
'How to have a borderless form show in the taskbar with an icon
'and respond to the WIN+M key.
'
'by Joe Hart (bghost@ti.cz)
'
'This module contains all of the declare statements and constants
'that this project needs.  Normally all of my code uses a type library
'for Windows that I wrote to eliminate most of the declare statements
'constants and types you would ever need using the API.
'you can download that type library from planet source code.

Public Const SC_MOVE = &HF010&
Public Const SC_SIZE = &HF000&
Public Const SC_MAXIMIZE = &HF030&
Public Const MF_BYCOMMAND = 0&
Public Const GWL_HWNDPARENT = -8&

Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Sub Main()
   'show the main form
   frmMain.Show
   'and show the owner.  The owner sets itself up as the parent of frmmain.  We could do all that here though.
   frmOwner.Show
End Sub
