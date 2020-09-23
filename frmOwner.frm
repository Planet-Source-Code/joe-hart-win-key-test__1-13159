VERSION 5.00
Begin VB.Form frmOwner 
   BackColor       =   &H00000000&
   Caption         =   "WinKey Test"
   ClientHeight    =   30
   ClientLeft      =   6390
   ClientTop       =   -660
   ClientWidth     =   2265
   Icon            =   "frmOwner.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   30
   ScaleWidth      =   2265
End
Attribute VB_Name = "frmOwner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the owner form.  It hides off the screen so it isn't visible, but allows the
'program to have an icon and respond to the WIN key.  You COULD get an icon for
'your borderless form by changing the style (with the API) but it still wouldn't have
'respond to the WIN key.

'The caption of THIS form and it's icon will show up in the taskbar, so make sure
'you change them if you plan on using this in a real project.
'
'Also note that you need to replace each occurance of frmMain with the name
'of your form.

Private m_hWndParent As Long 'holds the handle of the old parent of the main form

Private Sub Form_Load()
   Dim hMenu& 'the handle of the system menu
   
   'Get the handle to the system menu
   hMenu = GetSystemMenu(Me.hwnd, 0&)
    
    'delete inapropriate members on the system menu
   Call DeleteMenu(hMenu, SC_MAXIMIZE, MF_BYCOMMAND)
   Call DeleteMenu(hMenu, SC_SIZE, MF_BYCOMMAND)
   
   'we COULD use the move, but it is a lot more complicated and requires subclassing
   'because the MOVE moves this form and not the one the user expects to move.
   Call DeleteMenu(hMenu, SC_MOVE, MF_BYCOMMAND)
   
   'make this form the parent of the main form and get the old parent at the same time
   m_hWndParent = SetWindowLong(frmMain.hwnd, GWL_HWNDPARENT, Me.hwnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'must set the parent back to what it used to be
   Call SetWindowLong(frmMain.hwnd, GWL_HWNDPARENT, m_hWndParent)
   'and unload the main form.
   Unload frmMain
End Sub
