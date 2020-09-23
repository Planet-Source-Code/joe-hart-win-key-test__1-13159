VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2685
   ClientLeft      =   4080
   ClientTop       =   2715
   ClientWidth     =   5610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   4080
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "This is the main form.  Notice that it is borderless, but it can respond to the WIN+M key.  It  also has an icon on the taskbar."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the main form.

'In this example program, it doesn't do anything, but notice that it is a borderless form...

Private Sub cmdClose_Click()
   'we unload the owner form, not this one.  Closing the owner form will also close this one.
   Unload frmOwner
End Sub

