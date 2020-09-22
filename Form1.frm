VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Scroll A picture box inside another."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   2175
      Left            =   4080
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   3975
   End
   Begin VB.PictureBox picOuter 
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2115
      ScaleWidth      =   3915
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox picInner 
         BackColor       =   &H00000000&
         Height          =   7575
         Left            =   0
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   7515
         ScaleWidth      =   6315
         TabIndex        =   1
         Top             =   0
         Width           =   6375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim S As cPicScroll  ' reference the the class file.

Private Sub cmdAbout_Click()
   Dim Msg As String, Cap As String
   Cap = "gh0ul@hotmail.com"
   Msg = " Scroll a picture box of any size inside" & vbCrLf
   Msg = Msg & " another much smaller one using this VB5 " & vbCrLf
   Msg = Msg & " Class Module. "
   MsgBox Msg, , Cap
End Sub

Private Sub cmdExit_Click()
   End
End Sub

Private Sub Form_Initialize()
    ' create scrolling picturebox object.
    Set S = New cPicScroll
    
    ' set up the Scrollbar Properties
    S.SetUpScrollBars HScroll1, VScroll1, picInner.Height - picOuter.Height, 80, picInner.Height / 3, _
                      picInner.Width - picOuter.Width, 40, picInner.Width / 2
End Sub


Private Sub HScroll1_Change()
    S.MoveH HScroll1, picInner
End Sub

Private Sub VScroll1_Change()
    S.MoveV VScroll1, picInner
End Sub
