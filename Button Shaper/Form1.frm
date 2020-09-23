VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Button Changer"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   5430
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Shape 
      AutoRedraw      =   -1  'True
      Height          =   735
      Left            =   0
      MousePointer    =   99  'Custom
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Height          =   2040
      Left            =   990
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Button Shaper"
      Top             =   1620
      Width           =   2715
   End
   Begin VB.Label Label2 
      Caption         =   "Its how you can change the command buttons"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   11.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   45
      TabIndex        =   3
      Top             =   900
      Width           =   5370
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":17384
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   225
      TabIndex        =   2
      Top             =   4410
      Width           =   4875
   End
   Begin VB.Image p2 
      Height          =   480
      Left            =   -1755
      Picture         =   "Form1.frx":17420
      Top             =   3285
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #######################################
' ##  A button changer demo            ##
' ##  by Gaurav Sharma                 ##
' ##  mail your views                  ##
' ##  fascinating_guy007@yahoo.com     ##
' ##  If you like this,please Vote me! ##
' #######################################

Function CreateShape()
Dim WindowRegion As Long
Shape.ScaleMode = vbPixels
Shape.AutoRedraw = True
Shape.AutoSize = True
Shape.BorderStyle = vbBSNone

Command1.Width = Shape.Width
Command1.Height = Shape.Height
WindowRegion = CreateREgion(Shape)
SetWindowRgn Command1.hwnd, WindowRegion, True

End Function

Private Sub Command1_Click()
MsgBox "Hope you like this button shaper"
End Sub

Private Sub Form_Load()

Shape.Picture = Command1.Picture
CreateShape
Shape.Visible = False
End Sub

