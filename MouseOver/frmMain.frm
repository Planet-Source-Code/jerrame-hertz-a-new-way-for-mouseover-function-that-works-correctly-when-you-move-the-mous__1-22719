VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "MouseOver"
   ClientHeight    =   2985
   ClientLeft      =   7710
   ClientTop       =   2655
   ClientWidth     =   5865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2985
   ScaleWidth      =   5865
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   855
      Left            =   3480
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2400
      Top             =   1440
   End
   Begin VB.Label Label3 
      Caption         =   "PictureBox"
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Label"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Timer1_Timer()
    
    ' FormName, ControlName, Control.Left In Pixels, Control.Right In Pixels, Control.Top In Pixels, Control.Bottom In Pixels
    MouseMove Me, Label2, 27, 76, 63, 80
    MouseMove Me, Picture1, 236, 332, 64, 119
    ReturnPixels Me, Label1
    
End Sub
