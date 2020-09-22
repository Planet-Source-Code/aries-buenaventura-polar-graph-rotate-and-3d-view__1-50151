VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polar Graph"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   2880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraFrame 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   2835
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test #5"
         Height          =   375
         Index           =   4
         Left            =   480
         TabIndex        =   6
         Top             =   3600
         Width           =   1875
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test #4"
         Height          =   375
         Index           =   3
         Left            =   480
         TabIndex        =   5
         Top             =   3180
         Width           =   1875
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test #3"
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   2760
         Width           =   1875
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test #2"
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   3
         Top             =   2340
         Width           =   1875
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "Test #1"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   2
         Top             =   1920
         Width           =   1875
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0000
         Height          =   1575
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' email: ravemasterharuglory@yahoo.com

Private Sub cmdTest_Click(Index As Integer)
    Select Case Index
    Case Is = 0
        frmRotate.Show vbModal
    Case Is = 1
        frmPG3D.Show vbModal
    Case Is = 2
        frmRot3D.Show vbModal
    Case Is = 3
        frmAni01.Show vbModal
    Case Is = 4
        frmAni02.Show vbModal
    End Select
End Sub
