VERSION 5.00
Begin VB.Form frmRotate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rotating Polar Graph"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   3015
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto"
      Height          =   285
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Width           =   495
   End
   Begin VB.TextBox txtAngle 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "0"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Top             =   3120
      Width           =   435
   End
   Begin VB.Frame fraHolder 
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   3015
      Begin VB.OptionButton optDir 
         Caption         =   "Counter-Clockwise"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Clockwise"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   3075
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      Begin VB.Timer tmrAni 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   60
         Top             =   60
      End
   End
   Begin VB.Label lblAngle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Angle = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   60
      TabIndex        =   6
      Top             =   3120
      Width           =   600
   End
End
Attribute VB_Name = "frmRotate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RADIUS = 70

Dim cx    As Integer ' center x
Dim cy    As Integer ' center y
Dim Angle As Single  ' angle of polar graph

Private Sub Render()
    Dim i   As Integer
    Dim px  As Single
    Dim py  As Single
    Dim vx  As Single
    Dim vy  As Single
    
    picViewer.Cls
    
    For i = 0 To 2500 Step 100
        px = RADIUS * Cos(Radians(i))
        py = RADIUS * Sin(Radians(i))
        
        vx = px * Cos(Radians(Angle)) - py * Sin(Radians(Angle))
        vy = px * Sin(Radians(Angle)) + py * Cos(Radians(Angle))
        
        If i = 0 Then
            picViewer.CurrentX = cx + vx
            picViewer.CurrentY = cy - vy
        End If
        
        picViewer.Line (picViewer.CurrentX, picViewer.CurrentY)- _
                       (cx + vx, cy - vy), vbWhite
    Next i
End Sub

Private Sub cmdOk_Click()
    If IsNumeric(txtAngle.Text) Then
        Angle = CSng(txtAngle.Text)
        Call Render
    End If
End Sub

Private Sub Form_Load()
    cx = picViewer.ScaleWidth / 2
    cy = picViewer.ScaleHeight / 2
    
    Call Render
End Sub

Private Sub chkAuto_Click()
    tmrAni.Enabled = CBool(chkAuto.Value)
End Sub

Private Sub tmrAni_Timer()
    If optDir(0).Value Then
        Angle = (Angle + 1) Mod 360
    Else
        Angle = (Angle - 1) Mod 360
    End If
    
    txtAngle.Text = Angle
    
    Call Render
End Sub

