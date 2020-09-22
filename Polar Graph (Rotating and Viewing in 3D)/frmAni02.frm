VERSION 5.00
Begin VB.Form frmAni02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation #2"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3030
   StartUpPosition =   3  'Windows Default
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
      Left            =   2400
      TabIndex        =   10
      Top             =   3120
      Width           =   615
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
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Text            =   "0"
      Top             =   3180
      Width           =   1215
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
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Text            =   "90"
      Top             =   3480
      Width           =   1215
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
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Text            =   "0"
      Top             =   3780
      Width           =   1215
   End
   Begin VB.Frame fraHolder 
      Caption         =   "Direction"
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   3015
      Begin VB.OptionButton optDir 
         Caption         =   "Counter-Clockwise"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   1755
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Clockwise"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   1035
      End
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      Height          =   3075
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Timer tmrAni 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   60
         Top             =   60
      End
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X = "
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
      Left            =   0
      TabIndex        =   9
      Top             =   3180
      Width           =   285
   End
   Begin VB.Label lblY 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Y = "
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
      Left            =   0
      TabIndex        =   8
      Top             =   3480
      Width           =   300
   End
   Begin VB.Label lblZ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z = "
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
      Left            =   0
      TabIndex        =   7
      Top             =   3780
      Width           =   285
   End
End
Attribute VB_Name = "frmAni02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cx    As Integer ' center x
Dim cy    As Integer ' center y

Dim Angle   As Single ' Angle of rotation
Dim x_angle As Single ' x angle of polar graph
Dim y_angle As Single ' y angle of polar graph
Dim z_angle As Single ' z angle of polar graph

Private Sub Render()
    Dim i   As Integer
    Dim px  As Single
    Dim py  As Single
    Dim pz  As Single
    
    Dim vx  As Single
    Dim vy  As Single
    
    picViewer.Cls
    For i = 0 To 1440
        px = 100 * Sin(0.05 * Radians(i)) * Cos(Radians(i))
        py = 100 * Sin(0.05 * Radians(i)) * Sin(Radians(i))
        pz = 0
        
        vx = px * Cos(Radians(Angle)) - _
             py * Sin(Radians(Angle))
        vy = px * Sin(Radians(Angle)) + _
             py * Cos(Radians(Angle))
             
        vx = vx * Cos(Radians(x_angle)) + _
             vy * Cos(Radians(y_angle)) + _
             pz * Cos(Radians(z_angle))
        vy = vx * Sin(Radians(x_angle)) + _
             vy * Sin(Radians(y_angle)) + _
             pz * Sin(Radians(z_angle))
             
        If i Mod 90 = 0 Then
            picViewer.ForeColor = vbRed
        ElseIf i Mod 90 = 15 Then
            picViewer.ForeColor = vbGreen
        ElseIf i Mod 90 = 30 Then
            picViewer.ForeColor = vbBlue
        ElseIf i Mod 90 = 60 Then
            picViewer.ForeColor = vbYellow
        End If
        
        If i = 0 Then
            picViewer.CurrentX = cx + vx
            picViewer.CurrentY = cy + vy
        End If
        
        picViewer.Line (picViewer.CurrentX, picViewer.CurrentY)- _
                       (cx + vx, cy - vy)
    Next i
    
    picViewer.ForeColor = vbWhite
    picViewer.FillColor = vbBlack
    picViewer.Circle (cx + vx, cx - vy), 3
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    
    x_angle = CSng(txtAngle(0).Text) Mod 360
    y_angle = CSng(txtAngle(1).Text) Mod 360
    z_angle = CSng(txtAngle(2).Text) Mod 360
End Sub

Private Sub Form_Load()
    cx = picViewer.ScaleWidth / 2
    cy = picViewer.ScaleHeight / 2
    
    x_angle = 0
    y_angle = 90
    z_angle = 0
    
    tmrAni.Enabled = True
End Sub

Private Sub tmrAni_Timer()
    If optDir(0).Value Then
        Angle = (Angle + 10) Mod 360
    Else
        Angle = (Angle - 10) Mod 360
    End If
    
    Call Render
End Sub



