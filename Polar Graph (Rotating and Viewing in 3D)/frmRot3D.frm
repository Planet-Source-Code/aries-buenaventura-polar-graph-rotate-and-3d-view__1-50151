VERSION 5.00
Begin VB.Form frmRot3D 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polar Graph"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto"
      Height          =   285
      Index           =   0
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   495
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto"
      Height          =   285
      Index           =   1
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4320
      Width           =   615
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
      Index           =   1
      Left            =   2340
      TabIndex        =   18
      Top             =   4020
      Width           =   615
   End
   Begin VB.Frame fraHolder 
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   15
      Top             =   3420
      Width           =   3015
      Begin VB.OptionButton optDir 
         Caption         =   "Clockwise"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   17
         Top             =   180
         Width           =   1035
      End
      Begin VB.OptionButton optDir 
         Caption         =   "Counter-Clockwise"
         Height          =   195
         Index           =   1
         Left            =   1200
         TabIndex        =   16
         Top             =   180
         Value           =   -1  'True
         Width           =   1755
      End
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
      Index           =   3
      Left            =   660
      TabIndex        =   13
      Text            =   "0"
      Top             =   3120
      Width           =   1275
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3075
      Left            =   0
      ScaleHeight     =   201
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   9
      Top             =   0
      Width           =   3015
      Begin VB.Timer tmrAni 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1260
         Top             =   1080
      End
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
      Height          =   315
      Index           =   0
      Left            =   660
      TabIndex        =   8
      Text            =   "0"
      Top             =   4020
      Width           =   1275
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
      Height          =   315
      Index           =   1
      Left            =   660
      TabIndex        =   7
      Text            =   "0"
      Top             =   4320
      Width           =   1275
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
      Height          =   315
      Index           =   2
      Left            =   660
      TabIndex        =   6
      Text            =   "0"
      Top             =   4620
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
      Index           =   0
      Left            =   1980
      TabIndex        =   5
      Top             =   3120
      Width           =   435
   End
   Begin VB.Frame fraHolder 
      Caption         =   "Rotate"
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   4980
      Width           =   3015
      Begin VB.OptionButton optAxis 
         Caption         =   "X-axis"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "Y-axis"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   2
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "Z-axis"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   1
         Top             =   300
         Width           =   855
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
      TabIndex        =   14
      Top             =   3120
      Width           =   600
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
      Left            =   360
      TabIndex        =   12
      Top             =   4020
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
      Left            =   360
      TabIndex        =   11
      Top             =   4320
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
      Left            =   360
      TabIndex        =   10
      Top             =   4620
      Width           =   285
   End
End
Attribute VB_Name = "frmRot3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RADIUS = 70

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
    
    For i = 0 To 2500 Step 100
        px = RADIUS * Cos(Radians(i))
        py = RADIUS * Sin(Radians(i))
        pz = 0
        
        ' rotate
        vx = px * Cos(Radians(Angle)) - _
             py * Sin(Radians(Angle))
        vy = px * Sin(Radians(Angle)) + _
             py * Cos(Radians(Angle))
        
        ' 3D
        vx = vx * Cos(Radians(x_angle)) + _
             vy * Cos(Radians(y_angle)) + _
             pz * Cos(Radians(z_angle))
        vy = vx * Sin(Radians(x_angle)) + _
             vy * Sin(Radians(y_angle)) + _
             pz * Sin(Radians(z_angle))
             
        If i = 0 Then
            picViewer.CurrentX = cx + vx
            picViewer.CurrentY = cy - vy
        End If
                
        If i Mod 2000 = 500 Then
            picViewer.ForeColor = vbRed
        ElseIf i Mod 2000 = 1000 Then
            picViewer.ForeColor = vbYellow
        ElseIf i Mod 2000 = 1500 Then
            picViewer.ForeColor = vbGreen
        End If
        
        picViewer.Line (picViewer.CurrentX, picViewer.CurrentY)- _
                       (cx + vx, cy - vy)
    Next i
End Sub

Private Sub chkAuto_Click(Index As Integer)
    If CBool(chkAuto(0).Value) Then
        tmrAni.Enabled = True
    ElseIf CBool(chkAuto(1).Value) Then
        tmrAni.Enabled = True
    Else
        tmrAni.Enabled = False
    End If
End Sub

Private Sub cmdOk_Click(Index As Integer)
    On Error GoTo ErrHandler
    
    x_angle = CSng(txtAngle(0).Text)
    y_angle = CSng(txtAngle(1).Text)
    z_angle = CSng(txtAngle(2).Text)
    
    Angle = CSng(txtAngle(3).Text)
    
    Call Render
    
ErrHandler:
End Sub

Private Sub Form_Load()
    cx = picViewer.ScaleWidth / 2
    cy = picViewer.ScaleHeight / 2
    
    x_angle = 0
    y_angle = 90
    z_angle = 0
    
    Call Render
End Sub

Private Sub tmrAni_Timer()
    If CBool(chkAuto(0).Value) Then
        If optDir(0).Value Then
            Angle = (Angle + 1) Mod 360
        Else
            Angle = (Angle - 1) Mod 360
        End If
    
        txtAngle(3).Text = Angle
    End If
    
    If CBool(chkAuto(1).Value) Then
        If optAxis(0).Value Then
            x_angle = (x_angle + 1) Mod 360
        ElseIf optAxis(1).Value Then
            y_angle = (y_angle + 1) Mod 360
        Else
            z_angle = (z_angle + 1) Mod 360
        End If
        
        txtAngle(0).Text = x_angle
        txtAngle(1).Text = y_angle
        txtAngle(2).Text = z_angle
    End If
    
    Call Render
End Sub

