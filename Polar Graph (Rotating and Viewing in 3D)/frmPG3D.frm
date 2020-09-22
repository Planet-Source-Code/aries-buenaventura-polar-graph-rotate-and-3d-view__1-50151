VERSION 5.00
Begin VB.Form frmPG3D 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Polar Graph in 3D"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraHolder 
      Caption         =   "Rotate"
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   4020
      Width           =   3015
      Begin VB.OptionButton optAxis 
         Caption         =   "Z-axis"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   12
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "Y-axis"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "X-axis"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "Auto"
      Height          =   285
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3420
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
      Left            =   2340
      TabIndex        =   7
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
      Index           =   2
      Left            =   420
      TabIndex        =   6
      Text            =   "0"
      Top             =   3720
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
      Left            =   420
      TabIndex        =   5
      Text            =   "0"
      Top             =   3420
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
      Index           =   0
      Left            =   420
      TabIndex        =   4
      Text            =   "0"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.PictureBox picViewer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
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
         Left            =   1260
         Top             =   1500
      End
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
      Left            =   60
      TabIndex        =   3
      Top             =   3720
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
      Left            =   60
      TabIndex        =   2
      Top             =   3420
      Width           =   300
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
      Left            =   60
      TabIndex        =   1
      Top             =   3120
      Width           =   285
   End
End
Attribute VB_Name = "frmPG3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RADIUS = 70

Dim cx    As Integer ' center x
Dim cy    As Integer ' center y

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
        
        vx = px * Cos(Radians(x_angle)) + _
             py * Cos(Radians(y_angle)) + _
             pz * Cos(Radians(z_angle))
        vy = px * Sin(Radians(x_angle)) + _
             py * Sin(Radians(y_angle)) + _
             pz * Sin(Radians(z_angle))
        
        If i = 0 Then
            picViewer.CurrentX = cx + vx
            picViewer.CurrentY = cy - vy
        End If
        
        picViewer.Line (picViewer.CurrentX, picViewer.CurrentY)- _
                       (cx + vx, cy - vy), vbWhite
    Next i
End Sub

Private Sub chkAuto_Click()
    tmrAni.Enabled = CBool(chkAuto.Value)
End Sub

Private Sub cmdOk_Click()
    x_angle = CSng(txtAngle(0).Text)
    y_angle = CSng(txtAngle(1).Text)
    z_angle = CSng(txtAngle(2).Text)
    
    Call Render
End Sub

Private Sub Form_Load()
    cx = picViewer.ScaleWidth / 2
    cy = picViewer.ScaleHeight / 2
    
    Call Render
End Sub

Private Sub tmrAni_Timer()
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
    
    Call Render
End Sub
