VERSION 5.00
Begin VB.Form frmAni01 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation #1"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   3000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraHolder 
      Caption         =   "Rotate"
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   3015
      Begin VB.OptionButton optAxis 
         Caption         =   "X-axis"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "Y-axis"
         Height          =   195
         Index           =   1
         Left            =   1020
         TabIndex        =   3
         Top             =   300
         Width           =   855
      End
      Begin VB.OptionButton optAxis 
         Caption         =   "Z-axis"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   1920
         TabIndex        =   2
         Top             =   300
         Width           =   855
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
      TabIndex        =   0
      Top             =   0
      Width           =   3015
      Begin VB.Timer tmrAni 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1260
         Top             =   1080
      End
   End
End
Attribute VB_Name = "frmAni01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const RADIUS = 70

Private Type PolarGraphInfo
    X     As Single
    Y     As Single
    Z     As Single
    Color As Long
End Type
    
Dim cx    As Integer
Dim cy    As Integer
Dim PG(1 To 4) As PolarGraphInfo

Private Sub Render()
    Dim i      As Integer
    Dim px     As Single
    Dim py     As Single
    Dim pz     As Single
    Dim vx     As Single
    Dim vy     As Single
    Dim curpos As Integer
    
    picViewer.Cls
    For curpos = LBound(PG()) To UBound(PG())
        For i = 0 To 360 Step 15
            px = RADIUS * Cos(Radians(i))
            py = RADIUS * Sin(Radians(i))
            pz = 0
            
            vx = px * Cos(Radians(PG(curpos).X)) + _
                 py * Cos(Radians(PG(curpos).Y)) + _
                 pz * Cos(Radians(PG(curpos).Z))
            vy = px * Sin(Radians(PG(curpos).X)) + _
                 py * Sin(Radians(PG(curpos).Y)) + _
                 pz * Sin(Radians(PG(curpos).Z))
                 
            If i = 0 Then
                picViewer.CurrentX = cx + vx
                picViewer.CurrentY = cy - vy
            End If
                    
            picViewer.ForeColor = PG(curpos).Color
            picViewer.Line (cx, cy)-(cx + vx, cy - vy)
            
            picViewer.ForeColor = PG((curpos + 1) Mod UBound(PG()) + 1).Color
            picViewer.FillColor = PG((curpos + 1) Mod UBound(PG()) + 1).Color
            picViewer.Circle (cx + vx, cy - vy), 2
        Next i
    Next curpos
End Sub

Private Sub Form_Load()
    cx = picViewer.ScaleWidth / 2
    cy = picViewer.ScaleHeight / 2
    
    PG(1).X = 0
    PG(1).Y = 90
    PG(1).Z = 0
    PG(1).Color = vbRed
    
    PG(2).X = 90
    PG(2).Y = 90
    PG(2).Z = 0
    PG(2).Color = vbYellow
    
    PG(3).X = 45
    PG(3).Y = 90
    PG(3).Z = 0
    PG(3).Color = vbGreen
    
    PG(4).X = 135
    PG(4).Y = 90
    PG(4).Z = 0
    PG(4).Color = vbCyan
    
    tmrAni.Enabled = True
End Sub

Private Sub tmrAni_Timer()
    If optAxis(0).Value Then
        PG(1).X = (PG(1).X + 1) Mod 360
        PG(2).X = (PG(2).X + 1) Mod 360
        PG(3).X = (PG(3).X + 1) Mod 360
        PG(4).X = (PG(4).X + 1) Mod 360
    ElseIf optAxis(1).Value Then
        PG(1).Y = (PG(1).Y + 1) Mod 360
        PG(2).Y = (PG(2).Y + 1) Mod 360
        PG(3).Y = (PG(3).Y + 1) Mod 360
        PG(4).Y = (PG(4).Y + 1) Mod 360
    Else
        PG(1).Z = (PG(1).Z + 1) Mod 360
        PG(2).Z = (PG(2).Z + 1) Mod 360
        PG(3).Z = (PG(3).Z + 1) Mod 360
        PG(4).Z = (PG(4).Z + 1) Mod 360
    End If
        
    Call Render
End Sub
