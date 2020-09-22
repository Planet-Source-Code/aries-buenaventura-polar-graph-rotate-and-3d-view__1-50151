Attribute VB_Name = "modMath"
Option Explicit

Public Function PI() As Single
    PI = Atn(1) * 4
End Function

Public Function Radians(ByVal Degrees As Single)
    Radians = PI * Degrees / 180
End Function

