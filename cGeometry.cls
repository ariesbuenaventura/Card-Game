VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGeometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Angle As Single

Dim oldx As Single, oldy As Single, oldz As Single

Public Sub rotabouty(ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    oldx = X
    oldy = Y
    oldz = z
    
    X = (oldx * Cos(Angle)) - (oldz * Sin(Angle))
    Y = Y
    z = (oldx * Sin(Angle)) + (oldz * Cos(Angle))
End Sub

Public Sub rotaboutx(ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    oldx = X
    oldy = Y
    oldz = z
    
    X = X
    Y = (oldz * Sin(Angle)) + (oldy * Cos(Angle))
    z = (oldz * Cos(Angle)) - (oldy * Sin(Angle))
End Sub

Public Sub rotaboutz(ByRef X As Single, ByRef Y As Single, ByRef z As Single)
    oldx = X
    oldy = Y
    oldz = z
    
    X = (oldx * Sin(Angle)) + (oldy * Cos(Angle))
    Y = (oldx * Cos(Angle)) - (oldy * Sin(Angle))
    z = z
End Sub

Private Sub Class_Initialize()
    Angle = Rads(10) ' start at 10 degs.
End Sub
