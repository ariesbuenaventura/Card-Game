VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Signature   As String
Public Title       As String
Public CardSizeOp  As CardSizeConstants
Public AllowResize As Boolean
Public DataColl    As New Collection
Public TempColl    As New Collection
Public ValHolder1  As Variant
Public ValHolder2  As Variant
Public ValHolder3  As Variant
Public Score       As Long
Public StockColl   As New Collection
Public WasteColl   As New Collection
Public Speed       As Integer
Public HelpID      As Integer

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub gm_DoEvents()
    DoEvents
End Sub

Public Sub gm_Delay(Interval As Long)
    Sleep Interval
End Sub

Public Function gm_TwipsPerPixelX()
    gm_TwipsPerPixelX = Screen.TwipsPerPixelX
End Function

Public Function gm_TwipsPerPixelY()
    gm_TwipsPerPixelY = Screen.TwipsPerPixelY
End Function

Public Sub PlaySound(ID As Integer)
    BeginPlaySound ID, False
End Sub

Private Sub Class_Initialize()
    Score = 0
    ValHolder1 = 0
    ValHolder2 = 0
    ValHolder3 = 0
    Speed = 5
End Sub
