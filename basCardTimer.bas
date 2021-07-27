Attribute VB_Name = "basCardTimer"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const TimerAniID = 100
Public Const TimerHotTrackingID = 200

Public Sub TimerCallBack(ByVal hWnd As Long, ByVal uMsg As Long, _
                         ByVal idEvent As Long, ByVal dwTime As Long)
    
    Dim thisCard As Card
        
    CopyMemory thisCard, GetProp(hWnd, "ClassID"), &H4
    If idEvent = TimerAniID Then
        thisCard.TimerAniUpdate
    Else
        thisCard.TimerHotTrackingUpdate
    End If
    CopyMemory thisCard, &H0, &H4
End Sub
