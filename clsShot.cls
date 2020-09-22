VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private lneShot     As Line

Private sX          As Long
Private sY          As Long
Private lBearing    As Long
Private bAlive      As Boolean
Private bKillMe     As Boolean

Const Speed As Long = 150

Public Sub CreateShot(lneShotLoad, ShotX As Long, ShotY As Long, Bearing As Long)
    Dim lLoadOffset As Long
    
    Load lneShotLoad(lneShotLoad.Count)
    
    Set lneShot = lneShotLoad(lneShotLoad.Count - 1)
    
    sX = ShotX
    sY = ShotY
    lBearing = Bearing
    
    lneShot.Visible = True
    bAlive = True
    
    DrawShot
End Sub

Public Property Get Alive() As Boolean
    Alive = sX > -100 And sX < Screen.Width + 100 And _
        sY > -100 And sY < Screen.Height + 100 And Not bKillMe
End Property

Public Sub DrawShot()
    On Error Resume Next
    Static iQBColour As Integer
    lneShot.X1 = sX
    lneShot.X2 = sX
    lneShot.Y1 = sY
    lneShot.Y2 = sY
    iQBColour = iQBColour + 1
    If iQBColour > 15 Then iQBColour = 1
    lneShot.BorderColor = QBColor(iQBColour)
End Sub

Public Sub MoveShot()
    sX = GetDestinationX(CLng(sX), Speed, lBearing)
    sY = GetDestinationY(CLng(sY), Speed, lBearing)
    
    DrawShot
End Sub

Public Property Get ShotLine() As Line
    Set ShotLine = lneShot
End Property

Public Property Get CollisionRadius() As Long
    CollisionRadius = 10
End Property

Public Property Get X() As Long
    X = sX
End Property

Public Property Get Y() As Long
    Y = sY
End Property

Public Sub KillShotOff()
    bKillMe = True
End Sub
