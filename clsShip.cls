VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
 Option Explicit

Private lnePortThrust   As Line
Private lneStarThrust   As Line
Private lnePort         As Line
Private lneStar         As Line
Private lneAft          As Line

Private lShip_X         As Long
Private lShip_Y         As Long
Private lShipBearing    As Long
Private lMoveBearing    As Long
Private lSpeed          As Long
Private lLimitX         As Long
Private lLimitY         As Long

Public Sub CreateShip(ShipX As Long, ShipY As Long, _
    Bearing As Long, BaseLine, LimitX As Long, LimitY As Long)
    
    Dim X As Integer
    
    For X = 1 To 5
        Load BaseLine(BaseLine.Count)
    Next X
    
    Set lnePort = BaseLine(1)
    Set lneStar = BaseLine(2)
    Set lneAft = BaseLine(3)
    Set lnePortThrust = BaseLine(4)
    Set lneStarThrust = BaseLine(5)
    
    lShip_X = ShipX
    lShip_Y = ShipY
    lLimitX = LimitX
    lLimitY = LimitY
    lShipBearing = Bearing
    
    lnePort.Visible = True
    lneStar.Visible = True
    lneAft.Visible = True
    lnePortThrust.Visible = False
    lneStarThrust.Visible = False
    
    lSpeed = 0
    
    DrawShip
End Sub

Public Property Let Speed(SetSpeed As Long)
    lSpeed = SetSpeed
End Property

Public Property Get Speed() As Long
    Speed = lSpeed
End Property

Public Sub Thrust()
    Dim lY                  As Long
    Dim lX                  As Long
    Dim lThrustX            As Long
    Dim lThrustY            As Long
    Dim lResolvedX          As Long
    Dim lResolvedY          As Long
    Dim lOldMoveBearing     As Long
    
    'Get Thrust Vectors
    lThrustX = GetDestinationX(0, 1, lShipBearing)
    lThrustY = GetDestinationY(0, 1, lShipBearing)
    
    'Get current vectors
    lX = GetDestinationX(0, lSpeed, lMoveBearing)
    lY = GetDestinationY(0, lSpeed, lMoveBearing)
    
    lResolvedX = lX + lThrustX
    lResolvedY = lY + lThrustY
    
    lOldMoveBearing = lMoveBearing
    lMoveBearing = GetBearing(0, 0, lResolvedX, lResolvedY)
    
    'Course correction
    If lMoveBearing = lOldMoveBearing And lOldMoveBearing <> lShipBearing Then
        If Force360(lShipBearing - lMoveBearing) < 180 Then
            'Right Turn
            lMoveBearing = Force360(lMoveBearing + 1)
        Else
            'Left Turn
            lMoveBearing = Force360(lMoveBearing - 1)
        End If
    End If
    
    lSpeed = Sqr(lResolvedX ^ 2 + lResolvedY ^ 2)
    If lSpeed > 100 Then lSpeed = 100
    
    Move True
End Sub

Public Sub Move(Optional bThrusting As Boolean = False)
    Static lSpeedDown As Long
    
    lShip_X = GetDestinationX(lShip_X, lSpeed, lMoveBearing)
    If lShip_X < 0 Then lShip_X = lLimitX
    If lShip_X > lLimitX Then lShip_X = 0
    
    lShip_Y = GetDestinationY(lShip_Y, lSpeed, lMoveBearing)
    If lShip_Y < 0 Then lShip_Y = lLimitY
    If lShip_Y > lLimitY Then lShip_Y = 0
    
    DrawShip bThrusting
    lSpeedDown = lSpeedDown + 1
    If lSpeedDown = 8 Then
        lSpeed = lSpeed - 1
        lSpeedDown = 0
    End If
    If lSpeed < 0 Then lSpeed = 0
End Sub

Public Sub DrawShip(Optional bThrusting As Boolean = False)
    Dim lDrawColour As Long
    
    lnePort.X1 = (200 * Cos(Rads(Force360(lShipBearing - 45))) * -1) + lShip_X
    lnePort.Y1 = (200 * Sin(Rads(Force360(lShipBearing - 45))) * -1) + lShip_Y
    lnePort.X2 = (450 * Sin(Rads(lShipBearing))) + lShip_X
    lnePort.Y2 = (450 * Cos(Rads(lShipBearing)) * -1) + lShip_Y
    
    If bThrusting Then
        lnePortThrust.X1 = (170 * Cos(Rads(Force360(lShipBearing - 75))) * -1) + lShip_X
        lnePortThrust.Y1 = (170 * Sin(Rads(Force360(lShipBearing - 75))) * -1) + lShip_Y
        lnePortThrust.X2 = (350 * Sin(Rads(Force360(lShipBearing + 180)))) + lShip_X
        lnePortThrust.Y2 = (350 * Cos(Rads(Force360(lShipBearing + 180))) * -1) + lShip_Y
        
        lneStarThrust.X1 = (170 * Cos(Rads(Force360(lShipBearing + 75)))) + lShip_X
        lneStarThrust.Y1 = (170 * Sin(Rads(Force360(lShipBearing + 75)))) + lShip_Y
        lneStarThrust.X2 = lnePortThrust.X2
        lneStarThrust.Y2 = lnePortThrust.Y2
    End If
    
    lnePortThrust.Visible = bThrusting
    lneStarThrust.Visible = bThrusting
    
    lneStar.X1 = (200 * Cos(Rads(Force360(lShipBearing + 45)))) + lShip_X
    lneStar.Y1 = (200 * Sin(Rads(Force360(lShipBearing + 45)))) + lShip_Y
    lneStar.X2 = lnePort.X2
    lneStar.Y2 = lnePort.Y2
    
    lneAft.X1 = lnePort.X1
    lneAft.Y1 = lnePort.Y1
    lneAft.X2 = lneStar.X1
    lneAft.Y2 = lneStar.Y1
    
    If Not frmMain.tmrInvulnerable.Enabled Then
        lDrawColour = 15
    Else
        lDrawColour = 12
    End If
    
    lneStar.BorderColor = QBColor(lDrawColour)
    lnePort.BorderColor = QBColor(lDrawColour)
    lneAft.BorderColor = QBColor(lDrawColour)
End Sub

Public Property Let Bearing(SetBearing As Long)
    lShipBearing = Force360(SetBearing)
    DrawShip
End Property

Public Property Get Bearing() As Long
    Bearing = lShipBearing
End Property

Public Property Get MoveBearing() As Long
    MoveBearing = lMoveBearing
End Property

Public Property Get CollisionRadius() As Long
    CollisionRadius = 300
End Property

Public Property Get X() As Long
    X = lShip_X
End Property

Public Property Get Y() As Long
    Y = lShip_Y
End Property
