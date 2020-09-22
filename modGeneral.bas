Attribute VB_Name = "modGeneral"
Option Explicit

Public Const PI             As Double = 3.14159265358979
Public Const PI_OVER_180    As Double = PI / 180
Public Const C180_OVER_PI   As Double = 180 / PI

Public Const RIGHT_ANGLE    As Long = 90
Public Const HALF_CIRCLE    As Long = 180
Public Const FULL_CIRCLE    As Long = 360

Public Const R_DEFAULT_HISCORE  As String = "0"
Public Const R_ASTEROIDS        As String = "Asteroids"
Public Const R_SCORES           As String = "Scores"
Public Const R_HISCORE          As String = "HiScore"
Public Const R_HISCORER         As String = "ScoredBy"

Public Function Force360(StartValue As Long)
    'Forces a value into the 0-359 range
    Dim lWhole360 As Long
    Dim lCalculator As Long
    
    lCalculator = StartValue
    If StartValue < 0 Then
        Do Until lCalculator > 0
            lCalculator = FULL_CIRCLE + lCalculator
        Loop
    Else
        lWhole360 = Int(Abs(StartValue / FULL_CIRCLE))
        lCalculator = Abs(StartValue) - lWhole360 * FULL_CIRCLE
        If StartValue < 0 Then lCalculator = FULL_CIRCLE - lCalculator
    End If
    Force360 = lCalculator
End Function

Public Function Arcsin(X)
    'Sin^-1(x)
    Arcsin = Atn(X / Sqr(-X * X + 1))
End Function

Public Function GetBearing(FromX As Long, FromY As Long, ToX As Long, ToY As Long) As Long
    Dim dXDiff          As Double
    Dim dYDiff          As Double
    Dim dHypoteneuse    As Double
    Dim lBearing        As Long
    
    dXDiff = ToX - FromX
    dYDiff = ToY - FromY
    
    If dYDiff = 0 Then
        
        Select Case dXDiff
            Case 0
                GetBearing = 0 'Default value when FromX & Y match ToX and Y
            Case Is > 0
                GetBearing = 90
            Case Is < 0
                GetBearing = 270
        End Select
        
    ElseIf dXDiff = 0 Then
        
        'Arbitrary values
        Select Case dYDiff
            Case Is > 0
                GetBearing = HALF_CIRCLE
            Case Is < 0
                GetBearing = 0
        End Select
        
    Else
        'This bit calculates the bearing and range
        dHypoteneuse = Sqr(dXDiff ^ 2 + dYDiff ^ 2) 'Distance between two objects
        GetBearing = Arcsin(dXDiff / dHypoteneuse) * C180_OVER_PI
        
        If dXDiff > 0 And dYDiff > 0 Then GetBearing = HALF_CIRCLE - GetBearing
        If dXDiff < 0 And dYDiff > 0 Then GetBearing = HALF_CIRCLE + Abs(GetBearing)
        If dXDiff < 0 And dYDiff < 0 Then GetBearing = FULL_CIRCLE + GetBearing
    End If
    
    'Make sure the result's in the 0-359 range
    GetBearing = Force360(GetBearing)
End Function


Public Function XChange(Speed As Long, Bearing As Long) As Long
    Dim lBear As Long
    
    lBear = Force360(Bearing)
    XChange = (Sin(lBear * PI_OVER_180) * Speed)
End Function

Public Function YChange(Speed As Long, Bearing As Long) As Long
    Dim lBear As Long
    
    lBear = Force360(Bearing)
    YChange = Sin(Force360(lBear - RIGHT_ANGLE) * PI_OVER_180) * Speed
End Function

Public Function GetDistance(FromX As Long, FromY As Long, ToX As Long, ToY As Long) As Long
    Dim dXDiff          As Double
    Dim dYDiff          As Double
    Dim dHypoteneuse    As Double
    
    dXDiff = Abs(FromX - ToX)
    dYDiff = Abs(FromY - ToY)
    dHypoteneuse = Sqr(dXDiff ^ 2 + dYDiff ^ 2) 'Distance between two objects
End Function

Public Function GetDestinationX(FromX As Long, Distance As Long, Bearing As Long)
    GetDestinationX = FromX + Sin(Rads(Bearing)) * Distance
End Function

Public Function GetDestinationY(FromY As Long, Distance As Long, Bearing As Long)
    GetDestinationY = FromY + (Cos(Rads(Bearing)) * Distance) * -1
End Function

Public Function Rads(Degrees As Long) As Double
    Rads = PI_OVER_180 * Degrees
End Function
