VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAsteroid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private m_Shape         As Shape
Private mvarDiameter    As Long
Private mvarX           As Long
Private mvarY           As Long
Private mvarSpeed       As Long
Private mvarBearing     As Long
Private m_bDestroyed    As Boolean

Public Sub CreateAsteroid(shpBase, Optional StartX As Long = 0, Optional StartY As Long = 0, _
    Optional StartBearing As Long = -1, Optional StartSpeed As Long = -1, _
    Optional StartDiameter As Long = 0)
    
    If StartX > 0 Then
        mvarX = StartX
    Else
        mvarX = Int(Rnd * Screen.Width)
    End If
    
    If StartY > 0 Then
        mvarY = StartY
    Else
        mvarY = Int(Rnd * Screen.Height)
    End If
        
    If StartSpeed >= 0 Then
        mvarSpeed = StartSpeed
    Else
        mvarSpeed = Int(Rnd * 40) + 10
    End If
    
    If StartBearing >= 0 Then
        mvarBearing = StartBearing
    Else
        mvarBearing = Int(Rnd * 359)
    End If
    
    If StartDiameter > 0 Then
        mvarDiameter = StartDiameter
    Else
        mvarDiameter = Int(Rnd * 1000) + 1000
    End If
    
    Load shpBase(shpBase.Count)
    Set m_Shape = shpBase(shpBase.Count - 1)
    m_Shape.Visible = True
    
    DrawAsteroid
End Sub

Public Sub MoveAsteroid()
    mvarX = GetDestinationX(mvarX, mvarSpeed, mvarBearing)
    mvarY = GetDestinationY(mvarY, mvarSpeed, mvarBearing)
    If mvarX < 0 Then mvarX = Screen.Width
    If mvarX > Screen.Width Then mvarX = 0
    If mvarY < 0 Then mvarY = Screen.Height
    If mvarY > Screen.Height Then mvarY = 0
    DrawAsteroid
End Sub

Public Sub DrawAsteroid()
    m_Shape.Width = mvarDiameter
    m_Shape.Height = mvarDiameter
    m_Shape.Left = mvarX - mvarDiameter / 2
    m_Shape.Top = mvarY - mvarDiameter / 2
End Sub

Public Property Let Bearing(ByVal vData As Long)
    mvarBearing = vData
End Property

Public Property Get Bearing() As Long
    Bearing = mvarBearing
End Property

Public Property Let Speed(ByVal vData As Long)
    mvarSpeed = vData
End Property

Public Property Get Speed() As Long
    Speed = mvarSpeed
End Property

Public Property Let Y(ByVal vData As Long)
    mvarY = vData
End Property

Public Property Get Y() As Long
    Y = mvarY
End Property

Public Property Let X(ByVal vData As Long)
    mvarX = vData
End Property

Public Property Get X() As Long
    X = mvarX
End Property

Public Property Let Diameter(ByVal vData As Long)
    mvarDiameter = vData
End Property

Public Property Get Diameter() As Long
    Diameter = mvarDiameter
End Property

Public Property Get CollisionRadius() As Long
    CollisionRadius = mvarDiameter / 2
End Property

Public Property Let Destroyed(SetDestroyed As Boolean)
    m_bDestroyed = SetDestroyed
End Property

Public Property Get Destroyed() As Boolean
    Destroyed = m_bDestroyed
End Property

Public Property Get Shape()
    Set Shape = m_Shape
End Property
