VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12300
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8895
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrInvulnerable 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   165
      Top             =   1500
   End
   Begin VB.Timer tmrShotRecycle 
      Interval        =   200
      Left            =   165
      Top             =   900
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   135
      Top             =   165
   End
   Begin VB.Label lblPause 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PAUSED"
      BeginProperty Font 
         Name            =   "Optimum"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   3390
      TabIndex        =   3
      Top             =   4080
      Width           =   2025
   End
   Begin VB.Label lblLives 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Optimum"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1665
      TabIndex        =   2
      Top             =   5220
      Width           =   6825
   End
   Begin VB.Label lblHiScore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Optimum"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1215
      TabIndex        =   1
      Top             =   2745
      Width           =   6825
   End
   Begin VB.Label lblScore 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Optimum"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   1470
      TabIndex        =   0
      Top             =   495
      Width           =   6570
   End
   Begin VB.Shape Asteroid 
      BorderColor     =   &H00FFFFFF&
      Height          =   1755
      Index           =   0
      Left            =   -5000
      Shape           =   3  'Circle
      Top             =   240
      Width           =   1830
   End
   Begin VB.Line lneShot 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   5
      Index           =   0
      X1              =   -2685
      X2              =   -2850
      Y1              =   2910
      Y2              =   3315
   End
   Begin VB.Line lneBase 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   -315
      X2              =   -1635
      Y1              =   450
      Y2              =   1170
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Private MyShip          As clsShip
Private colShots        As Collection
Private colAsteroids    As Collection

Private sTurning    As String
Private bThrust     As Boolean
Private lScore      As Long
Private lLives      As Long
Private lHiScore    As Long

Private Const S_LEFT            As String = "LEFT"
Private Const S_RIGHT           As String = "RIGHT"
Private Const MIN_AST_DIAMETER  As Long = 250

Private Sub Form_Load()
    'Seed the random number generator
    Randomize Timer
        
    'Setup the form and statistics display
    SetupScreen
    
    'Set up objects and readouts
    Restart
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bPause As Boolean
    
    Select Case KeyCode
        Case vbKeyLeft
            'Set turning string flag, to be picked up by timer
            sTurning = S_LEFT
            
        Case vbKeyRight
            'Set turning string flag, to be picked up by timer
            sTurning = S_RIGHT
            
        Case vbKeyUp
            'Set thrusting flag, to be picked up by timer
            bThrust = True
            
        Case vbKeyEscape
            'End immediately
            Unload Me
            End
            
        Case vbKeyP
            bPause = Not Me.tmrMove.Enabled
            
            Me.lblPause.Visible = Not bPause
            Me.tmrMove.Enabled = bPause
            Me.tmrInvulnerable = bPause
            Me.tmrShotRecycle = bPause
         
        Case vbKeySpace
            If Not tmrShotRecycle.Enabled Then
                Dim shotNew As clsShot
                
                'Create new shot
                Set shotNew = New clsShot
                shotNew.CreateShot Me.lneShot, MyShip.X, _
                    MyShip.Y, MyShip.Bearing
                colShots.Add shotNew
                
                'Delay before allow next shot
                tmrShotRecycle.Enabled = True
            End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyLeft, vbKeyRight
            'Reset turning string flag
            sTurning = vbNullString
            
        Case vbKeyUp
            'Reset thrusting flag
            bThrust = False
            
    End Select
End Sub

Private Sub Restart()
    On Error Resume Next
    
    Dim ctlTMP As Control
    Dim astTMP As clsAsteroid
    
    Dim lLoop As Long
    Dim X As Long
    
    'Tidy up shapes and collections from previous game
    For Each ctlTMP In Me.Controls
        If TypeOf ctlTMP Is Shape Then
            Unload ctlTMP
        End If
    Next
    
    If Not colShots Is Nothing Then
        For lLoop = colShots.Count To 1 Step -1
            colShots(lLoop).ShotLine.Visible = False
            Unload colShots(lLoop).ShotLine
            colShots.Remove lLoop
        Next lLoop
    End If
    
    If Not colAsteroids Is Nothing Then
        For lLoop = colAsteroids.Count To 1 Step -1
            colAsteroids(lLoop).Shape.Visible = False
            Unload colAsteroids(lLoop).Shape 'Redundant
            colAsteroids.Remove lLoop
        Next lLoop
    End If
       
    Set colShots = New Collection
    Set colAsteroids = New Collection
    
    For lLoop = 1 To 5
        Unload Me.lneBase(lLoop)
    Next lLoop
    
    'Create ship
    Set MyShip = New clsShip
      
    MyShip.CreateShip Screen.Width / 2, Screen.Height / 2, 0, Me.lneBase, _
        Screen.Width, Screen.Height
        
    'Create asteroids
    For X = 1 To 4
        Set astTMP = New clsAsteroid
        astTMP.CreateAsteroid Me.Asteroid
        colAsteroids.Add astTMP
    Next X
    
    'Update statistics
    lLives = 3
    lScore = 0
    UpdateHiScore
    UpdateScore
    UpdateLives
    
    'Reset movement flags
    sTurning = vbNullString
    bThrust = False
    
    'Start movement again
    tmrMove.Enabled = True
End Sub

Private Sub SetupScreen()
    'Maximise the form
    Me.WindowState = vbMaximized
    
    'Hide cursor
    ShowCursor False
    
    'Setup score display
    Me.lblScore.Left = 0
    Me.lblScore.Top = 0
    
    'Setup HiScore display
    Me.lblHiScore.Left = Screen.Width - Me.lblHiScore.Width
    Me.lblHiScore.Top = 0
    
    'Setup Lives display
    Me.lblLives.Left = 0
    Me.lblLives.Top = Screen.Height - Me.lblLives.Height
    
    Me.lblPause.Left = Screen.Width / 2 - Me.lblPause.Width / 2
    Me.lblPause.Top = Screen.Height / 2 - Me.lblPause.Height / 2
    Me.lblPause.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    'Show the mouse pointer again
    ShowCursor 1
End Sub

Private Sub tmrInvulnerable_Timer()
    'Invulnerability has worn off
    Me.tmrInvulnerable.Enabled = False
End Sub

Private Sub tmrMove_Timer()
    On Error Resume Next
    
    Dim shtTMP  As clsShot
    Dim astTMP  As clsAsteroid
    
    Dim lShotCount  As Long
    Dim lAstCount   As Long
    
    'Move ship
    If sTurning = S_RIGHT Then MyShip.Bearing = MyShip.Bearing + 3
    If sTurning = S_LEFT Then MyShip.Bearing = MyShip.Bearing - 3
    
    If bThrust Then
        MyShip.Thrust
    Else
        MyShip.Move
    End If
    
    'Move Shots
    For Each shtTMP In colShots
        If shtTMP.Alive Then
             shtTMP.MoveShot
        End If
        
        'Shot/Asteroid collision
        For Each astTMP In colAsteroids
            AsteroidAndShotCollisionCheck shtTMP, astTMP
        Next
    Next
    
    'Move Asteroids
    For Each astTMP In colAsteroids
        astTMP.MoveAsteroid
        
        'Asteroid/Ship collision
        If Not Me.tmrInvulnerable.Enabled Then
            AsteroidAndShipCollisionCheck astTMP
        End If
    Next
    
    'Remove Asteroids
    For lAstCount = colAsteroids.Count To 1 Step -1
        If colAsteroids(lAstCount).Destroyed Then
            colAsteroids(lAstCount).Shape.Visible = False
            colAsteroids.Remove lAstCount
        End If
    Next lAstCount
    
    'Remove shots
    For lShotCount = colShots.Count To 1 Step -1
        If Not colShots(lShotCount).Alive Then
            colShots(lShotCount).ShotLine.Visible = False
            colShots.Remove lShotCount
        End If
    Next
    
    'Add an asteroid or two
    If colAsteroids.Count < 4 Then
        For lAstCount = 1 To Int(Rnd * 2) + 1
            Set astTMP = New clsAsteroid
            astTMP.CreateAsteroid Me.Asteroid
            colAsteroids.Add astTMP
        Next lAstCount
    End If
End Sub

Private Sub tmrShotRecycle_Timer()
    'Allow next shot
    tmrShotRecycle.Enabled = False
End Sub

Public Sub AsteroidAndShotCollisionCheck(ByRef shtCheck As clsShot, _
    ByRef astCheck As clsAsteroid)
    
    Dim lCollisionRadius As Long
    Dim lDistX As Long
    Dim lDistY As Long
    Dim lDistance As Long
    
    lCollisionRadius = shtCheck.CollisionRadius + astCheck.CollisionRadius
    
    lDistX = Abs(shtCheck.X - astCheck.X)
    
    If lDistX <= lCollisionRadius Then
    
        lDistY = Abs(shtCheck.Y - astCheck.Y)
        
        If lDistY <= lCollisionRadius Then
            
            lDistance = Sqr(lDistX ^ 2 + lDistY ^ 2)
    
            If lDistance < lCollisionRadius Then
                shtCheck.KillShotOff
                AsteroidHit astCheck, GetBearing(shtCheck.X, shtCheck.Y, _
                    astCheck.X, astCheck.Y)
                lScore = lScore + astCheck.Diameter
                UpdateScore
            End If
            
        End If
    End If
End Sub

Public Function AsteroidAndShipCollisionCheck(ByRef astCheck As clsAsteroid)
    Dim lCollisionRadius As Long
    Dim lDistX As Long
    Dim lDistY As Long
    Dim lDistance As Long

    lCollisionRadius = MyShip.CollisionRadius + astCheck.CollisionRadius
    
    lDistX = Abs(MyShip.X - astCheck.X)
    
    If lDistX <= lCollisionRadius Then
    
        lDistY = Abs(MyShip.Y - astCheck.Y)
        
        If lDistY <= lCollisionRadius Then
        
            lDistance = Sqr(lDistX ^ 2 + lDistY ^ 2)

            If lDistance < lCollisionRadius Then

                lLives = lLives - 1
                UpdateLives
                If lLives = 0 Then EndGame
                RestartShip

            End If
        End If
    End If
End Function

Private Function AsteroidHit(astHit As clsAsteroid, lFromBearing As Long) As Boolean
    Dim astNew As clsAsteroid
    
    astHit.Bearing = lFromBearing
    astHit.Diameter = astHit.Diameter / 2
    
    If astHit.Diameter < MIN_AST_DIAMETER Then
        
        astHit.Destroyed = True
        
    Else
        
        Set astNew = New clsAsteroid
        astNew.CreateAsteroid Me.Asteroid, astHit.X, astHit.Y, _
            Force360(lFromBearing + 90), astHit.Speed, astHit.Diameter
        colAsteroids.Add astNew
        
    End If
End Function

Private Sub RestartShip()
    Dim lLoop As Long
    
    For lLoop = 1 To 5
        Unload Me.lneBase(lLoop)
    Next lLoop
       
    Me.tmrInvulnerable.Enabled = True
    
    MyShip.CreateShip Me.Width / 2, Me.Height / 2, 0, Me.lneBase, _
        Screen.Width, Screen.Height
End Sub

Private Sub UpdateScore()
    Me.lblScore.Caption = Format(lScore)
End Sub

Private Sub UpdateLives()
    Me.lblLives.Caption = "Ships: " & Format(lLives)
End Sub

Private Sub UpdateHiScore()
    Dim sScorer As String
    
    lHiScore = CLng(GetSetting(R_ASTEROIDS, R_SCORES, R_HISCORE, R_DEFAULT_HISCORE))
    sScorer = GetSetting(R_ASTEROIDS, R_SCORES, R_HISCORER, "MYSTERY MACHINE")
    Me.lblHiScore.Caption = "High Score: " & sScorer & " " & Format(lHiScore)
End Sub


Private Sub EndGame()
    Dim sScoreRating As String
    Dim sScorerTMP As String
    
    tmrMove.Enabled = False
    
    Select Case lScore
        Case Is < 10000
            sScoreRating = "You very bad player! Go home and suck eggs! Never touch PC again!"
        Case Is < 20000
            sScoreRating = "Feeble - my blind aunt play this game better than you!"
        Case Is < 40000
            sScoreRating = "You merely adequate! Overall, undistinguished performance."
        Case Is < 80000
            sScoreRating = "Pretty good, come back when you ready for real game!"
        Case Is < 160000
            sScoreRating = "Classy stuff, Mr Stone Shooter. But you not the best!"
        Case Is < 320000
            sScoreRating = "You really good at this. But enlightenment is journey, not destination."
        Case Else
            sScoreRating = "YOU GOD-LIKE!"
    End Select
    MsgBox "Game Over: " & sScoreRating, vbOKOnly, "Game Over"
    
    If lScore > lHiScore Then
        sScorerTMP = InputBox("You got a new High Score! Please enter your name below:", _
            "High Score!", "Mystery Machine")
        
        SaveSetting R_ASTEROIDS, R_SCORES, R_HISCORE, Format(lScore)
        SaveSetting R_ASTEROIDS, R_SCORES, R_HISCORER, UCase(sScorerTMP)
    End If
    
    lLives = 3
    lScore = 0
    UpdateHiScore
    UpdateScore
    UpdateLives
    Restart
End Sub
