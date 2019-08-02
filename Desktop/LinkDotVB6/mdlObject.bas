Attribute VB_Name = "mdlObject"
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

'TYPE IDs:
Global Const PLAYER = 0
Global Const DOT = 2
Global Const LINK = 3
Global Const ENEMY1 = 5
Global Const SLASH = 8
Global Const STARPLATINUM = 9
Global Const SPHITBOXOBJ = 10
Global Const BEACON = 11
Global Const PORTALOBJ = 12

Global Const PI = 3.14159265358979

Global Const MAXENEMYKNOCKBACKSTUN = 50

Global Const SPELLANGLETHRESHOLD = 15

Global Const MAXDOTRANGE = 238

Global Const PLAYERWIDTH = 32
Global Const PLAYERHEIGHT = 32

Const ENEMYSPEED1 = 20

Global PlayerX As Long
Global PlayerY As Long

Global MouseX As Long
Global MouseY As Long

Global Running As Boolean

Global ActiveLink As Boolean

Global LinkDots() As GameObject

'Spell Angle Arrays
Dim StarPlatinumAngle(4) As Single
Dim PillarAngle(5) As Single

'Spell Dimension Arrays
Dim StarPlatinumDimension(4) As Single
Dim PillarDimension(5) As Single

Public Type Point
    X As Long
    Y As Long
End Type

Public Type EuclideanLine
    Point1 As Point
    Point2 As Point
End Type

Public Type EuclideanRectangle
    Points(1 To 4) As Point
End Type

Public Type GameObject
    ID As Integer
    TypeID As Integer
    LinkID As Double
    X As Long
    Y As Long
    DataTag As String * 50
    CustomDraw As Boolean
    Removed As Boolean
    SpriteID As Integer
    SpriteFrame As Long
    SpriteStrip As Long
    KnockBackX As Double
    KnockBackY As Double
    AnimationSpeed As Long
    Width As Integer
    Height As Integer
    OnCreate As Boolean
End Type

Public Type SpellCorrect
    Theta As Double
    Length As Double
End Type

Public Function CreateLine(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As EuclideanLine
    Dim Point1 As Point
    Dim Point2 As Point
    Dim ReturnLine As EuclideanLine
    
    Point1.X = X1
    Point1.Y = Y1
    Point2.X = X2
    Point2.Y = Y2
    
    ReturnLine.Point1 = Point1
    ReturnLine.Point2 = Point2
    
    CreateLine = ReturnLine
End Function

Public Function ToDegrees(ByVal Rad As Single) As Double
    ToDegrees = (Rad * 180) / PI
End Function

Public Function ToRadians(ByVal Degrees As Single) As Double
    ToRadians = (Degrees * PI) / 180
End Function

Public Sub ObjectModuleInit()
    StarPlatinumAngle(0) = 90
    StarPlatinumAngle(1) = 90
    StarPlatinumAngle(2) = 90
    StarPlatinumAngle(3) = 90
    
    StarPlatinumDimension(0) = 1
    StarPlatinumDimension(1) = 1
    StarPlatinumDimension(2) = 1
    StarPlatinumDimension(3) = 1
    
    PillarAngle(0) = 108
    PillarAngle(1) = 108
    PillarAngle(2) = 108
    PillarAngle(3) = 108
    PillarAngle(4) = 108
    
    PillarDimension(0) = 1
    PillarDimension(1) = 1
    PillarDimension(2) = 1
    PillarDimension(3) = 1
    PillarDimension(4) = 1
End Sub

Public Function GetTheta(A As Point, B As Point, C As Point) As Double 'Returns angle <ABC
    Dim ThetaPrimitive As Double
    
    Dim BA As Point
    Dim BC As Point

    Dim DotProduct As Double
    
    BA.X = A.X - B.X
    BA.Y = A.Y - B.Y
    
    BC.X = C.X - B.X
    BC.Y = C.Y - B.Y
    
    DotProduct = BA.X * BC.X + BA.Y * BC.Y
    
    If (Abs(DotProduct) > 0.1 And (Sqr(BA.X ^ 2 + BA.Y ^ 2) * Sqr(BC.X ^ 2 + BC.Y ^ 2)) ^ 2 - DotProduct ^ 2 > 0) Then
        ThetaPrimitive = ToDegrees(Atn(Sqr(((Sqr(BA.X ^ 2 + BA.Y ^ 2) * Sqr(BC.X ^ 2 + BC.Y ^ 2)) ^ 2 - DotProduct ^ 2)) / DotProduct))
    Else
        ThetaPrimitive = 90
    End If
    'VB6 Not having inverse cos is making me revisit identities...
    'As if actually having to use vectors wasn't good enough
    'For now, just use primitive theta
    
    If (ThetaPrimitive < 0) Then
        ThetaPrimitive = 180 - Abs(ThetaPrimitive)
    End If
    
    GetTheta = ThetaPrimitive
End Function

Public Function OppositeDirection(ByVal Direction As String) As String
    Select Case Direction
        Case "U"
            OppositeDirection = "D"
        Case "D"
            OppositeDirection = "U"
        Case "R"
            OppositeDirection = "L"
        Case "L"
            OppositeDirection = "R"
    End Select
End Function

Public Function BackTrack(ByVal Direction As String) As Boolean
    If (CurrentRoom.Spot > 1) Then
        If (VBA.Mid$(CurrentRoom.Path, CurrentRoom.Spot - 1, 1) = OppositeDirection(Direction)) Then
            CurrentRoom.Spot = CurrentRoom.Spot - 1
            BackTrack = True
        Else
            BackTrack = False
        End If
    Else
        BackTrack = False
    End If
End Function

Public Sub Create(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Object.LinkID = Clock
    Select Case TypeID
        Case PLAYER
            Object.SpriteID = WIZARD
            Object.SpriteFrame = 0
            Object.DataTag = "000|500|300"
        Case DOT
            Object.SpriteID = BASICDOT
            Object.SpriteFrame = 0
        Case LINK
            Object.CustomDraw = True
        Case BEACON
            Object.SpriteID = BEACONSPRITE
        Case ENEMY1
            Object.SpriteID = ENEMYWILLO
            Object.SpriteFrame = 0
            Object.DataTag = "300"
        Case SPHITBOXOBJ
            Object.SpriteID = SPHITBOXSPRITE
        Case PORTALOBJ
            Object.SpriteID = PORTALSPRITE
        Case STARPLATINUM
            Object.SpriteID = STARPLATINUMSPRITE
            Object.SpriteFrame = 0
           
            Dim Theta As Single
            Theta = VBA.Val(Object.DataTag)
            If (Theta > 45 And Theta < 135) Then
                Object.SpriteStrip = 3
            ElseIf (Theta >= 135 And Theta <= 225) Then
                Object.SpriteStrip = 1
            ElseIf (Theta > 225 And Theta <= 315) Then
                Object.SpriteStrip = 0
            Else
                Object.SpriteStrip = 2
            End If
            
        Case SLASH
            Object.SpriteID = SLASHSPRITE
            Object.SpriteFrame = 0
            If (MouseAngle >= 337.5 Or MouseAngle < 22.5) Then
                Object.SpriteStrip = 0
                'Object.X = Object.X
            ElseIf (MouseAngle >= 22.5 And MouseAngle < 67.5) Then
                Object.SpriteStrip = 5
                Object.X = Object.X + GetSprite(SLASHSPRITE, 5).FrameWidth - ToTwips(50)
                Object.Y = Object.Y - GetSprite(SLASHSPRITE, 5).FrameHeight + ToTwips(32)
            ElseIf (MouseAngle >= 67.5 And MouseAngle <= 112.5) Then
                Object.SpriteStrip = 2
                Object.Y = Object.Y - ToTwips(32)
            ElseIf (MouseAngle > 112.5 And MouseAngle < 157.5) Then
                Object.SpriteStrip = 4
                Object.X = Object.X - GetSprite(SLASHSPRITE, 5).FrameWidth + ToTwips(32)
                Object.Y = Object.Y - GetSprite(SLASHSPRITE, 5).FrameHeight + ToTwips(32)
            ElseIf (MouseAngle >= 157.5 And MouseAngle < 202.5) Then
                Object.SpriteStrip = 1
                Object.X = Object.X - GetSprite(SLASHSPRITE, 1).FrameWidth + ToTwips(32)
            ElseIf (MouseAngle >= 202.5 And MouseAngle < 247.5) Then
                Object.SpriteStrip = 6
                Object.X = Object.X + GetSprite(SLASHSPRITE, 5).FrameWidth - ToTwips(80)
                Object.Y = Object.Y - GetSprite(SLASHSPRITE, 5).FrameHeight + ToTwips(60)
            ElseIf (MouseAngle >= 247.5 And MouseAngle < 292.5) Then
                Object.SpriteStrip = 3
                Object.Y = Object.Y
            ElseIf (MouseAngle >= 292.5 And MouseAngle < 337.5) Then
                Object.SpriteStrip = 7
            End If
    End Select
    Object.TypeID = TypeID
    AddObject Object
End Sub

Public Sub CreateQueue(Object As GameObject, ByVal X As Long, ByVal Y As Long, ByVal TypeID As Integer)
    Object.X = X
    Object.Y = Y
    Select Case TypeID
        Case PLAYER
            Object.SpriteID = PLAYER
            Object.SpriteFrame = 0
    End Select
    Object.TypeID = TypeID
    AddObjectQueue Object
End Sub

Public Function GetLines(ByVal X As Long, ByVal Y As Long, ByVal Width As Long, ByVal Height As Long) As EuclideanLine()
    Dim Lines(4) As EuclideanLine
    
    With Lines(0)
        .Point1.X = X
        .Point1.Y = Y
        .Point2.X = X
        .Point2.Y = Y + Height
    End With
    
    With Lines(1)
        .Point1.X = X + Width
        .Point1.Y = Y
        .Point2.X = X + Width
        .Point2.Y = Y + Height
    End With
    
    With Lines(2)
        .Point1.X = X
        .Point1.Y = Y
        .Point2.X = X + Width
        .Point2.Y = Y
    End With
    
    With Lines(3)
        .Point1.X = X
        .Point1.Y = Y + Height
        .Point2.X = X + Width
        .Point2.Y = Y + Height
    End With
    
    GetLines = Lines
End Function

Public Function CollisionCheck(Object1 As GameObject, Object2 As GameObject) As Boolean
    CollisionCheck = (Abs(Object1.X - Object2.X) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteStrip).FrameWidth + _
                    GetSprite(Object2.SpriteID, Object2.SpriteStrip).FrameWidth)) And _
                    (Abs(Object1.Y - Object2.Y) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteStrip).FrameWidth + GetSprite(Object2.SpriteID, Object2.SpriteStrip).FrameHeight))
End Function

'Where the second object does not have an image
Public Function CollisionCheckImg1(Object1 As GameObject, Object2 As GameObject) As Boolean
    CollisionCheckImg1 = (Abs(Object1.X - Object2.X) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteStrip).FrameWidth + _
                    GetSprite(Object2.SpriteID, Object2.SpriteStrip).FrameWidth)) And _
                    (Abs(Object1.Y - Object2.Y) * 2 < (GetSprite(Object1.SpriteID, Object1.SpriteStrip).FrameWidth + GetSprite(Object2.SpriteID, Object2.SpriteStrip).FrameHeight))
End Function

Public Function CollisionCheckLineLine(Line1 As EuclideanLine, Line2 As EuclideanLine) As Boolean
    Dim UA As Double
    Dim UB As Double
    
    If (((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                    - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y)) <> 0 _
                    And ((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                        - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y)) <> 0) Then
    
        UA = ((Line2.Point2.X - Line2.Point1.X) * (Line1.Point1.Y - Line2.Point1.Y) _
                - (Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point1.X - Line2.Point1.X)) _
                    / ((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                        - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y))
                        
                        
        UB = ((Line1.Point2.X - Line1.Point1.X) * (Line1.Point1.Y - Line2.Point1.Y) _
                - (Line1.Point2.Y - Line1.Point1.Y) * (Line1.Point1.X - Line2.Point1.X)) _
                    / ((Line2.Point2.Y - Line2.Point1.Y) * (Line1.Point2.X - Line1.Point1.X) _
                        - (Line2.Point2.X - Line2.Point1.X) * (Line1.Point2.Y - Line1.Point1.Y))
                        
        If (UA >= 0 And UA <= 1 And UB >= 0 And UB <= 1) Then
            CollisionCheckLineLine = True
        Else
            CollisionCheckLineLine = False
        End If
    
    End If
End Function

Public Function CollisionCheckObjectLine(ObjectTest As GameObject, LineTest As EuclideanLine) As Boolean
    Dim Left As Boolean
    Dim Right As Boolean
    Dim Up As Boolean
    Dim Down As Boolean
    
    Dim ObjectLine1 As EuclideanLine
    Dim ObjectLine2 As EuclideanLine
    Dim ObjectLine3 As EuclideanLine
    Dim ObjectLine4 As EuclideanLine

    Dim ObjSprite As Sprite
    
    ObjSprite = GetSprite(ObjectTest.SpriteID, ObjectTest.SpriteStrip)

    With ObjectLine1
        .Point1.X = ObjectTest.X
        .Point1.Y = ObjectTest.Y
        .Point2.X = ObjectTest.X
        .Point2.Y = ObjectTest.Y + ObjSprite.FrameHeight
    End With
    
    With ObjectLine2
        .Point1.X = ObjectTest.X + ObjSprite.FrameWidth
        .Point1.Y = ObjectTest.Y
        .Point2.X = ObjectTest.X + ObjSprite.FrameWidth
        .Point2.Y = ObjectTest.Y + ObjSprite.FrameHeight
    End With
    
    With ObjectLine3
        .Point1.X = ObjectTest.X
        .Point1.Y = ObjectTest.Y
        .Point2.X = ObjectTest.X + ObjSprite.FrameWidth
        .Point2.Y = ObjectTest.Y
    End With
    
    With ObjectLine4
        .Point1.X = ObjectTest.X
        .Point1.Y = ObjectTest.Y + ObjSprite.FrameHeight
        .Point2.X = ObjectTest.X + ObjSprite.FrameWidth
        .Point2.Y = ObjectTest.Y + ObjSprite.FrameHeight
    End With
    
    Left = CollisionCheckLineLine(LineTest, ObjectLine1)
    Right = CollisionCheckLineLine(LineTest, ObjectLine2)
    Up = CollisionCheckLineLine(LineTest, ObjectLine3)
    Down = CollisionCheckLineLine(LineTest, ObjectLine4)
    
    If (Left Or Right Or Up Or Down _
        Or (((LineTest.Point1.X >= ObjectTest.X) And (LineTest.Point1.X <= ObjectTest.X + ObjSprite.FrameWidth)) _
            And (LineTest.Point1.Y >= ObjectTest.Y) And (LineTest.Point1.Y <= ObjectTest.Y + ObjSprite.FrameHeight) _
            And (LineTest.Point2.X >= ObjectTest.X) And (LineTest.Point2.X <= ObjectTest.X + ObjSprite.FrameWidth)) _
            And (LineTest.Point2.Y >= ObjectTest.Y) And (LineTest.Point2.Y <= ObjectTest.Y + ObjSprite.FrameHeight)) Then
        CollisionCheckObjectLine = True
    Else
        CollisionCheckObjectLine = False
    End If
End Function

Public Function WillCast(PlayerAngles() As Single, PlayerDimensions() As Single, SpellAngles() As Single, SpellDimensions() As Single, NumAngles As Integer, InitialD As SpellCorrect) As Boolean
    Dim Cast As Boolean
    Dim I As Integer
    Cast = True
    
    For I = 0 To NumAngles - 1
        If (Abs(PlayerAngles(I) - SpellAngles(I)) > SPELLANGLETHRESHOLD) Then
            Cast = False
        End If
    Next I
    
    For I = 0 To NumAngles - 1 'NumAngles = NumDimensions
        If (Abs(PlayerDimensions(I) / InitialD.Length - SpellDimensions(I)) > 0.3) Then
            Cast = False
        End If
    Next I
    
    WillCast = Cast
End Function

Public Sub UpdateObject(Object As GameObject)

    Dim I As Integer

    Select Case Object.TypeID
        Case PLAYER
            Dim HSpeed As Integer
            Dim VSpeed As Integer
            
            Dim Pos As Integer '0 = Up, 1 = Right, 2 = Down, 3 = Left
            
            Dim MacroXPrev As Integer
            Dim MacroYPrev As Integer
            
            Dim XPrev As Long
            Dim YPrev As Long
            
            Dim Mana As Integer
            Dim Health As Integer
            Dim DotState As String
            Dim SlashState As String
            
            Dim CanPlace As Boolean
            Dim CanSlash As Boolean
            
            CanPlace = False

            If (VBA.Val(VBA.Mid$(VBA.Trim$(Split(Object.DataTag, "|")(0)), 1, 1)) = 0) Then
                CanPlace = True
            End If
            
            If (VBA.Val(VBA.Mid$(VBA.Trim$(Split(Object.DataTag, "|")(0)), 2, 1)) = 0) Then
                CanSlash = True
            End If
            
            If (VBA.Val(Split(Object.DataTag, "|")(1)) <= 0 And VBA.Len(VBA.Trim$(Split(Object.DataTag, "|")(1))) >= 1) Then
                Dim Name As String
                Object.Removed = True
                If (LevelsCompleted = 1) Then
                    MsgBox "Game Over! You completed " & LevelsCompleted & " level.", vbCritical, "Game Over!"
                Else
                    MsgBox "Game Over! You completed " & LevelsCompleted & " levels.", vbCritical, "Game Over!"
                End If
                
                If (LevelsCompleted >= frmScores.GetLevelsCompleted(frmScores.GetMaxEntries)) Then
                    Name = InputBox$("Enter your name: ", "Leaderboard")
                    
                    If (Name = "") Then
                        Name = "Anonymous"
                    End If
                    
                    frmScores.AddEntry Name, LevelsCompleted
                End If
                
                Running = False
            End If
            
            Mana = VBA.Val(Split(Object.DataTag, "|")(2))
            Health = VBA.Val(Split(Object.DataTag, "|")(1))
            DotState = VBA.Mid$(Object.DataTag, 1, 1)
            SlashState = VBA.Mid$(Object.DataTag, 2, 1)
            Pos = VBA.Mid$(Object.DataTag, 3, 1)
            
            If (Mana < 300) Then
                Mana = Mana + 1
            End If
            
            Object.KnockBackX = (Object.KnockBackX / 1.5)
            Object.KnockBackY = (Object.KnockBackY / 1.5)
            
            XPrev = Object.X
            YPrev = Object.Y
            MacroXPrev = MacroX
            MacroYPrev = MacroY
            
            HSpeed = 0
            VSpeed = 0
            
            If (IsKeyDown(vbKeyW)) Then
                VSpeed = -50
                Pos = 0
            End If
            If (IsKeyDown(vbKeyA)) Then
                HSpeed = -50
                Pos = 3
            End If
            If (IsKeyDown(vbKeyS)) Then
                VSpeed = 50
                Pos = 2
            End If
            If (IsKeyDown(vbKeyD)) Then
                HSpeed = 50
                Pos = 1
            End If
            
            If (IsMouseDown(vbLeftButton) And CanPlace And Not ActiveLink And UBound(Dots) < 100 And Mana > 50 And Sqr(((MouseX - CamCorrectX(Object.X))) ^ 2 + (MouseY - CamCorrectY(Object.Y)) ^ 2) <= ToTwips(MAXDOTRANGE) / 2) Then
                Dim DotObj As GameObject
                CanPlace = False
                DotState = "1"
                DotObj.DataTag = VBA.Format$(CurrentRoom.Spot, "0")
                CreateQueue DotObj, AbsoluteCorrectX(MouseX), AbsoluteCorrectY(MouseY), DOT
                Mana = Mana - 50
            End If
            
            If (IsKeyDown(vbKeySpace) And CanSlash) Then
                Dim SlashObj As GameObject
                CanSlash = False
                SlashState = "1"
                CreateQueue SlashObj, Object.X, Object.Y, SLASH
            End If
            
            If (IsMouseDown(vbRightButton) And CanPlace And Not ActiveLink) Then
                ReDim LinkDots(UBound(Dots))
                Dim LinkObj As GameObject
                Dim Life As String
                
                Life = VBA.Format$(10, "0")
                
                CanPlace = False
                Object.DataTag = "1" & VBA.Right$(Object.DataTag, VBA.Len(Object.DataTag) - 1)
                For I = 0 To UBound(Dots) - 1
                    LinkDots(I) = Dots(I)
                Next I
                
                LinkObj.DataTag = Life
                
                CreateQueue LinkObj, 0, 0, LINK
            End If
            
            If (Not IsMouseDown(vbLeftButton) And Not IsMouseDown(vbRightButton)) Then
                CanPlace = True
                DotState = "0"
            End If
            
            If (Not IsKeyDown(vbKeySpace)) Then
                CanSlash = True
                SlashState = "0"
            End If
            
            If (Abs(HSpeed) > 0 Or Abs(VSpeed) > 0) Then
                Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 5) + 1
                Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 5))) Mod 3
            Else
                Object.SpriteFrame = 0
            End If
            
            Object.SpriteStrip = Pos
            
            Object.X = Object.X + HSpeed + Object.KnockBackX
            Object.Y = Object.Y + VSpeed + Object.KnockBackY

            PlayerX = Object.X
            PlayerY = Object.Y
            
            MacroX = Int(Object.X / ToTwips(800))
            MacroY = Int(Object.Y / ToTwips(800))
            
            Dim Direction As String

            If (MacroX - MacroXPrev <> 0 Or MacroY - MacroYPrev <> 0) Then
                If (MacroX > MacroXPrev) Then
                    Direction = "R"
                ElseIf (MacroX < MacroXPrev) Then
                    Direction = "L"
                ElseIf (MacroY > MacroYPrev) Then
                    Direction = "D"
                ElseIf (MacroY < MacroYPrev) Then
                    Direction = "U"
                End If
                
                Dim RelativeX As Long
                Dim RelativeY As Long
                
                RelativeX = 0
                RelativeY = 0
                
                If (Not BreakTime) Then
                    If (Direction = VBA.Mid$(CurrentRoom.Path, CurrentRoom.Spot, 1) And CurrentRoom.Spot <> VBA.Len(VBA.Trim$(CurrentRoom.Path))) Then
                        If (CurrentRoom.Spot < VBA.Len(VBA.Trim$(CurrentRoom.Path))) Then
                            CurrentRoom.Spot = CurrentRoom.Spot + 1
                            
                            Dim NumEnemies As Integer
                            
                            NumEnemies = Int((Rnd() * (3 + Int(Difficulty / 3)))) + 1
                            
                            For I = 0 To NumEnemies - 1
                                Dim EnemyObject As GameObject
                                CreateQueue EnemyObject, Object.X + ToTwips(Int(Rnd() * 800)), Object.Y + ToTwips(Int(Rnd() * 800)), ENEMY1
                            Next I
                        End If
                    ElseIf (Not BackTrack(Direction)) Then
                        Select Case Direction
                            Case "R"
                                Object.X = Object.X - ToTwips(800)
                                RelativeX = -ToTwips(800)
                            Case "L"
                                Object.X = Object.X + ToTwips(800)
                                RelativeX = ToTwips(800)
                            Case "D"
                                Object.Y = Object.Y - ToTwips(800)
                                RelativeY = -ToTwips(800)
                            Case "U"
                                Object.Y = Object.Y + ToTwips(800)
                                RelativeY = ToTwips(800)
                        End Select
                        
                        For I = 0 To UBound(Objects) - 1
                            If (Objects(I).ID <> Object.ID And Objects(I).TypeID <> BEACON And Objects(I).TypeID <> PORTALOBJ) Then  '
                                Objects(I).X = Objects(I).X + RelativeX
                                Objects(I).Y = Objects(I).Y + RelativeY
                            End If
                        Next I
                        
                        For I = 0 To UBound(Dots) - 1
                            Dots(I).X = Dots(I).X + RelativeX
                            Dots(I).Y = Dots(I).Y + RelativeY
                        Next I
                        
                        MacroX = MacroXPrev
                        MacroY = MacroYPrev
                    End If
                End If
            
            End If
            
            Dim Temp As GameObject
            Dim OriX As Long
            Dim OriY As Long
            Temp = Object
            
            OriX = Object.X
            OriY = Object.Y

            For I = 0 To UBound(MapCollision)
                If (CollisionCheckObjectLine(Temp, MapCollision(I))) Then
                    Temp.X = XPrev
                End If
                If (CollisionCheckObjectLine(Temp, MapCollision(I))) Then
                    Temp.Y = YPrev
                End If
                If (CollisionCheckObjectLine(Object, MapCollision(I))) Then
                    Object.Y = YPrev
                End If
                If (CollisionCheckObjectLine(Object, MapCollision(I))) Then
                    Object.X = XPrev
                End If
                If (Object.X <> Temp.X) Then
                    Object.X = OriX
                End If
                If (Object.Y <> Temp.Y) Then
                    Object.Y = OriY
                End If
            Next I

            For I = 0 To UBound(Objects) - 1
                If (Objects(I).TypeID = ENEMY1 And CollisionCheck(Object, Objects(I))) Then

                    Dim Theta As Double

                    Health = Health - 50
                    Objects(I).DataTag = VBA.Format$(VBA.Val(Objects(I).DataTag) - 15, "0")
                    
                    Object.KnockBackX = (Object.X - Objects(I).X) * 0.75
                    Object.KnockBackY = (Object.Y - Objects(I).Y) * 0.75
                    
                    Objects(I).KnockBackX = -Object.KnockBackX * 0.75
                    Objects(I).KnockBackY = -Object.KnockBackY * 0.75
                    
                ElseIf (Objects(I).TypeID = PORTALOBJ And CollisionCheck(Object, Objects(I)) And IsKeyDown(vbKeyUp)) Then
                    Objects(I).OnCreate = True
                End If
            Next I
            
            Object.DataTag = DotState & SlashState & VBA.Format$(Pos, "0") & "|" & VBA.Format$(Health, "0") & "|" & VBA.Format$(Mana, "0") & "|"
            
            CamX = Object.X + CAMXOFFSET
            CamY = Object.Y + CAMYOFFSET
        Case PORTALOBJ
        
            If (Object.OnCreate) Then
                Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 10) + 1
                Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 10))) Mod 4
                
                If (Object.SpriteFrame >= 3 And VBA.Trim$(Object.DataTag) = "R") Then
                    LoadNewLevel
                    frmMain.ResetKeys
                ElseIf (Object.SpriteFrame >= 3 And VBA.Trim$(Object.DataTag) = "B") Then
                    LevelsCompleted = LevelsCompleted + 1
                    LoadBreakRoom
                    frmMain.ResetKeys
                End If
            End If
        
        Case STARPLATINUM

            Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 4) + 1
            Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 4))) Mod 50

            If (Not Object.OnCreate And Object.SpriteFrame >= 18) Then
                Object.OnCreate = True
                Dim HitBox As GameObject
                HitBox.DataTag = VBA.Format$(Object.LinkID, "0")
                Select Case Object.SpriteStrip
                    Case 0
                        HitBox.SpriteStrip = 1
                        CreateQueue HitBox, Object.X + ToTwips(20), Object.Y + ToTwips(60), SPHITBOXOBJ
                    Case 1
                        HitBox.SpriteStrip = 0
                        CreateQueue HitBox, Object.X, Object.Y, SPHITBOXOBJ
                    Case 2
                        HitBox.SpriteStrip = 0
                        CreateQueue HitBox, Object.X + ToTwips(75), Object.Y, SPHITBOXOBJ
                    Case 3
                        HitBox.SpriteStrip = 1
                        CreateQueue HitBox, Object.X + ToTwips(15), Object.Y - ToTwips(5), SPHITBOXOBJ
                End Select
                Object.DataTag = "0"
            End If
            
            If (Object.SpriteFrame >= 33 And VBA.Val(Object.DataTag) < 8) Then
                Object.SpriteFrame = 18
                Object.DataTag = VBA.Format$(VBA.Val(Object.DataTag) + 1, "0")
            End If

            If (Object.SpriteFrame >= 49 And VBA.Val(Object.DataTag) >= 8) Then
                For I = 0 To UBound(Objects) - 1
                    If (Objects(I).TypeID = SPHITBOXOBJ And Abs(VBA.Val(VBA.Trim$(Objects(I).DataTag)) - Object.LinkID) <= 4) Then
                        Objects(I).Removed = True
                    End If
                Next I
                Object.Removed = True
            End If
            
            
        Case LINK
            ActiveLink = True

            Object.DataTag = VBA.Format$((VBA.Val(Object.DataTag) - 1), "0")
            
            Dim InitialD As SpellCorrect
            Dim AveragePosition As Point
            
            If (Not Object.OnCreate) Then
            
                Dim Shout As Boolean
                Dim NumAngles As Integer
                
                NumAngles = UBound(Dots)
                
                Dim Angles(100) As Single
                Dim Dimensions(100) As Single
                
                Shout = True
            
                For I = 0 To UBound(Dots) - 1
                    Dim A As Point
                    Dim B As Point
                    Dim C As Point
                    
                    A.X = Dots(I).X
                    A.Y = Dots(I).Y
                    
                    B.X = Dots((I + 1) Mod UBound(Dots)).X
                    B.Y = Dots((I + 1) Mod UBound(Dots)).Y
                    
                    C.X = Dots((I + 2) Mod UBound(Dots)).X
                    C.Y = Dots((I + 2) Mod UBound(Dots)).Y
                    
                    Dimensions(I) = Sqr((B.X - A.X) ^ 2 + (B.Y - A.Y) ^ 2)
                    Angles(I) = GetTheta(A, B, C)
                    
                    AveragePosition.X = AveragePosition.X + A.X
                    AveragePosition.Y = AveragePosition.Y + A.Y
                    
                    If (I = 0) Then
                        InitialD.Length = Dimensions(I)
                        If (Abs(B.X - A.X) > 0.1) Then
                            InitialD.Theta = 90 - ToDegrees(Atn((B.Y - A.Y) / (B.X - A.X)))
                        Else
                            InitialD.Theta = 90
                        End If
                    End If
                Next I
                
                If (UBound(Dots) <> 0) Then
                    AveragePosition.X = AveragePosition.X / UBound(Dots)
                    AveragePosition.Y = AveragePosition.Y / UBound(Dots)
                End If
                
                Dim Spell As GameObject
                Select Case NumAngles
                    Case 4
                        If (WillCast(Angles, Dimensions, StarPlatinumAngle, StarPlatinumDimension, NumAngles, InitialD)) Then
                            Spell.DataTag = MouseAngle
                            CreateQueue Spell, AveragePosition.X - GetSprite(STARPLATINUMSPRITE, 0).FrameWidth / 2, AveragePosition.Y - GetSprite(STARPLATINUMSPRITE, 0).FrameHeight / 2, STARPLATINUM
                        End If
                    Case 5
                        If (WillCast(Angles, Dimensions, PillarAngle, PillarDimension, NumAngles, InitialD)) Then
                            CreateQueue Spell, AveragePosition.X - GetSprite(STARPLATINUMSPRITE, 0).FrameWidth / 2, AveragePosition.Y - GetSprite(STARPLATINUMSPRITE, 0).FrameHeight / 2, BEACON
                        End If
                End Select
            
                Object.OnCreate = True
            End If
            
            If (VBA.Val(Object.DataTag) <= 0) Then
                Object.Removed = True
                ActiveLink = False
                For I = 0 To UBound(Objects) - 1
                    If (Objects(I).TypeID = DOT) Then
                        Objects(I).Removed = True
                    End If
                Next
                ReDim Dots(0)
                ReDim LinkDots(0)
            End If
        Case DOT
        
            Dim DistanceToPlayer As Long
        
            Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 5) + 1
            Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 5))) Mod 6
            
            DistanceToPlayer = Sqr((Object.X - PlayerX) ^ 2 + (Object.Y - PlayerY) ^ 2)
            
        Case BEACON
        
            Dim LifeSpan As Integer
            
            LifeSpan = VBA.Val(Object.DataTag)
            
            If (LifeSpan >= 10000) Then
                Object.Removed = True
            End If
            
            Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 1) + 1
            Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 1))) Mod 40
            Object.DataTag = VBA.Format$(LifeSpan + 1, "0")
            
        Case SLASH
            Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 2) + 1
            Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 2))) Mod 6
            
            If (Object.SpriteFrame = 5) Then
                Object.Removed = True
            End If
        Case ENEMY1
            If (Object.SpriteStrip = 0) Then
                Object.AnimationSpeed = ((Object.AnimationSpeed + 1) Mod 10) + 1
                Object.SpriteFrame = (Object.SpriteFrame + (Int(Object.AnimationSpeed / 10))) Mod 4
            Else
                Object.SpriteFrame = 0
            End If
        
            Object.KnockBackX = Object.KnockBackX / 1.5
            Object.KnockBackY = Object.KnockBackY / 1.5
            
            If (Int(Abs(Object.KnockBackX)) > 0 Or Int(Abs(Object.KnockBackY)) > 0) Then
                Object.SpriteStrip = 1
                Object.SpriteFrame = 0
            Else
                Object.SpriteStrip = 0
            End If
            
            XPrev = Object.X
            YPrev = Object.Y
        
            If (VBA.Val(VBA.Trim$(Object.DataTag)) <= 0) Then
                Object.Removed = True
            End If
            
            Dim DeltaX As Double
            Dim DeltaY As Double
            Dim ScaleDelta As Double
            
            For I = 0 To UBound(Objects) - 1
                If (CollisionCheck(Object, Objects(I)) And (Objects(I).TypeID = ENEMY1) And I <> Object.ID) Then
                    Object.KnockBackX = -(Objects(I).X - Object.X)
                    Object.KnockBackY = -(Objects(I).Y - Object.Y)
                    Objects(I).KnockBackX = (Objects(I).X - Object.X)
                    Objects(I).KnockBackY = (Objects(I).Y - Object.Y)
                    Object.DataTag = VBA.Format$(VBA.Val(Object.DataTag) - 30, "0")
                    Objects(I).DataTag = VBA.Format$(VBA.Val(Objects(I).DataTag) - 30, "0")
                End If
                
                If (CollisionCheck(Object, Objects(I)) And (Objects(I).TypeID = SLASH) And I <> Object.ID) Then
                    Object.KnockBackX = MAXENEMYKNOCKBACKSTUN * Cos(ToRadians(MouseAngle)) * 20
                    Object.KnockBackY = -MAXENEMYKNOCKBACKSTUN * Sin(ToRadians(MouseAngle)) * 20
                    Object.DataTag = VBA.Format$(VBA.Val(Object.DataTag) - 50, "0")
                End If
                
                If (CollisionCheck(Object, Objects(I)) And (Objects(I).TypeID = SPHITBOXOBJ)) Then
                    Object.KnockBackX = -(Objects(I).X - Object.X) * 0.1
                    Object.KnockBackY = -(Objects(I).Y - Object.Y) * 0.1
                    Object.DataTag = VBA.Format$(VBA.Val(Object.DataTag) - 20, "0")
                End If
                
                'Follow
            
                If (Objects(I).TypeID = PLAYER) Then
                
                    DeltaX = Object.X - Objects(I).X
                    DeltaY = Object.Y - Objects(I).Y
                    
                    If (DeltaX <> 0 Or DeltaY <> 0) Then
                        ScaleDelta = ENEMYSPEED1 * 1.25 / (Sqr(DeltaX ^ 2 + DeltaY ^ 2))
                    End If
                    
                    Object.X = Object.X - (DeltaX * ScaleDelta)
                    Object.Y = Object.Y - (DeltaY * ScaleDelta)
                End If
            Next I
            
            Object.X = Object.X + Object.KnockBackX
            Object.Y = Object.Y + Object.KnockBackY
            
            If (ActiveLink) Then
                Dim Point1 As Point
                Dim Point2 As Point
                Dim LinkLine As EuclideanLine
                For I = 0 To UBound(Dots) - 1
                    Point1.X = Dots(I).X
                    Point1.Y = Dots(I).Y
                    Point2.X = Dots((I + 1) Mod UBound(Dots)).X
                    Point2.Y = Dots((I + 1) Mod UBound(Dots)).Y
                    LinkLine.Point1 = Point1
                    LinkLine.Point2 = Point2
                    If (CollisionCheckObjectLine(Object, LinkLine)) Then
                        Object.DataTag = VBA.Format$(VBA.Val(Object.DataTag) - 20, "0")
                        Object.KnockBackX = (XPrev - Object.X) * ToTwips(1.5) * (MAXENEMYKNOCKBACKSTUN / (Sqr((XPrev - Object.X) ^ 2 + (YPrev - Object.Y) ^ 2)))
                        Object.KnockBackY = (YPrev - Object.Y) * ToTwips(1.5) * (MAXENEMYKNOCKBACKSTUN / (Sqr((XPrev - Object.X) ^ 2 + (YPrev - Object.Y) ^ 2)))
                    End If
                Next
            End If
    End Select
End Sub
