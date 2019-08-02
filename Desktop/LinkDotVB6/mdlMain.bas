Attribute VB_Name = "mdlMain"
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

Global Objects() As GameObject
Global ObjCreateQueue() As GameObject
Global ObjRemoveQueue() As GameObject
Global DotRemoveQueue() As GameObject
Global Dots() As GameObject
Global CamX, CamY As Long
Global MouseAngle As Double

Global Const INTEGERLIMIT = 32767
Global Const LONGLIMIT = 2147483647
Global Const CAMXOFFSET = 200
Global Const CAMYOFFSET = 700

Global Clock As Double
Global Difficulty As Integer
Global LevelsCompleted As Integer
Global BackgroundID As Integer

Global CurrentRoom As SeamlessRoom
Global LoadingNext As Boolean
Global LoadingBreak As Boolean
Global BreakTime As Boolean

Global MacroX As Integer
Global MacroY As Integer

Public Type RecordScore
    Name As String * 30
    LevelsComplete As Integer
End Type

Public Sub BubbleSort(RecArray() As RecordScore, ByVal NumRecs As Integer)
    Dim I As Integer
    Dim J As Integer
    Dim Temp As RecordScore
    
    For I = 1 To NumRecs - 1
        For J = 1 To NumRecs - 1
            If (RecArray(J).LevelsComplete < RecArray(J + 1).LevelsComplete) Then
                Temp = RecArray(J)
                RecArray(J) = RecArray(J + 1)
                RecArray(J + 1) = Temp
            End If
        Next J
    Next I
End Sub

Public Function ToPixels(ByVal Twips As Long) As Long
    ToPixels = (Twips / TILELENGTH) * 32
End Function

Public Function ToTwips(ByVal Pixels As Long) As Long
    ToTwips = (Pixels * TILELENGTH) / 32
End Function

Public Function CamCorrectX(ByVal X As Long) As Long
    CamCorrectX = (frmMain.Width / 2) + (X - CamX)
End Function

Public Function CamCorrectY(ByVal Y As Long) As Long
    CamCorrectY = (frmMain.Height / 2) + (Y - CamY)
End Function

Public Function AbsoluteCorrectX(ByVal X As Long) As Long
    AbsoluteCorrectX = (X - (frmMain.Width / 2)) + CamX
End Function

Public Function AbsoluteCorrectY(ByVal Y As Long) As Long
    AbsoluteCorrectY = (Y - (frmMain.Height / 2)) + CamY
End Function

Public Sub Init()
    ReDim Objects(0)
    ReDim ObjCreateQueue(0)
    ReDim ObjRemoveQueue(0)
    ReDim DotRemoveQueue(0)
    ReDim Dots(0)
    CamX = 3500
    CamY = 1700
    
    Clock = 0
    Difficulty = 0
    
    mdlObject.ObjectModuleInit
End Sub

Public Function IsKeyDown(ByVal Key As Integer) As Boolean
    IsKeyDown = frmMain.IsKeyDownMAIN(Key)
End Function

Public Function IsMouseDown(ByVal Button As Integer) As Boolean
    IsMouseDown = frmMain.IsMouseDownMAIN(Button)
End Function

Public Sub AddDot(DOT As GameObject)
    Dim Upper As Integer
    Upper = UBound(Dots) + 1
    
    ReDim Preserve Dots(Upper)

    Dots(UBound(Dots) - 1) = DOT
End Sub

Public Sub AddObject(GObject As GameObject)
    Dim Upper As Integer
    Upper = UBound(Objects) + 1
    
    ReDim Preserve Objects(Upper)
    
    GObject.ID = UBound(Objects) - 1

    Objects(UBound(Objects) - 1) = GObject
End Sub

Public Sub RemoveObjectQueue(GObject As GameObject)
    Dim Upper As Integer
    Upper = UBound(ObjRemoveQueue) + 1
    
    ReDim Preserve ObjRemoveQueue(Upper)

    ObjRemoveQueue(UBound(ObjRemoveQueue) - 1) = GObject
End Sub

Public Sub AddObjectQueue(GObject As GameObject)
    Dim Upper As Integer
    Upper = UBound(ObjCreateQueue) + 1
    
    ReDim Preserve ObjCreateQueue(Upper)
    
    ObjCreateQueue(UBound(ObjCreateQueue) - 1) = GObject
End Sub

Public Sub RemoveObject(Object As GameObject)
    Dim I As Integer
    Dim Lower As Integer
    Dim Index As Integer
    
    Index = Object.ID

    For I = Index To UBound(Objects) - 1
        Objects(I) = Objects(I + 1)
    Next I
    Lower = UBound(Objects) - 1
    ReDim Preserve Objects(Lower)
End Sub

Public Sub Update()
    Dim I As Integer
    Dim J As Integer

    For I = 0 To UBound(Objects) - 1
        Objects(I).ID = I
        UpdateObject Objects(I)
        If (Objects(I).Removed) Then
            RemoveObjectQueue Objects(I)
        End If
    Next I
    
    Dim PrevID As Integer
    For I = 0 To UBound(ObjRemoveQueue) - 1
        For J = 0 To UBound(ObjRemoveQueue) - 1
            If (ObjRemoveQueue(J).ID > ObjRemoveQueue(I).ID) Then
                ObjRemoveQueue(J).ID = ObjRemoveQueue(J).ID - 1
            End If
        Next J
        RemoveObject ObjRemoveQueue(I)
    Next I
    
    For I = 0 To UBound(ObjCreateQueue) - 1
        Create ObjCreateQueue(I), ObjCreateQueue(I).X, ObjCreateQueue(I).Y, ObjCreateQueue(I).TypeID
        If (ObjCreateQueue(I).TypeID = DOT) Then
            AddDot ObjCreateQueue(I)
        End If
    Next I
    
    If (LoadingBreak) Then
    
        Dim PlayerObject As GameObject
        Dim Portal As GameObject
    
        ReDim Objects(0)
        ReDim Dots(0)
        ReDim LinkDots(0)
        ReDim MapCollisions(0)
        
        BreakTime = True
        
        Portal.DataTag = "R"
        
        Create Portal, ToTwips(350), ToTwips(175), PORTALOBJ
        Create PlayerObject, ToTwips(400), ToTwips(400), PLAYER
        
        BackgroundID = -1
    
        LoadingBreak = False
    End If
    
    If (LoadingNext) Then

        Dim NumEnemies As Integer
        
        Dim X As Integer
        Dim Y As Integer
        
        BreakTime = False
        
        X = 0
        Y = 0
        
        For I = 1 To VBA.Len(VBA.Trim$(CurrentRoom.Path))
            Select Case (VBA.Mid$(CurrentRoom.Path, I, 1))
                Case "L"
                    X = X - 1
                Case "R"
                    X = X + 1
                Case "U"
                    Y = Y - 1
                Case "D"
                    Y = Y + 1
            End Select
        Next I
        
        ReDim Objects(0)
        ReDim Dots(0)
        ReDim LinkDots(0)
        
        Portal.DataTag = "B"
        
        Create Portal, ToTwips(X * 800 + 400), ToTwips(Y * 800 + 400), PORTALOBJ
        Create PlayerObject, ToTwips(400), ToTwips(400), PLAYER
        NumEnemies = Int(Rnd() * 3) + 1
        
        For I = 0 To NumEnemies - 1
            Dim EnemyObject As GameObject
            Create EnemyObject, ToTwips(Rnd() * 800), ToTwips(Rnd() * 800), ENEMY1
        Next I
        
        LoadingNext = False
    End If
    
    ReDim ObjCreateQueue(0)
    ReDim ObjRemoveQueue(0)
End Sub

Public Sub Render()

    Dim I As Integer
    
    If (Not BreakTime) Then
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800) + 1 * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800) + 1 * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800) + 1 * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800) + 1 * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800) - 1 * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800) - 1 * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800) - 1 * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800) - 1 * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800) - 1 * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800) + 1 * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
        frmMain.picDisplay.PaintPicture frmImages.imgEndlessBackground(BackgroundID), CamCorrectX(MacroX * ToTwips(800) + 1 * ToTwips(800)), CamCorrectY(MacroY * ToTwips(800) - 1 * ToTwips(800)), frmImages.imgEndlessBackground(BackgroundID).Width, frmImages.imgEndlessBackground(BackgroundID).Height
    Else
        frmMain.picDisplay.PaintPicture frmImages.picBreak(0).Image, CamCorrectX(0), CamCorrectY(0), frmImages.picBreak(0).Width, frmImages.picBreak(0).Height
    End If
    
    SetMapCollisions

    Dim Obj As GameObject
    For I = 0 To UBound(Objects) - 1
        Obj = Objects(I)
        If (Not Obj.CustomDraw) Then
            Dim DrawX, DrawY As Long
            DrawX = CamCorrectX(Obj.X)
            DrawY = CamCorrectY(Obj.Y)
            
            If (DrawX < frmMain.Width + ToTwips(100) And DrawX > 0 - ToTwips(100) And DrawY < frmMain.Height + ToTwips(100) And DrawY > 0 - ToTwips(100)) Then
                DrawImage GetSprite(Obj.SpriteID, Obj.SpriteStrip), Obj.SpriteFrame, DrawX, DrawY
            End If
        End If
        RenderObject Obj
    Next I
    
    For I = 0 To UBound(Objects) - 1
        Obj = Objects(I)
        RenderObjectEnd Obj
    Next I
End Sub
