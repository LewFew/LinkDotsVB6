Attribute VB_Name = "mdlGraphics"
Option Explicit

'Author: William Chan
'Date: May 17th, 2019
'Purpose: ICS4U Culminating Assignment

'SpriteIDs

Global Const WIZARD = 0
Global Const ENEMYWILLO = 4
Global Const BASICDOT = 5
Global Const RANGEINDICATOR = 6
Global Const SLASHSPRITE = 7
Global Const STARPLATINUMSPRITE = 8
Global Const SPHITBOXSPRITE = 9
Global Const BEACONSPRITE = 11
Global Const PORTALSPRITE = 12

Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean

Public Type Sprite
    Width As Long
    Height As Long
    FrameWidth As Long
    FrameHeight As Long
    Pic As Picture
    hDC As Long 'Pointer Value
End Type
Public Sub ResetDraw()
    frmMain.picDisplay.FillColor = vbBlack
    frmMain.picDisplay.ForeColor = vbBlack
    frmMain.picDisplay.DrawWidth = 1
End Sub

Public Function Max(ByVal A As Long, ByVal B As Long) As Long
    If (A >= B) Then
        Max = A
    Else
        Max = B
    End If
End Function

Public Function Min(ByVal A As Long, ByVal B As Long) As Long
    If (A <= B) Then
        Min = A
    Else
        Min = B
    End If
End Function

Public Sub RenderObjectEnd(Object As GameObject)
    Select Case Object.TypeID
        Case PLAYER
            Dim RangeInd As Sprite
            Dim Dark As Sprite
        
            RangeInd = GetSprite(RANGEINDICATOR, 0)

            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbBlack
            frmMain.picDisplay.Line (0, 0)-(500 * 8, 150), , B
            frmMain.picDisplay.FillColor = vbGreen
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (0, 0)-(VBA.Val(Split(VBA.Trim$(Object.DataTag), "|")(1)) * 8, 75), , B
            frmMain.picDisplay.FillColor = frmImages.shpManaColour.FillColor
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (0, 75)-(VBA.Val(Split(VBA.Trim$(Object.DataTag), "|")(2)) * 13.33333, 150), , B
            DrawImage RangeInd, 0, frmMain.Width / 2 - RangeInd.Width / 2, frmMain.Height / 2 - RangeInd.Height / 1.6
        Case ENEMY1
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbBlack
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X - 300), CamCorrectY(Object.Y - 50))-(CamCorrectX(Object.X + 300 * 4), CamCorrectY(Object.Y - 100)), , B
            frmMain.picDisplay.FillStyle = vbFSSolid
            frmMain.picDisplay.FillColor = vbGreen
            frmMain.picDisplay.ForeColor = vbBlack
            frmMain.picDisplay.Line (CamCorrectX(Object.X - 300), CamCorrectY(Object.Y - 50))-(CamCorrectX(Object.X - 300 + VBA.Val(Object.DataTag) * 5), CamCorrectY(Object.Y - 100)), , B
    End Select
End Sub

Public Sub RenderObject(Object As GameObject)
    Dim I As Integer
    Select Case Object.TypeID
        Case LINK
            Dim SizeOffset As Integer
            If (VBA.Trim$(Object.DataTag) <> 0) Then
                frmMain.picDisplay.DrawWidth = VBA.Val(VBA.Trim$(Object.DataTag))
                frmMain.picDisplay.ForeColor = frmImages.shpBasicLinkColour.FillColor
                SizeOffset = frmMain.picDisplay.DrawWidth - 1
                For I = 0 To UBound(LinkDots) - 1
                    frmMain.picDisplay.Line (CamCorrectX(LinkDots(I).X + ToTwips(10) + SizeOffset), CamCorrectY(LinkDots(I).Y + ToTwips(10) + SizeOffset))-(CamCorrectX(LinkDots((I + 1) Mod UBound(LinkDots)).X + ToTwips(10) + SizeOffset), CamCorrectY(LinkDots((I + 1) Mod UBound(LinkDots)).Y + ToTwips(10) + SizeOffset))
                Next I
            End If
    End Select
    ResetDraw
End Sub

Public Sub DrawImage(Img As Sprite, ByVal Frame As Long, ByVal X As Integer, ByVal Y As Integer)
    TransparentBlt frmMain.picDisplay.hDC, ToPixels(X), ToPixels(Y), ToPixels(Img.FrameWidth), ToPixels(Img.FrameHeight), Img.hDC, Frame * ToPixels(Img.FrameWidth), 0, ToPixels(Img.FrameWidth), ToPixels(Img.FrameHeight), frmImages.shpTransparent.FillColor
End Sub

Public Function GetSprite(ByVal SpriteID As Integer, ByVal Strip As Integer) As Sprite
    Dim ReturnImage As Sprite
    Select Case SpriteID
        Case WIZARD
            Set ReturnImage.Pic = frmImages.picWizard(Strip).Picture
            ReturnImage.Width = frmImages.picWizard(Strip).Width
            ReturnImage.Height = frmImages.picWizard(Strip).Height
            ReturnImage.FrameWidth = ToTwips(32)
            ReturnImage.FrameHeight = ToTwips(32)
            ReturnImage.hDC = frmImages.picWizard(Strip).hDC
        Case PORTALSPRITE
            Set ReturnImage.Pic = frmImages.picPortal(Strip).Picture
            ReturnImage.Width = frmImages.picPortal(Strip).Width
            ReturnImage.Height = frmImages.picPortal(Strip).Height
            ReturnImage.FrameWidth = ToTwips(46)
            ReturnImage.FrameHeight = ToTwips(51)
            ReturnImage.hDC = frmImages.picPortal(Strip).hDC
        Case BASICDOT
            Set ReturnImage.Pic = frmImages.picDot(Strip).Picture
            ReturnImage.Width = frmImages.picDot(Strip).Width
            ReturnImage.Height = frmImages.picDot(Strip).Height
            ReturnImage.FrameWidth = ToTwips(20)
            ReturnImage.FrameHeight = ToTwips(20)
            ReturnImage.hDC = frmImages.picDot(Strip).hDC
        Case SPHITBOXSPRITE
            Set ReturnImage.Pic = frmImages.picSPHitbox(Strip).Picture
            ReturnImage.Width = frmImages.picSPHitbox(Strip).Width
            ReturnImage.Height = frmImages.picSPHitbox(Strip).Height
            If (Strip = 0) Then
                ReturnImage.FrameWidth = ToTwips(32)
                ReturnImage.FrameHeight = ToTwips(60)
            Else
                ReturnImage.FrameWidth = ToTwips(60)
                ReturnImage.FrameHeight = ToTwips(32)
            End If
            ReturnImage.hDC = frmImages.picSPHitbox(Strip).hDC
        Case ENEMYWILLO
            Set ReturnImage.Pic = frmImages.picWilloWisp(Strip).Picture
            ReturnImage.Width = frmImages.picWilloWisp(Strip).Width
            ReturnImage.Height = frmImages.picWilloWisp(Strip).Height
            If (Strip = 0) Then
                ReturnImage.FrameWidth = ToTwips(35)
                ReturnImage.FrameHeight = ToTwips(44)
            ElseIf (Strip = 1) Then
                ReturnImage.FrameWidth = ToTwips(43)
                ReturnImage.FrameHeight = ToTwips(35)
            End If
            ReturnImage.hDC = frmImages.picWilloWisp(Strip).hDC
        Case RANGEINDICATOR
            Set ReturnImage.Pic = frmImages.picRange(Strip).Picture
            ReturnImage.Width = frmImages.picRange(Strip).Width
            ReturnImage.Height = frmImages.picRange(Strip).Height
            ReturnImage.FrameWidth = ToTwips(250)
            ReturnImage.FrameHeight = ToTwips(250)
            ReturnImage.hDC = frmImages.picRange(Strip).hDC
        Case BEACONSPRITE
            Set ReturnImage.Pic = frmImages.picBeacon(Strip).Picture
            ReturnImage.Width = frmImages.picBeacon(Strip).Width
            ReturnImage.Height = frmImages.picBeacon(Strip).Height
            ReturnImage.FrameWidth = ToTwips(42)
            ReturnImage.FrameHeight = ToTwips(42)
            ReturnImage.hDC = frmImages.picBeacon(Strip).hDC
        Case SLASHSPRITE
            Set ReturnImage.Pic = frmImages.picSlash(Strip).Picture
            ReturnImage.Width = frmImages.picSlash(Strip).Width
            ReturnImage.Height = frmImages.picSlash(Strip).Height
            If (Strip = 0 Or Strip = 1) Then
                ReturnImage.FrameWidth = ToTwips(66)
                ReturnImage.FrameHeight = ToTwips(36)
            ElseIf (Strip = 2 Or Strip = 3) Then
                ReturnImage.FrameWidth = ToTwips(36)
                ReturnImage.FrameHeight = ToTwips(66)
            Else
                ReturnImage.FrameWidth = ToTwips(55)
                ReturnImage.FrameHeight = ToTwips(60)
            End If
            ReturnImage.hDC = frmImages.picSlash(Strip).hDC
        Case STARPLATINUMSPRITE
            Set ReturnImage.Pic = frmImages.picStarPlatinumRush(Strip).Picture
            ReturnImage.Width = frmImages.picStarPlatinumRush(Strip).Width
            ReturnImage.Height = frmImages.picStarPlatinumRush(Strip).Height
            If (Strip = 0) Then
                ReturnImage.FrameWidth = ToTwips(100)
                ReturnImage.FrameHeight = ToTwips(100)
            Else
                ReturnImage.FrameWidth = ToTwips(100)
                ReturnImage.FrameHeight = ToTwips(70)
            End If
            ReturnImage.hDC = frmImages.picStarPlatinumRush(Strip).hDC
    End Select
    GetSprite = ReturnImage
End Function
