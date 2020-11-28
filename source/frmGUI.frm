VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmGUI 
   Caption         =   "CDR2CFF: Конвертация в CFF"
   ClientHeight    =   3228
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4788
   OleObjectBlob   =   "frmGUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ResizeBoundingBox_Click()

    Dim Doc As Document
    Dim DieCut As Shape
    Dim DieCutWidth As Double, DieCutHeight As Double
    Dim Rect As Shape
    
    Set Doc = ActiveDocument
    Doc.Unit = cdrMillimeter
    
    ActivePage.Shapes.All.CreateSelection
    Set DieCut = ActiveSelection.Group
    DieCut.GetSize DieCutWidth, DieCutHeight
    
    Doc.MasterPage.SetSize DieCutWidth + Point, DieCutHeight + Point
    
    Set Rect = ActiveLayer.CreateRectangle(ActivePage.LeftX, ActivePage.TopY, ActivePage.RightX, ActivePage.BottomY)
    With Rect
        .Outline.Color.CMYKAssign 0, 0, 0, 0
        .OrderToBack
    End With
    
    With DieCut
        .AlignToShape cdrAlignVCenter, Rect
        .AlignToShape cdrAlignHCenter, Rect
        .Ungroup
    End With
        
    Rect.Delete
    
    ActiveDocument.ClearSelection

End Sub

Private Sub GetCoordinates0_Click()
    
    Dim CoordX As Double, CoordY As Double
    Dim ShWidth As Double, ShHeight As Double
    Dim CommaToPointX As String, CommaToPointY As String
    Dim HorizontalCoordinate As Double, VerticalCoordinate As Double
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    ActiveDocument.BeginCommandGroup
    
    For Each Knife In ActiveSelection.Shapes
        With Knife
            .GetBoundingBox CoordX, CoordY, ShWidth, ShHeight
            .Outline.Color.CMYKAssign 0, 0, 0, 50
        End With
        
        HorizontalCoordinate = CoordX - HalfPoint
        VerticalCoordinate = CoordY - HalfPoint
        
        CommaToPointX = Replace(CStr(HorizontalCoordinate), ",", ".")
        CommaToPointY = Replace(CStr(VerticalCoordinate), ",", ".")
          
        Set TextCoor = ActiveLayer.CreateParagraphText(-210, 0, 0, ActivePage.SizeHeight, _
            "C,Die_0, " & CommaToPointX & ", " & CommaToPointY & ", 0.00,1,1", , , _
            "Courier New", 24, cdrFalse, cdrFalse, , cdrLeftAlignment)
    Next Knife
    
    ActiveDocument.ClearSelection
    
    ActiveDocument.EndCommandGroup
        
End Sub

Private Sub GetCoordinates90_Click()
    
    Dim CoordX As Double, CoordY As Double
    Dim ShWidth As Double, ShHeight As Double
    Dim CommaToPointX As String, CommaToPointY As String
    Dim HorizontalCoordinate As Double, VerticalCoordinate As Double
        
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    ActiveDocument.BeginCommandGroup
    
    For Each Knife In ActiveSelection.Shapes
        With Knife
            .GetBoundingBox CoordX, CoordY, ShWidth, ShHeight
            .Outline.Color.CMYKAssign 0, 0, 0, 50
        End With
        
        HorizontalCoordinate = CoordX + Knife.SizeWidth + HalfPoint
        CommaToPointX = Replace(CStr(HorizontalCoordinate), ",", ".")
        VerticalCoordinate = CoordY - HalfPoint
        CommaToPointY = Replace(CStr(VerticalCoordinate), ",", ".")
        
        Set TextCoor = ActiveLayer.CreateParagraphText(-210, 0, 0, ActivePage.SizeHeight, _
            "C,Die_0, " & CommaToPointX & ", " & CommaToPointY & ", 90.00,1,1", , , _
            "Courier New", 24, cdrFalse, cdrFalse, , cdrLeftAlignment)
            
        ActiveDocument.ClearSelection
    Next Knife
  
    ActiveDocument.ClearSelection
    
    ActiveDocument.EndCommandGroup
 
End Sub

Private Sub GetCoordinates180_Click()
    
    Dim CoordX As Double, CoordY As Double
    Dim ShWidth As Double, ShHeight As Double
    Dim CommaToPointX As String, CommaToPointY As String
    Dim HorizontalCoordinate As Double, VerticalCoordinate As Double
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    ActiveDocument.BeginCommandGroup
    
    For Each Knife In ActiveSelection.Shapes
        With Knife
            .GetBoundingBox CoordX, CoordY, ShWidth, ShHeight
            .Outline.Color.CMYKAssign 0, 0, 0, 50
        End With
        
        HorizontalCoordinate = CoordX + Knife.SizeWidth + HalfPoint
        CommaToPointX = Replace(CStr(HorizontalCoordinate), ",", ".")
        VerticalCoordinate = CoordY + Knife.SizeHeight + HalfPoint
        CommaToPointY = Replace(CStr(VerticalCoordinate), ",", ".")
        
        Set TextCoor = ActiveLayer.CreateParagraphText(-210, 0, 0, ActivePage.SizeHeight, _
            "C,Die_0, " & CommaToPointX & ", " & CommaToPointY & ", 180.00,1,1", , , _
            "Courier New", 24, cdrFalse, cdrFalse, , cdrLeftAlignment)
            
    Next Knife
    
    ActiveDocument.ClearSelection
    
    ActiveDocument.EndCommandGroup
  
End Sub

Private Sub GetCoordinates270_Click()
    
    Dim CoordX As Double, CoordY As Double
    Dim ShWidth As Double, ShHeight As Double
    Dim CommaToPointX As String, CommaToPointY As String
    Dim HorizontalCoordinate As Double, VerticalCoordinate As Double
    
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrBottomLeft
    
    ActiveDocument.BeginCommandGroup
    
    For Each Knife In ActiveSelection.Shapes
        With Knife
            .GetBoundingBox CoordX, CoordY, ShWidth, ShHeight
            .Outline.Color.CMYKAssign 0, 0, 0, 50
        End With
        
        HorizontalCoordinate = CoordX - HalfPoint
        CommaToPointX = Replace(CStr(CoordX), ",", ".")
        VerticalCoordinate = CoordY + Knife.SizeHeight + HalfPoint
        CommaToPointY = Replace(CStr(VerticalCoordinate), ",", ".")
        
        Set TextCoor = ActiveLayer.CreateParagraphText(-210, 0, 0, ActivePage.SizeHeight, _
            "C,Die_0, " & CommaToPointX & ", " & CommaToPointY & ", 270.00,1,1", , , _
            "Courier New", 24, cdrFalse, cdrFalse, , cdrLeftAlignment)
        ActiveDocument.ClearSelection
    Next Knife
  
    ActiveDocument.ClearSelection
    
    ActiveDocument.EndCommandGroup
  
End Sub

Private Sub ListGenerator_Click()

    Dim sr As ShapeRange
    
    Set PageSizeTxt = ActiveLayer.CreateParagraphText(-210, 0, 0, ActivePage.SizeHeight, _
        "UR, " & Replace(CStr(ActivePage.SizeWidth), ",", ".") & ", " & Replace(CStr(ActivePage.SizeHeight), ",", ".") & vbNewLine & "SCALE,1,1", , , _
        "Courier New", 24, cdrFalse, cdrFalse, , cdrLeftAlignment)
    PageSizeTxt.OrderToBack
    Set sr = ActiveLayer.FindShapes(Type:=cdrTextShape)
    sr.Combine

End Sub
