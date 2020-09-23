Attribute VB_Name = "modGradient"
Option Explicit

Public Type ColorGradient
    Color As Long
    Place As Single
End Type

Public Enum GradientTypes
    Linear = 0
    Radial = 1
End Enum

Public Type GradientFiles
    Points() As ColorGradient
    FillType As GradientTypes
End Type

Public objWorkSpace As Workspace
Public GradientType As GradientTypes
Public PointColors() As ColorGradient
Public CustomGradient() As GradientFiles


Public Sub PaintGradient(Picture As PictureBox, Pointers() As ColorGradient, FillType As GradientTypes, Optional ByVal Angle As Single, Optional ByVal X As Long, Optional ByVal Y As Long)

Dim Alfa As Double
Dim InitialColor As Long
Dim ResultColor As Long
Dim Steps As Long
Dim Start As Long
Dim intCount As Integer
Dim intCount2 As Integer
Dim intCount3 As Integer
Dim NewRed As Double
Dim NewGreen As Double
Dim NewBlue As Double
Dim Colors() As ColorGradient
Dim intLowness As Integer
Dim intIndex As Integer
Dim ColorPointers() As ColorGradient
Dim Correction As RGBColor
Dim OrigScaleMode As Integer
Dim AngleDif As Single
Dim Temp As Long
Dim New_X, New_Y As Long
Dim MaxDis As Long
Dim intCorr As Long

    OrigScaleMode = Picture.ScaleMode
    Picture.ScaleMode = 3 ' pixel
    Alfa = Angle
    ColorPointers = Pointers
    Picture.Cls
    If UBound(ColorPointers) - LBound(ColorPointers) = 0 Then
        Picture.BackColor = ColorPointers(UBound(ColorPointers)).Color
    End If
    
    ReDim Colors(LBound(ColorPointers) To UBound(ColorPointers) - 1)
    
    'order the pointers so they are in ascent order
    For intCount = LBound(Colors) To UBound(Colors)
        intLowness = 101
        For intCount2 = LBound(ColorPointers) To UBound(ColorPointers) - 1
            If ColorPointers(intCount2).Place < intLowness Then
                intLowness = ColorPointers(intCount2).Place
                intIndex = intCount2
            End If
        Next
        Colors(intCount) = ColorPointers(intIndex)
        ColorPointers(intIndex) = ColorPointers(UBound(ColorPointers) - 1)
        ReDim Preserve ColorPointers(LBound(ColorPointers) To UBound(ColorPointers) - 1)
        
    Next
            
    If FillType = Linear Then
        
        'if the angle > 180 then the colors and places are inverted
        'this becous the angle functions in VB work with radial not degree
        If Alfa Mod 360 >= 180 Then
            For intCount = LBound(Colors) To Round((UBound(Colors) - LBound(Colors)) / 2 + 0.01)
                Temp = Colors(intCount).Color
                Colors(intCount).Color = Colors(UBound(Colors) - (intCount - LBound(Colors))).Color
                Colors(UBound(Colors) - (intCount - LBound(Colors))).Color = Temp
                
                Temp = Colors(intCount).Place
                Colors(intCount).Place = 100 - Colors(UBound(Colors) - (intCount - LBound(Colors))).Place
                Colors(UBound(Colors) - (intCount - LBound(Colors))).Place = 100 - Temp
            Next
            
            If (UBound(Colors) - LBound(Colors) + 1) Mod 2 <> 0 Then
                Colors(UBound(Colors) - LBound(Colors)).Place = 100 - Colors(UBound(Colors) - LBound(Colors)).Place
            End If
            
        End If
        
        If Colors(LBound(Colors)).Place > 0 Then
            ReDim Preserve Colors(LBound(Colors) To UBound(Colors) + 1)
            For intCount = UBound(Colors) - 1 To LBound(Colors) Step -1
                Colors(intCount + 1) = Colors(intCount)
            Next
            Colors(LBound(Colors)).Color = Colors(LBound(Colors) + 1).Color
            Colors(LBound(Colors)).Place = 0
        End If

        If Colors(UBound(Colors)).Place < 100 Then
            ReDim Preserve Colors(LBound(Colors) To UBound(Colors) + 1)
            Colors(UBound(Colors)).Color = Colors(UBound(Colors) - 1).Color
            Colors(UBound(Colors)).Place = 100
        End If
        
        If Alfa Mod 180 <= 45 Then
            AngleDif = Picture.ScaleHeight * Tan((Alfa Mod 180) * Trans)
            If Alfa > 90 And Alfa < 180 Then intCorr = Abs(AngleDif)
            
            For intCount = LBound(Colors) To UBound(Colors) - 1
                
                Start = (Picture.ScaleWidth + Abs(AngleDif)) / 100 * Colors(intCount).Place
                Steps = (Picture.ScaleWidth + Abs(AngleDif)) / 100 * Colors(intCount + 1).Place - Start - 1
                
                InitialColor = Colors(intCount).Color
                ResultColor = Colors(intCount + 1).Color
                
                NewRed = ((DefineRGB(ResultColor).Red - DefineRGB(InitialColor).Red) / Steps)
                NewGreen = ((DefineRGB(ResultColor).Green - DefineRGB(InitialColor).Green) / Steps)
                NewBlue = ((DefineRGB(ResultColor).Blue - DefineRGB(InitialColor).Blue) / Steps)
                
                Correction.Red = DefineRGB(InitialColor).Red
                Correction.Green = DefineRGB(InitialColor).Green
                Correction.Blue = DefineRGB(InitialColor).Blue
                
                For intCount2 = 0 To Steps
                    Picture.Line (Start + intCount2 - intCorr, 0)-(Start + intCount2 - Round(AngleDif) + intCorr, Picture.ScaleHeight), RGB(Correction.Red + NewRed * intCount2, Correction.Green + NewGreen * intCount2, Correction.Blue + NewBlue * intCount2)
                Next
                
            Next
        Else
            '------------------------------------
            Alfa = Alfa - 90
            
            AngleDif = Picture.ScaleWidth * Tan((Alfa) * Trans)
            If Alfa > 0 And Alfa < 90 Or Alfa > 180 Then intCorr = Abs(AngleDif)
            
            For intCount = LBound(Colors) To UBound(Colors) - 1
                
                Start = (Picture.ScaleHeight + Abs(AngleDif)) / 100 * Colors(intCount).Place
                Steps = (Picture.ScaleHeight + Abs(AngleDif)) / 100 * Colors(intCount + 1).Place - Start - 1
                
                InitialColor = Colors(intCount).Color
                ResultColor = Colors(intCount + 1).Color
                
                NewRed = ((DefineRGB(ResultColor).Red - DefineRGB(InitialColor).Red) / Steps)
                NewGreen = ((DefineRGB(ResultColor).Green - DefineRGB(InitialColor).Green) / Steps)
                NewBlue = ((DefineRGB(ResultColor).Blue - DefineRGB(InitialColor).Blue) / Steps)
                
                Correction.Red = DefineRGB(InitialColor).Red
                Correction.Green = DefineRGB(InitialColor).Green
                Correction.Blue = DefineRGB(InitialColor).Blue
                
                For intCount2 = 0 To Steps
                    
                    Picture.Line (0, Start + intCount2 - intCorr)-(Picture.ScaleWidth, intCount2 + Start - Abs(AngleDif) + intCorr), RGB(Correction.Red + NewRed * intCount2, Correction.Green + NewGreen * intCount2, Correction.Blue + NewBlue * intCount2)
                    
                Next
            Next
            '---------------------------------------------
        End If
    ElseIf FillType = Radial Then
        
        New_X = IIf(X = 0, (Picture.ScaleWidth - 1) / 2, X)
        New_Y = IIf(Y = 0, (Picture.ScaleHeight - 1) / 2, Y)
        
        MaxDis = IIf(New_X > (Picture.ScaleWidth - New_X), New_X, Picture.ScaleWidth - New_X)
        Temp = IIf(New_Y > (Picture.ScaleHeight - New_Y), New_Y, Picture.ScaleHeight - New_Y)
        If Temp > MaxDis Then MaxDis = Temp
        Picture.FillStyle = 0
        
        Steps = Sqr(Picture.ScaleWidth ^ 2 + Picture.ScaleHeight ^ 2)
        Picture.FillColor = Colors(LBound(Colors)).Color
        Picture.Circle (New_X, New_Y), Steps, Colors(UBound(Colors)).Color
        
        For intCount = LBound(Colors) To UBound(Colors) - 1
            
            Start = (MaxDis / 100) * (100 - Colors(intCount + 1).Place)
            Steps = ((MaxDis / 100) * (100 - Colors(intCount).Place)) - Start
            
            InitialColor = Colors(intCount).Color
            ResultColor = Colors(intCount + 1).Color
            
            NewRed = ((DefineRGB(ResultColor).Red - DefineRGB(InitialColor).Red) / (Steps))
            NewGreen = ((DefineRGB(ResultColor).Green - DefineRGB(InitialColor).Green) / (Steps))
            NewBlue = ((DefineRGB(ResultColor).Blue - DefineRGB(InitialColor).Blue) / (Steps))
            
            If NewRed <= 0 Or DefineRGB(InitialColor).Red < DefineRGB(ResultColor).Red Then Correction.Red = DefineRGB(InitialColor).Red
            If NewGreen <= 0 Or DefineRGB(InitialColor).Green < DefineRGB(ResultColor).Green Then Correction.Green = DefineRGB(InitialColor).Green
            If NewBlue <= 0 Or DefineRGB(InitialColor).Blue < DefineRGB(ResultColor).Blue Then Correction.Blue = DefineRGB(InitialColor).Blue
            
            For intCount2 = 0 To Steps
                Picture.FillColor = RGB(Correction.Red + NewRed * intCount2, Correction.Green + NewGreen * intCount2, Correction.Blue + NewBlue * intCount2)
                Picture.Circle (New_X, New_Y), (Steps - intCount2) + Start, Picture.FillColor
            Next
        Next
        
    End If
    
    Picture.ScaleMode = OrigScaleMode
    
End Sub

Public Sub SaveGradient(Location As String, Gradients() As GradientFiles)
Dim objDTB As Database
Dim objRec As Recordset
Dim objTable As TableDef
Dim intCount, intCount2 As Integer

    ' Make sure there isn't already a file with the name of
    ' the new database.
    If Dir(Location) <> "" Then Kill Location

    ' Create a new database
    Set objDTB = objWorkSpace.CreateDatabase(Location, dbLangGeneral, dbEncrypt)
        
    For intCount = LBound(Gradients) To UBound(Gradients)
        
        ' Create a new TableDef object.
        Set objTable = objDTB.CreateTableDef("Point " & intCount)
    
        With objTable
            ' Create fields and append them to the new TableDef
            ' object. This must be done before appending the
            ' TableDef object to the TableDefs collection
            
            .Fields.Append .CreateField("Color", dbLong)
            .Fields.Append .CreateField("Place", dbLong)
            
            ' Append the new TableDef object to the database.
            objDTB.TableDefs.Append objTable
            
            Set objRec = objDTB.OpenRecordset("Point " & intCount, dbOpenDynaset)
            
            For intCount2 = LBound(Gradients(1).Points) To UBound(Gradients(1).Points) - 1
            
                objRec.AddNew
                objRec("Color") = Gradients(intCount).Points(intCount2).Color
                objRec("Place") = Gradients(intCount).Points(intCount2).Place
                objRec.Update
                
            Next
            
        End With
    
    Next
    
    Set objTable = objDTB.CreateTableDef("FillTypes")
    
    With objTable
        ' Create fields and append them to the new TableDef
        ' object. This must be done before appending the
        ' TableDef object to the TableDefs collection
        .Fields.Append .CreateField("FillType", dbLong)
    End With
    
    ' Append the new TableDef object to the database.
    objDTB.TableDefs.Append objTable
            
    Set objRec = objDTB.OpenRecordset("FillTypes", dbOpenDynaset)
    
    For intCount = LBound(Gradients) To UBound(Gradients)
            
        objRec.AddNew
        objRec("FillType") = Gradients(intCount).FillType
        objRec.Update
        
    Next
    
    objDTB.Close

End Sub

Public Function LoadGradient(Location As String, TempGradient() As GradientFiles)
Dim objDTB As Database
Dim objRec As Recordset
Dim objTable As TableDef
Dim intCount, intCount2 As Integer

    ' Open database
    Set objDTB = objWorkSpace.OpenDatabase(Location, dbEncrypt)
        
    For Each objTable In objDTB.TableDefs
        If Left$(objTable.Name, 5) = "Point" Then
            
            intCount = intCount + 1
            
            ReDim Preserve TempGradient(1 To intCount)
                
                Set objRec = objTable.OpenRecordset()
                
                For intCount2 = 1 To objTable.RecordCount
                    
                    objRec.Move intCount2 - 1
                    ReDim Preserve TempGradient(intCount).Points(1 To intCount2)
                    
                    TempGradient(intCount).Points(intCount2).Color = objRec.Fields("Color").Value
                    TempGradient(intCount).Points(intCount2).Place = objRec.Fields("Place").Value
                
                Next
        
        ReDim Preserve TempGradient(intCount).Points(1 To intCount2)
        
        End If
        
    Next

    Set objTable = objDTB.TableDefs("FillTypes")

    For intCount = 1 To objTable.RecordCount
        
        Set objRec = objTable.OpenRecordset()
        objRec.Move intCount - 1
        TempGradient(intCount).FillType = objRec.Fields("FillType").Value
        
    Next

    objDTB.Close
    
End Function

Public Sub AddToCustom(Colors() As ColorGradient, FillType As GradientTypes)
    On Error GoTo Errorhandler
    ReDim Preserve CustomGradient(1 To UBound(CustomGradient) + 1)
1:  CustomGradient(UBound(CustomGradient)).Points = Colors
    CustomGradient(UBound(CustomGradient)).FillType = FillType
    Load frmGradient.picIcon(UBound(CustomGradient))
    
    frmGradient.picIcon(UBound(CustomGradient)).Left = _
      ((UBound(CustomGradient) - 1) Mod 7) * (frmGradient.picIcon(UBound(CustomGradient)).Width + 60) + 60
    frmGradient.picIcon(UBound(CustomGradient)).Top = _
      ((UBound(CustomGradient) - 1) \ 7) * (frmGradient.picIcon(UBound(CustomGradient)).Height + 60) + 60
    frmGradient.picIconHolder.Left = frmGradient.picIcon(UBound(CustomGradient)).Left - 36
    frmGradient.picIconHolder.Top = frmGradient.picIcon(UBound(CustomGradient)).Top - 36
    frmGradient.picIconHolder.ZOrder 1
    frmGradient.picIconHolder.Visible = True
    PaintGradient frmGradient.picIcon(UBound(CustomGradient)), Colors, FillType
    frmGradient.picIcon(UBound(CustomGradient)).Visible = True
    frmGradient.intSelCustom = UBound(CustomGradient)
    
    If frmGradient.picIcon(UBound(CustomGradient)).Top + _
     frmGradient.picIcon(UBound(CustomGradient)).Height + 120 > frmGradient.picCustom.Height Then
        frmGradient.picCustom.Height = frmGradient.picIcon(UBound(CustomGradient)).Top + _
        frmGradient.picIcon(UBound(CustomGradient)).Height + 120
        frmGradient.vscrCustom.Min = 0
        frmGradient.vscrCustom.Max = frmGradient.vscrCustom.Height - frmGradient.picCustom.Height
        frmGradient.vscrCustom.Value = frmGradient.vscrCustom.Max
        frmGradient.vscrCustom.Enabled = True
        frmGradient.picCustom.Top = frmGradient.vscrCustom.Top + frmGradient.vscrCustom.Max
    End If
    
Exit Sub
Errorhandler:
    If Err.Number = 9 Then 'Subscript out of range
        ReDim CustomGradient(1 To 1)
        GoTo 1
    End If
End Sub

Public Sub LoadCustom(GradientPoints() As GradientFiles)
Dim intCount As Integer
    frmGradient.NewCustom
    For intCount = LBound(GradientPoints) To UBound(GradientPoints)
        AddToCustom GradientPoints(intCount).Points, GradientPoints(intCount).FillType
    Next
    frmGradient.picIcon_Click 1
End Sub


