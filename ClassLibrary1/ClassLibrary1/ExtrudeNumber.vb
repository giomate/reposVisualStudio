Imports Inventor

Public Sub ExtrudeNumber(value As Integer, oPartCompDef As PartComponentDefinition, wp1 As WorkPlane, face As Integer)


    Dim oSketch As Sketch
    oSketch = oPartCompDef.Sketches.Add(wp1)

    Dim numero As Integer

    numero = value

    Dim oTG As TransientGeometry
    oTG = ThisApplication.TransientGeometry
    Dim positionNumber As Point2d


    If face = 1 Then
        positionNumber = oTG.CreatePoint2d(-0.2, -2.5)
    Else
        positionNumber = oTG.CreatePoint2d(2, 4)
    End If
    oSketch.Edit


    Dim oTextBox As TextBox
    Dim oStyle As TextStyle




    Dim sText As String
    sText = CStr(numero)

    oTextBox = oSketch.TextBoxes.AddFitted(positionNumber, sText)
    oStyle = oSketch.TextBoxes.Item(1).Style
    oStyle.FontSize = 0.81197733879089
    oStyle.Bold = True
    oTextBox = oSketch.TextBoxes.AddFitted(positionNumber, sText, oStyle)
    Dim x As Double
    x = 4 * Atn(1) * 1.5
    oTextBox.Rotation = x






    oSketch.ExitEdit

    Dim oPaths As ObjectCollection
    oPaths = ThisApplication.TransientObjects.CreateObjectCollection
    oPaths.Add(oTextBox)

    Dim oProfile As Profile
    oProfile = oSketch.Profiles.AddForSolid(False, oPaths)

    Dim oExtrudeDef As ExtrudeDefinition
    oExtrudeDef = oPartCompDef.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, kJoinOperation)
    Call oExtrudeDef.SetDistanceExtent(0.12, kPositiveExtentDirection)
    Dim oExtrude As ExtrudeFeature
    oExtrude = oPartCompDef.Features.ExtrudeFeatures.Add(oExtrudeDef)


End Sub