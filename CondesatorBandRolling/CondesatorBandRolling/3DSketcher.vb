Imports Inventor

Public Class Class3DSketch
    Public Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Integer
        Public Dmax As Integer
        Public Dmin As Integer

    End Structure

    Dim Tr As Double
    Dim Cr As Double
    Dim oApp As Inventor.Application
    Dim oSk3D As Sketch3D
    Dim DP As DesignParam
    Dim oTransGeom As TransientGeometry
    Dim oSketchLine As SketchLine3D


    Public Sub New()


        DP.p = 13
        DP.q = 31
        DP.b = 25
        DP.Dmax = 200
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        oTransGeom = oApp.TransientGeometry


    End Sub




    Function DrawFirstFace(oDoc As PartDocument) As Sketch3D
        oApp = oDoc.Parent
        oSk3D = oDoc.ComponentDefinition.Sketches3D.Add
        oSketchLine = bendingLine(1)

    End Function
    Function bendingLine(n As Integer) As SketchLine3D
        Return oSk3D.SketchLines3D.AddByTwoPoints(torousPoint(0), torousPoint(n * 4 * Math.PI / (DP.p * DP.q)))

    End Function
    Function getPoints12(oSk3D As Sketch3D) As Point[2]

        
        Dim point1 As Point
        point1 = oTransGeom.CreatePoint()

        Return 0
    End Function

    Function torousPoint(phi As Double) As Point
        Dim x As Double
        Dim y As Double
        Dim z As Double
        x = (Tr + Cr * Math.Cos(phi * DP.q)) * Math.Cos(phi * DP.p)
        y = (Tr + Cr * Math.Cos(phi * DP.q)) * Math.Sin(phi * DP.p)
        z = (Cr * Math.Sin(phi * DP.q))

        Return oTransGeom.CreatePoint(x, y, z)
    End Function




    Function getDoc(oApp As Inventor.Application) As PartDocument
        Return oApp.ActiveDocument
    End Function

    Function is3DActive(oApp As Inventor.Application) As Boolean
        If Not TypeOf oApp.ActiveEditObject Is Sketch3D Then
            MsgBox("A 3d sketch must be active.")
            Return False
        Else
            Return True
        End If

    End Function



End Class
