Imports Inventor
Imports System

Public Class Curves3D
    Public oDoc As PartDocument
    Public sk3D As Sketch3D
    Dim line3D As SketchLines3D
    Public curve As SketchEquationCurve3D
    Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Double
        Public Dmax As Double
        Public Dmin As Double

    End Structure
    Public DP As DesignParam
    Public Tr As Double
    Public Cr As Double
    Public Sub New(docu As Inventor.Document)
        oDoc = docu
        DP.Dmax = 200
        DP.Dmin = 1
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 13
        DP.q = 31
        DP.b = 25
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.q, UnitsTypeEnum.kUnitlessUnits, "q")
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.p, UnitsTypeEnum.kUnitlessUnits, "p")
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.b, UnitsTypeEnum.kMillimeterLengthUnits, "b")
    End Sub
    Public Function DefineTrobinaParameters(docu As Inventor.Document) As Parameter
        Dim p As Parameter
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.q, UnitsTypeEnum.kUnitlessUnits, "q")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.p, UnitsTypeEnum.kUnitlessUnits, "p")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.b / 10, UnitsTypeEnum.kMillimeterLengthUnits, "b")
        p = oDoc.ComponentDefinition.Parameters.ReferenceParameters.Item(oDoc.ComponentDefinition.Parameters.ReferenceParameters.Count)
        Return p
    End Function
    Public Function DrawTrobinaCurveAlone() As Sketch3D
        sk3D = oDoc.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = "trobinaCurve"
        DrawTrobinaCurve(sk3D)

        Return sk3D
    End Function
    Public Function DrawTrobinaCurve(sk As Sketch3D) As SketchEquationCurve3D

        Return DrawTrobinaCurve(sk, 0)
    End Function
    Public Function DrawTrobinaCurve(sk As Sketch3D, q As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(sk, q, 1.0)
    End Function
    Public Function DrawTrobinaCurve(sk As Sketch3D, q As Integer, d As Double) As SketchEquationCurve3D

        sk.Edit()
        Dim r, z As String
        Dim t As Double = 2 * Math.PI * q / (DP.q)
        r = String.Concat(Tr.ToString() & " + " & Cr.ToString() & "mm * cos( t * q * 1rad )")
        z = String.Concat("- " & Cr.ToString() & "mm * sin( t * q * 1rad )")
        curve = sk.SketchEquationCurves3D.Add(CoordinateSystemTypeEnum.kCylindrical, r, " t * 1rad * p", z, -4 * Math.PI / (DP.p * 4) * d + t, 4 * Math.PI / (DP.p * 4) * d + t)
        curve.Construction = True
        sk.ExitEdit()

        Return curve
    End Function
    Public Function DrawLowerRing(sk As Sketch3D) As SketchEquationCurve3D


        Dim r, z As String
        Dim t As Double = Math.PI
        r = String.Concat(Tr.ToString())
        z = String.Concat("- " & Cr.ToString())
        curve = sk.SketchEquationCurves3D.Add(CoordinateSystemTypeEnum.kCylindrical, r, " t * 1rad", z, -1 * t, t)
        curve.Construction = True


        Return curve
    End Function
    Public Function DrawTrobinaCurve(q As Integer, s As String) As SketchEquationCurve3D

        sk3D = oDoc.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = DrawTrobinaCurve(sk3D, q)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
    Public Function DrawTrobinaCurve(q As Integer, s As String, d As Double) As SketchEquationCurve3D

        sk3D = oDoc.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = DrawTrobinaCurve(sk3D, q, d)
        sk3D.GeometricConstraints3D.AddGround(curve)
        Return curve
    End Function
End Class
