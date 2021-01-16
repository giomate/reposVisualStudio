Imports Inventor
Imports System

Public Class Curves3D
    Public oDoc As PartDocument
    Public sk3D As Sketch3D
    Dim line3D As SketchLines3D
    Public curve As SketchEquationCurve3D
    Public direction As Integer
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
        DP.p = 11
        DP.q = 23
        DP.b = 25
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.q, UnitsTypeEnum.kUnitlessUnits, "q")
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.p, UnitsTypeEnum.kUnitlessUnits, "p")
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.b, UnitsTypeEnum.kMillimeterLengthUnits, "b")
    End Sub
    Public Function DefineTrobinaParameters(docu As Inventor.Document) As Parameter
        Dim p As Parameter
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(direction, UnitsTypeEnum.kUnitlessUnits, "direction")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.q, UnitsTypeEnum.kUnitlessUnits, "q")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.p, UnitsTypeEnum.kUnitlessUnits, "p")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.b / 10, UnitsTypeEnum.kMillimeterLengthUnits, "b")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.Dmax / 10, UnitsTypeEnum.kMillimeterLengthUnits, "DMax")
        docu.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.Dmin / 10, UnitsTypeEnum.kMillimeterLengthUnits, "DMin")

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
        Dim dir As Integer = GetParameter("direction")._Value
        Return DrawTrobinaCurve(sk, 0, dir)
    End Function
    Public Function DrawTrobinaCurve(sk As Sketch3D, q As Integer, direction As Integer) As SketchEquationCurve3D
        Dim fn As Integer
        Dim s() As String
        s = Strings.Split(sk.Name, "s")
        fn = CInt(s(1))
        Return DrawTrobinaCurve(sk, q, 1.0, fn, direction)
    End Function
    Public Function DrawTrobinaCurve(sk As Sketch3D, q As Integer, f As Integer, direction As Integer) As SketchEquationCurve3D


        Return DrawTrobinaCurve(sk, q, 1.0, f, direction)
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
    Public Function DrawTrobinaCurve(sk As Sketch3D, q As Integer, d As Double, f As Integer, direction As Integer) As SketchEquationCurve3D
        Dim g As Integer = 0
        If f > 6 Then
            If f > 7 Then
                If f > 8 Then
                    g = f - 4
                Else
                    g = f - 3
                End If
            Else
                g = f - 1
            End If

        End If
        sk.Edit()
        Dim r, z, y As String
        Dim t As Double = 2 * Math.PI * (q / (DP.q))
        r = String.Concat(Tr.ToString() & " + " & Cr.ToString() & "mm * cos( t * q * 1rad )")
        z = String.Concat("- " & Cr.ToString() & "mm * sin( t * q * 1rad )")
        If direction > 0 Then
            y = " t * 1rad * p"
        Else
            y = "- t * 1rad * p"
        End If

        curve = sk.SketchEquationCurves3D.Add(CoordinateSystemTypeEnum.kCylindrical, r, y, z, (f - 8) * Math.PI / (DP.p * 16) * d + t, (f + 2 + g * 4 / 3) * Math.PI / (DP.p * 16) * d + t)
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
        Dim dir As Integer = GetParameter("direction")._Value
        sk3D = oDoc.ComponentDefinition.Sketches3D.Add()
        sk3D.Name = s
        curve = DrawTrobinaCurve(sk3D, q, dir)
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
    Public Function GetParameter(name As String) As Parameter
        Dim p As Parameter = Nothing
        Try
            p = oDoc.ComponentDefinition.Parameters.ModelParameters.Item(name)
        Catch ex As Exception
            Try
                p = oDoc.ComponentDefinition.Parameters.ReferenceParameters.Item(name)
            Catch ex1 As Exception
                Try
                    p = oDoc.ComponentDefinition.Parameters.UserParameters.Item(name)
                Catch ex2 As Exception
                    MsgBox(ex2.ToString())
                    MsgBox("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function
End Class
