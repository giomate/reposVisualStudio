﻿Imports Inventor
Imports System

Public Class Curve3D
    Public oDoc As PartDocument
    Public sk3D As Sketch3D
    Dim line3D As SketchLines3D
    Dim curve As SketchEquationCurve3D
    Public Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Double
        Public Dmax As Double
        Public Dmin As Double

    End Structure
    Dim DP As DesignParam
    Dim Tr As Double
    Dim Cr As Double
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
        'oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.p, UnitsTypeEnum.kUnitlessUnits, "p")
        ' oDoc.ComponentDefinition.Parameters.ReferenceParameters.AddByValue(DP.b, UnitsTypeEnum.kMillimeterLengthUnits, "b")
    End Sub
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

        sk.Edit()
        Dim r, z As String
        Dim t As Double = 2 * Math.PI * q / (DP.q)
        r = String.Concat(Tr.ToString() & " + " & Cr.ToString() & "mm * cos( t * q * 1rad )")
        z = String.Concat("- " & Cr.ToString() & "mm * sin( t * q * 1rad )")
        curve = sk.SketchEquationCurves3D.Add(CoordinateSystemTypeEnum.kCylindrical, r, " t * 1rad * p", z, -6 * Math.PI / (DP.q * 4) + t, 6 * Math.PI / (DP.q * 4) + t)
        'sk3D.SketchEquationCurves3D.Add(CoordinateSystemTypeEnum.kCylindrical, String.Concat(Tr.ToString & " + " & Cr.ToString & " * cos(" & DP.q.ToString & " * t * 1rad)"), String.Concat(" t * 1rad * " & DP.p.ToString), String.Concat("-" & Cr.ToString & " * sin(" & DP.q.ToString & " * t * 1rad)"), -4 * Math.PI / (DP.p * 4), 4 * Math.PI / (DP.p * 4))
        curve.Construction = True
        sk.ExitEdit()

        Return curve
    End Function

End Class
