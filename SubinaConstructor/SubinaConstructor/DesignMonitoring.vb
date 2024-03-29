﻿Imports Inventor

Public Class DesignMonitoring
    Dim partDoc As PartDocument
    Dim Problems As Boolean
    Public sickFeature As PartFeature
    Public sickDimension As DimensionConstraint3D

    Public Sub New(docu As Inventor.Document)
        partDoc = docu
    End Sub

    Public Sub TestProblems()

        If PartHasProblems(partDoc) Then
            MsgBox("Part has problems")
        Else
            MsgBox("Part is OK")
        End If
    End Sub

    Public Function PartHasProblems(partDoc As PartDocument) As Boolean
        Dim feature As PartFeature
        For Each feature In partDoc.ComponentDefinition.Features
            If feature.HealthStatus <> HealthStatusEnum.kUpToDateHealth And
               feature.HealthStatus <> HealthStatusEnum.kBeyondStopNodeHealth And
               feature.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
                sickFeature = feature
                PartHasProblems = True
                Exit Function

            End If
        Next

        PartHasProblems = False
    End Function
    Public Function PartHasProblems() As Boolean


        Return partHasProblems(partDoc)

    End Function
    Public Function AreDimensionsHealthy(sketch As Sketch3D) As Boolean

        For Each dimension As DimensionConstraint3D In sketch.DimensionConstraints3D
            If dimension.Parameter.HealthStatus <> HealthStatusEnum.kUpToDateHealth And
               dimension.Parameter.HealthStatus <> HealthStatusEnum.kBeyondStopNodeHealth And
               dimension.Parameter.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
                sickDimension = dimension
                AreDimensionsHealthy = False
                Exit Function

            End If
        Next
        AreDimensionsHealthy = True
    End Function
    Public Function AreDimensions2DHealthy(ps As PlanarSketch) As Boolean

        For Each dimension As DimensionConstraint In ps.DimensionConstraints
            If dimension.Parameter.HealthStatus <> HealthStatusEnum.kUpToDateHealth And
               dimension.Parameter.HealthStatus <> HealthStatusEnum.kBeyondStopNodeHealth And
               dimension.Parameter.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
                sickDimension = dimension
                AreDimensions2DHealthy = False
                Exit Function

            End If
        Next
        AreDimensions2DHealthy = True
    End Function
    Public Function AreConstrainsHealthy(sketch As Sketch3D) As Boolean
        For Each constrain As GeometricConstraint3D In sketch.GeometricConstraints3D
            If Not constrain.Deletable Then
                AreConstrainsHealthy = False
                Exit Function

            End If
        Next
            AreConstrainsHealthy = True
    End Function
    Public Function IsFeatureHealthy(feature As PartFeature) As Boolean

        If feature.HealthStatus <> HealthStatusEnum.kUpToDateHealth And
               feature.HealthStatus <> HealthStatusEnum.kBeyondStopNodeHealth And
               feature.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
            sickFeature = feature
            Return False
        Else
            Return True
        End If

    End Function
    Public Function IsDerivedPartHealthy(feature As DerivedPartComponent) As Boolean

        If feature.HealthStatus <> HealthStatusEnum.kUpToDateHealth And
               feature.HealthStatus <> HealthStatusEnum.kBeyondStopNodeHealth And
               feature.HealthStatus <> HealthStatusEnum.kSuppressedHealth Then
            sickFeature = feature
            Return False
        Else
            Return True
        End If

    End Function
    Function IsSketch3dhHealthy(sk3D As Sketch3D) As Boolean
        Dim b As Boolean
        If sk3D.HealthStatus = HealthStatusEnum.kOutOfDateHealth Or
     sk3D.HealthStatus = HealthStatusEnum.kUpToDateHealth Then
            If AreDimensionsHealthy(sk3D) Then
                b = True

            Else
                b = False

            End If
        Else
            b = False
        End If
        Return b
    End Function
    Function IsSketch2DhHealthy(ps As PlanarSketch) As Boolean
        Dim b As Boolean
        If ps.HealthStatus = HealthStatusEnum.kOutOfDateHealth Or
     ps.HealthStatus = HealthStatusEnum.kUpToDateHealth Then
            If AreDimensions2DHealthy(ps) Then
                b = True

            Else
                b = False

            End If
        Else
            b = False
        End If
        Return b
    End Function

End Class
