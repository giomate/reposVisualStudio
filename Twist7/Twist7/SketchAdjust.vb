﻿Imports Inventor


Public Class SketchAdjust

    Dim oApp As Inventor.Application
    Dim oDesignProjectMgr As DesignProjectManager
    Dim oPartDoc As PartDocument
    Dim oSk3D As Sketch3D
    Dim delta, gain, sp, obj, SetpointCorrector As Double
    Dim oTheta, oGap, oTecho, p, k, posible As Parameter
    Dim resolution, maxRes, minRes As Integer
    Dim dimension As DimensionConstraint3D
    Dim kapput, prio As Boolean
    Dim variables() As String = {"angulo", "techo", "gap1", "gap2", "foldez", "doblez"}
    Dim ErrorOptimizer(variables.Length) As Double
    Dim counter As Integer
    Dim monitor As DesignMonitoring
    Dim comando As Commands


    Public Structure DesignParam
        Public p As Integer
        Public q As Integer
        Public b As Integer
        Public Dmax As Double
        Public Dmin As Double

    End Structure

    Dim DP As DesignParam
    Dim Tr As Double
    Dim Cr As Double
    Public Sub New(docu As Inventor.Document)
        oApp = docu.Parent
        oDesignProjectMgr = oApp.DesignProjectManager
        oPartDoc = docu
        monitor = New DesignMonitoring(oPartDoc)
        DP.Dmax = 200
        DP.Dmin = 32
        Tr = (DP.Dmax + DP.Dmin) / 4
        Cr = (DP.Dmax - DP.Dmin) / 4
        DP.p = 17
        DP.q = 37
        k = Nothing
        prio = False
        ErrorOptimizer.Initialize()
        resolution = 100
        minRes = 4
        maxRes = 1000000
        counter = 0
        comando = New Commands(oApp)

    End Sub
    Function UpdateDocu(d As PartDocument) As Integer
        oPartDoc = d
        oSk3D = d.ComponentDefinition.Sketches3D.Item(d.ComponentDefinition.Sketches3D.Count)
        Return True
    End Function
    Function openFile(fileName As String) As PartDocument
        Try
            If oApp.Documents.Count > 0 Then
                If Not (oApp.ActiveDocument.FullFileName = createFileName(fileName)) Then
                    oPartDoc = oApp.Documents.Open(fileName, True)
                End If
            Else
                oPartDoc = oApp.Documents.Open(fileName, True)
            End If

            oPartDoc = oApp.ActiveDocument
        Catch ex3 As Exception
            Debug.Print(ex3.ToString())
            Debug.Print("Unable to find Document")
        End Try

        ' Conversions.SetUnitsToMetric(oPartDoc)
        Return oPartDoc
    End Function

    Function createFileName(fileName As String) As String
        Dim strFilePath As String
        strFilePath = oDesignProjectMgr.ActiveDesignProject.WorkspacePath
        ' Dim strFileName As String
        'strFileName = "Embossed" & CStr(I) & ".ipt"
        Dim strFullFileName As String
        strFullFileName = strFilePath & "\" & fileName
        Return strFullFileName
    End Function

    Function adjust(fileName As String, theta As Double) As Double
        openFile(createFileName(fileName))
        'calculateGain(theta, "angulo")
        sp = theta
        Try
            changeParameter(openMainSketch(oPartDoc))
            checkBuilder()
            Debug.Print("done!!")

        Catch ex As Exception
            Call UndoCommand()
            makeallDriven()
            Debug.Print(ex.ToString())
        End Try

        Return oTheta._Value
    End Function
    Function openMainSketch(oDoc As PartDocument) As Sketch3D

        oSk3D = oDoc.ComponentDefinition.Sketches3D.Item("MainSketch")
        Return oSk3D
    End Function
    Public Sub changeParameter(oSk3D As Sketch3D)
        Try

            oTheta = getParameter("angulo")

            oSk3D.Edit()
            checkOtherVariables("angulo")
            calculateGain(sp, "angulo")

            While (objetiveZero() > sp / resolution)
                p = iterate("angulo", sp)
                While kapput
                    checkBuilder()
                    If k.Name = Nothing Then
                        checkGaps(p.Name)
                    Else
                        checkGaps(k.Name)
                    End If

                End While
                checkOtherVariables("angulo")
                calculateGain(sp, "angulo")




            End While

            oSk3D.Solve()
            oSk3D.ExitEdit()

        Catch ex4 As Exception
            Call UndoCommand()
            makeallDriven()
            If resolution < maxRes Then
                resolution = resolution * 2
            End If
            Debug.Print(ex4.ToString())
            Debug.Print("Fail adjusting " & oTheta.Name & " ...last value:" & oTheta.Value.ToString)
            checkOtherVariables(oTheta.Name)
            Exit Sub
        End Try


    End Sub

    Public Sub makeallDriven()
        GetDimension("techo").Driven = True
        GetDimension("doblez").Driven = True
        GetDimension("foldez").Driven = True
        GetDimension("gap2").Driven = True
        GetDimension("gap1").Driven = True
        'getDimension("angulo").Driven = True



    End Sub
    Public Sub makeallConstrained()
        For Each variable As String In variables
            GetDimension(variable).Driven = False
        Next



    End Sub
    Public Function getParameter(name As String) As Parameter
        Try
            p = oPartDoc.ComponentDefinition.Parameters.ModelParameters.Item(name)
        Catch ex As Exception
            Try
                p = oPartDoc.ComponentDefinition.Parameters.ReferenceParameters.Item(name)
            Catch ex1 As Exception
                Try
                    p = oPartDoc.ComponentDefinition.Parameters.UserParameters.Item(name)
                Catch ex2 As Exception
                    Debug.Print(ex2.ToString())
                    Debug.Print("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function

    Function calculateGain(setValue As Double, name As String) As Double
        Try
            p = getParameter(name)
            delta = (setValue - p._Value) / (setValue * resolution)
            gain = Math.Exp(delta)
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Fail Calculating " & p.Name)
        End Try

        Return gain
    End Function
    Function calculateGainForMinimun(setValue As Double, name As String) As Double
        Try
            p = getParameter(name)
            Dim k As Double
            delta = (setValue - p._Value) / (resolution)
            If GetDimension(name).Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
                k = 4
            Else
                k = 1
            End If
            gain = Math.Pow(Math.Exp(delta), 1 / k)
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Fail Calculating " & p.Name)
        End Try

        Return gain
    End Function
    Public Sub UndoCommand()
        Dim oCommandMgr As CommandManager
        oCommandMgr = oApp.CommandManager

        ' Get control definition for the line command. 
        Dim oControlDef As ControlDefinition
        oControlDef = oCommandMgr.ControlDefinitions.Item("AppUndoCmd")
        ' Execute the command. 
        Call oControlDef.Execute()
    End Sub
    Public Sub checkGaps(name As String)

        Try
            Select Case name
                Case "techo"
                    checkAngulos(name, 2.7, 3.1, Math.Max(2.7, getParameter(name)._Value * 0.9))
                Case "gap1"
                    checkDimension(name, 3, 10, 10 * Math.Max(0.3, getParameter(name)._Value / 2))
                Case "gap2"
                    checkDimension(name, 2, 9, 10 * Math.Max(0.2, getParameter(name)._Value / 2))
                Case "foldez"
                    checkDimension(name, getParameter("doblez")._Value * 10, 2, 10 * Math.Max(getParameter("doblez")._Value, getParameter(name)._Value / 2))
                Case "doblez"
                    checkDimension(name, 0.01, 0.9, 10 * Math.Max(0.001, getParameter(name)._Value / 2))
                    'kapput = False
            End Select



        Catch ex As Exception
            UndoCommand()
            makeallDriven()
            If resolution < maxRes Then
                resolution = resolution * 2
            End If
            Debug.Print(ex.ToString())
            Debug.Print("Fail Iteration  checkGaps: " & p.Value.ToString)
        End Try

    End Sub

    Public Function iterate(name As String, setpoint As Double) As Parameter
        Dim pit As Parameter
        pit = getParameter(name)

        Try
            calculateGain(setpoint, name)
            makeallDriven()
            GetDimension(name).Driven = False
            While ((Math.Abs(delta * resolution)) > (setpoint / resolution) And (Not kapput))
                pit.Value = pit.Value * calculateGain(setpoint, name)
                Debug.Print("iterating  " & pit.Name & " = " & pit.Value.ToString)
                checkOtherVariables(pit.Name)
                checkBuilder()
                If resolution > minRes Then
                    resolution = resolution - 1
                    Debug.Print("Resolution:  " & resolution.ToString)
                End If

            End While
            Debug.Print("adjusting " & pit.Name & " = " & pit.Value.ToString)
            Debug.Print("Resolution:  " & resolution.ToString)
            If kapput Then
                If resolution < maxRes Then
                    resolution = resolution + 10
                End If
                UndoCommand()
                'getDimension("techo").Driven = False
                Return k
            Else
                If resolution > minRes Then
                    resolution = resolution - 1
                End If

            End If
            GetDimension("techo").Driven = False

        Catch ex As Exception
            UndoCommand()
            If resolution < maxRes Then
                resolution = resolution * 2
            End If
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)
            GetDimension(pit.Name).Driven = True
            checkOtherVariables(pit.Name)
            Return pit

        End Try

        Return p
    End Function
    Public Function AdjustDimensionSmothly(name As String, setpoint As Double) As Boolean
        Dim pit As Parameter
        Dim b As Boolean = False
        Dim c, r As Double
        Dim dc As DimensionConstraint3D
        pit = getParameter(name)

        Try
            calculateGain(setpoint, name)
            SetpointCorrector = AdjustResolution(name, setpoint)
            dc = GetDimension(name)
            dc.Driven = False
            While (((Math.Abs(delta * resolution * SetpointCorrector)) > (setpoint / resolution)) And (monitor.IsSketch3dhHealthy(oSk3D)))
                r = pit.Value * calculateGain(setpoint, name)
                pit.Value = r
                oPartDoc.Update2()
                AdjustResolution(name, pit._Value)
                Debug.Print("Adjusting " & pit.Name & " = " & pit.Value.ToString)
                counter = counter + 1

            End While
            counter = 0
            If Not monitor.IsSketch3dhHealthy(oSk3D) Then
                RecoveryUnhealthySketch(pit)
                Return False
            Else
                b = True
                Debug.Print("adjusted " & pit.Name & " = " & pit.Value.ToString)
                Debug.Print("Resolution:  " & resolution.ToString)
            End If





        Catch ex As Exception

            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)

            RecoveryUnhealthySketch(pit)
            Return False

        End Try

        Return b
    End Function
    Public Function AdjustGapSmothly(gap As DimensionConstraint3D, setpoint As Double, angle As DimensionConstraint3D) As Boolean
        Dim pit As Parameter
        Dim b As Boolean = False


        pit = getParameter(gap.Parameter.Name)

        Try
            calculateGain(setpoint, gap.Parameter.Name)
            SetpointCorrector = AdjustResolution(gap.Parameter.Name, setpoint) * 8
            While (((Math.Abs(delta * resolution * SetpointCorrector)) > (setpoint / resolution)) And (monitor.IsSketch3dhHealthy(oSk3D)) And (Not IsLastAngleOk(angle)))
                pit.Value = pit.Value * Math.Pow(calculateGain(setpoint, gap.Parameter.Name), 1 / 8)
                oPartDoc.Update2()
                AdjustResolution(gap.Parameter.Name, pit._Value)
                Debug.Print("iterating  " & pit.Name & " = " & pit.Value.ToString)


            End While

            If Not monitor.IsSketch3dhHealthy(oSk3D) Then
                RecoveryUnhealthySketch(pit)
                Return False
            Else
                If ((Math.Abs(delta * resolution * SetpointCorrector)) > (setpoint / resolution)) Then
                    If IsLastAngleOk(angle) Then
                        b = True
                    Else
                        angle.Driven = True
                        b = False
                    End If
                End If
            End If


            b = True
            Debug.Print("adjusted " & pit.Name & " = " & pit.Value.ToString)
            Debug.Print("Resolution:  " & resolution.ToString)

            Return b
        Catch ex As Exception
            UndoCommand()
            If resolution < maxRes Then
                resolution = resolution * 2
            End If
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)

            RecoveryUnhealthySketch(pit)
            Return False

        End Try

        Return b
    End Function

    Public Function GetMinimalDimension(dc As DimensionConstraint3D) As Boolean
        Dim pit As Parameter

        Try

            Dim b As Boolean = False
            Dim setPoint As Double = 0
            Dim climit As Integer = 8
            pit = getParameter(dc.Parameter.Name)
            calculateGainForMinimun(setPoint, dc.Parameter.Name)
            dc.Driven = False
            If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
                SetpointCorrector = Math.Pow(resolution, 4)
                climit = 16
            Else
                SetpointCorrector = 1

            End If
            While (((Math.Abs(delta * resolution) * SetpointCorrector) > (Math.Pow(10, -3) / resolution)) And (monitor.IsSketch3dhHealthy(oSk3D)) And counter < climit)
                pit.Value = pit.Value * calculateGainForMinimun(setPoint, dc.Parameter.Name)
                oPartDoc.Update2()
                counter = counter + 1
                Debug.Print("iterating Number:  " & counter.ToString() & "  " & pit.Name & " = " & pit.Value.ToString)
            End While

            If Not monitor.IsSketch3dhHealthy(oSk3D) Then
                RecoveryUnhealthySketch(pit)
                If monitor.IsSketch3dhHealthy(oSk3D) Then
                    If counter < climit Then
                        pit.Value = pit.Value / calculateGainForMinimun(setPoint, dc.Parameter.Name)
                        resolution = resolution * (1 + 1 / CDbl(climit))
                        GetMinimalDimension(dc)
                    Else
                        counter = 0
                        Return True
                    End If
                End If
            Else
                Return True
            End If
            b = True
            Debug.Print("adjusted " & pit.Name & " = " & pit.Value.ToString)
            Debug.Print("Resolution:  " & resolution.ToString)

            Return b
        Catch ex As Exception
            comando.UndoCommand()
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)

            RecoveryUnhealthySketch(pit)
            Return False

        End Try


    End Function
    Function IsLastAngleOk(ac As DimensionConstraint3D) As Boolean
        If (ac.Parameter._Value < 0.05 Or ac.Parameter._Value > (Math.PI - 0.05)) Then
            Return True
        End If
        Return False
    End Function
    Function RecoveryUnhealthySketch(p As Parameter) As Parameter
        Try

            While Not monitor.IsSketch3dhHealthy(oSk3D)
                comando.UndoCommand()
                posible = p

            End While
            Return p
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Fail Recovering " & p.Name & " ...last value:" & p.Value.ToString)
            Return RecoveryUnhealthySketch(p)
        End Try



    End Function
    Function AdjustResolution(name As String, s As Double) As Double
        Dim c, r As Double
        Dim dc As DimensionConstraint3D
        dc = GetDimension(name)
        If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
            c = s * dc.Parameter._Value * 128
            r = (1 / (s + Math.Pow(10, -6)) + resolution - 1) / ((counter / 128) + 1)
            If r < 32 Then
                resolution = 32
            Else
                resolution = r

            End If

        Else
            c = s * dc.Parameter._Value * 32
            r = (1 / (s + Math.Pow(10, -6)) + resolution - 1) / (counter + 1)
            If r < 4 Then
                resolution = 4
            Else
                resolution = r

            End If
        End If
        Return c
    End Function
    Public Sub checkOtherVariables(name As String)
        Dim variable As String
        prio = False
        For Each variable In variables
            If prio Then
                checkGaps(variable)
            ElseIf name = variable Then
                prio = True
                If name = variables.Last Then
                    kapput = False
                End If
            End If


        Next




    End Sub
    Public Sub checkBuilder()
        Try
            oSk3D.Solve()
            oSk3D.ExitEdit()
            oPartDoc.Update()
            oSk3D.Edit()

        Catch ex As Exception
            UndoCommand()
            'makeallDriven()
            'kapput = True
            k = p
            If resolution < maxRes Then
                resolution = resolution * 2
            End If

            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & p.Name & " ...last value:" & p.Value.ToString)
            GetDimension(k.Name).Driven = True
            checkOtherVariables(k.Name)
            Exit Sub
        End Try

    End Sub
    Public Function checkDimension(name As String, a As Double, b As Double, setpoint As Double) As Parameter
        p = getParameter(name)

        If (p._Value < a / 10 Or p._Value > b / 10) Then
            k = p
            GetDimension(name).Driven = False
            iterate(name, setpoint / 10)
            GetDimension(name).Driven = True
        Else
            kapput = False
        End If

        Return p
    End Function
    Public Function checkAngulos(name As String, a As Double, b As Double, setpoint As Double) As Parameter
        p = getParameter(name)

        If (p._Value < a Or p._Value > b) Then
            k = p
            GetDimension(name).Driven = False
            iterate(name, setpoint)
            GetDimension(name).Driven = True
        Else
            kapput = False
        End If

        Return p
    End Function

    Function GetDimension(name As String) As DimensionConstraint3D
        For Each dimension In oSk3D.DimensionConstraints3D
            If dimension.Parameter.Name = name Then
                Return dimension
            End If
        Next
        Return Nothing
    End Function

    Function objetiveZero() As Double
        obj = 0
        counter = 0
        For Each variable As String In variables
            obj = obj + Math.Pow(ErrorOptimizer(counter), 2) * Math.Exp(-counter)
            counter = counter + 1
        Next


        Return obj
    End Function

    Function AdjustLineLenghtSmothly(dc As LineLengthDimConstraint3D, v As Double) As Boolean

        Dim dName As String
        Dim b As Boolean = False

        dName = dc.Parameter.Name

        If oPartDoc.Update2() Then
            If UpdateDocu(oPartDoc) Then
                If AdjustDimensionSmothly(dName, v) Then
                    b = True
                End If
            End If
        End If



        Return b
    End Function
    Function AdjustTwoLineAngleSmothly(dc As TwoLineAngleDimConstraint3D, v As Double) As Boolean

        Dim dName As String
        Dim b As Boolean = False

        dName = dc.Parameter.Name

        If dc.Parameter._Value > Math.PI / 2 Then
            If AdjustDimensionSmothly(dName, v) Then
                b = True
            End If
        Else
            If AdjustDimensionSmothly(dName, v) Then
                b = True
            End If
        End If


        Return b
    End Function
    Function AdjustTwoPointsSmothly(dc As TwoPointDistanceDimConstraint3D, v As Double) As Boolean

        Dim dName As String
        Dim b As Boolean = False

        dName = dc.Parameter.Name

        If oPartDoc.Update2() Then
            If UpdateDocu(oPartDoc) Then
                If AdjustDimensionSmothly(dName, v) Then
                    b = True
                End If
            End If
        End If



        Return b
    End Function

    Public Function AdjustDimensionConstraint3DSmothly(dc As DimensionConstraint3D, v As Double) As Boolean
        Dim b As Boolean
        dc.Driven = False
        Select Case dc.Type
            Case ObjectTypeEnum.kLineLengthDimConstraint3DObject

                b = AdjustLineLenghtSmothly(dc, v)
            Case ObjectTypeEnum.kTwoPointDistanceDimConstraint3DObject

                b = AdjustTwoPointsSmothly(dc, v)
            Case ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject
                b = AdjustTwoLineAngleSmothly(dc, v)
            Case Else
                b = AdjustTwoPointsSmothly(dc, v)
        End Select
        Return b
    End Function


End Class
