Imports Inventor


Public Class SketchAdjust

    Dim oApp As Inventor.Application
    Dim oDesignProjectMgr As DesignProjectManager
    Dim oPartDoc As PartDocument
    Dim oSk3D As Sketch3D
    Dim delta, gain, sp, obj, SetpointCorrector As Double
    Dim oTheta, oGap, oTecho, p, k, posible As Parameter
    Dim resolution, maxRes, minRes As Double
    Dim dimension As DimensionConstraint3D
    Dim kapput, prio As Boolean
    Dim variables() As String = {"angulo", "techo", "gap1", "gap2", "foldez", "doblez"}
    Dim ErrorOptimizer(variables.Length) As Double
    Dim counter, skcounter, errorCounter As Integer
    Dim monitor As DesignMonitoring
    Dim comando As Commands
    Dim compDef As SheetMetalComponentDefinition


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
        compDef = oPartDoc.ComponentDefinition
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
        skcounter = 0
        errorCounter = 0

    End Sub
    Function UpdateDocu(d As PartDocument) As Integer
        oPartDoc = d
        oSk3D = d.ComponentDefinition.Sketches3D.Item(d.ComponentDefinition.Sketches3D.Count)
        compDef = d.ComponentDefinition
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
            MsgBox(ex3.ToString())
            MsgBox("Unable to find Document")
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


    Function openMainSketch(oDoc As PartDocument) As Sketch3D

        oSk3D = oDoc.ComponentDefinition.Sketches3D.Item("MainSketch")
        Return oSk3D
    End Function


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
                    MsgBox(ex2.ToString())
                    MsgBox("Parameter not found: " & name)
                End Try

            End Try

        End Try

        Return p
    End Function

    Function calculateGain(setValue As Double, dc As DimensionConstraint3D) As Double
        Try
            p = dc.Parameter
            delta = (setValue - p._Value) / (setValue * resolution)
            gain = Math.Exp(delta)
        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Fail Calculating " & p.Name)
        End Try

        Return gain
    End Function
    Function calculateGain(setValue As Double, dc As DimensionConstraint) As Double
        Try
            p = dc.Parameter
            delta = (setValue - p._Value) / (setValue * resolution)
            gain = Math.Exp(delta)
        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Fail Calculating " & p.Name)
        End Try

        Return gain
    End Function
    Function calculateGainForMinimun(setValue As Double, dc As DimensionConstraint3D) As Double
        Try
            p = dc.Parameter
            Dim k As Double
            delta = (setValue - p._Value) / (resolution)
            If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
                k = 4
            Else
                k = 2
            End If
            gain = Math.Pow(Math.Exp(delta), 1 / k)
        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Fail Calculating " & p.Name)
        End Try

        Return gain
    End Function
    Function calculateGainForMaximun(setValue As Double, dc As DimensionConstraint3D) As Double
        Try
            p = dc.Parameter
            Dim k As Double
            delta = (setValue - p._Value) / (resolution)
            If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
                k = 4
            Else
                k = 1
            End If
            gain = Math.Pow(Math.Exp(delta), 1 / k)
        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Fail Calculating " & p.Name)
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



    Function AdjustDimConstrain3DSmothly(dci As DimensionConstraint3D, setpoint As Double) As Boolean
        Dim pit As Parameter
        Dim b As Boolean = False
        Dim c, r As Double
        oSk3D = dci.Parent
        pit = dci.Parameter

        Try
            calculateGain(setpoint, dci)
            SetpointCorrector = AdjustResolution(dci, setpoint)

            dci.Driven = False
            While (((Math.Abs(delta * resolution * SetpointCorrector)) > (setpoint / resolution)) And (monitor.IsSketch3DHealthy(oSk3D)))
                r = pit._Value * calculateGain(setpoint, dci)
                dci.Parameter._Value = r
                oSk3D.Solve()
                AdjustResolution(dci, pit._Value)

                counter = counter + 1

            End While
            counter = 0
            If Not monitor.IsSketch3DHealthy(oSk3D) Then
                RecoveryUnhealthySketch(oSk3D)
                Return False
            Else
                b = True
                Debug.Print("adjusted " & pit.Name & " = " & pit.Value.ToString)
                Debug.Print("Resolution:  " & resolution.ToString)
            End If
        Catch ex As Exception

            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)

            RecoveryUnhealthySketch(oSk3D)
            Return False

        End Try

        Return b
    End Function
    Public Function AdjustGapSmothly(gap As DimensionConstraint3D, setpoint As Double, angle As DimensionConstraint3D) As Boolean
        Dim pit As Parameter
        Dim b As Boolean = False


        Try
            oSk3D = gap.Parent
            gap.Driven = False
            pit = gap.Parameter
            calculateGain(setpoint, gap)
            SetpointCorrector = AdjustResolution(gap, setpoint)
            While (((Math.Abs(delta * resolution * SetpointCorrector)) > (setpoint / resolution)) And (monitor.IsSketch3DHealthy(oSk3D)) And (Not IsLastAngleOk(angle)) And (counter < Math.Pow(4 * 4, 4)))
                gap.Parameter._Value = gap.Parameter._Value * Math.Pow(calculateGain(setpoint, gap), 1 / 16)
                oSk3D.Solve()
                AdjustResolution(gap, gap.Parameter._Value)

                counter = counter + 1
            End While
            counter = 0
            If Not monitor.IsSketch3DHealthy(oSk3D) Then
                RecoveryUnhealthySketch(oSk3D)
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


            Return b
        Catch ex As Exception
            UndoCommand()
            If resolution < maxRes Then
                resolution = resolution * 2
            End If
            MsgBox(ex.ToString())
            MsgBox("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)

            RecoveryUnhealthySketch(oSk3D)
            Return False

        End Try

        Return b
    End Function
    Public Function AdjustGapSmothly(gap As DimensionConstraint3D, setpoint As Double, angle As DimensionConstraint3D, angleLimit As Double) As Boolean
        Dim pit As Parameter
        Dim b As Boolean = False


        Try
            oSk3D = gap.Parent
            gap.Driven = False
            pit = gap.Parameter
            calculateGain(setpoint, gap)
            SetpointCorrector = AdjustResolution(gap, setpoint) * 4
            While (((Math.Abs(delta * resolution * SetpointCorrector)) > (setpoint / resolution)) And (monitor.IsSketch3DHealthy(oSk3D)) And (Not IsLastAngleOk(angle, angleLimit)) And (counter < Math.Pow(4 * 4, 4)))
                gap.Parameter._Value = pit._Value * Math.Pow(calculateGain(setpoint, gap), 1 / 4)
                oSk3D.Solve()
                AdjustResolution(gap, pit._Value)

                counter = counter + 1
            End While
            counter = 0
            If Not monitor.IsSketch3DHealthy(oSk3D) Then
                RecoveryUnhealthySketch(oSk3D)
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


            Return b
        Catch ex As Exception
            UndoCommand()
            If resolution < maxRes Then
                resolution = resolution * 2
            End If
            MsgBox(ex.ToString())
            MsgBox("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)

            RecoveryUnhealthySketch(oSk3D)
            Return False

        End Try

        Return b
    End Function

    Public Function GetMinimalDimension(dc As DimensionConstraint3D) As Boolean
        Dim pit As Parameter

        Try
            oSk3D = dc.Parent
            Dim b As Boolean = False
            Dim setPoint As Double = 0
            Dim climit As Integer = 4
            pit = dc.Parameter
            calculateGainForMinimun(setPoint, dc)
            dc.Driven = False
            If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
                SetpointCorrector = Math.Pow(resolution, 2)
                climit = 6
            Else
                SetpointCorrector = 1

            End If
            While (((Math.Abs(delta * resolution) * SetpointCorrector) > (Math.Pow(10, -3) / resolution)) And (monitor.IsSketch3DHealthy(oSk3D)) And (errorCounter < climit) And (counter < climit * 4))
                dc.Parameter._Value = pit._Value * calculateGainForMinimun(setPoint, dc)
                oSk3D.Solve()
                AdjustResolution(dc, pit._Value)
                counter = counter + 1

            End While
            counter = 0
            If Not monitor.IsSketch3DHealthy(oSk3D) Then
                'dc.Driven = True
                RecoveryUnhealthySketch(oSk3D)
                If monitor.IsSketch3DHealthy(oSk3D) Then
                    If errorCounter < climit Then
                        errorCounter = errorCounter + 1
                        ' dc.Driven = False
                        pit.Value = pit.Value / calculateGainForMinimun(setPoint, dc)
                        resolution = resolution * (1 + 4 / CDbl(climit))
                        GetMinimalDimension(dc)
                    Else

                        errorCounter = 0
                        Return True
                    End If
                Else
                    RecoveryUnhealthySketch(oSk3D)
                End If
            Else
                errorCounter = 0

                Return True
            End If
            b = True


            Return b
        Catch ex As Exception
            'comando.UndoCommand()
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)
            RecoveryUnhealthySketch(oSk3D)
            errorCounter = errorCounter + 1
            GetMinimalDimension(dc)
            Return False

        End Try


    End Function
    Public Function GetMaximalDimension(dc As DimensionConstraint3D) As Boolean
        Dim pit As Parameter
        Dim climit As Double = 8
        Dim setPoint As Double = dc.Parameter._Value * (1 + 2 / climit)
        Try

            Dim b As Boolean = False
            pit = dc.Parameter
            oSk3D = dc.Parent
            calculateGainForMaximun(setPoint, dc)
            dc.Driven = False
            If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
                SetpointCorrector = Math.Pow(resolution, 4)
                climit = 4
            Else
                SetpointCorrector = 1

            End If
            While (((Math.Abs(delta * resolution) * SetpointCorrector) > (Math.Pow(10, -3) / resolution)) And (monitor.IsSketch3DHealthy(oSk3D)) And (errorCounter < climit))
                dc.Parameter._Value = pit._Value * calculateGainForMaximun(setPoint, dc)
                oSk3D.Solve()
                AdjustResolution(dc, pit._Value)
                counter = counter + 1
                Debug.Print("iterating Number:  " & counter.ToString() & "  " & pit.Name & " = " & pit.Value.ToString)

            End While
            counter = 0
            If Not monitor.IsSketch3DHealthy(oSk3D) Then
                'dc.Driven = True
                RecoveryUnhealthySketch(oSk3D)
                If monitor.IsSketch3DHealthy(oSk3D) Then
                    If errorCounter < climit Then
                        errorCounter = errorCounter + 1
                        ' dc.Driven = False
                        pit.Value = pit.Value / calculateGainForMaximun(setPoint, dc.Parameter.Name)
                        resolution = resolution * (1 + 4 / CDbl(climit))
                        GetMaximalDimension(dc)
                    Else
                        errorCounter = 0
                        Return True
                    End If
                Else
                    RecoveryUnhealthySketch(oSk3D)
                End If
            Else
                If errorCounter < climit Then
                    GetMaximalDimension(dc)
                Else
                    errorCounter = 0
                    Return True
                End If



            End If
            b = True
            Debug.Print("adjusted " & pit.Name & " = " & pit.Value.ToString)
            Debug.Print("Resolution:  " & resolution.ToString)

            Return b
        Catch ex As Exception
            'comando.UndoCommand()
            Debug.Print(ex.ToString())
            Debug.Print("Fail adjusting " & pit.Name & " ...last value:" & pit.Value.ToString)
            RecoveryUnhealthySketch(oSk3D)
            If errorCounter < climit Then
                errorCounter = errorCounter + 1
                ' dc.Driven = False
                pit.Value = pit.Value / calculateGainForMaximun(setPoint, dc.Parameter.Name)
                resolution = resolution * (1 + 4 / CDbl(climit))
                GetMaximalDimension(dc)
            Else

                errorCounter = 0
                Return True
            End If

            Return False

        End Try


    End Function
    Function IsLastAngleOk(ac As DimensionConstraint3D) As Boolean
        Dim limit As Double = 0.08

        Return IsLastAngleOk(ac, limit)
    End Function
    Function IsLastAngleOk(ac As DimensionConstraint3D, limit As Double) As Boolean
        If ac.Parameter._Value > Math.PI / 2 Then
            If Math.Abs(ac.Parameter._Value - (Math.PI - limit)) < 1 / 64 Then
                Return True
            Else
                Return False
            End If
        Else
            If Math.Abs(ac.Parameter._Value - limit) < 1 / 64 Then
                Return True
            Else
                Return False
            End If
        End If


        Return False
    End Function

    Function RecoveryUnhealthySketch(p As Parameter) As Parameter
        Try
            oPartDoc = p.Parent
            While Not monitor.IsSketch3DHealthy(oSk3D)
                comando.UndoCommand(oPartDoc)

                If Not monitor.IsSketch3DHealthy(oSk3D) Then
                    oPartDoc.Update()
                    comando.UndoCommand()
                End If
                posible = p

            End While
            Return p
        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Fail Recovering " & p.Name & " ...last value:" & p.Value.ToString)
            Return RecoveryUnhealthySketch(p)
        End Try



    End Function
    Public Function RecoveryUnhealthySketch(sk As Sketch3D) As Sketch3D
        Try
            compDef = sk.Parent
            oPartDoc = compDef.Document
            Dim skerror As Integer = skcounter
            While Not monitor.IsSketch3DHealthy(sk)
                For index = 1 To skerror + 1
                    comando.UndoCommand()
                Next
                sk.Solve()
                ' oPartDoc.Update()

                If Not monitor.IsSketch3DHealthy(sk) Then
                    skcounter = skcounter + 1
                    RecoveryUnhealthySketch(sk)
                Else
                    skerror = 0
                End If




            End While
            skcounter = 0

            Return sk
        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Fail Recovering " & sk.Name)
            Return RecoveryUnhealthySketch(sk)
        End Try



    End Function
    Function AdjustResolution(dc As DimensionConstraint3D, s As Double) As Double
        Dim c, r As Double

        If dc.Type = ObjectTypeEnum.kTwoLineAngleDimConstraint3DObject Then
            c = s * dc.Parameter._Value * 16
            r = (1 / (s + Math.Pow(10, -6)) + resolution - 1) / ((counter / 2) + 1)
            If r < 8 Then
                resolution = 8
            Else
                resolution = r

            End If

        Else
            c = s * dc.Parameter._Value * 16
            r = (1 / (s + Math.Pow(10, -6)) + resolution - 1) / (counter + 1)
            If r < 4 Then
                resolution = 4
            Else
                resolution = r

            End If
        End If
        Return c
    End Function





    Function GetDimension(name As String) As DimensionConstraint3D
        Try
            For Each dimension In oSk3D.DimensionConstraints3D
                If dimension.Parameter.Name = name Then
                    Return dimension
                End If
            Next
        Catch ex As Exception
            oSk3D = compDef.Sketches3D.Item(compDef.Sketches3D.Count)
            For Each d As DimensionConstraint3D In oSk3D.DimensionConstraints3D
                If d.Parameter.Name = name Then
                    Return d
                End If
            Next
        End Try

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





    Public Function AdjustDimensionConstraint3DSmothly(dc As DimensionConstraint3D, v As Double) As Boolean
        Dim b As Boolean
        dc.Driven = False


        b = AdjustDimConstrain3DSmothly(dc, v)

        Return b
    End Function


End Class
