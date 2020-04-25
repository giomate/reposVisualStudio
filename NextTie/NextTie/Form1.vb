Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports System
Imports System.IO
Imports System.Text
Imports System.IO.Directory
Imports Inventor
Imports GetInitialConditions
Public Class Form1
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Dim started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim iteration As Integer

    Dim tie1, tie0, nextTie As TieMaker1

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        iteration = 8
        ' Add any initialization after the InitializeComponent() call.
        Try

            oApp = Marshal.GetActiveObject("Inventor.Application")
            started = True

        Catch ex As Exception
            Try
                Dim oInvAppType As Type = GetTypeFromProgID("Inventor.Application")

                oApp = CreateInstance(oInvAppType)
                oApp.Visible = True

                'Note: if you shut down the Inventor session that was started
                'this(way) there is still an Inventor.exe running. We will use
                'this Boolean to test whether or not the Inventor App  will
                'need to be shut down.
                started = True


            Catch ex2 As Exception
                MsgBox(ex2.ToString())
                MsgBox("Unable to get or start Inventor")
            End Try
        End Try
    End Sub
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If started Then
            'oApp.ActiveDocument.Close()
            MsgBox("Closing Session")
        End If
        oApp = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        oApp = Nothing
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        done = MakeNextTie()
    End Sub

    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try

            done = KeepMakingTies()
            If done Then
                Me.Close()
            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Unable to find Document")
        End Try
    End Sub
    Public Function MakeNextTie() As Boolean
        Dim b As Boolean
        Try
            If (started And (Not running)) Then
                invDoc = New InventorFile(oApp)
                tie1 = New TieMaker1(invDoc.OpenSheetMetalFile("Band0.ipt"))
                If tie1.MakeNextTie().ComponentDefinition.Features.Count > 5 Then
                    If tie1.compDef.Sketches3D.Item("last").SketchLines3D.Count > 0 Then
                        b = tie1.done
                    End If

                End If

            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Public Function KeepMakingTies() As Boolean
        Dim b As Boolean = True
        Dim q As Integer
        Dim s As String
        Try
            If (started And (Not running)) Then
                invDoc = New InventorFile(oApp)
                tie0 = New TieMaker1(invDoc.OpenSheetMetalFile("Band0.ipt"))
                nextTie = New TieMaker1(tie0.FindLastTie())
                q = tie0.qLastTie
                If q > 0 Then
                    tie0.doku.Save2(True)
                    tie0.doku.Close(True)
                End If

                While (q < tie0.trobinaCurve.DP.q And b)
                    If nextTie.MakeNextTie().ComponentDefinition.Features.Count > 5 Then
                        If nextTie.compDef.Sketches3D.Item("nextIntroLine").SketchLines3D.Count > 0 Then
                            b = nextTie.done
                            If b Then
                                oDoc = nextTie.doku
                                s = nextTie.fullFileName
                                oDoc.Update2(True)
                                oDoc.Save2(True)
                                oApp.Documents.CloseAll()
                                tie0 = New TieMaker1(invDoc.OpenFullFileName(s))
                                nextTie = New TieMaker1(tie0.FindLastTie())
                                q = tie0.qLastTie

                            End If
                        End If

                    End If
                End While
                If b Then
                    If nextTie.MoveAllTies(iteration) > 0 Then
                        iteration = iteration + 1
                        b = True
                    End If
                End If

            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
End Class
