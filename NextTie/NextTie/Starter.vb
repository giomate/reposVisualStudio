Imports Inventor
Public Class Starter
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Public started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim iteration As Integer

    Dim tie1, tie0, nextTie As TieMaker1


    Public Sub New(app As Inventor.Application)
        oApp = app
        iteration = 21
        If oApp.Visible Then
            started = True
        End If
    End Sub
    Public Function AddInTies(i As Integer) As Boolean
        iteration = i
        done = KeepMakingTies()
        Return done
    End Function


    Public Function KeepMakingTies() As Boolean
        Dim b As Boolean = True
        Dim q As Integer
        Dim s As String

        Try
            If (started) Then
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
                        If nextTie.compDef.Sketches3D.Item("introLine").SketchLines3D.Count > 0 Then
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
