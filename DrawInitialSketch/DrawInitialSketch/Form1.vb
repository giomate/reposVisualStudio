﻿Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor
Public Class Form1
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Dim started, running, done As Boolean
    Dim oDesignProjectMgr As DesignProjectManager
    Dim invDoc As InventorFile
    Dim trobinaCurve As Curve3D
    Dim fullFileName, pattern, shortFileName As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

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
    Private Sub Form1_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Try

            done = StartDrawing()
            If done Then
                Me.Close()
            End If
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Debug.Print("Unable to find Document")
        End Try
    End Sub
    Public Function StartDrawing() As Boolean
        Dim b As Boolean
        Try
            If (started And (Not running)) Then
                invDoc = New InventorFile(oApp)
                trobinaCurve = New Curve3D(invDoc.CreateSheetMetalFile("Band0.ipt"))

                b = trobinaCurve.DrawTrobinaCurveAlone().SketchEquationCurves3D.Count > 0
            End If

        Catch ex As Exception
            Debug.Print(ex.ToString())
            b = False
        End Try
        Return b
    End Function
End Class
