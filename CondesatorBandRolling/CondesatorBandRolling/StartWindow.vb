﻿Imports System.Type
Imports System.Activator
Imports System.Runtime.InteropServices
Imports Inventor


Public Class StartWindow
    Dim oApp As Inventor.Application
    Dim oDoc As Inventor.Document
    Dim started As Boolean
    Dim oDesignProjectMgr As DesignProjectManager

    Public Sub New()


        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        Try

            oApp = Marshal.GetActiveObject("Inventor.Application")
            oDesignProjectMgr = oApp.DesignProjectManager

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

    Private Sub StartWindow_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If started Then
            oApp.Quit()
            MsgBox("Closing Session")
        End If

        oApp = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call SetActiveProject(oApp)
        BandSheetMetal.CreateFile(oApp)
        BandSheetMetal.CreateFaceFeature(oApp.ActiveDocument)







    End Sub
End Class




