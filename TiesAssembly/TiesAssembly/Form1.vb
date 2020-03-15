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
    Dim banda As BandAssembly

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        iteration = 1
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

            done = MakeDonnut()
            If done Then
                Me.Close()
            End If


        Catch ex As Exception
            MsgBox(ex.ToString())
            MsgBox("Unable to find Document")
        End Try
    End Sub
    Public Function AssemblyAllTies() As Boolean
        Dim b As Boolean
        Try
            If (started And (Not running)) Then
                banda = New BandAssembly(oApp)
                banda.PutTies()
                b = banda.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Public Function MakeDonnut() As Boolean
        Dim b As Boolean
        Try
            If (started And (Not running)) Then
                banda = New BandAssembly(oApp)
                banda.TrimDonnut()
                b = banda.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function
    Public Function MakeBand() As Boolean
        Dim b As Boolean
        Try
            If (started And (Not running)) Then
                banda = New BandAssembly(oApp)
                banda.DeriveAllTies()
                b = banda.done
            End If
            Return b
        Catch ex As Exception
            MsgBox(ex.ToString())
            Return False
        End Try

    End Function

End Class
