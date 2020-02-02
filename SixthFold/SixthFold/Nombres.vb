Imports System
Imports System.Text.RegularExpressions
Imports Inventor
Public Class Nombres
    Dim newName, currentName, oldName, datei As String
    Dim doku As PartDocument
    Dim app As Application
    Dim compDef As SheetMetalComponentDefinition

    Public Sub New()
        newName = "noname"
        datei = ".ipt"
    End Sub
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        compDef = doku.ComponentDefinition
    End Sub
    Public Function IncrementLabel(name As String, pattern As String) As String
        Dim p() As String

        Try
            p = Strings.Split(name, pattern)
            newName = String.Concat(p(0), pattern, CStr(CInt(p(1)) + 1))
        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return name
        End Try

        Return newName
    End Function
    Public Function IncrementLabelIpt(name As String, pattern As String) As String
        Dim p() As String

        Try
            p = Strings.Split(name, datei)
            newName = String.Concat(IncrementLabel(p(0), pattern), datei)

        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return name
        End Try

        Return newName
    End Function
    Public Function CreateIptLabel(name As String, pattern As String) As String
        Dim p() As String

        Try
            p = Strings.Split(name, datei)
            newName = String.Concat(IncrementLabel(p(0), pattern), datei)

        Catch ex As Exception
            Debug.Print(ex.ToString())
            Return name
        End Try

        Return newName
    End Function
    Public Function ContainSFirst(s As String) As Boolean
        Dim b As Boolean
        Dim pattern = String.Concat("\b", "s", "\d+", "\w*")
        b = Regex.IsMatch(s, pattern)
        Return b
    End Function
    Public Function GetQNumber(docu As Inventor.Document) As Integer
        Dim s() As String
        s = Strings.Split(docu.FullFileName, "Band")
        s = Strings.Split(s(1), ".ipt")

        Return CInt(s(0)) + 1
    End Function
    Public Function GetNextSketchName(docu As Inventor.Document) As String
        Dim s(), sn As String
        sn = compDef.Sketches3D.Item(compDef.Sketches3D.Count).Name
        s = Strings.Split(sn, "s")
        sn = String.Concat("s" & (CInt(s(1)) + 1).ToString)

        Return sn
    End Function
End Class
