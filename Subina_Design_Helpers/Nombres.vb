Imports System
Imports System.Text.RegularExpressions
Imports Inventor
Public Class Nombres
    Dim newName, currentName, oldName, datei As String
    Dim doku As PartDocument
    Dim app As Application
    Dim compDef As ComponentDefinition

    Public Sub New()
        newName = "noname"
        datei = ".ipt"
    End Sub
    Public Sub New(docu As Inventor.Document)
        doku = docu
        app = doku.Parent
        'get the document sub-type


        compDef = doku.ComponentDefinition


    End Sub
    Public Function IncrementLabel(name As String, pattern As String) As String
        Dim p() As String

        Try
            p = Strings.Split(name, pattern)
            newName = String.Concat(p(0), pattern, CStr(CInt(p(1)) + 1))
        Catch ex As Exception
            MsgBox(ex.ToString())
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
            MsgBox(ex.ToString())
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
            MsgBox(ex.ToString())
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
    Public Function Contain_F_First(s As String) As Boolean
        Dim b As Boolean
        Dim pattern = String.Concat("\b", "f", "\d+", "\w*")
        b = Regex.IsMatch(s, pattern)
        Return b
    End Function
    Public Function GetQNumber(docu As Inventor.Document) As Integer
        Dim s() As String
        s = Strings.Split(docu.FullFileName, "Band")
        s = Strings.Split(s(1), ".ipt")

        Return CInt(s(0)) + 1
    End Function
    Public Function GetFileNumber(fullName As String) As Integer
        Dim s() As String
        Dim pattern = String.Concat("Band", "\d+", "\w*")
        If Regex.IsMatch(fullName, pattern) Then
            s = Strings.Split(fullName, "Band")
            s = Strings.Split(s(1), ".ipt")

            Return CInt(s(0))
        End If

        Return 0
    End Function
    Public Function GetNextSketchName(docu As Inventor.Document) As String
        Dim s(), sn As String
        Dim i As Integer = 0

        While Not ContainSFirst(compDef.Sketches3D.Item(compDef.Sketches3D.Count - i).Name)
            i = i + 1
        End While
        sn = compDef.Sketches3D.Item(compDef.Sketches3D.Count - i).Name
        s = Strings.Split(sn, "s")
        sn = String.Concat("s" & (CInt(s(1)) + 1).ToString)

        Return sn
    End Function
    Public Function GetCurrentSketchNumber(docu As Inventor.Document) As Integer
        Dim s(), sn As String
        Dim i As Integer = 0

        While Not ContainSFirst(compDef.Sketches3D.Item(compDef.Sketches3D.Count - i).Name)
            i = i + 1
        End While
        sn = compDef.Sketches3D.Item(compDef.Sketches3D.Count - i).Name
        s = Strings.Split(sn, "s")


        Return CInt(s(1))
    End Function
    Public Function GetFeatureNumber(f As FoldFeature) As Integer
        Dim s(), sn As String
        Dim i As Integer = 0

        If Contain_F_First(f.Name) Then
            s = Strings.Split(f.Name, "f")
            i = CInt(s(1))
        End If


        Return i
    End Function
    Public Function MakeNextFileName(docu As Document) As String
        Dim s(), sn As String
        Dim bn As String = "Band"
        Dim en As String = ".ipt"
        Dim nq As Integer
        s = Strings.Split(docu.FullFileName, bn)
        nq = GetQNumber(doku)
        sn = String.Concat(bn, nq.ToString, en)

        Return sn
    End Function
    Public Function GetQNumberString(ffn As String) As Integer
        Dim s() As String
        s = Strings.Split(ffn, "Band")
        s = Strings.Split(s(1), ".ipt")

        Return CInt(s(0))
    End Function
    Public Function ConvertQNumberLetter(q As Integer) As String
        Dim c As Char
        Dim m As Integer
        Dim s As String
        Math.DivRem(q, 94, m)
        c = ChrW(m + 33)
        s = c.ToString
        Return s
    End Function
    Public Function ContainBand(s As String) As Boolean
        Dim b As Boolean
        Dim pattern = String.Concat("Band")
        b = Regex.IsMatch(s, pattern)
        Return b
    End Function
    Public Function MakeSkeletonFileName(ffn As String) As String
        Dim s(), sn As String
        Dim bn As String = "Band"
        s = Strings.Split(ffn, bn)
        sn = String.Concat(s(0), "Skeleton", s(1))

        Return sn
    End Function
    Public Function MakeRibFileName(dn As String, i As Integer) As String

        Dim sn As String = String.Concat(dn, "\Rib", i, ".ipt")

        Return sn
    End Function
    Public Function ContainCorto(s As String) As Boolean
        Dim b As Boolean
        Dim pattern = String.Concat("shortRod")
        b = Regex.IsMatch(s, pattern)
        Return b
    End Function
    Public Function ContainEgg(s As String) As Boolean
        Dim b As Boolean
        Dim pattern = String.Concat("egg")
        b = Regex.IsMatch(s, pattern)
        Return b
    End Function
    Public Function GetRodNumber(rf As RevolveFeature) As Integer
        Dim s() As String
        s = Strings.Split(rf.Name, "shortRod")
        Return CInt(s(1))
    End Function
    Public Function MakeWedgeFileName(ffn As String) As String
        Dim s(), sn As String
        Dim bn As String = "Band"
        s = Strings.Split(ffn, bn)
        sn = String.Concat(s(0), "Wedge", s(1))

        Return sn
    End Function
End Class
