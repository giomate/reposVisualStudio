Imports Microsoft.Office.Interop.Excel
Imports System.IO
Public Class ExcelInterface
    Dim oExcel As Application
    Dim oBook As Workbook
    Dim oSheet As Worksheet
    Dim path As String
    Dim col() As Object


    Public Sub New(name As String)
        oExcel = CreateObject("Excel.Application")
        oExcel.Visible = False
        path = name & "\keys.xlsx"
        If File.Exists(path) Then
            oBook = oExcel.Workbooks.Open(path)
        Else
            oBook = oExcel.Workbooks.Add

        End If


    End Sub
    Public Function SaveArray(tans() As Long, rods() As Long) As Integer
        Try
            Dim r As Range



            Try
                oSheet = oBook.Worksheets.Add
                r = oSheet.Cells(1, 1)
                r.Value = "tans"
                WriteArray(tans, 1)
                r = oSheet.Cells(1, 2)
                r.Value = "rods"
                WriteArray(rods, 2)
                If File.Exists(path) Then
                    oBook.Save()
                Else
                    oBook.SaveAs(path)
                End If


            Catch ex As Exception

            End Try

        Catch ex As Exception

        End Try

        Return 0
    End Function
    Function WriteArray(a() As Long, c As Long) As Integer
        Dim arr() As Long
        Dim r As Range
        Try
        ReDim arr(a.Length - 1)
            arr = a
            For i = 1 To arr.Length
                r = oSheet.Cells(i + 1, c)
                r.Value = arr(i - 1)
            Next
        Catch ex As Exception

        End Try
        Return 0
    End Function
    Function Long2Variant(a() As Long) As Integer
        For i = 0 To a.Length - 1
            col(i) = a(i)
        Next
        Return 0
    End Function
End Class
