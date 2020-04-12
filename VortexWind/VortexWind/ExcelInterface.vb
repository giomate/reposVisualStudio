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

        path = name & "\keys.xls"
        If File.Exists(path) Then
            oBook = oExcel.Workbooks.Open(path)
        Else
            oBook = oExcel.Workbooks.Add
            oBook.SaveAs(path)
        End If


    End Sub
    Public Function SaveArray(tans() As Long, rods() As Long) As Integer
        Try
            Dim h() As String = {"tans", "rods"}
            Dim t As Object = tans.ToList
            oSheet = oBook.Worksheets.Add
            oSheet.Range("A1").Value = "tans"
            oSheet.Range("B2").Value = "rods"
            Try
                oSheet.Range("A2").Resize(tans.Length, 1).Value = t
            Catch ex As Exception
                Long2Variant(tans)
                oSheet.Range("A2").Resize(tans.Length, 1).Value = col
            End Try
            oBook.Save()
            oExcel.Quit()
            'oSheet.Range("B2").Resize(rods.Length, 1).Value = rods
        Catch ex As Exception
            MsgBox(ex.ToString())
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
