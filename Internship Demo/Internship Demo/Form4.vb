Imports System.IO
Imports itextsharp.text.pdf
Imports itextsharp.text
Imports System.Reflection.Metadata
Imports Document = iTextSharp.text.Document
Imports Microsoft.Office.Interop.Excel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Font = iTextSharp.text.Font

Public Class Form4
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim writer As New StreamWriter("D:\sample projects\Internship Demo\Exported txt\info on io.txt")

        For i As Integer = 0 To Form3.DataGridView1.Rows.Count - 2 Step +1

            For j As Integer = 0 To Form3.DataGridView1.Columns.Count - 1 Step +1

                ' if last column
                If j = Form3.DataGridView1.Columns.Count - 1 Then
                    writer.Write(vbTab & Form3.DataGridView1.Rows(i).Cells(j).Value.ToString())
                Else
                    writer.Write(vbTab & Form3.DataGridView1.Rows(i).Cells(j).Value.ToString() & vbTab & "|")
                End If


            Next j

            writer.WriteLine("")

        Next i

        writer.Close()



        Dim writer1 As New StreamWriter("D:\sample projects\Internship Demo\Exported txt\input.txt")

        For i As Integer = 0 To Form3.DataGridView2.Rows.Count - 2 Step +1

            For j As Integer = 0 To Form3.DataGridView2.Columns.Count - 1 Step +1

                ' if last column
                If j = Form3.DataGridView1.Columns.Count - 1 Then
                    writer1.Write(vbTab & Form3.DataGridView2.Rows(i).Cells(j).Value.ToString())
                Else
                    writer1.Write(vbTab & Form3.DataGridView2.Rows(i).Cells(j).Value.ToString() & vbTab & "|")
                End If


            Next j

            writer1.WriteLine("")

        Next i

        writer1.Close()



        Dim writer2 As New StreamWriter("D:\sample projects\Internship Demo\Exported txt\output.txt")

        For i As Integer = 0 To Form3.DataGridView3.Rows.Count - 2 Step +1

            For j As Integer = 0 To Form3.DataGridView3.Columns.Count - 1 Step +1

                ' if last column
                If j = Form3.DataGridView1.Columns.Count - 1 Then
                    writer2.Write(vbTab & Form3.DataGridView3.Rows(i).Cells(j).Value.ToString())
                Else
                    writer2.Write(vbTab & Form3.DataGridView3.Rows(i).Cells(j).Value.ToString() & vbTab & "|")
                End If


            Next j

            writer2.WriteLine("")

        Next i

        writer2.Close()
        MessageBox.Show("Data Exported")


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim xlApp As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue As Object = System.Reflection.Missing.Value
        Dim i As Integer
        Dim j As Integer
        xlApp = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook = xlApp.Workbooks.Add(misValue)
        xlWorkSheet = xlWorkBook.Sheets("sheet1")
        For i = 0 To Form3.DataGridView1.RowCount - 2
            For j = 0 To Form3.DataGridView1.ColumnCount - 1
                For k As Integer = 1 To Form3.DataGridView1.Columns.Count
                    xlWorkSheet.Cells(1, k) = Form3.DataGridView1.Columns(k - 1).HeaderText
                    xlWorkSheet.Cells(i + 2, j + 1) = Form3.DataGridView1(j, i).Value.ToString()
                Next
            Next
        Next
        xlWorkSheet.SaveAs("D:\sample projects\Internship Demo\Exported excel\info on io.xlsx")
        xlWorkBook.Close()
        xlApp.Quit()
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)


        Dim xlApp1 As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook1 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue1 As Object = System.Reflection.Missing.Value
        Dim i1 As Integer
        Dim j1 As Integer
        xlApp1 = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook1 = xlApp1.Workbooks.Add(misValue1)
        xlWorkSheet1 = xlWorkBook1.Sheets("sheet1")
        For i1 = 0 To Form3.DataGridView2.RowCount - 2
            For j1 = 0 To Form3.DataGridView2.ColumnCount - 1
                For k As Integer = 1 To Form3.DataGridView2.Columns.Count
                    xlWorkSheet1.Cells(1, k) = Form3.DataGridView2.Columns(k - 1).HeaderText
                    xlWorkSheet1.Cells(i1 + 2, j1 + 1) = Form3.DataGridView2(j1, i1).Value.ToString()
                Next
            Next
        Next
        xlWorkSheet1.SaveAs("D:\sample projects\Internship Demo\Exported excel\input.xlsx")
        xlWorkBook1.Close()
        xlApp1.Quit()
        releaseObject(xlApp1)
        releaseObject(xlWorkBook1)
        releaseObject(xlWorkSheet1)




        Dim xlApp2 As Microsoft.Office.Interop.Excel.Application
        Dim xlWorkBook2 As Microsoft.Office.Interop.Excel.Workbook
        Dim xlWorkSheet2 As Microsoft.Office.Interop.Excel.Worksheet
        Dim misValue2 As Object = System.Reflection.Missing.Value
        Dim i2 As Integer
        Dim j2 As Integer
        xlApp2 = New Microsoft.Office.Interop.Excel.Application
        xlWorkBook2 = xlApp2.Workbooks.Add(misValue2)
        xlWorkSheet2 = xlWorkBook2.Sheets("sheet1")
        For i2 = 0 To Form3.DataGridView3.RowCount - 2
            For j2 = 0 To Form3.DataGridView3.ColumnCount - 1
                For k As Integer = 1 To Form3.DataGridView3.Columns.Count
                    xlWorkSheet2.Cells(1, k) = Form3.DataGridView3.Columns(k - 1).HeaderText
                    xlWorkSheet2.Cells(i2 + 2, j2 + 1) = Form3.DataGridView3(j2, i2).Value.ToString()
                Next
            Next
        Next
        xlWorkSheet2.SaveAs("D:\sample projects\Internship Demo\Exported excel\output.xlsx")
        xlWorkBook2.Close()
        xlApp2.Quit()
        releaseObject(xlApp2)
        releaseObject(xlWorkBook2)
        releaseObject(xlWorkSheet2)
        MsgBox("Succesfully Exported As Excel")

    End Sub
    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Paragraph As New Paragraph
        Dim pdfFile As New Document(PageSize.A4, 40, 40, 40, 20)
        pdfFile.AddTitle("info")
        Dim write As PdfWriter = PdfWriter.GetInstance(pdfFile, New FileStream("D:\sample projects\Internship Demo\Exported pdf\info on i-o.pdf", FileMode.Create))
        pdfFile.Open()

        Dim pTitle As New Font(iTextSharp.text.Font.BOLD)
        Dim pTable As New Font(iTextSharp.text.Font.NORMAL)

        Paragraph = New Paragraph(New Chunk("Info on I-O", pTitle))
        Paragraph.Alignment = Element.ALIGN_CENTER
        Paragraph.SpacingAfter = 5.0F

        pdfFile.Add(Paragraph)

        Dim pdftable As New PdfPTable(Form3.DataGridView1.Columns.Count)
        pdftable.TotalWidth = 500.0F
        pdftable.LockedWidth = True

        Dim widths(0 To Form3.DataGridView1.Columns.Count - 1) As Single
        For i As Integer = 0 To Form3.DataGridView1.Columns.Count - 1
            widths(i) = 1.0F

        Next

        pdftable.SetWidths(widths)
        pdftable.HorizontalAlignment = 0
        pdftable.SpacingBefore = 5.0F


        Dim pdfcell As PdfPCell = New PdfPCell

        For i As Integer = 0 To Form3.DataGridView1.Columns.Count - 1
            pdfcell = New PdfPCell(New Phrase(New Chunk(Form3.DataGridView1.Columns(i).HeaderText, pTable)))

            pdfcell.HorizontalAlignment = PdfPCell.ALIGN_LEFT
            pdftable.AddCell(pdfcell)

        Next
        For i As Integer = 0 To Form3.DataGridView1.Rows.Count - 2
            For j As Integer = 0 To Form3.DataGridView1.Columns.Count - 1
                pdfcell = New PdfPCell(New Phrase(Form3.DataGridView1(j, i).Value.ToString(), pTable))
                pdftable.HorizontalAlignment = PdfPCell.ALIGN_LEFT
                pdftable.AddCell(pdfcell)
            Next
        Next
        pdfFile.Add(pdftable)
        pdfFile.Close()




        Dim Paragraph1 As New Paragraph
        Dim pdfFile1 As New Document(PageSize.A4, 40, 40, 40, 20)
        pdfFile1.AddTitle("input")
        Dim write1 As PdfWriter = PdfWriter.GetInstance(pdfFile1, New FileStream("D:\sample projects\Internship Demo\Exported pdf\input.pdf", FileMode.Create))
        pdfFile1.Open()

        Dim pTitle1 As New Font(iTextSharp.text.Font.BOLD)
        Dim pTable1 As New Font(iTextSharp.text.Font.NORMAL)

        Paragraph = New Paragraph(New Chunk("INPUT", pTitle1))
        Paragraph.Alignment = Element.ALIGN_CENTER
        Paragraph.SpacingAfter = 5.0F

        pdfFile1.Add(Paragraph)

        Dim pdftable1 As New PdfPTable(Form3.DataGridView2.Columns.Count)
        pdftable1.TotalWidth = 500.0F
        pdftable1.LockedWidth = True

        Dim widths1(0 To Form3.DataGridView2.Columns.Count - 1) As Single
        For i As Integer = 0 To Form3.DataGridView2.Columns.Count - 1
            widths1(i) = 1.0F

        Next

        pdftable1.SetWidths(widths1)
        pdftable1.HorizontalAlignment = 0
        pdftable1.SpacingBefore = 5.0F


        Dim pdfcell1 As PdfPCell = New PdfPCell

        For i As Integer = 0 To Form3.DataGridView2.Columns.Count - 1
            pdfcell1 = New PdfPCell(New Phrase(New Chunk(Form3.DataGridView2.Columns(i).HeaderText, pTable1)))

            pdfcell1.HorizontalAlignment = PdfPCell.ALIGN_LEFT
            pdftable1.AddCell(pdfcell1)

        Next
        For i As Integer = 0 To Form3.DataGridView2.Rows.Count - 2
            For j As Integer = 0 To Form3.DataGridView2.Columns.Count - 1
                pdfcell1 = New PdfPCell(New Phrase(Form3.DataGridView2(j, i).Value.ToString(), pTable1))
                pdftable1.HorizontalAlignment = PdfPCell.ALIGN_LEFT
                pdftable1.AddCell(pdfcell1)
            Next
        Next
        pdfFile1.Add(pdftable1)
        pdfFile1.Close()





        Dim Paragraph2 As New Paragraph
        Dim pdfFile2 As New Document(PageSize.A4, 40, 40, 40, 20)
        pdfFile2.AddTitle("output")
        Dim write2 As PdfWriter = PdfWriter.GetInstance(pdfFile2, New FileStream("D:\sample projects\Internship Demo\Exported pdf\output.pdf", FileMode.Create))
        pdfFile2.Open()

        Dim pTitle2 As New Font(iTextSharp.text.Font.BOLD)
        Dim pTable2 As New Font(iTextSharp.text.Font.NORMAL)

        Paragraph = New Paragraph(New Chunk("OUTPUT", pTitle2))
        Paragraph.Alignment = Element.ALIGN_CENTER
        Paragraph.SpacingAfter = 5.0F

        pdfFile2.Add(Paragraph)

        Dim pdftable2 As New PdfPTable(Form3.DataGridView3.Columns.Count)
        pdftable2.TotalWidth = 500.0F
        pdftable2.LockedWidth = True

        Dim widths2(0 To Form3.DataGridView3.Columns.Count - 1) As Single
        For i As Integer = 0 To Form3.DataGridView3.Columns.Count - 1
            widths2(i) = 1.0F

        Next

        pdftable2.SetWidths(widths2)
        pdftable2.HorizontalAlignment = 0
        pdftable2.SpacingBefore = 5.0F


        Dim pdfcell2 As PdfPCell = New PdfPCell

        For i As Integer = 0 To Form3.DataGridView3.Columns.Count - 1
            pdfcell2 = New PdfPCell(New Phrase(New Chunk(Form3.DataGridView3.Columns(i).HeaderText, pTable2)))

            pdfcell2.HorizontalAlignment = PdfPCell.ALIGN_LEFT
            pdftable2.AddCell(pdfcell2)

        Next
        For i As Integer = 0 To Form3.DataGridView3.Rows.Count - 2
            For j As Integer = 0 To Form3.DataGridView3.Columns.Count - 1
                pdfcell2 = New PdfPCell(New Phrase(Form3.DataGridView3(j, i).Value.ToString(), pTable2))
                pdftable2.HorizontalAlignment = PdfPCell.ALIGN_LEFT
                pdftable2.AddCell(pdfcell2)
            Next
        Next
        pdfFile2.Add(pdftable2)
        pdfFile2.Close()

        MessageBox.Show("PDF Format Succes Exported !", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)


    End Sub
End Class