
Imports ClosedXML.Excel
Imports System.IO
Imports System.Data
Imports System.Reflection
Imports iTextSharp.text
Imports iTextSharp.text.pdf


Imports System.Linq
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAutoFormat = Microsoft.Office.Interop.Excel.XlRangeAutoFormat
Imports Microsoft.Office.Interop
Imports System.Xml.XPath
Imports System.Xml
Imports Microsoft.Office.Interop.Excel
Imports System.Windows

Public Class Form3

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox1.Text = Form2.ComboBox1.Text
        TextBox2.Text = Form2.TextBox1.Text
        TextBox3.Text = Form2.TextBox3.Text
        TextBox4.Text = Form2.ComboBox2.Text
        TextBox5.Text = Form2.TextBox2.Text
        TextBox6.Text = Form2.TextBox4.Text

        Me.DataGridView1.Rows.Add("+9V", "+9±0.1", "9.086", "13 ± 1", "13.2")
        Me.DataGridView1.Rows.Add("-9V", "- 9 ± 0.1", "-9.087", "4 ± 1", "4.5")
        Me.DataGridView1.Rows.Add("BIAS PIN-4", "+0.63 ± 0.01", "0.628")


        Me.DataGridView2.Rows.Add("0", "Short", "0.04±0.03", "0.047", "0.044", "0.044", "0.048", "0.048")
        Me.DataGridView2.Rows.Add("1", "1200", "0.15±0.03", "0.156", "0.153", "0.153", "0.156", "0.155")
        Me.DataGridView2.Rows.Add("2", "4700", "0.45±0.03", "0.452", "0.448", "0.449", "0.452", "0.452")
        Me.DataGridView2.Rows.Add("3", "1k", "0.85±0.03", "0.855", "0.851", "0.851", "0.857", "0.855")
        Me.DataGridView2.Rows.Add("4", "1.2K", "0.99±0.03", "0.991", "0.992", "0.993", "0.994", "0.992")
        Me.DataGridView2.Rows.Add("5", "1.5K", "1.18±0.03", "1.189", "1.19", "1.189", "1.19", "1.187")
        Me.DataGridView2.Rows.Add("6", "1.8K", "1.37±0.03", "1.37", "1.369", "1.369", "1.37", "1.367")
        Me.DataGridView2.Rows.Add("7", "ΟΡΕΝ", "8.36±0.10", "8.37", "8.376", "8.37", "8.37", "8.37")



        Me.DataGridView3.Rows.Add("0", "5.69±0.1", "5.7", "5.7", "5.73", "5.73", "5.73", "5.73", "5.7", "5.7", "5.7", "5.7")
        Me.DataGridView3.Rows.Add("1", "4.74+0.1", "4.75", "4.75", "4.78", "4.78", "4.78", "4.78", "4.75", "4.75", "4.76", "4.76")
        Me.DataGridView3.Rows.Add("2", "2.14±0.1", "2.16", "2.16", "2.19", "2.19", "2.19", "2.19", "2.16", "2.16", "2.16", "2.16")
        Me.DataGridView3.Rows.Add("3", "- 1.38±0.1", "-1.37", "-1.37", "-1.34", "-1.34", "-1.34", "-1.34", "-1.39", "-1.39", "-1.36", "-1.36")
        Me.DataGridView3.Rows.Add("4", "- 2.59±0.1", "-2.56", "-2.56", "-2.58", "-2.58", "-2.58", "-2.58", "-2.59", "-2.59", "-2.56", "-2.56")
        Me.DataGridView3.Rows.Add("5", "- 4.3±0.1", "-4.3", "-4.3", "-4.31", "-4.31", "-4.29", "-4.29", "-4.31", "-4.31", "-4.26", "-4.26")
        Me.DataGridView3.Rows.Add("6", "- 5.9±0.1", "-5.86", "-5.86", "-5.88", "-5.88", "-5.86", "-5.86", "-5.88", "-5.88", "-5.85", "-5.85")
        Me.DataGridView3.Rows.Add("7", "- 7.04 + 0.2", "- 7.02", "- 7.02", "- 7.04", "- 7.04", "- 7.03", "- 7.03", "- 7.03", "- 7.03", "- 7.02", "- 7.02")















    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub




    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Form4.Show()
        Me.Hide()
    End Sub

End Class