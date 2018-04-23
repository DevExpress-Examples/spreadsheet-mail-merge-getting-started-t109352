Imports DevExpress.XtraBars.Ribbon
Imports DevExpress.Spreadsheet

Partial Public Class Form1
    Inherits RibbonForm

    Private dataSet As nwindDataSet
    Private adapter As nwindDataSetTableAdapters.CategoriesTableAdapter
    Private template As IWorkbook

	Public Sub New()
        InitializeComponent()

        dataSet = New nwindDataSet()
        adapter = New nwindDataSetTableAdapters.CategoriesTableAdapter()
        adapter.Fill(dataSet.Categories)

        template = SpreadsheetControl1.Document

        template.MailMergeDataSource = dataSet
        template.MailMergeDataMember = "Categories"
	End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        SpreadsheetControl1.LoadDocument("Documents\MailMergeTemplate.xlsx")
    End Sub

    Private Sub BarButtonItem1_ItemClick(sender As Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
        Dim resultWorkbooks As IList(Of IWorkbook) = SpreadsheetControl1.Document.GenerateMailMergeDocuments()

        Dim fileName As String
        Dim index As Integer = 0

        For Each workbook As IWorkbook In resultWorkbooks
            fileName = String.Format("C:\TEMP\SavedDocument{0}" & ".xlsx", index)
            index += 1
            workbook.SaveDocument(fileName, DocumentFormat.OpenXml)
        Next workbook
    End Sub
End Class
