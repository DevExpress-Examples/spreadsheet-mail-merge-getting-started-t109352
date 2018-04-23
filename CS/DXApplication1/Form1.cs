using System.Collections.Generic;
using DevExpress.XtraBars.Ribbon;
using DevExpress.Spreadsheet;

namespace DXApplication1 {
    public partial class Form1 : RibbonForm {

        nwindDataSet dataSet;
        nwindDataSetTableAdapters.CategoriesTableAdapter adapter;
        IWorkbook template;

        public Form1() {
            InitializeComponent();

            dataSet = new nwindDataSet();
            adapter = new nwindDataSetTableAdapters.CategoriesTableAdapter();
            adapter.Fill(dataSet.Categories);

            template = spreadsheetControl1.Document;

            template.MailMergeDataSource = dataSet;
            template.MailMergeDataMember = "Categories";
        }

        private void Form1_Load(object sender, System.EventArgs e) {
            spreadsheetControl1.LoadDocument("Documents\\MailMergeTemplate.xlsx");
        }

        private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            IList<IWorkbook> resultWorkbooks = spreadsheetControl1.Document.GenerateMailMergeDocuments();

            string fileName;
            int index = 0;

            foreach (IWorkbook workbook in resultWorkbooks) {
                fileName = string.Format("Documents\\SavedDocument{0}" + ".xlsx", index++);
                workbook.SaveDocument(fileName, DocumentFormat.OpenXml);
            }
        }
    }
}
