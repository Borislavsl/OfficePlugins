using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using AddInUtilities;

namespace ExcelAddInForExportToPDF
{
    public partial class ExportToPdf
    {
        private void ExportToPdf_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void exportButton_Click(object sender, RibbonControlEventArgs e)
        {
            var exportType = (sender as RibbonButton).Label;
            string outputFolder = Util.GetOutputFolder(exportType);

            Application excelApplication = Globals.ThisAddIn.Application;
            Workbook excelWorkbook = excelApplication.ActiveWorkbook;

            string outputPath = Util.SaveAsPDFFileDialog(outputFolder, excelWorkbook.Name, exportType);
            if (string.IsNullOrEmpty(outputPath))
                return;

            string exportResult;
            if (Util.UseAspose(exportType))
                exportResult = AsposeExportWorkbookToPdf(excelWorkbook, outputPath);
            else
                exportResult = VSTOExportWorkbookToPdf(excelWorkbook, outputPath);

            Util.ShowExportResult(exportResult, "Workbook");
        }
    }
}
