using System;
using System.IO;
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

        private string VSTOExportWorkbookToPdf(Workbook excelWorkbook, string outputPath)
        {
            var resultMessage = string.Empty;

            try
            {
                excelWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, outputPath);
            }
            catch (Exception ex)
            {
                resultMessage = ex.Message;
            }

            return resultMessage;
        }

        private string AsposeExportWorkbookToPdf(Workbook excelWorkbook, string outputPath)
        {
            var resultMessage = string.Empty;
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            try
            {
                excelWorkbook.SaveCopyAs(tempFile);

                var asposeWorkbook = new Aspose.Cells.Workbook(tempFile);
                asposeWorkbook.Save(outputPath, Aspose.Cells.SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                resultMessage = ex.Message;
            }
            finally
            {
                File.Delete(tempFile);
            }

            return resultMessage;
        }
    }
}
