using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using AddInUtilities;

namespace WordAddInForExportToPDF
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

            Application wordApplication = Globals.ThisAddIn.Application;
            Document wordDocument = wordApplication.ActiveDocument;

            string outputPath = Util.SaveAsPDFFileDialog(outputFolder, wordDocument.Name, exportType);
            if (string.IsNullOrEmpty(outputPath))
                return;

            string exportResult;
            if (Util.UseAspose(exportType))
                exportResult = AsposeExportDocumentToPdf(wordDocument, outputPath);
            else
                exportResult = VSTOExportDocumentToPdf(wordDocument, outputPath);

            Util.ShowExportResult(exportResult, "Document");
        }

        private string VSTOExportDocumentToPdf(Document wordDocument, string outputPath)
        {
            var resultMessage = string.Empty;

            try
            {
                wordDocument.ExportAsFixedFormat(outputPath, WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                resultMessage = ex.Message;
            }

            return resultMessage;
        }

        private string AsposeExportDocumentToPdf(Document wordDocument, string outputPath)
        {
            var resultMessage = string.Empty;
            string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");

            try
            {
                (wordDocument as IPersistFile).Save(tempFile, false);

                var asposeDocument = new Aspose.Words.Document(tempFile);
                asposeDocument.Save(outputPath, Aspose.Words.SaveFormat.Pdf);
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
