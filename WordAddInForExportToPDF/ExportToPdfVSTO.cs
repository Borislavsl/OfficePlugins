using System;
using Microsoft.Office.Interop.Word;

namespace WordAddInForExportToPDF
{
    public partial class ExportToPdf
    {
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
    }
}