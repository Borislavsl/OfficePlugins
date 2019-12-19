using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using Microsoft.Office.Interop.Word;

namespace WordAddInForExportToPDF
{
    public partial class ExportToPdf
    {
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