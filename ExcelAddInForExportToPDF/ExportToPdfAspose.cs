using System;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInForExportToPDF
{
    public partial class ExportToPdf
    {
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
