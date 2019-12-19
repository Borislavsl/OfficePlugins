using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelAddInForExportToPDF
{
    public partial class ExportToPdf
    {
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
    }
}
