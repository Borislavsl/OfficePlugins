using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace AddInUtilities
{
    public static class Util
    {
        private const string ASPOSE = "ASPOSE";        

        private static string successMessage = " is successfully exported to PDF";
        private static string errorMessage = "Exporting to PDF got the error: ";
        private static string dialogTitle = " Save As PDF";

        public static string GetOutputFolder(string exportType)
        {
            var uri = new System.Uri(Assembly.GetExecutingAssembly().CodeBase).AbsolutePath;
            var projectFolder = new DirectoryInfo(uri).Parent.Parent.Parent.FullName;
            var outputFolder = Path.Combine(projectFolder, "Output", exportType);

            return outputFolder;
        }

        public static string SaveAsPDFFileDialog(string folderName, string fileName, string exportType = "")
        {
            var saveFileDialog = new SaveFileDialog()
            {
                DefaultExt = "*.pdf",
                Filter = "PDF Files (*.pdf)|*.pdf",
                InitialDirectory = folderName,
                FileName = fileName,
                Title = exportType + dialogTitle
            };

            var result = saveFileDialog.ShowDialog();
            if (result == DialogResult.OK)
                return saveFileDialog.FileName;

            return null;
        }

        public static bool UseAspose(string exportType) => exportType.ToUpper().Contains(ASPOSE);

        public static void ShowExportResult(string exportResult, string item)
        {
            if (string.IsNullOrEmpty(exportResult))
                MessageBox.Show(item + successMessage, "SUCCESS", MessageBoxButtons.OK, MessageBoxIcon.Information);
            else
                MessageBox.Show(errorMessage + exportResult, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
