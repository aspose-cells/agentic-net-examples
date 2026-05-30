using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    public class ExportWithoutProperties
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("HTML export completed successfully.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Data");
            sheet.Cells["B2"].PutValue(123);

            // (Optional) Set some document properties that we will later omit from the HTML output
            workbook.BuiltInDocumentProperties.Author = "John Doe";
            workbook.BuiltInDocumentProperties.Title = "Demo Workbook";

            // Configure HTML save options to exclude document and worksheet properties
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportDocumentProperties = false,   // Do not export document properties
                ExportWorksheetProperties = false   // Do not export worksheet properties
            };

            // Define output file path
            string outputPath = "ExportWithoutProps.html";

            // Ensure the directory exists (in case a relative path is used)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, options);
        }
    }
}