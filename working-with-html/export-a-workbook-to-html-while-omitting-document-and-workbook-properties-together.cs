using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportHtmlWithoutProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello World!");
                sheet.Cells["B2"].PutValue(12345);

                // Set some document and workbook properties (these will be omitted in the HTML output)
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

                // Configure HTML save options to exclude both document and workbook properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false, // Omit document properties
                    ExportWorkbookProperties = false   // Omit workbook properties
                };

                string outputPath = "OutputWithoutProperties.html";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as an HTML file using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"Workbook exported to HTML without document and workbook properties: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportHtmlWithoutProperties.Run();
        }
    }
}