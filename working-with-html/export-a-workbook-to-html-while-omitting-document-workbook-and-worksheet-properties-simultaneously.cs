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
                sheet.Cells["A1"].PutValue("Sample Data");

                // (Optional) Set some document and workbook properties to demonstrate they will be omitted
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Title = "Demo Workbook";

                // Configure HTML save options to exclude all property exports
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false,   // Omit document properties
                    ExportWorkbookProperties = false,   // Omit workbook properties
                    ExportWorksheetProperties = false   // Omit worksheet properties
                };

                // Define output file path
                string outputPath = "output_without_properties.html";

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file saved without properties to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
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