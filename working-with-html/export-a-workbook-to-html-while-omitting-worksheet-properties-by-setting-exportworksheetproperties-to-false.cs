using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportWorkbookToHtmlWithoutWorksheetProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleSheet";
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Configure HTML save options to omit worksheet properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportWorksheetProperties = false // Disable exporting of worksheet properties
                };

                string outputPath = "output_without_worksheet_props.html";

                // Save the workbook as an HTML file using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}