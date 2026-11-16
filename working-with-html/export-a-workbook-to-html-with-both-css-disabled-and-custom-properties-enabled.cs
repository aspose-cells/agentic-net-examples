// Title: Export an Excel workbook to HTML without CSS and with custom document properties using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML without generating CSS files | How to include custom document properties when saving Excel as HTML with Aspose.Cells | Disable external stylesheet in HtmlSaveOptions while preserving custom properties in .NET | Save Excel workbook as HTML with inline styles only using Aspose.Cells C# | Export Excel to HTML and keep custom properties accessible in the output using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions disable external CSS | export workbook to HTML with custom document properties | C# Aspose.Cells HTML export without stylesheet | custom document properties in Aspose.Cells HTML output | HTML export options for Excel using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorkbook
{
    // The example creates a workbook, adds sample data and a custom document property, configures HtmlSaveOptions for HTML format, ensures the output directory exists, and saves the workbook as an HTML file without generating external CSS files.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook = new Workbook();

                // Populate the workbook with sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Text");
                sheet.Cells["B1"].PutValue(12345);

                // Add a custom document property
                workbook.CustomDocumentProperties.Add("MyCustomProperty", "CustomValue");

                // Configure HTML export options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
                // (Optional) Set additional options here if needed

                // Define output file path
                string outputFile = "ExportedWorkbook.html";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Export the workbook to HTML
                workbook.Save(outputFile, htmlOptions);
                Console.WriteLine($"Workbook successfully exported to: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
