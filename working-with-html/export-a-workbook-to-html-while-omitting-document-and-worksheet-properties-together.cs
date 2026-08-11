// Title: Export Aspose.Cells Workbook to HTML without Document or Worksheet Properties (C#)
// Description: Demonstrates how to save a workbook as HTML using Aspose.Cells while disabling ExportDocumentProperties, ExportWorksheetProperties, and ExportWorkbookProperties, producing clean HTML that contains no workbook metadata.
// Keywords: Aspose.Cells HTML export | ExportDocumentProperties false | ExportWorksheetProperties false | ExportWorkbookProperties false | C# save Excel as HTML without metadata | HtmlSaveOptions hide properties
// Common Searches: Aspose.Cells export to HTML without document properties C# | How to hide worksheet properties in HTML output Aspose.Cells | C# save Excel workbook as HTML without metadata | HtmlSaveOptions exclude workbook properties
// Developer Intent: Generate an HTML file from a workbook while stripping all document, worksheet, and workbook metadata.
// Use Cases: Create lightweight HTML reports that do not reveal author or title information. | Embed Excel data in web pages or emails without exposing internal metadata. | Produce clean HTML snippets for web applications where property data is unnecessary.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with ExportDocumentProperties, ExportWorksheetProperties, and ExportWorkbookProperties set to false. | Explain the effect of each HtmlSaveOptions flag (ExportDocumentProperties, ExportWorksheetProperties, ExportWorkbookProperties) on the resulting HTML. | Suggest additional HtmlSaveOptions settings (e.g., hide gridlines, apply custom CSS) that keep all properties excluded.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to save a workbook as HTML using Aspose.Cells while disabling ExportDocumentProperties, ExportWorksheetProperties, and ExportWorkbookProperties, producing clean HTML that contains no workbook metadata.
    public class ExportHtmlWithoutProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add some sample data
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello");
                sheet.Cells["B1"].PutValue("World");

                // Set some document properties (these will be omitted in the HTML output)
                workbook.BuiltInDocumentProperties.Author = "Test Author";
                workbook.BuiltInDocumentProperties.Title = "Test Title";

                // Configure HTML save options to exclude document and worksheet properties
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false,   // Omit document properties
                    ExportWorksheetProperties = false, // Omit worksheet properties
                    ExportWorkbookProperties = false   // Omit workbook properties (optional)
                };

                string outputPath = "output_without_properties.html";

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, options);
                Console.WriteLine($"HTML file saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
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
