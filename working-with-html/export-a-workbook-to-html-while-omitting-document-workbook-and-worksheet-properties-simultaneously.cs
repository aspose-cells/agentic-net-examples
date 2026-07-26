// Title: Export Aspose.Cells Workbook to HTML without Document, Workbook, or Worksheet Properties (C#)
// Description: Shows how to save an Aspose.Cells workbook as HTML while suppressing all built‑in document, workbook, and worksheet properties using HtmlSaveOptions in C#.
// Keywords: Aspose.Cells HTML export | C# HtmlSaveOptions | ExportDocumentProperties false | ExportWorkbookProperties false | ExportWorksheetProperties false | remove Excel metadata HTML | save workbook as clean HTML | Aspose.Cells hide properties | Aspose.Cells .NET | Excel to HTML without metadata
// Common Searches: Aspose.Cells export HTML without properties | How to hide document properties in HTML output using Aspose.Cells | C# save Excel as HTML without workbook metadata | HtmlSaveOptions exclude worksheet properties | Remove author title from HTML export Aspose.Cells
// Developer Intent: Create an HTML file from a workbook that contains no document, workbook, or worksheet metadata.
// Use Cases: Generating lightweight web reports where metadata must be omitted for privacy. | Archiving spreadsheet snapshots as clean HTML without exposing author or title information. | Embedding Excel data in emails or web pages without revealing internal workbook properties. | Preparing public‑download HTML files that hide proprietary spreadsheet metadata.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with ExportDocumentProperties, ExportWorkbookProperties, and ExportWorksheetProperties set to false. | Explain step‑by‑step how HtmlSaveOptions can be configured to exclude all metadata when converting an Excel file to HTML in .NET. | Provide a minimal example that preserves cell formatting while stripping all document and worksheet properties during HTML export with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Shows how to save an Aspose.Cells workbook as HTML while suppressing all built‑in document, workbook, and worksheet properties using HtmlSaveOptions in C#.
    public class ExportWithoutProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "SampleSheet";
                sheet.Cells["A1"].PutValue("Hello World!");

                // Set some document and workbook properties (these will be omitted in the HTML output)
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

                // Configure HTML save options to exclude document, workbook, and worksheet properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportDocumentProperties = false,   // Omit document properties
                    ExportWorkbookProperties = false,   // Omit workbook properties
                    ExportWorksheetProperties = false   // Omit worksheet properties
                };

                // Save the workbook as HTML with the specified options
                string outputPath = "OutputWithoutProperties.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file saved without document, workbook, and worksheet properties: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during export: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWithoutProperties.Run();
        }
    }
}
