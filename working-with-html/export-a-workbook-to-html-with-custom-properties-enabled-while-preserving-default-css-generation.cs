// Title: Export Aspose.Cells Workbook to HTML with Document Properties and Default CSS (C#)
// Description: Shows how to create a workbook, set built‑in properties (Author, Title), enable ExportWorkbookProperties in HtmlSaveOptions, and save the file as HTML while keeping Aspose.Cells' default CSS styling.
// Keywords: Aspose.Cells | C# HTML export | ExportWorkbookProperties | document properties | default CSS | HtmlSaveOptions | Workbook to HTML | Excel metadata in HTML | web publishing spreadsheet | Aspose.Cells tutorial
// Common Searches: Aspose.Cells export workbook to HTML with properties | C# HtmlSaveOptions ExportWorkbookProperties true | keep default CSS when saving Excel as HTML Aspose | include author and title in HTML output Aspose.Cells | how to export Excel metadata to HTML using Aspose
// Developer Intent: Generate an HTML file from a workbook that contains built‑in document metadata and uses the library’s standard CSS generation.
// Use Cases: Web‑based reports that need author and title information displayed in the HTML view. | Emailing spreadsheet snapshots without applying custom styles. | Automated batch conversion pipelines that preserve default styling and embedded metadata. | Embedding Excel metadata in web pages for compliance or audit trails.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to HTML, enable ExportWorkbookProperties, and retain the default CSS. | Demonstrate adding custom document properties and including them in the HTML output using Aspose.Cells. | Explain the difference between default CSS generation and inline styling when exporting to HTML with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AspiseCellsExamples
{
    // Shows how to create a workbook, set built‑in properties (Author, Title), enable ExportWorkbookProperties in HtmlSaveOptions, and save the file as HTML while keeping Aspose.Cells' default CSS styling.
    public class ExportWorkbookToHtmlWithProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data to the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello World!");

                // Set built‑in document properties (author and title)
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

                // Create HTML save options
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    // Ensure workbook properties are exported (default is true)
                    ExportWorkbookProperties = true
                };

                // Save the workbook as HTML, keeping default CSS generation
                workbook.Save("ExportedWorkbook.html", options);
                Console.WriteLine("Workbook exported successfully to ExportedWorkbook.html");
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
            ExportWorkbookToHtmlWithProperties.Run();
        }
    }
}
