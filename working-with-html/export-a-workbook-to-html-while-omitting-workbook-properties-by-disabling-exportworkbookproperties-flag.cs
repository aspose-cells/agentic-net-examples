// Title: Export Aspose.Cells Workbook to HTML without Workbook Properties (ExportWorkbookProperties = false)
// Description: Creates a workbook, adds data, sets built‑in properties, disables ExportWorkbookProperties in HtmlSaveOptions, and saves the file as HTML so the output contains no author, title, or other workbook metadata.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportWorkbookProperties | disable workbook metadata | HTML export | save workbook as HTML | remove document properties | Aspose.Cells .NET
// Common Searches: Aspose.Cells export HTML without metadata | HtmlSaveOptions ExportWorkbookProperties false example | remove author and title from HTML output Aspose.Cells | save Excel workbook as HTML without document properties
// Developer Intent: Generate HTML from a workbook while omitting all embedded document properties.
// Use Cases: Publish spreadsheet data on a public website without exposing author or title information. | Create email‑friendly HTML snapshots of workbooks that preserve only cell content. | Automate bulk conversion of Excel files to clean HTML for archival or reporting purposes.
// AI Prompts: Write C# code that exports an Aspose.Cells workbook to HTML with ExportWorkbookProperties set to false and applies a custom stylesheet. | Explain the impact of the ExportWorkbookProperties flag on the generated HTML and how to confirm that properties are excluded. | Show how to combine other HtmlSaveOptions (e.g., embedding images, setting page margins) while disabling workbook property export.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Creates a workbook, adds data, sets built‑in properties, disables ExportWorkbookProperties in HtmlSaveOptions, and saves the file as HTML so the output contains no author, title, or other workbook metadata.
    public class ExportWithoutWorkbookProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and add some sample data
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Hello World!");

                // Set some workbook properties (these will be omitted in the HTML output)
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

                // Create HTML save options and disable exporting of workbook properties
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportWorkbookProperties = false // Omit workbook properties in the HTML file
                };

                // Save the workbook as HTML using the configured options
                string outputPath = "output_without_workbook_props.html";
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file saved without workbook properties: {outputPath}");
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
            ExportWithoutWorkbookProperties.Run();
        }
    }
}
