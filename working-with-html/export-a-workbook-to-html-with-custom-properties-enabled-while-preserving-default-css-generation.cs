// Title: Export Workbook to HTML with Document & Workbook Properties Using Aspose.Cells (C#)
// Description: Shows how to create a workbook, set built‑in document properties (Author, Title), enable ExportWorkbookProperties and ExportDocumentProperties in HtmlSaveOptions, and save the file as HTML while preserving Aspose.Cells' default CSS styling.
// Keywords: Aspose.Cells | C# HTML export | ExportWorkbookProperties | ExportDocumentProperties | HtmlSaveOptions | default CSS | Excel to HTML | custom document properties | workbook metadata | Aspose.Cells .NET
// Common Searches: Aspose.Cells export HTML with document properties C# | How to keep default CSS when saving Excel as HTML Aspose | Enable ExportWorkbookProperties in HtmlSaveOptions | Save Excel workbook to HTML with author and title metadata | Aspose.Cells HtmlSaveOptions default style preservation
// Developer Intent: Generate an HTML file from a workbook that includes author/title metadata and retains the library’s standard CSS.
// Use Cases: Publish Excel data on a website with SEO‑friendly metadata. | Archive spreadsheets as HTML while preserving styling and property information. | Feed HTML reports into a content management system that reads embedded workbook properties. | Automate batch conversion of Excel templates to HTML with consistent look and metadata.
// AI Prompts: Write C# code using Aspose.Cells to export a workbook to HTML with ExportWorkbookProperties and ExportDocumentProperties turned on, without disabling the default CSS. | Explain which HtmlSaveOptions flags control property export and CSS generation, and show how to verify the properties appear in the resulting HTML. | Provide guidance on customizing the output path, file name, and embedding additional custom document properties while keeping default styling.

using System;
using Aspose.Cells;

namespace AsposeCellsExportHtml
{
    // Shows how to create a workbook, set built‑in document properties (Author, Title), enable ExportWorkbookProperties and ExportDocumentProperties in HtmlSaveOptions, and save the file as HTML while preserving Aspose.Cells' default CSS styling.
    public class ExportWithCustomProperties
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add some sample data
                worksheet.Cells["A1"].PutValue("Hello World!");

                // Set custom document properties (author and title)
                workbook.BuiltInDocumentProperties.Author = "John Doe";
                workbook.BuiltInDocumentProperties.Title = "Sample Workbook";

                // Create HTML save options
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    // Enable exporting of workbook and document properties (default is true)
                    ExportWorkbookProperties = true,
                    ExportDocumentProperties = true
                };

                // Save the workbook as HTML with the specified options
                workbook.Save("SampleOutput.html", options);
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
            ExportWithCustomProperties.Run();
        }
    }
}
