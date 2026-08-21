// Title: Export Worksheet Names as HTML Headings with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, name its worksheets, set the page header placeholder (&A) to show the sheet name, enable ExportPageHeaders, and save all sheets as a single HTML file where each worksheet name is rendered as a heading element.
// Keywords: Aspose.Cells | C# | HtmlSaveOptions | ExportPageHeaders | worksheet name heading | HTML export | single file output | page header placeholder &A | multiple sheets to HTML | Aspose.Cells example
// Common Searches: Aspose.Cells export worksheet names as HTML headings | HtmlSaveOptions ExportPageHeaders C# example | Save multiple Excel sheets to one HTML file with headings | How to use &A placeholder for sheet name in HTML output | Generate <h1> tags from Excel sheet names using Aspose.Cells
// Developer Intent: Generate a single HTML document where each worksheet name appears as a heading element.
// Use Cases: Publish a multi‑sheet Excel workbook as a web‑ready report with clear section titles. | Create documentation that lists each worksheet with its name as a visible heading. | Build a single‑page HTML dashboard that groups data by worksheet, using headings for navigation.
// AI Prompts: Show how to change the heading level (e.g., <h2> instead of <h1>) when ExportPageHeaders is enabled. | Explain how to export only selected worksheets as a single HTML file with headings for each sheet. | Provide code to apply custom CSS styles to the heading elements generated from worksheet names.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, name its worksheets, set the page header placeholder (&A) to show the sheet name, enable ExportPageHeaders, and save all sheets as a single HTML file where each worksheet name is rendered as a heading element.
    public class HtmlExportWorksheetNamesAsHeadings
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("HTML file generated successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and add a few worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("Data");
            workbook.Worksheets.Add("Report");

            // Set page header to display sheet name using &A placeholder
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.SetHeader(1, "&A"); // Center section will show the sheet name
            }

            // Configure HTML save options to export page headers as headings
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPageHeaders = true,
                SaveAsSingleFile = true,
                ShowAllSheets = true
            };

            // Save the workbook as HTML
            string outputPath = "WorksheetsWithHeadings.html";
            workbook.Save(outputPath, htmlOptions);
        }
    }
}
