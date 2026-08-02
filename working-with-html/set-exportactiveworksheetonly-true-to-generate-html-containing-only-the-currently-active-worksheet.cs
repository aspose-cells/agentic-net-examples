// Title: Export Active Worksheet to HTML with Aspose.Cells for .NET (C#)
// Description: Shows how to set HtmlSaveOptions.ExportActiveWorksheetOnly to true so that saving a Workbook generates an HTML file that includes only the currently active sheet.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportActiveWorksheetOnly | C# HTML export | single worksheet HTML | active sheet only | Aspose.Cells .NET | Workbook to HTML | save active worksheet | Aspose.Cells example
// Common Searches: Aspose.Cells export only active sheet to HTML C# | HtmlSaveOptions ExportActiveWorksheetOnly usage | Generate HTML from a specific worksheet using Aspose.Cells | C# save workbook as HTML with active worksheet only | How to hide other sheets when exporting to HTML Aspose
// Developer Intent: Produce an HTML file that contains just the active worksheet from a workbook that has multiple sheets.
// Use Cases: Display a preview of the user‑selected sheet on a web page without loading other worksheets. | Create a lightweight HTML report for email that reveals data from only one sheet. | Export a confidential worksheet to HTML while keeping the remaining sheets private.
// AI Prompts: Provide a C# code snippet that sets ExportActiveWorksheetOnly to true and saves only the active sheet as HTML using Aspose.Cells. | Explain the impact of HtmlSaveOptions.ExportActiveWorksheetOnly on the generated HTML and what other options may need adjustment. | Show how to programmatically select a worksheet as active and then export just that sheet to HTML with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExportActiveWorksheetOnlyDemo
{
    // Shows how to set HtmlSaveOptions.ExportActiveWorksheetOnly to true so that saving a Workbook generates an HTML file that includes only the currently active sheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets.Add("SecondSheet");

            // Add sample data to both sheets
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in First Sheet");
            workbook.Worksheets[1].Cells["A1"].PutValue("Data in Second Sheet");

            // Set the first worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportActiveWorksheetOnly = true;

            // Define output path (adjust as needed)
            string outputPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ActiveSheetOnly.html");

            // Save the workbook as HTML using the configured options
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {outputPath}");
        }
    }
}
