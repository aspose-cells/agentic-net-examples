// Title: Convert an Excel workbook to a single HTML file with a worksheet navigation pane using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a .xlsx file and saves it as one HTML document containing a navigation pane for each worksheet using Aspose.Cells. | Show how to configure HtmlSaveOptions in Aspose.Cells to export all worksheets (ExportActiveWorksheetOnly = false) when saving to HTML. | Add robust file‑existence checking and exception handling to the Excel‑to‑HTML conversion example.
// Common Searches: how to export all worksheets to a single html file with navigation using aspose.cells c# | aspnet convert excel workbook to html with sheet links aspnet core | set ExportActiveWorksheetOnly false in Aspose.Cells HtmlSaveOptions example | c# generate html navigation pane for multiple Excel sheets using Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | export Excel workbook to single HTML Aspose.Cells | HTML navigation pane for Excel worksheets C# | C# Aspose.Cells convert workbook to HTML | save multiple worksheets as HTML Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// Loads input.xlsx, sets HtmlSaveOptions.ExportActiveWorksheetOnly = false to include all sheets, and saves the workbook as output.html with a navigation pane linking each worksheet; includes file‑existence check and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options: export all worksheets to generate a navigation pane
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as a single HTML file with a navigation pane
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
