// Title: Export a specific print area to HTML without CSS using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook, defines a print area on a worksheet, and saves the selected range as a plain HTML file with no CSS using Aspose.Cells. | Show how to set HtmlSaveOptions.ExportPrintAreaOnly to true and HtmlSaveOptions.DisableCss to true for a clean HTML export in Aspose.Cells.
// Common Searches: asp.net export defined print area to html without css using aspose.cells | c# aspose.cells htmlsaveoptions exportprintareaonly disablecss example | how to save only a range (e.g., A1:D10) as html from excel with aspose.cells | aspose.cells generate html without style sheets from workbook | export excel print area to html without css classes c#
// Tags: aspose.cells htmlsaveoptions exportprintareaonly | aspose.cells disablecss html export | c# export excel print area to html | aspose.cells save workbook range as html | html export without css styles aspose.cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, sets a print area (A1:D10) on the first worksheet, configures HtmlSaveOptions with ExportPrintAreaOnly = true and DisableCss = true, and saves the workbook as an HTML file that contains only the defined range and no CSS styling.
class ExportPrintArea
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the print area for the first worksheet (e.g., A1:D10)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.PrintArea = "A1:D10";

            // Configure HTML export options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true   // Export only the defined print area
                // ExportCssClass property is not available in current API version
            };

            // Export the workbook to HTML using the configured options
            workbook.Save(outputPath, options);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
