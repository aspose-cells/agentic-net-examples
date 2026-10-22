// Title: Export Excel workbook to HTML with worksheet gridlines as CSS borders using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets HtmlSaveOptions.ExportGridLines to true, and saves the workbook as an HTML file with gridlines rendered as CSS borders using Aspose.Cells. | Show how to add file‑existence validation and exception handling around an Aspose.Cells Excel‑to‑HTML conversion that preserves worksheet gridlines. | Demonstrate configuring Aspose.Cells HtmlSaveOptions to export all worksheets into a single HTML document while keeping the default gridline color.
// Common Searches: Aspose.Cells C# export Excel to HTML preserving gridlines as CSS borders | How to keep Excel worksheet gridlines when converting to HTML with Aspose.Cells | HtmlSaveOptions ExportGridLines true example for multiple worksheets | C# code sample converting .xlsx to .html with gridlines using Aspose.Cells | Export Excel file to HTML with default gridline color using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# Excel to HTML conversion with gridlines | preserve worksheet gridlines Aspose.Cells | export multiple worksheets to single HTML | HTML CSS borders from Excel gridlines

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// Loads an .xlsx file, configures HtmlSaveOptions to export gridlines, and saves the workbook as an HTML file, including file‑existence checking and error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to export gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,                 // Preserve worksheet gridlines
                ExportActiveWorksheetOnly = false      // Export all worksheets
                // GridLineColor property is not available in this version; default color will be used
            };

            // Save the workbook to an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
