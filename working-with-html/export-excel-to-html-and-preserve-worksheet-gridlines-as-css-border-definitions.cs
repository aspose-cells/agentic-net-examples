// Title: Export an Excel workbook to HTML with worksheet gridlines rendered as CSS borders using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file and saves it as HTML with gridlines exported as CSS borders using Aspose.Cells. | Adjust the export to include only the active worksheet while keeping the gridlines visible in the HTML output. | Insert custom CSS into the generated HTML without disabling the ExportGridLines option.
// Common Searches: Aspose.Cells export Excel to HTML with gridlines as CSS borders | C# HtmlSaveOptions ExportGridLines true example | How to keep worksheet gridlines when converting .xlsx to HTML using Aspose.Cells | Save all worksheets to a single HTML file with borders using Aspose.Cells .NET | Preserve Excel cell borders in HTML output with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | C# export Excel to HTML with CSS borders | Preserve worksheet gridlines in HTML output | Export all worksheets to single HTML file Aspose.Cells | Convert .xlsx to HTML with borders .NET

using System;
using System.IO;
using Aspose.Cells;

// The program loads input.xlsx, enables HtmlSaveOptions.ExportGridLines, and saves the workbook as output.html, preserving worksheet gridlines as CSS borders.
class ExportExcelToHtml
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options to preserve gridlines
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
