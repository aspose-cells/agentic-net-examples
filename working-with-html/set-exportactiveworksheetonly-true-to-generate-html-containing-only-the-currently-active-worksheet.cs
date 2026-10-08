// Title: Export only the active worksheet to HTML using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel workbook, marks a specific worksheet as active, and saves only that sheet as an HTML file with Aspose.Cells. | Show how to set up Aspose.Cells HtmlSaveOptions so that ExportActiveWorksheetOnly is enabled during HTML conversion.
// Common Searches: Aspose.Cells export active worksheet only to HTML in C# | C# HtmlSaveOptions ExportActiveWorksheetOnly property example | How to save a single Excel sheet as HTML with Aspose.Cells | Set active sheet index before HTML conversion using Aspose.Cells | Generate HTML for one worksheet from a workbook using Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportActiveWorksheetOnly | C# active worksheet HTML export | Aspose.Cells set active sheet index | single worksheet HTML conversion Aspose.Cells | Aspose.Cells save single sheet as HTML

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, sets the desired worksheet as active, configures HtmlSaveOptions with ExportActiveWorksheetOnly = true, and saves the workbook as an HTML file that contains only the active worksheet.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the first worksheet as the active sheet
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure HTML save options to export only the active worksheet
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportActiveWorksheetOnly = true
            };

            // Save the workbook as HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
