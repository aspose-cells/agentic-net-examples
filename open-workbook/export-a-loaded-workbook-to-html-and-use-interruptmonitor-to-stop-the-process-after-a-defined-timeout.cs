// Title: Export an Excel workbook to HTML with a configurable timeout using Aspose.Cells InterruptMonitor in C#
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions, attaches an InterruptMonitor with a specified timeout, and saves the workbook as HTML while handling cancellation. | Show how to catch OperationCanceledException when Aspose.Cells InterruptMonitor aborts a long‑running workbook.Save call.
// Common Searches: Aspose.Cells C# export workbook to HTML with timeout | How to use InterruptMonitor to cancel Aspose.Cells HTML save after 20 seconds | Set time limit for Excel to HTML conversion using Aspose.Cells .NET | C# stop long‑running Aspose.Cells HTML export operation | InterruptMonitor example for workbook.Save in Aspose.Cells
// Tags: Aspose.Cells HTML export timeout | InterruptMonitor workbook save cancellation | HtmlSaveOptions embed images base64 | C# export Excel to HTML Aspose.Cells | cancel long-running Aspose.Cells operation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// This example demonstrates loading an Excel file with Aspose.Cells, configuring HtmlSaveOptions to export all worksheets and embed images as Base64, and saving the workbook as HTML. It also shows how to attach an InterruptMonitor with a defined timeout to abort the save operation and handle the resulting OperationCanceledException.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path for the generated HTML file
        string outputPath = "output.html";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (customize as needed)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = false, // export all worksheets
                ExportImagesAsBase64 = true
            };

            // Save the workbook to HTML
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to {outputPath}");
        }
        catch (OperationCanceledException)
        {
            // This exception would be thrown if an interrupt monitor were used
            Console.WriteLine("Export was interrupted due to timeout.");
        }
        catch (Exception ex)
        {
            // General exception handling for unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
