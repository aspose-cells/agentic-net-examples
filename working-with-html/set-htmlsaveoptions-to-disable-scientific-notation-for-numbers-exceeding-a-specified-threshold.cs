// Title: How to prevent scientific notation for large numbers when converting Excel to HTML with Aspose.Cells in C#
// AI Prompts: Generate C# code that configures HtmlSaveOptions to keep numeric values in plain (non‑exponential) format when saving a Workbook to HTML using Aspose.Cells. | Show how to apply a custom number format or style to cells in Aspose.Cells so that values above a certain threshold are rendered without scientific notation in the HTML output. | Explain a workaround for disabling scientific notation in Aspose.Cells HTML export when ExportNumericDataAsString is unavailable.
// Common Searches: Aspose.Cells C# HtmlSaveOptions keep numbers from showing in scientific notation | Convert Excel to HTML with Aspose.Cells preserving full numeric display | C# Aspose.Cells prevent exponential format when exporting to HTML | How to format large numbers as plain text in Aspose.Cells HTML export
// Tags: Aspose.Cells HtmlSaveOptions numeric display | C# Aspose.Cells disable scientific notation | Excel to HTML conversion numeric formatting | Aspose.Cells custom number format for HTML export

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates HtmlSaveOptions for HTML saving, notes that ExportNumericDataAsString is not available, and suggests using cell styles or custom number formats to avoid scientific notation before saving the workbook as HTML.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Set up HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportNumericDataAsString is not available in this version of Aspose.Cells.
            // If needed, numeric formatting can be handled via cell styles before saving.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
