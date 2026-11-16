// Title: Enable original numeric formatting when converting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# keep Excel number formatting when saving as HTML | How to preserve numeric cell styles during Excel to HTML conversion using Aspose.Cells | HtmlSaveOptions to retain original number format in HTML output from .NET | Convert Excel to HTML with Aspose.Cells while maintaining numeric formats | C# Aspose.Cells export workbook to HTML without losing number formatting
// Tags: Aspose.Cells HtmlSaveOptions numeric formatting | Excel to HTML conversion preserve number style | C# save workbook as HTML with original cell formatting | retain numeric format Aspose.Cells export | HTML export keep Excel number format .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, configures Aspose.Cells HtmlSaveOptions (which preserve numeric formatting by default), and saves the workbook as an HTML file, demonstrating how to keep original number formats in the HTML output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options (numeric format is preserved by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as an HTML file with the specified options
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
