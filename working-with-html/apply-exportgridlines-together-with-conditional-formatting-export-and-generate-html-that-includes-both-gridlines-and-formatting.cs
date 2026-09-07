// Title: Export Excel workbook to HTML with visible gridlines and conditional formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx workbook, enables grid line rendering in the HTML output, and attempts to preserve cell style rules, with graceful handling if the style‑export option is missing. | Demonstrate how to verify the source file, configure HtmlSaveOptions to render grid lines in the HTML, and provide a fallback notice when cell style rules cannot be exported by the current Aspose.Cells library. | Create a self‑contained console program that converts an Excel file to HTML, includes visible grid lines, logs success or errors, and documents the limitation regarding style‑rule export.
// Common Searches: Aspose.Cells C# export Excel to HTML with grid lines and conditional formatting | How to include Excel conditional formatting when saving as HTML using Aspose.Cells .NET | HtmlSaveOptions ExportGridLines true example C# | Missing ExportConditionalFormatting property in Aspose.Cells version | C# code to convert .xlsx to .html preserving cell styles with Aspose.Cells
// Tags: Aspose.Cells HTML export display cell borders | Aspose.Cells cell style rules HTML export | C# convert Excel to HTML with borders | fallback handling for unsupported style‑rule export | Excel to HTML conversion preserving cell styles .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx workbook, checks that the file exists, configures HtmlSaveOptions to enable grid line rendering, notes that ExportConditionalFormatting is unavailable in the current Aspose.Cells release, and saves the workbook as an HTML file while handling file‑not‑found and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Include grid lines in the generated HTML
                ExportGridLines = true
                // Note: ExportConditionalFormatting is not available in this version of Aspose.Cells
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
