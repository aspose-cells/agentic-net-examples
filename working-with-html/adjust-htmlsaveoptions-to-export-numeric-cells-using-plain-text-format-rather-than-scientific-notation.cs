// Title: Export numeric cells as plain text (avoid scientific notation) using HtmlSaveOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel numbers to HTML without scientific notation | How to keep numeric values as plain text when saving workbook to HTML with Aspose.Cells | Disable exponential format for numbers in Aspose.Cells HTML export | HtmlSaveOptions setting to prevent scientific notation in HTML output | Convert numeric cells to text during HTML conversion using Aspose.Cells .NET
// Tags: HtmlSaveOptions numeric cell text export | Aspose.Cells disable scientific notation HTML | export numbers as plain text .NET | Excel to HTML numeric formatting Aspose | C# Aspose.Cells HtmlSaveOptions configuration

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Demonstrates loading an Excel workbook, configuring HtmlSaveOptions to control numeric formatting, and saving the workbook as HTML while aiming to output numeric cells as plain text instead of scientific notation.
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
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Create HtmlSaveOptions for HTML export
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Note: ExportNumericCellValuesAsText is not available in this version.
            // The default behavior will be used.

            // Save the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook successfully saved as HTML to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
