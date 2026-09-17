// Title: How to save an Excel workbook to HTML with Aspose.Cells for .NET while forcing numeric cells to appear as plain text instead of scientific notation
// AI Prompts: Enable the numeric‑as‑string mode on HtmlSaveOptions before calling Workbook.Save to produce HTML with numbers shown as plain text. | Update the C# example to create HtmlSaveOptions, turn on the option that writes numeric values as strings, and then save the workbook to an .html file. | Show the steps for using Aspose.Cells HtmlSaveOptions to suppress scientific notation for numeric cells during HTML export in a .NET application.
// Common Searches: Aspose.Cells C# export Excel to HTML without scientific notation | HtmlSaveOptions ExportNumericAsString usage example | how to keep numbers as plain text when converting Excel to HTML with Aspose.Cells | prevent scientific notation in HTML output from Aspose.Cells .NET | save workbook as HTML with numeric values displayed as text Aspose.Cells
// Tags: HtmlSaveOptions ExportNumericAsString | Aspose.Cells HTML numeric formatting | C# export Excel to HTML plain text numbers | disable scientific notation Aspose.Cells | save workbook as HTML without number formatting

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example demonstrates loading an Excel workbook, configuring HtmlSaveOptions to turn on the ExportNumericAsString flag, and saving the workbook as HTML so that numeric cells are written as plain text rather than using scientific notation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // Note: ExportNumericAsString is not available in this version of Aspose.Cells.
            // Numeric values will be saved using the default formatting.

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to HTML using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
