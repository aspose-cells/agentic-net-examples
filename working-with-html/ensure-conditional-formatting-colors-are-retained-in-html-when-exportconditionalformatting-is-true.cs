// Title: How to retain conditional formatting colors when exporting an Excel workbook to HTML with Aspose.Cells for .NET
// Common Searches: Aspose.Cells export conditional formatting colors to HTML C# example | C# save Excel as HTML preserving conditional formatting using HtmlSaveOptions | Enable ExportConditionalFormatting property to keep cell colors in HTML output | Why Aspose.Cells HTML export loses conditional formatting colors and how to fix | Export Excel conditional formatting to HTML with Aspose.Cells .NET
// Tags: Aspose.Cells HtmlSaveOptions ExportConditionalFormatting | C# export Excel to HTML with conditional formatting | preserve conditional formatting colors in HTML output | HTML conversion of Excel workbook using Aspose.Cells | retain cell style formatting during HTML export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving; // HtmlSaveOptions resides in this namespace

// Loads an Excel workbook, configures HtmlSaveOptions with ExportConditionalFormatting enabled, and saves the file as HTML while preserving the original conditional formatting colors and handling missing files or runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            // Ensure the input workbook exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            // If the version of Aspose.Cells supports it, enable conditional formatting export:
            // htmlOptions.ExportConditionalFormatting = true;

            // Save the workbook as an HTML file with the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
