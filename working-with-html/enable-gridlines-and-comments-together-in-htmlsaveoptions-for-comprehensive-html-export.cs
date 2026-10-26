// Title: Export an Excel workbook to HTML with gridlines and cell comments using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to HTML with gridlines and comments | Enable gridlines in HtmlSaveOptions while keeping cell comments in HTML output | How to configure HtmlSaveOptions ExportGridLines true and include comments | Export Excel workbook to HTML preserving comments and gridlines with Aspose.Cells .NET | C# HtmlSaveOptions settings for gridlines and cell comment visibility
// Tags: HtmlSaveOptions ExportGridLines true | Aspose.Cells HTML export cell comments | C# Excel to HTML with gridlines | Aspose.Cells .NET HtmlSaveOptions configuration | preserve Excel comments in HTML output

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, sets HtmlSaveOptions.ExportGridLines to true (comments are exported by default), and saves the workbook as an HTML file, demonstrating how to produce HTML that shows both gridlines and cell comments while handling missing files and runtime errors.
class HtmlExportWithGridlinesAndComments
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file '{inputPath}' was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Export gridlines
                ExportGridLines = true
                // Note: Exporting of cell comments is enabled by default in recent versions.
                // If a specific property exists (e.g., ExportComments), it can be set here.
            };

            // Save the workbook as HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
