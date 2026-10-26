// Title: Export an Excel worksheet to HTML with gridlines using Aspose.Cells for .NET (cell comments not supported in this version)
// Common Searches: Aspose.Cells export worksheet to HTML with gridlines C# | How to include cell comments when saving Excel as HTML with Aspose.Cells .NET | Enable gridlines in HTML output from Aspose.Cells | Export Excel comments to HTML using Aspose.Cells version limitation | HtmlSaveOptions ExportGridLines property example in C# | Aspose.Cells HTML export without cell comments limitation
// Tags: Aspose.Cells HTML export gridlines | Aspose.Cells export cell comments HTML limitation | HtmlSaveOptions ExportGridLines C# | Excel to HTML conversion Aspose.Cells .NET | HTML output with gridlines Aspose.Cells | cell comments export not available Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel workbook, selects the first worksheet, configures HtmlSaveOptions to enable gridlines, notes that the ExportCellComments property is unavailable in the current Aspose.Cells version, and saves the worksheet as an HTML file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.html";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Optional: work with a specific worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Configure HTML export options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Enable gridlines in the exported HTML
                ExportGridLines = true

                // Note: ExportCellComments property is not available in this version of Aspose.Cells.
            };

            // Export the workbook (or the active worksheet) to HTML with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
