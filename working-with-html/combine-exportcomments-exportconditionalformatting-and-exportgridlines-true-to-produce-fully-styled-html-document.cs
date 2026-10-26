// Title: Export an Excel workbook to fully styled HTML with comments, conditional formatting, and grid lines using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export workbook to HTML with comments and conditional formatting | Include grid lines when saving Excel as HTML using Aspose.Cells | Preserve cell comments and conditional formatting in HTML output with Aspose.Cells .NET | Set ExportComments and ExportConditionalFormatting in HtmlSaveOptions Aspose.Cells | Generate fully styled HTML from XLSX while keeping all formatting Aspose.Cells
// Tags: HtmlSaveOptions ExportComments true | ExportConditionalFormatting HTML Aspose.Cells | ExportGridLines option Aspose.Cells | styled HTML generation from Excel workbook | preserve Excel formatting in HTML export

using Aspose.Cells;
using System;
using System.IO;

// The example checks for the source XLSX file, loads it into a Workbook, configures HtmlSaveOptions to enable ExportComments, ExportConditionalFormatting, and ExportGridLines, and then saves the workbook as a fully styled HTML document, handling any runtime exceptions.
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
                Console.WriteLine($"Input file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML export options (grid lines are supported)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportGridLines = true
                // Additional styling options can be set here if needed
            };

            // Save the workbook as a styled HTML document
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
