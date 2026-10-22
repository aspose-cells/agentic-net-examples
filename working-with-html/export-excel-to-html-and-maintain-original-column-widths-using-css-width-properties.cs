// Title: Export an Excel workbook to HTML with original column widths rendered as CSS width styles using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions to embed images as Base64, and saves the workbook as HTML where each column's width is expressed with CSS width attributes. | Show how to check for the Excel file's existence, handle errors, and use Aspose.Cells to generate a self‑contained HTML file while attempting to retain the worksheet's column dimensions.
// Common Searches: Aspose.Cells .NET keep Excel column widths in HTML export | C# save workbook as HTML with CSS column width styling using Aspose.Cells | How to embed images as Base64 and preserve column dimensions when converting Excel to HTML with Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions column width CSS | C# export Excel to HTML with embedded Base64 images | preserve Excel column dimensions in generated HTML | self-contained HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, configures HtmlSaveOptions to embed images as Base64, and saves the workbook as output.html. Although the specific ExportCustomColumnWidth option is unavailable, the code demonstrates the standard approach for creating HTML with CSS‑styled column widths and handling any runtime exceptions.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Embed images directly into the HTML using Base64 (keeps the HTML self‑contained)
                ExportImagesAsBase64 = true
                // Note: ExportCustomColumnWidth and ExportCustomRowHeight are not available in this version
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
