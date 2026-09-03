// Title: Export a single Excel worksheet to HTML with cell styles and merged cells preserved using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, configures HtmlSaveOptions to export only the active sheet with grid lines, and saves it as an HTML file while keeping all cell formatting and merged ranges using Aspose.Cells. | Show how to set up Aspose.Cells HtmlSaveOptions in a .NET application to retain cell styles and merged cell structures when converting a worksheet to HTML.
// Common Searches: Aspose.Cells .NET export active worksheet to HTML with original formatting | How to keep merged cells when saving Excel as HTML using Aspose.Cells | C# HtmlSaveOptions preserve cell styles and gridlines in HTML output | Convert specific sheet from XLSX to HTML with Aspose.Cells preserving layout | Save Excel workbook as HTML file while retaining merged ranges Aspose.Cells
// Tags: HtmlSaveOptions export active worksheet only | preserve cell formatting in HTML conversion Aspose.Cells | merged cell support in Aspose.Cells HTML output | C# export Excel to HTML with gridlines | Aspose.Cells HTMLSaveOptions styling retention

using System;
using System.IO;
using Aspose.Cells;

// The example loads 'input.xlsx', configures HtmlSaveOptions to export only the active worksheet with grid lines, and saves the result as 'output.html'. Cell styles and merged cell structures are retained automatically by Aspose.Cells for .NET.
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportActiveWorksheetOnly = true, // Export only the active sheet
                ExportGridLines = true            // Keep grid lines for visual fidelity
                // Merged cells and cell styles are exported by default
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
