// Title: Convert an Excel workbook to HTML with percentage-based column widths using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets HtmlSaveOptions.WidthScalable to true, and saves the workbook as an HTML file where column widths are expressed as percentages. | Show how to handle missing input files and configure Aspose.Cells HTML save options to produce proportionally scaled column widths in the generated HTML.
// Common Searches: Aspose.Cells how to export Excel to HTML with scalable column widths in C# | C# HtmlSaveOptions WidthScalable property example for percentage column widths | Save workbook as HTML preserving column width percentages using Aspose.Cells .NET | Convert .xlsx to .html with column width scaling Aspose.Cells tutorial
// Tags: Aspose.Cells HTML export using scalable widths | C# HtmlSaveOptions column width percentage rendering | Excel to HTML conversion with Aspose.Cells .NET | column widths rendered as percentages in HTML | configure HTML save options for column width scaling

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, configures HtmlSaveOptions with WidthScalable enabled to output column widths as percentages, and saves the workbook as an HTML file, including basic error handling for missing input files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.html";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                // Enable scalable column widths
                WidthScalable = true
                // Note: ColumnWidthType property is not available in this version of Aspose.Cells
            };

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
