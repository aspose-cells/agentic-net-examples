// Title: Convert an Excel workbook with ColorScale conditional formatting to HTML while preserving gradient colors using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file containing ColorScale conditional formatting and saves it as an HTML file using Aspose.Cells, ensuring the gradient colors are retained. | Show how to configure Aspose.Cells HtmlSaveOptions to export conditional formatting, specifically ColorScale rules, when converting a workbook to HTML in a .NET application.
// Common Searches: Aspose.Cells preserve ColorScale gradients when exporting Excel to HTML in C# | How to export conditional formatting ColorScale to HTML using Aspose.Cells .NET | C# convert workbook with ColorScale rules to HTML with Aspose.Cells | HtmlSaveOptions settings for keeping Excel color scales in HTML output
// Tags: Aspose.Cells HtmlSaveOptions export ColorScale | C# Excel to HTML conversion preserving conditional formatting | ColorScale gradient export Aspose.Cells .NET | SaveFormat.Html with conditional formatting | Excel workbook HTML output color scales

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// Loads 'ColorScaleWorkbook.xlsx', applies default HtmlSaveOptions, and saves as 'ColorScaleWorkbook.html' while preserving the ColorScale conditional formatting gradients.
class Program
{
    static void Main()
    {
        const string inputFile = "ColorScaleWorkbook.xlsx";
        const string outputFile = "ColorScaleWorkbook.html";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Error: Input file '{inputFile}' not found.");
                return;
            }

            // Load the workbook that contains ColorScale conditional formatting rules
            Workbook workbook = new Workbook(inputFile);

            // Configure HTML save options (conditional formatting is exported by default)
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);

            // Save the workbook as HTML with gradient colors reflected
            workbook.Save(outputFile, htmlOptions);
            Console.WriteLine($"Workbook successfully saved as HTML to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
