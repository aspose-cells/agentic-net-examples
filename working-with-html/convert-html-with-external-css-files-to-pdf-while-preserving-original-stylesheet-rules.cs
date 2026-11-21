// Title: How to convert an HTML page with external CSS to PDF using Aspose.Cells in C# while preserving stylesheet formatting
// Common Searches: C# Aspose.Cells convert HTML file that references external CSS to PDF preserving layout | Load HTML with linked stylesheet into Aspose.Cells workbook and export to PDF | Aspose.Cells HtmlLoadOptions example for preserving CSS styles during PDF conversion | Convert HTML with external CSS to PDF using Aspose.Cells SaveFormat.Pdf in C# | Handling missing HTML file error when converting to PDF with Aspose.Cells
// Tags: Aspose.Cells HtmlLoadOptions CSS handling | HTML to PDF conversion with external stylesheet | C# Aspose.Cells SaveFormat.Pdf preserving styles | Workbook load HTML external CSS | PDF export preserving HTML formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input HTML file, loads it into an Aspose.Cells Workbook using HtmlLoadOptions (which automatically processes linked CSS files), and saves the workbook as a PDF with SaveFormat.Pdf, ensuring the original HTML layout and stylesheet rules are retained while handling potential runtime errors.
class HtmlToPdfConverter
{
    static void Main()
    {
        string inputPath = "input.html";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input HTML file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the HTML file into a workbook using HtmlLoadOptions
            HtmlLoadOptions loadOptions = new HtmlLoadOptions();
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook as PDF, preserving the HTML formatting
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
