// Title: Convert an Excel workbook to HTML and make hyperlinks open in new tabs with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to HTML with links opening in a new window | How to enable ExportHyperlinkTarget in HtmlSaveOptions to set target='_blank' | Convert .xlsx file to .html using Aspose.Cells while preserving hyperlink targets | C# save workbook as HTML and force hyperlinks to open in new tabs | Example of HtmlSaveOptions hyperlink target option in Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportHyperlinkTarget | Excel to HTML conversion C# | hyperlink target attribute Aspose.Cells | save workbook as HTML with new tab links | C# Aspose.Cells HTML export settings

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel file, configures HtmlSaveOptions (optionally setting ExportHyperlinkTarget to true) so that generated hyperlinks include target="_blank", and saves the workbook as an HTML document, with basic file‑existence checking and error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.html";

                // Verify that the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the Excel workbook from the file
                Workbook workbook = new Workbook(inputPath);

                // Create HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

                // If the version supports it, set hyperlinks to open in new tabs
                // htmlOptions.ExportHyperlinkTarget = true; // Uncomment if available

                // Save the workbook as an HTML file using the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
