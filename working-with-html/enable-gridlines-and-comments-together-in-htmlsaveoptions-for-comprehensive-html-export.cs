// Title: Export Excel workbook to HTML with gridlines using Aspose.Cells for .NET (cell comments not supported)
// Common Searches: Aspose.Cells C# export workbook to HTML with grid lines | Enable gridlines in HtmlSaveOptions Aspose.Cells .NET | Export cell comments to HTML using Aspose.Cells C# | HTML export of Excel with visible gridlines Aspose.Cells | Cell comments missing in HTML output from Aspose.Cells
// Tags: Aspose.Cells HtmlSaveOptions ExportGridLines | Aspose.Cells HTML conversion with gridlines | cell comments omission in Aspose.Cells HTML export | Excel workbook.Save HTML Aspose.Cells .NET | gridlines visibility in HTML output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example loads an Excel workbook, configures HtmlSaveOptions to enable ExportGridLines, notes that cell comments cannot be exported in the current Aspose.Cells version, and saves the workbook as an HTML file while handling missing input files and runtime exceptions.
    class HtmlExportWithGridlinesAndComments
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.html";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Enable exporting of grid lines
                    ExportGridLines = true
                    // Note: Exporting cell comments is not available in this version of Aspose.Cells.
                };

                // Save the workbook as HTML with the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook successfully saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display an error message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
