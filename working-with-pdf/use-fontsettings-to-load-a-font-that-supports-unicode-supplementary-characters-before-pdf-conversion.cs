// Title: Load an Excel workbook, apply a Unicode‑compatible default font, and save it as PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an .xlsx file using Aspose.Cells, sets Workbook.DefaultStyle.Font to a Unicode‑capable typeface such as Arial Unicode MS, and saves the workbook as a PDF. | Show how to configure FontSettings or the default style in Aspose.Cells for .NET so that supplementary Unicode characters render correctly during Excel‑to‑PDF conversion.
// Common Searches: how to use a Unicode font with Aspose.Cells when converting Excel to PDF in C# | Aspose.Cells .NET set default workbook font for supplementary characters before PDF export | C# example of loading Arial Unicode MS in Aspose.Cells to preserve emojis in PDF | configure FontSettings in Aspose.Cells to support Unicode characters during PDF conversion
// Tags: Aspose.Cells set default font Unicode | Excel to PDF conversion Unicode support | C# Aspose.Cells FontSettings PDF | load Arial Unicode MS Aspose.Cells | supplementary Unicode characters PDF export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an existing Excel workbook, changes its DefaultStyle.Font to a Unicode‑compatible typeface (e.g., Arial Unicode MS) to ensure proper rendering of supplementary characters, and then saves the workbook as a PDF using Aspose.Cells for .NET, including basic file existence checks and error handling.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The input file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Set a default font that supports Unicode characters (e.g., Arial Unicode MS)
                // Use the DefaultStyle's Font property to specify the default font.
                workbook.DefaultStyle.Font.Name = "Arial Unicode MS";

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"PDF file has been successfully created at \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
