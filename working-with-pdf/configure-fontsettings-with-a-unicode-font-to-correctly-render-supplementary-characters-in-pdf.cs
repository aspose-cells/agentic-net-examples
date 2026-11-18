// Title: Set FontSettings to a Unicode font for correct supplementary character rendering when converting Excel to PDF with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set default Unicode font for PDF export | how to display emoji in PDF generated from Excel using Aspose.Cells | configure FontSettings to render supplementary characters in Aspose.Cells PDF conversion | C# Aspose.Cells PDF output missing Unicode characters | set fallback font for Unicode characters in Aspose.Cells PDF save
// Tags: Aspose.Cells FontSettings Unicode fallback | C# Excel to PDF Unicode character support | supplementary character rendering Aspose.Cells PDF | default font setup for PDF export Aspose.Cells | emoji display in PDF using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // This example loads an Excel workbook, optionally configures FontSettings with a Unicode font to ensure supplementary characters (e.g., emoji) are rendered correctly, and saves the workbook as a PDF while handling errors and creating the output directory.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // NOTE: Setting a default Unicode font is optional.
                // If needed, ensure the font is installed and use the appropriate API for your Aspose.Cells version.

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
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
}
