// Title: Set a Unicode‑compatible default font in Aspose.Cells C# before converting an Excel workbook to PDF
// Common Searches: how to specify a Unicode font for PDF export in Aspose.Cells C# | Aspose.Cells convert Excel to PDF with supplementary characters support | C# load Noto Sans as default workbook font for PDF generation using Aspose.Cells | prevent missing characters when saving Excel as PDF with Aspose.Cells | configure default style font for Unicode in Aspose.Cells PDF conversion
// Tags: Aspose.Cells set default font for PDF | C# Unicode font for Excel to PDF conversion | supplementary character support in PDF export | load Noto Sans with Aspose.Cells | Excel workbook default style font Unicode

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, assigns a Unicode‑compatible default font (e.g., Noto Sans) to the workbook's style, ensures the output directory exists, and then saves the workbook as a PDF, handling any errors that occur.
    class Program
    {
        static void Main()
        {
            try
            {
                // Verify that the input workbook exists before loading
                string inputPath = @"C:\Data\input.xlsx";
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"Input file not found: {inputPath}");

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Set a default font that supports Unicode characters (must be installed on the system)
                workbook.DefaultStyle.Font.Name = "Noto Sans";

                // Define output path and ensure the directory exists
                string outputPath = @"C:\Data\output.pdf";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Convert and save the workbook to PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"PDF saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
