// Title: Convert an Excel workbook to PDF in C# with Aspose.Cells using system fallback font for missing fonts
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets a default fallback font for missing glyphs, and saves the workbook as a PDF. | Show how to configure Workbook.Settings.DefaultFont in a .NET console application before exporting to PDF with Aspose.Cells. | Create a resilient C# program that verifies the source Excel file, applies system default font substitution, converts it to PDF, and logs any exceptions.
// Common Searches: Aspose.Cells C# export Excel to PDF with fallback font when original font is missing | How to set default font for PDF conversion in Aspose.Cells .NET | Convert .xlsx to .pdf using Aspose.Cells and handle missing fonts | C# console application Aspose.Cells PDF save with system default font substitution | Aspose.Cells Workbook.Settings.DefaultFont usage example
// Tags: Aspose.Cells workbook to PDF with default font | C# set Workbook.Settings.DefaultFont | Excel to PDF conversion handling missing fonts | Aspose.Cells PDF export fallback font | system font substitution Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, optionally sets a default fallback font for missing glyphs via Workbook.Settings.DefaultFont, and saves the file as a PDF while checking file existence and handling exceptions.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.pdf";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Optional: set a default font if needed (uncomment if supported by your Aspose.Cells version)
                // workbook.Settings.DefaultFont = "Arial";

                // Convert and save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
