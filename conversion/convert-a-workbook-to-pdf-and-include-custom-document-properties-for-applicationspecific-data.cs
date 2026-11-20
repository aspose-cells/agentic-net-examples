// Title: Convert an Excel workbook to PDF and embed custom document properties with Aspose.Cells for .NET (C#)
// Common Searches: how to add custom document properties to an Excel file before converting to PDF using Aspose.Cells C# | Aspose.Cells export workbook to PDF with custom metadata example | C# save .xlsx as PDF and include custom properties via Aspose.Cells | set custom document properties in Aspose.Cells workbook then convert to PDF | Aspose.Cells PDF conversion with custom metadata in .NET
// Tags: Aspose.Cells workbook custom metadata addition | PDF export of Excel with Aspose.Cells C# | using SaveFormat.Pdf in Aspose.Cells | pre-conversion file existence check | ensure output folder before saving PDF

using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    // The sample loads an XLSX workbook, adds custom document properties (ProjectId and Version), verifies the input file and creates the output directory if needed, then saves the workbook as a PDF using Aspose.Cells' SaveFormat.Pdf, with handling for password‑protected files and general errors.
    class Program
    {
        static void Main()
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (no password supplied)
                Workbook workbook = new Workbook(inputPath);

                // Add custom document properties
                workbook.CustomDocumentProperties.Add("ProjectId", "PRJ-2023-001");
                workbook.CustomDocumentProperties.Add("Version", 3);

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files or other Cells-specific errors
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is password protected. Provide a valid password to load it.");
                }
                else
                {
                    Console.WriteLine($"Aspose.Cells error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General error handling
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
