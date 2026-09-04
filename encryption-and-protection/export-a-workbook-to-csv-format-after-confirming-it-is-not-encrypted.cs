// Title: Export a non‑encrypted Excel workbook to CSV using Aspose.Cells for .NET
// AI Prompts: Create a C# console program that loads an .xlsx file with Aspose.Cells, verifies the workbook is not encrypted, and saves it as a CSV file, providing clear messages for success or encryption errors. | Write code that catches CellsException when an Excel workbook is password‑protected, logs a user‑friendly warning, and only proceeds with the CSV export if the workbook is unencrypted.
// Common Searches: Aspose.Cells C# export XLSX to CSV while handling password‑protected workbooks | How to check if an Excel workbook is encrypted before saving as CSV with Aspose.Cells | C# catch CellsException for encrypted workbook when converting to CSV using Aspose.Cells | Save non‑encrypted Excel file as CSV with Aspose.Cells and display custom error messages
// Tags: export workbook to csv Aspose.Cells | detect encrypted workbook Aspose.Cells | handle CellsException encryption C# | save xlsx as csv Aspose.Cells | check workbook encryption before export

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing XLSX file, verifies the file exists, attempts to save it as a CSV using Aspose.Cells, and catches CellsException to identify encrypted workbooks, outputting appropriate console messages for success, encryption issues, or other errors.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Save the workbook as CSV
                workbook.Save(outputPath, SaveFormat.Csv);
                Console.WriteLine($"Workbook successfully exported to \"{outputPath}\".");
            }
            catch (CellsException ex)
            {
                // Handle encrypted workbook scenario based on exception message
                if (ex.Message != null && ex.Message.IndexOf("encrypted", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is encrypted and cannot be exported.");
                }
                else
                {
                    Console.WriteLine($"Aspose.Cells error: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                // General exception handling for unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
