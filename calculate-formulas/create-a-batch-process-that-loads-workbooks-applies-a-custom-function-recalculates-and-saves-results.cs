// Title: Batch recalculate formulas in multiple .xlsx workbooks with Aspose.Cells for .NET (C#)
// Common Searches: C# batch recalculate formulas in a folder of Excel .xlsx files using Aspose.Cells | How to load, calculate all formulas, and save multiple workbooks with Aspose.Cells .NET | Automate Excel formula evaluation for many files in a directory using Aspose.Cells | Aspose.Cells example for processing multiple workbooks and forcing full calculation
// Tags: batch calculate formulas Aspose.Cells | process multiple xlsx workbooks C# | load workbook recalculate save Aspose.Cells | directory based Excel processing .NET | automated formula evaluation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsBatchProcess
{
    // The sample scans an input folder for .xlsx files, loads each workbook with Aspose.Cells, forces a full formula calculation, and saves the updated workbook to an output folder, handling missing files and runtime exceptions.
    class Program
    {
        static void Main(string[] args)
        {
            // Folder containing source workbooks.
            string inputFolder = @"C:\InputWorkbooks";
            // Folder where processed workbooks will be saved.
            string outputFolder = @"C:\OutputWorkbooks";

            // Ensure output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Get all Excel files in the input folder (top level only).
            string[] files = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

            foreach (string filePath in files)
            {
                // Verify the source file exists before attempting to load.
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Source file not found: {filePath}");
                    continue;
                }

                try
                {
                    // Load the workbook.
                    Workbook workbook = new Workbook(filePath);

                    // Force full calculation of all formulas.
                    workbook.CalculateFormula();

                    // Build the output file path.
                    string fileName = Path.GetFileName(filePath);
                    string outputPath = Path.Combine(outputFolder, fileName);

                    // Save the updated workbook.
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed and saved: {outputPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
