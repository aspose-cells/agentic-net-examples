using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace BatchSharedFormula
{
    class Program
    {
        static void Main(string[] args)
        {
            // List of workbook file paths to process
            List<string> workbookPaths = new List<string>
            {
                "Book1.xlsx",
                "Book2.xlsx",
                "Book3.xlsx"
                // Add more paths as needed
            };

            // Define the shared formula and the range size
            string sharedFormula = "=A1*2"; // Example: double the value in column A
            int rowsToPopulate = 100;      // Number of rows the shared formula will cover
            int columnsToPopulate = 1;     // Number of columns (only column B in this case)

            foreach (string path in workbookPaths)
            {
                try
                {
                    // Verify that the input file exists
                    if (!File.Exists(path))
                    {
                        Console.WriteLine($"Input file not found: '{path}'. Skipping.");
                        continue;
                    }

                    // Load the workbook
                    Workbook workbook = new Workbook(path);

                    // Access the first worksheet
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Apply the shared formula starting at cell B1
                    cells["B1"].SetSharedFormula(sharedFormula, rowsToPopulate, columnsToPopulate);

                    // Recalculate formulas so that the new values are materialized
                    workbook.CalculateFormula();

                    // Determine a safe output path (use current directory if none is provided)
                    string fullInputPath = Path.GetFullPath(path);
                    string directory = Path.GetDirectoryName(fullInputPath);
                    if (string.IsNullOrEmpty(directory))
                    {
                        directory = Directory.GetCurrentDirectory();
                    }

                    string outputFileName = Path.GetFileNameWithoutExtension(fullInputPath) + "_SharedFormula.xlsx";
                    string outputPath = Path.Combine(directory, outputFileName);

                    // Save the modified workbook
                    workbook.Save(outputPath);

                    Console.WriteLine($"Processed '{path}' and saved as '{outputPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{path}': {ex.Message}");
                }
            }
        }
    }
}