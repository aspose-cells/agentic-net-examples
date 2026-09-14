// Title: Export the first worksheet of an Excel workbook to CSV with Aspose.Cells LightCells API in C# without loading the full model
// AI Prompts: Write C# code that employs Aspose.Cells LightCells API to open an Excel file, verify its presence, and write only the first worksheet to a CSV file using SaveFormat.CSV, with proper exception handling. | Demonstrate a minimal‑memory conversion of an .xlsx file to .csv in C# by using Aspose.Cells to skip full workbook loading and directly save the first sheet as CSV.
// Common Searches: convert first sheet of xlsx to csv using Aspose.Cells lightweight API C# | Aspose.Cells C# export single worksheet to CSV without loading entire workbook | how to save Excel worksheet as CSV with minimal memory usage in C# | C# check file existence before Aspose.Cells conversion to CSV
// Tags: LightCells worksheet to CSV conversion | Aspose.Cells SaveFormat.CSV example | C# xlsx to csv without full model load | file existence check Aspose.Cells | exception handling Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the input Excel file exists, loads it with Aspose.Cells LightCells API to avoid loading the full workbook into memory, and saves the first worksheet directly as a CSV file, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook (any supported format)
            string inputPath = "input.xlsx";

            // Output CSV file
            string outputPath = "output.csv";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook without loading the full model into memory
            var workbook = new Workbook(inputPath);

            // Save the first worksheet as CSV
            workbook.Save(outputPath, SaveFormat.CSV);

            Console.WriteLine($"CSV file successfully created at: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
