// Title: Calculate a running total column in a CSV file with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# compute cumulative sum while reading CSV | Add a total column to a CSV using Aspose.Cells workbook | Load CSV, calculate running total, and save with TxtSaveOptions in .NET | C# example for cumulative column calculation in CSV with Aspose.Cells | Preserve header row when adding running totals to CSV using Aspose.Cells
// Tags: running total calculation Aspose.Cells CSV | cumulative total column Aspose.Cells .NET | load CSV workbook Aspose.Cells | TxtSaveOptions CSV output Aspose.Cells | header detection numeric column Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads a CSV file with Aspose.Cells, detects an optional header, iterates the first numeric column to compute a running total, writes each cumulative sum to a new second column, ensures the output directory exists, and saves the updated data back to CSV using TxtSaveOptions.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.csv";
        string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the CSV file
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);
            Workbook workbook = new Workbook(inputPath, loadOptions);
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Index of the column that contains the numeric values (0‑based)
            int valueColumnIndex = 0;
            // Index of the column where the running total will be written
            int totalColumnIndex = 1;

            double runningTotal = 0;

            // Determine if the first row is a header (non‑numeric)
            int startRow = 0;
            if (!double.TryParse(cells[0, valueColumnIndex].StringValue, out _))
            {
                startRow = 1; // Skip header row
            }

            // Iterate through all rows that contain data
            int lastRow = cells.MaxDataRow;
            for (int row = startRow; row <= lastRow; row++)
            {
                // Parse the numeric value; treat parsing failures as 0
                double value;
                if (!double.TryParse(cells[row, valueColumnIndex].StringValue, out value))
                {
                    value = 0;
                }

                // Update running total and write it to the new column
                runningTotal += value;
                cells[row, totalColumnIndex].PutValue(runningTotal);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook back to CSV, preserving the new column
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.CSV);
            workbook.Save(outputPath, saveOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
