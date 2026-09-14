// Title: Replace negative numbers with zero in the 'ProfitMargins' named range using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, locate the named range "ProfitMargins", and set any numeric cell whose value is less than 0 to 0 with Aspose.Cells in C#. | Iterate through all cells of a named range, clamp negative numeric values to zero, and save the modified workbook using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells how to set negative values to zero in a specific named range | replace values less than zero in Excel named range using Aspose.Cells .NET | iterate over cells in a named range and modify numeric values with Aspose.Cells C# | Aspose.Cells example correcting negative numbers in ProfitMargins range
// Tags: Aspose.Cells set negative cells to zero | C# named range value clamping Aspose.Cells | ProfitMargins range numeric correction Aspose.Cells | Excel workbook modify cells Aspose.Cells C# | iterate over Aspose.Cells range cells

using Aspose.Cells;
using System;
using System.IO;

// The program loads 'input.xlsx', retrieves the named range 'ProfitMargins', changes any numeric cell values below zero to zero, and saves the updated workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ProfitMargins"
            // Use the Names collection to get the named range (compatible with all versions)
            Name profitMargins = workbook.Worksheets.Names["ProfitMargins"];
            if (profitMargins == null)
            {
                Console.WriteLine("Named range 'ProfitMargins' not found.");
                return;
            }

            // Get the actual cell range represented by the named range
            Aspose.Cells.Range range = profitMargins.GetRange();

            // Iterate through each cell in the range
            foreach (Cell cell in range)
            {
                // Process only numeric cells
                if (cell.Type == CellValueType.IsNumeric)
                {
                    double value = cell.DoubleValue;
                    // Replace negative values with zero
                    if (value < 0)
                    {
                        cell.PutValue(0);
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
