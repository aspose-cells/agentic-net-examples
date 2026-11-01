// Title: Read only numeric cells from an XLSX worksheet into a high‑performance List<double> using Aspose.Cells for .NET (C#)
// Common Searches: how to extract only numeric values from an Excel file with Aspose.Cells in C# | Aspose.Cells iterate used range and collect numeric data into List<double> | C# read numeric cells from .xlsx workbook using Aspose.Cells without loading formulas | store Excel numeric values in memory efficiently with Aspose.Cells .NET | filter non‑numeric cells while reading Excel worksheet using Aspose.Cells C#
// Tags: Aspose.Cells numeric cell extraction C# | XLSX used range iteration Aspose.Cells | store numeric Excel data in List<double> | filter numeric cell types C# Aspose.Cells | high‑performance in‑memory numeric collection .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an XLSX workbook with Aspose.Cells, accesses the first worksheet, determines its used range, and iterates each cell. It checks each cell's Value for numeric types (double, int, float, long) and adds the values to a List<double>, providing a fast in‑memory collection of numeric data while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File not found – {inputPath}");
            return;
        }

        try
        {
            // Load the workbook with explicit format
            var loadOptions = new LoadOptions(LoadFormat.Xlsx);
            var workbook = new Workbook(inputPath, loadOptions);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];
            var cells = worksheet.Cells;

            // Collection for numeric values
            var numericValues = new List<double>();

            // Determine the used range
            int maxRow = cells.MaxDataRow;
            int maxCol = cells.MaxDataColumn;

            // Iterate through each cell in the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    var cellValue = cells[row, col].Value;

                    // Store only numeric values
                    switch (cellValue)
                    {
                        case double d:
                            numericValues.Add(d);
                            break;
                        case int i:
                            numericValues.Add(i);
                            break;
                        case float f:
                            numericValues.Add(f);
                            break;
                        case long l:
                            numericValues.Add(l);
                            break;
                        // Add other numeric types if needed
                    }
                }
            }

            // Output the result
            Console.WriteLine($"Total numeric values read: {numericValues.Count}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
