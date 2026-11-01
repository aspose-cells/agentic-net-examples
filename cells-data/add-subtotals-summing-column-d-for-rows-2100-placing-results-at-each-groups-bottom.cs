// Title: C# Aspose.Cells example: calculate group subtotals for column D (rows 2‑100) and write them below the worksheet data
// Common Searches: how to calculate group subtotals for a column using Aspose.Cells in C# | Aspose.Cells C# sum values in column D grouped by column A for rows 2 to 100 | insert subtotal rows after data range with Aspose.Cells .NET | C# code to aggregate Excel rows 2‑100 and write totals below using Aspose.Cells | group by column A and total column D in an existing workbook with Aspose.Cells
// Tags: group subtotal calculation Aspose.Cells | sum column D by column A Aspose.Cells | write subtotals below data range C# | aggregate rows 2-100 Excel worksheet Aspose.Cells | dictionary based grouping and sum C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads a workbook, iterates rows 2‑100, groups rows by the value in column A, sums the numeric values from column D for each group, writes a "Subtotals" label and each group's total below the original data range, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "{InputFilePath}";
            string outputPath = "{OutputFilePath}";

            try
            {
                // Ensure the input file exists before loading
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Get the first worksheet
                var sheet = workbook.Worksheets[0];

                // Define the range: rows 2‑100 (zero‑based indices 1‑99)
                int startRow = 1;   // Row 2
                int endRow = 99;    // Row 100

                // Column indices (zero‑based)
                int groupByColumn = 0;   // Column A
                int subtotalColumn = 3; // Column D

                // Dictionary to hold subtotal per group
                var groupSums = new Dictionary<string, double>();

                // Iterate through the defined range and accumulate sums
                for (int row = startRow; row <= endRow; row++)
                {
                    // Get group key from column A
                    var groupCell = sheet.Cells[row, groupByColumn];
                    string groupKey = groupCell?.StringValue ?? string.Empty;

                    // Get numeric value from column D
                    var valueCell = sheet.Cells[row, subtotalColumn];
                    double cellValue = 0;
                    if (valueCell != null && double.TryParse(valueCell.Value?.ToString(), out double parsed))
                    {
                        cellValue = parsed;
                    }

                    // Accumulate sum for the group
                    if (groupSums.ContainsKey(groupKey))
                        groupSums[groupKey] += cellValue;
                    else
                        groupSums[groupKey] = cellValue;
                }

                // Write subtotals below the original data range
                int writeRow = endRow + 2; // Leave one empty row
                sheet.Cells[writeRow - 1, groupByColumn].PutValue("Subtotals");

                foreach (var kvp in groupSums)
                {
                    sheet.Cells[writeRow, groupByColumn].PutValue(kvp.Key);
                    sheet.Cells[writeRow, subtotalColumn].PutValue(kvp.Value);
                    writeRow++;
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
