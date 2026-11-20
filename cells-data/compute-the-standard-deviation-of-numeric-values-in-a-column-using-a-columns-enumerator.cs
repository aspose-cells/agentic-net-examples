// Title: Calculate sample standard deviation of numeric values in a specific Excel column using Aspose.Cells for .NET (C#)
// Common Searches: asp.net calculate standard deviation of an Excel column with Aspose.Cells | c# read numeric values from the first column of an Excel file and compute sample standard deviation | how to iterate through an Excel column using Aspose.Cells for statistical analysis | sample standard deviation formula implementation with Aspose.Cells worksheet in C#
// Tags: Aspose.Cells column enumeration for statistical calculations | sample standard deviation computation in C# using Aspose.Cells | extract numeric values from Excel column .NET | worksheet data analysis with Aspose.Cells | calculate variance and standard deviation in Excel via Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example verifies the presence of an input.xlsx file, loads it with Aspose.Cells, enumerates all rows of the first worksheet's first column, collects numeric cell values, and then computes the sample standard deviation (using the N‑1 denominator) before printing the result to the console.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Choose the column to process (e.g., column A -> index 0)
            int columnIndex = 0;

            // Collect numeric values from the column
            List<double> values = new List<double>();

            // Determine the last row that contains data
            int maxRow = sheet.Cells.MaxDataRow;

            for (int row = 0; row <= maxRow; row++)
            {
                Cell cell = sheet.Cells[row, columnIndex];
                if (cell == null || cell.Value == null)
                    continue;

                // Try to convert the cell value to double
                if (double.TryParse(cell.Value.ToString(), out double numeric))
                {
                    values.Add(numeric);
                }
            }

            // Ensure there are enough data points
            if (values.Count < 2)
            {
                Console.WriteLine("Not enough numeric data to compute standard deviation.");
                return;
            }

            // Compute mean
            double sum = 0;
            foreach (double v in values)
                sum += v;
            double mean = sum / values.Count;

            // Compute sample variance (denominator N-1)
            double varianceSum = 0;
            foreach (double v in values)
            {
                double diff = v - mean;
                varianceSum += diff * diff;
            }
            double variance = varianceSum / (values.Count - 1);

            // Standard deviation is the square root of variance
            double stdDev = Math.Sqrt(variance);

            // Output the result
            Console.WriteLine($"Standard Deviation of column {columnIndex + 1}: {stdDev}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
