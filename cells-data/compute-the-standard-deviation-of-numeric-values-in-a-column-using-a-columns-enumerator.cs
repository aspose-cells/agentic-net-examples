// Title: Compute column standard deviation with Aspose.Cells Columns enumerator in C#
// Description: Creates a workbook, populates column A with numeric data, uses the Columns collection to locate a target column, extracts numeric values up to MaxDataRow, calculates the sample standard deviation, prints the result, and saves the workbook.
// Keywords: Aspose.Cells | C# | standard deviation | column statistics | Columns enumerator | numeric values | sample std dev | Excel data analysis
// Common Searches: Aspose.Cells calculate standard deviation column C# | How to read column values with Columns enumerator Aspose.Cells | Sample standard deviation using Aspose.Cells .NET | Extract numeric data from Excel column Aspose.Cells
// Developer Intent: Calculate the sample standard deviation of numbers in a specific worksheet column using Aspose.Cells.
// Use Cases: Summarize variability of sales figures stored in a column before charting. | Assess dispersion of sensor readings recorded in an Excel column. | Automate reports that compute test‑score variability and store the result in the workbook.
// AI Prompts: Show how to modify the code to compute population standard deviation instead of sample. | Provide an example that writes the calculated standard deviation into a designated cell. | Explain how to ignore non‑numeric cells and handle mixed data types when calculating standard deviation with Aspose.Cells.

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsStdDevExample
{
    // Creates a workbook, populates column A with numeric data, uses the Columns collection to locate a target column, extracts numeric values up to MaxDataRow, calculates the sample standard deviation, prints the result, and saves the workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate column A (index 0) with sample numeric data
            double[] sampleData = { 10, 20, 30, 40, 50 };
            for (int i = 0; i < sampleData.Length; i++)
            {
                cells[i, 0].PutValue(sampleData[i]); // Row i, Column 0
            }

            // Target column index for which we want to compute the standard deviation
            int targetColumnIndex = 0;

            // Collect numeric values from the target column using a Columns enumerator
            List<double> values = new List<double>();
            foreach (Column column in worksheet.Cells.Columns)
            {
                // The Column object does not expose its index directly, so we compare with the target index
                // by checking the first cell in the column.
                // This works because we know the column we are interested in (index 0).
                if (column.Index == targetColumnIndex) // Column.Index is available in Aspose.Cells
                {
                    // Determine the last row that contains data in the worksheet
                    int maxRow = worksheet.Cells.MaxDataRow;
                    for (int row = 0; row <= maxRow; row++)
                    {
                        object cellValue = cells[row, column.Index].Value;
                        if (cellValue is double d)
                        {
                            values.Add(d);
                        }
                        else if (cellValue is int i)
                        {
                            values.Add(i);
                        }
                        // Non‑numeric values are ignored
                    }
                    break; // Target column processed; exit the enumerator
                }
            }

            // Compute standard deviation (sample standard deviation)
            double stdDev = double.NaN;
            if (values.Count > 1)
            {
                double mean = values.Average();
                double variance = values.Sum(v => Math.Pow(v - mean, 2)) / (values.Count - 1);
                stdDev = Math.Sqrt(variance);
            }

            // Output the result
            Console.WriteLine($"Standard Deviation of column {targetColumnIndex}: {stdDev}");

            // Save the workbook (lifecycle rule: use create, then save)
            workbook.Save("StdDevResult.xlsx");
        }
    }
}
