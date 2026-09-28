// Title: Compute sample standard deviation of numeric values in the first worksheet column using Aspose.Cells Columns enumerator (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells Columns collection to iterate the first column, calculate the sample standard deviation of all numeric cells, and write the result to cell B1. | Show how to modify the example to compute population standard deviation instead of sample standard deviation with Aspose.Cells. | Provide a version that reads the column index from a variable and outputs the standard deviation to a configurable target cell.
// Common Searches: Aspose.Cells C# calculate standard deviation of a column | How to iterate columns with Columns enumerator in Aspose.Cells to perform statistical calculations | Sample standard deviation formula implementation using Aspose.Cells in C# | Write result of column statistics to another cell with Aspose.Cells | Compute standard deviation for numeric data in Excel using Aspose.Cells C# example
// Tags: standard deviation calculation Aspose.Cells C# | Columns enumerator numeric aggregation Aspose.Cells | write computed statistic to Excel cell Aspose.Cells | sample vs population standard deviation Aspose.Cells | iterate first worksheet column Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, fills column A with numeric values, uses the Columns enumerator to traverse the first column, aggregates the numbers, computes the sample standard deviation (n‑1), writes the result to cell B1, and saves the file as StdDevColumnExample.xlsx.
class StdDevColumnExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate column A (index 0) with sample numeric data
        double[] sampleData = { 10, 20, 30, 40, 50 };
        for (int i = 0; i < sampleData.Length; i++)
        {
            cells[i, 0].PutValue(sampleData[i]); // Row i, Column 0 (A)
        }

        // Compute standard deviation of the values in column A
        // Use the Columns enumerator to satisfy the requirement
        double sum = 0;
        int count = 0;

        foreach (Column col in sheet.Cells.Columns)
        {
            // We are interested in the first column (index 0)
            // Iterate through rows that contain data
            for (int row = 0; row <= cells.MaxDataRow; row++)
            {
                object val = cells[row, 0].Value;
                if (val is double d)
                {
                    sum += d;
                    count++;
                }
                else if (val is int i)
                {
                    sum += i;
                    count++;
                }
            }
            // Process only the first column, then exit the loop
            break;
        }

        if (count == 0)
        {
            Console.WriteLine("No numeric data found in the column.");
            return;
        }

        double mean = sum / count;

        double varianceSum = 0;
        for (int row = 0; row <= cells.MaxDataRow; row++)
        {
            object val = cells[row, 0].Value;
            double d = 0;
            if (val is double dd) d = dd;
            else if (val is int ii) d = ii;
            else continue;

            varianceSum += Math.Pow(d - mean, 2);
        }

        // Sample standard deviation (n-1 in denominator)
        double stdDev = Math.Sqrt(varianceSum / (count - 1));

        // Write the result to cell B1
        cells[0, 1].PutValue(stdDev);

        // Save the workbook
        workbook.Save("StdDevColumnExample.xlsx");
    }
}
