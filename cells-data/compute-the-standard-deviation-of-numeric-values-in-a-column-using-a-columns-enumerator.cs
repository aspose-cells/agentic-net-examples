using System;
using Aspose.Cells;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();               // create rule
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate column A (index 0) with sample numeric data
        double[] sampleData = { 10, 20, 30, 40, 50 };
        for (int i = 0; i < sampleData.Length; i++)
        {
            cells[i, 0].PutValue(sampleData[i]); // A1, A2, ...
        }

        // Variables for standard deviation calculation
        double sum = 0;
        double sumSq = 0;
        int count = 0;

        // Use the Columns enumerator to locate the target column (index 0)
        foreach (Column column in cells.Columns)
        {
            // Column.Index gives the zero‑based column index
            if (column.Index == 0) // target column A
            {
                // Iterate through all rows that contain data in this column
                for (int row = 0; row <= cells.MaxDataRow; row++)
                {
                    object val = cells[row, column.Index].Value;
                    if (val is double d)
                    {
                        sum += d;
                        sumSq += d * d;
                        count++;
                    }
                }
                break; // column found and processed
            }
        }

        // Compute sample standard deviation
        double mean = sum / count;
        double variance = (sumSq - (sum * sum) / count) / (count - 1); // unbiased estimator
        double stdDev = Math.Sqrt(variance);

        // Write the result to cell B1
        cells[0, 1].PutValue(stdDev); // column B, row 1

        // Save the workbook
        workbook.Save("StdDevColumnA.xlsx");               // save rule
    }
}