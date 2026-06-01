using System;
using Aspose.Cells;

namespace AsposeCellsStdDevExample
{
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
                cells[i, 0].PutValue(sampleData[i]); // Row i, Column 0 (A)
            }

            // Variables for standard deviation calculation
            double sum = 0;
            double sumSq = 0;
            int count = 0;

            // Enumerate through the Columns collection to locate column A (index 0)
            foreach (Column column in cells.Columns)
            {
                // Column.Index gives the zero‑based column number
                if (column.Index == 0) // Target column A
                {
                    // Determine the last row that contains data in this column
                    int lastRow = cells.MaxDataRow;

                    // Iterate over each cell in the column up to lastRow
                    for (int row = 0; row <= lastRow; row++)
                    {
                        object valObj = cells[row, column.Index].Value;
                        if (valObj != null && double.TryParse(valObj.ToString(), out double val))
                        {
                            sum += val;
                            sumSq += val * val;
                            count++;
                        }
                    }

                    // Once the target column is processed, break out of the loop
                    break;
                }
            }

            // Compute standard deviation (sample, using n‑1 denominator)
            double stdDev = 0;
            if (count > 1)
            {
                double variance = (sumSq - (sum * sum) / count) / (count - 1);
                stdDev = Math.Sqrt(variance);
            }

            // Output the result
            Console.WriteLine($"Standard Deviation of column A: {stdDev}");

            // Save the workbook (optional, just to demonstrate lifecycle compliance)
            workbook.Save("StdDevResult.xlsx");
        }
    }
}