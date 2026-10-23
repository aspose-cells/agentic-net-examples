// Title: Calculate the standard deviation of numeric values in a worksheet column using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells calculate standard deviation for a column of numbers | How to iterate over a column with MaxDataRow in Aspose.Cells and compute statistics | Compute variance and standard deviation in Excel using Aspose.Cells API | Sample code for standard deviation calculation on an Excel column with Aspose.Cells .NET | Retrieve numeric cells from a column and get standard deviation using Aspose.Cells
// Tags: standard deviation Aspose.Cells C# implementation | column iteration using MaxDataRow Aspose.Cells | numeric cell aggregation Aspose.Cells | statistical analysis workbook Aspose.Cells | export standard deviation result to XLSX Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a new workbook, fills column A with numeric data, uses MaxDataRow to enumerate the populated rows, aggregates sum and sum of squares for numeric cells, computes variance and standard deviation, prints the result, and saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (create rule)
            var workbook = new Workbook();

            // Get the first worksheet
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Sample numeric data placed in column A (index 0)
            double[] sampleValues = { 12.5, 15.0, 9.8, 20.3, 14.1 };
            for (int i = 0; i < sampleValues.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(sampleValues[i]);
            }

            // Calculate standard deviation for column A
            double sum = 0;
            double sumSq = 0;
            int count = 0;

            // Iterate through rows that contain data
            int maxRow = sheet.Cells.MaxDataRow;
            for (int row = 0; row <= maxRow; row++)
            {
                var cell = sheet.Cells[row, 0];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    double val = cell.DoubleValue;
                    sum += val;
                    sumSq += val * val;
                    count++;
                }
            }

            double stdDev = 0;
            if (count > 1)
            {
                double variance = (sumSq - (sum * sum) / count) / (count - 1);
                stdDev = Math.Sqrt(variance);
            }

            Console.WriteLine($"Standard Deviation of column A: {stdDev}");

            // Save the workbook (save rule)
            string outputPath = "StandardDeviationResult.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
