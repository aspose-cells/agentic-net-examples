// Title: Add a sum row below an Excel ListObject table using Aspose.Cells C# PutValue with row offset
// Common Searches: how to calculate and write a column total at the end of an Aspose.Cells ListObject in C# | Aspose.Cells C# insert sum row after table using worksheet.Cells.PutValue | C# sum values of a specific column in an Excel table and place total below the table with Aspose.Cells | using Aspose.Cells ListObject to get data range and add an aggregate row in C# | write aggregate value to a cell with offset from a ListObject data range Aspose.Cells
// Tags: Aspose.Cells PutValue with offset | C# add aggregate row to Excel table | ListObject data range sum in C# | Insert total cell below table Aspose | Worksheet cell write numeric total

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the first ListObject, computes the sum of numeric values in a chosen column, and writes the total one row below the table using the worksheet's Cells.PutValue method, then saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table
            if (worksheet.ListObjects.Count == 0)
            {
                Console.WriteLine("No tables found in the worksheet.");
                return;
            }

            // Get the first table (ListObject)
            ListObject table = worksheet.ListObjects[0];

            // Zero‑based index of the column to sum within the table (e.g., third column)
            int sumColumnIndex = 2;

            // Determine data range (excluding header)
            int firstDataRow = table.DataRange.FirstRow;
            int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
            int firstDataColumn = table.DataRange.FirstColumn;

            // Calculate sum of the selected column
            double total = 0;
            for (int r = firstDataRow; r <= lastDataRow; r++)
            {
                object cellValue = worksheet.Cells[r, firstDataColumn + sumColumnIndex].Value;
                if (cellValue != null && double.TryParse(cellValue.ToString(), out double d))
                {
                    total += d;
                }
            }

            // Row and column where the total will be placed (one row below data)
            int totalRow = lastDataRow + 1;
            int totalColumn = firstDataColumn + sumColumnIndex;

            // Insert the total directly into the worksheet cell
            worksheet.Cells[totalRow, totalColumn].PutValue(total);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
