// Title: How to enumerate rows with Aspose.Cells for .NET and sum numeric values in the first column
// AI Prompts: Write C# code that uses Aspose.Cells to obtain a row enumerator from a worksheet and accumulate the numeric values found in column A. | Show a safe way to retrieve the first cell of each row, verify it contains a number, and add it to a total using Aspose.Cells. | Demonstrate printing the calculated sum to the console and saving the workbook to an Excel file after processing.
// Common Searches: aspocells iterate rows and calculate sum of column A in C# | c# Aspose.Cells get rows enumerator and sum first column values | how to sum numeric cells in first column using Aspose.Cells Rows collection | enumerate worksheet rows with Aspose.Cells and compute column total | Aspose.Cells .NET example summing values in column A
// Tags: Aspose.Cells row enumeration .NET | sum first column Aspose.Cells | iterate worksheet rows C# | calculate column total Excel Aspose.Cells | Rows.GetEnumerator usage Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills cells A1‑A5 with numbers 1 through 5, obtains an IEnumerator for the worksheet's Rows collection, iterates each row, safely retrieves the first cell, checks for a numeric value, adds it to a running total, prints the sum, and saves the workbook as "SumFirstColumn.xlsx".
    public class SumFirstColumnDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample numeric data in the first column (A)
                for (int i = 0; i < 5; i++)
                {
                    // Put values 1,2,3,4,5 in cells A1..A5
                    worksheet.Cells[i, 0].PutValue(i + 1);
                }

                // Initialize sum accumulator
                double sum = 0;

                // Get an enumerator for the rows collection
                IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator();

                // Traverse each row
                while (rowEnumerator.MoveNext())
                {
                    // Cast the current element to Row
                    Row row = (Row)rowEnumerator.Current;

                    // Get the first cell in the row (column index 0)
                    Cell cell = row.GetCellOrNull(0);

                    // If the cell exists and contains a numeric value, add it to the sum
                    if (cell != null && cell.Value != null)
                    {
                        if (double.TryParse(cell.Value.ToString(), out double value))
                        {
                            sum += value;
                        }
                    }
                }

                // Output the result
                Console.WriteLine($"Sum of values in the first column: {sum}");

                // Save the workbook (demonstrates the required save lifecycle)
                string outputPath = "SumFirstColumn.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SumFirstColumnDemo.Run();
        }
    }
}
