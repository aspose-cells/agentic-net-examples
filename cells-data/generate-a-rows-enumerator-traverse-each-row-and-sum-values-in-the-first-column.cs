// Title: Sum numeric values in the first column with Aspose.Cells Row enumerator (C# .NET)
// Description: Creates a workbook, fills column A with numbers, text and nulls, obtains the RowCollection, uses its enumerator to walk each existing row, extracts the first cell, safely parses numeric values, accumulates a double total, prints the sum and saves the file.
// Keywords: Aspose.Cells row enumerator C# | iterate rows Aspose.Cells | sum first column Aspose.Cells | calculate column total .NET | handle non‑numeric cells Aspose.Cells | RowCollection enumeration | C# spreadsheet sum column
// Common Searches: Aspose.Cells iterate rows and sum column | C# sum first column using RowCollection | skip text cells when summing column Aspose.Cells | how to use GetEnumerator with Aspose.Cells rows | calculate column total in .NET spreadsheet library
// Developer Intent: The developer needs to loop through all rows of a worksheet and compute the sum of numeric entries in the first column while ignoring empty or non‑numeric cells.
// Use Cases: Generate a sales report and total the amounts listed in column A. | Validate that a column of measurements stays within a defined range. | Add a summary row that displays the aggregated total after data processing.
// AI Prompts: Show how to change the code to sum values in column B instead of column A. | Provide an example that writes the calculated sum back to cell C1. | Explain how to achieve the same result with LINQ over Aspose.Cells RowCollection.

using System;
using System.Collections;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills column A with numbers, text and nulls, obtains the RowCollection, uses its enumerator to walk each existing row, extracts the first cell, safely parses numeric values, accumulates a double total, prints the sum and saves the file.
    public class RowSumFirstColumnDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample numeric data in the first column (A)
                cells["A1"].PutValue(10);
                cells["A2"].PutValue(20);
                cells["A3"].PutValue(30);
                // Add some non‑numeric data to demonstrate safe handling
                cells["A4"].PutValue("Text");
                cells["A5"].PutValue(null);

                // Get the RowCollection from the worksheet
                RowCollection rows = cells.Rows;

                // Obtain an enumerator that iterates through all existing rows
                IEnumerator enumerator = rows.GetEnumerator();

                double sum = 0.0;

                // Traverse each row and sum the values in the first column (index 0)
                while (enumerator.MoveNext())
                {
                    Row row = (Row)enumerator.Current;

                    // Retrieve the first cell in the row; may be null if the cell does not exist
                    Cell cell = row.GetCellOrNull(0);
                    if (cell != null && cell.Value != null)
                    {
                        // Try to parse the cell value as a double; ignore if parsing fails
                        if (double.TryParse(cell.Value.ToString(), out double value))
                        {
                            sum += value;
                        }
                    }
                }

                Console.WriteLine($"Sum of values in the first column: {sum}");

                // Save the workbook (optional, just to demonstrate the save lifecycle)
                string outputPath = "RowSumFirstColumnDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RowSumFirstColumnDemo.Run();
        }
    }
}
