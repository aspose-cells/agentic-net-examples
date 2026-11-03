// Title: C# example: enumerate worksheet rows with Aspose.Cells to sum numeric values in the first column and write the result to another cell
// Common Searches: how to iterate over rows in an Excel worksheet using Aspose.Cells for .NET | C# code to calculate total of first column values with Aspose.Cells | write sum of a column back to a specific cell using Aspose.Cells API | example of summing numeric cells in the first column with Aspose.Cells | save modified workbook after updating a cell in Aspose.Cells C#
// Tags: row enumeration with Aspose.Cells C# | first column numeric aggregation Aspose.Cells | update cell B1 with computed sum Aspose.Cells | workbook load and save Aspose.Cells | exception handling for CellsException Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an Excel workbook, enumerates each row of the first worksheet, adds up numeric values from the first column, writes the calculated sum to cell B1, saves the workbook, and includes handling for Aspose.Cells‑specific and general exceptions.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (no password supplied)
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet and its cells
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                double sum = 0;

                // Iterate through rows that contain data
                foreach (Row row in cells.Rows)
                {
                    Cell cell = row[0]; // first column (index 0)

                    // Add numeric values only
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        sum += cell.DoubleValue;
                    }
                }

                // Output the result
                Console.WriteLine("Sum of first column: " + sum);

                // Write the sum back to cell B1
                cells["B1"].PutValue(sum);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files or other Aspose.Cells specific errors
                Console.WriteLine($"Aspose.Cells error: {ex.Message}");
            }
            catch (Exception ex)
            {
                // General exception handling
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
