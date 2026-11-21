// Title: Add a subtotal row for column D (rows 2‑100) in an Excel file with Aspose.Cells for .NET (C#)
// Common Searches: how to calculate a subtotal for column D rows 2 to 100 with Aspose.Cells in C# | Aspose.Cells C# write subtotal label and sum below a data range | C# sum numeric values in Excel column D and insert subtotal row using Aspose.Cells | using Aspose.Cells to add a total row after row 100 in an Excel worksheet | programmatically add a subtotal row in Excel via Aspose.Cells .NET API
// Tags: Aspose.Cells calculate column subtotal | C# write subtotal row in Excel worksheet | Excel column D aggregation with Aspose.Cells | insert total label below data range Aspose.Cells | Aspose.Cells sum numeric cells in range

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, iterates rows 2‑100 of column D to sum numeric values, writes the total and a "Subtotal" label in the row immediately after the range, and saves the file, handling any errors during processing.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Load existing workbook or create a new one if the file is missing
                Workbook workbook = File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range: rows 2‑100 (zero‑based indices 1‑99)
                int startRow = 1;   // Row 2
                int endRow   = 99;  // Row 100

                // Column D (zero‑based index 3)
                int targetColumn = 3;

                // Calculate sum for column D within the specified range
                double sum = 0.0;
                for (int row = startRow; row <= endRow; row++)
                {
                    try
                    {
                        var cell = worksheet.Cells[row, targetColumn];
                        if (cell.Value != null && double.TryParse(cell.Value.ToString(), out double val))
                        {
                            sum += val;
                        }
                    }
                    catch (Exception exRow)
                    {
                        Console.WriteLine($"Error processing row {row + 1}: {exRow.Message}");
                    }
                }

                // Write the subtotal below the range
                int subtotalRow = endRow + 1;
                worksheet.Cells[subtotalRow, targetColumn].PutValue(sum);
                worksheet.Cells[subtotalRow, targetColumn - 1].PutValue("Subtotal");

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
