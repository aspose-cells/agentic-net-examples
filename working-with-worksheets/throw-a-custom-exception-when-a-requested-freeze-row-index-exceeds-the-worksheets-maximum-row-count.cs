// Title: Throw a custom FreezeRowException in C# when the requested freeze row index exceeds the worksheet’s maximum row count using Aspose.Cells
// Common Searches: C# Aspose.Cells validate freeze row index before calling FreezePanes | how to create a custom exception for invalid freeze pane row in Aspose.Cells | check worksheet maximum row count with Aspose.Cells before applying freeze panes | Aspose.Cells throw exception when freeze row index is out of range in .NET | example of row index validation for FreezePanes in Aspose.Cells C#
// Tags: custom FreezeRowException Aspose.Cells | freeze panes row index validation C# | worksheet row count check Aspose.Cells | error handling for FreezePanes Aspose.Cells | Aspose.Cells maximum row limit handling

using System;
using System.IO;
using Aspose.Cells;

namespace FreezeRowExample
{
    // Custom exception for invalid freeze row index
    // The example defines a FreezeRowException, loads a workbook, retrieves the first worksheet, compares a requested zero‑based freeze row index with the worksheet's total row count (Cells.Rows.Count), throws the custom exception if the index exceeds the maximum, applies FreezePanes when valid, and saves the modified workbook.
    public class FreezeRowException : Exception
    {
        public FreezeRowException(string message) : base(message) { }
    }

    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Example: requested freeze row index (zero‑based)
                int requestedFreezeRow = 500000; // adjust as needed

                // Validate the requested index against the worksheet's maximum row count
                // Cells.Rows.Count returns the total number of rows currently in the sheet
                int maxRowCount = sheet.Cells.Rows.Count;
                if (requestedFreezeRow >= maxRowCount)
                {
                    throw new FreezeRowException(
                        $"Requested freeze row index {requestedFreezeRow} exceeds the worksheet's maximum row count of {maxRowCount}.");
                }

                // Apply freeze panes if the index is valid
                // Use the overload with four parameters to ensure compatibility
                sheet.FreezePanes(requestedFreezeRow, 0, requestedFreezeRow, 0);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (FreezeRowException fre)
            {
                Console.WriteLine($"Freeze row error: {fre.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}
