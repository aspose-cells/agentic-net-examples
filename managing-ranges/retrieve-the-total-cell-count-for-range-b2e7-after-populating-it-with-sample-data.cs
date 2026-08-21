// Title: Aspose.Cells for .NET (C#): Count Cells in Range B2:E7 After Filling Sample Data
// Description: Creates a new Workbook, defines the range B2:E7 with Aspose.Cells, populates each cell with a position label, calculates total cells using RowCount × ColumnCount, prints the result, and optionally saves the file.
// Keywords: Aspose.Cells C# count cells in range | retrieve range cell count .NET | RowCount ColumnCount Aspose.Cells | populate Excel range with sample data | B2:E7 cell total Aspose | C# Excel range size calculation | Aspose.Cells CreateRange example | total cells in Excel block
// Common Searches: how to get number of cells in a specific range using Aspose.Cells | C# example to fill B2:E7 and count cells | Aspose.Cells calculate total cells in range | retrieve cell count after populating range in .NET | count cells in Excel range programmatically
// Developer Intent: Find out how many cells exist in the B2:E7 block after inserting sample values with Aspose.Cells.
// Use Cases: Verify that a data block contains the expected number of cells before processing. | Generate a summary report showing how many cells were filled in a dynamic range. | Determine resource allocation for batch operations based on the total populated cells.
// AI Prompts: Write a C# function using Aspose.Cells that fills any given range with sequential identifiers and returns the total cell count. | Show how to safely compute RowCount × ColumnCount for a range and handle empty or null ranges in Aspose.Cells. | Explain how to modify the example to count only non‑empty cells after the range has been populated.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a new Workbook, defines the range B2:E7 with Aspose.Cells, populates each cell with a position label, calculates total cells using RowCount × ColumnCount, prints the result, and optionally saves the file.
    public class RetrieveCellCountDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Define the target range B2:E7
                AsposeRange targetRange = cells.CreateRange("B2", "E7");

                // Populate the range with sample data
                for (int row = 0; row < targetRange.RowCount; row++)
                {
                    for (int col = 0; col < targetRange.ColumnCount; col++)
                    {
                        // Put a simple string value indicating its position
                        targetRange[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Calculate total number of cells in the range
                int totalCellCount = targetRange.RowCount * targetRange.ColumnCount;

                // Output the result
                Console.WriteLine($"Total cells in range {targetRange.Address}: {totalCellCount}");

                // Save the workbook (optional)
                workbook.Save("RetrieveCellCountDemo.xlsx");
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
            RetrieveCellCountDemo.Run();
        }
    }
}
