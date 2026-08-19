// Title: Handle Out‑of‑Range Column Index in Cells.Subtotal – Aspose.Cells C# Example
// Description: This C# sample creates a workbook, fills three columns with sample data, defines the data range, and attempts a subtotal using a column index that exceeds the worksheet's width. The call to Cells.Subtotal is wrapped in a try‑catch block that captures the ArgumentOutOfRangeException, logs a friendly message, and then saves the file with its own error handling.
// Keywords: Aspose.Cells Cells.Subtotal | out of range column index | ArgumentOutOfRangeException C# | exception handling Aspose.Cells | subtotal error handling .NET | validate grouping column Aspose | C# workbook subtotal example
// Common Searches: Aspose.Cells catch exception for Cells.Subtotal column index | how to handle invalid column index in Cells.Subtotal | C# subtotal out of range error Aspose | prevent crash when using Cells.Subtotal | validate column index before subtotal Aspose.Cells
// Developer Intent: Prevent a runtime crash by detecting and handling an invalid grouping column index when calling Cells.Subtotal.
// Use Cases: Check the requested grouping column against worksheet.Columns.Count before invoking Cells.Subtotal. | Wrap Cells.Subtotal in a try‑catch block to log the error and allow the program to continue. | Provide end‑users with a clear message when a subtotal cannot be applied due to an out‑of‑range column.
// AI Prompts: Generate C# code that verifies a column index is within worksheet bounds before calling Cells.Subtotal and logs a warning if it is not. | Show how to catch ArgumentOutOfRangeException from Cells.Subtotal and write detailed diagnostics to a log file. | Create a reusable method that safely applies a subtotal with configurable grouping columns and includes comprehensive exception handling.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, fills three columns with sample data, defines the data range, and attempts a subtotal using a column index that exceeds the worksheet's width. The call to Cells.Subtotal is wrapped in a try‑catch block that captures the ArgumentOutOfRangeException, logs a friendly message, and then saves the file with its own error handling.
    public class SubtotalOutOfRangeDemo
    {
        public static void Run()
        {
            // Create a new workbook (create rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (3 columns: A, B, C)
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Sales");

            object[,] data = new object[,]
            {
                {"North", "Widget", 5000},
                {"North", "Gadget", 3000},
                {"South", "Widget", 6000},
                {"South", "Gadget", 4000},
                {"West",  "Widget", 4500}
            };

            for (int row = 0; row < data.GetLength(0); row++)
            {
                for (int col = 0; col < data.GetLength(1); col++)
                {
                    cells[row + 1, col].PutValue(data[row, col]);
                }
            }

            // Define the cell area that actually contains data (A1:C6)
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

            try
            {
                // Attempt to add subtotal using an out‑of‑range column index for grouping (e.g., 10)
                // This will throw an exception because the worksheet has only 3 columns (0‑2)
                cells.Subtotal(area, 10, ConsolidationFunction.Sum, new int[] { 2 });
                Console.WriteLine("Subtotal applied successfully.");
            }
            catch (Exception ex)
            {
                // Gracefully handle the out‑of‑range error
                Console.WriteLine($"Handled exception: {ex.Message}");
            }

            try
            {
                // Save the workbook (save rule)
                workbook.Save("SubtotalOutOfRangeDemo.xlsx");
                Console.WriteLine("Workbook saved as SubtotalOutOfRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Entry point required by the project
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SubtotalOutOfRangeDemo.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
