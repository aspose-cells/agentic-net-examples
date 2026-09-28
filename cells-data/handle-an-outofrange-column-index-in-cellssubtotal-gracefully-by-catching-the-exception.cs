// Title: Handle out-of-range column index when using Cells.Subtotal in Aspose.Cells for .NET with graceful exception handling
// AI Prompts: Write C# code that calls Cells.Subtotal inside a try‑catch block and logs the exception if the group‑by column index exceeds the worksheet bounds. | Show how to validate a column index before invoking Cells.Subtotal and fall back to a safe default grouping column in Aspose.Cells. | Create a reusable C# method that adds a subtotal with built‑in error handling for invalid group‑by column indices using Aspose.Cells.
// Common Searches: Aspose.Cells C# Cells.Subtotal throws exception for column index greater than worksheet columns | how to catch out of range column error when adding subtotal with Aspose.Cells | validate group by column index before calling Cells.Subtotal in .NET | graceful error handling for Cells.Subtotal method in Aspose.Cells example
// Tags: Cells.Subtotal exception handling C# | out-of-range column index Aspose.Cells | validate group-by column Aspose.Cells | subtotal operation error handling .NET | Excel subtotal with Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsSubtotalExample
{
    // The example creates a workbook, fills it with sample data, defines a range, then attempts to add a subtotal using Cells.Subtotal with an intentionally invalid group‑by column index. The code catches the resulting exception, logs the error, and still saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data (3 columns: A, B, C)
            cells["A1"].PutValue("Region");
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

            for (int i = 0; i < data.GetLength(0); i++)
            {
                cells[i + 1, 0].PutValue(data[i, 0]); // Column A
                cells[i + 1, 1].PutValue(data[i, 1]); // Column B
                cells[i + 1, 2].PutValue(data[i, 2]); // Column C
            }

            // Define the range that contains the data (A1:C6)
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

            try
            {
                // Intentionally use an out‑of‑range column index for grouping (e.g., 10)
                // The worksheet only has columns 0‑2, so this will throw an exception.
                int outOfRangeGroupBy = 10;

                // Attempt to add subtotals; the exception will be caught below.
                cells.Subtotal(area, outOfRangeGroupBy, ConsolidationFunction.Sum, new int[] { 2 });
                Console.WriteLine("Subtotal added successfully.");
            }
            catch (Exception ex)
            {
                // Handle the out‑of‑range error gracefully
                Console.WriteLine($"Failed to add subtotal: {ex.Message}");
            }

            // Save the workbook (output will be created even if subtotal failed)
            workbook.Save("SubtotalWithGracefulErrorHandling.xlsx");
        }
    }
}
