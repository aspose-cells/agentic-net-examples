using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SubtotalOutOfRangeDemo
    {
        public static void Run()
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
                cells[i + 1, 0].PutValue(data[i, 0]);
                cells[i + 1, 1].PutValue(data[i, 1]);
                cells[i + 1, 2].PutValue(data[i, 2]);
            }

            // Define the cell area covering the data (A1:C6)
            CellArea area = CellArea.CreateCellArea(0, 0, 5, 2);

            try
            {
                // Attempt to add a subtotal using an out‑of‑range column index (e.g., column 5)
                cells.Subtotal(
                    area,
                    0,                                 // Group by first column (Region)
                    ConsolidationFunction.Sum,         // Use SUM function
                    new int[] { 5 }                    // Invalid column index for subtotal
                );

                Console.WriteLine("Subtotal added successfully (unexpected).");
            }
            catch (Exception ex)
            {
                // Handle the out‑of‑range error gracefully
                Console.WriteLine($"Handled exception: {ex.Message}");
            }

            try
            {
                // Save the workbook
                workbook.Save("SubtotalOutOfRangeDemo.xlsx");
                Console.WriteLine("Workbook saved.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
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