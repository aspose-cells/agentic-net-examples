using System;
using Aspose.Cells;

class SubtotalOutOfRangeDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data (A1:C5)
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Product");
        cells["C1"].PutValue("Sales");

        object[,] data = new object[,]
        {
            { "North", "Widget", 5000 },
            { "North", "Gadget", 3000 },
            { "South", "Widget", 6000 },
            { "South", "Gadget", 4000 }
        };

        for (int i = 0; i < data.GetLength(0); i++)
        {
            cells[i + 1, 0].PutValue(data[i, 0]);
            cells[i + 1, 1].PutValue(data[i, 1]);
            cells[i + 1, 2].PutValue(data[i, 2]);
        }

        // Define the cell area covering the data (A1:C5)
        CellArea area = CellArea.CreateCellArea(0, 0, 4, 2);

        // Attempt to apply subtotal with an out‑of‑range column index (groupBy = 5)
        try
        {
            // This will throw because column index 5 does not exist in the defined area
            cells.Subtotal(area, 5, ConsolidationFunction.Sum, new int[] { 2 });
            Console.WriteLine("Subtotal applied successfully.");
        }
        catch (Exception ex)
        {
            // Gracefully handle the out‑of‑range exception
            Console.WriteLine($"Handled exception: {ex.Message}");
        }

        // Save the workbook
        workbook.Save("SubtotalOutOfRangeDemo.xlsx");
    }
}