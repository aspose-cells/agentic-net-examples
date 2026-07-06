using System;
using Aspose.Cells;

namespace AsposeCellsCellCalculateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (in-memory)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data that will be used by the formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Set a formula in cell B1 that depends on the above values
            Cell targetCell = cells["B1"];
            targetCell.Formula = "=SUM(A1:A3)";

            // At this point the formula has not been evaluated yet
            Console.WriteLine("Before calculation, B1 value: " + targetCell.Value); // Expected: null

            // Calculate only the target cell using Cell.Calculate
            targetCell.Calculate(new CalculationOptions());

            // After calculation the cell now holds the result of the formula
            Console.WriteLine("After calculation, B1 value: " + targetCell.Value); // Expected: 60

            // Verify that other cells are untouched (they retain their original values)
            Console.WriteLine("A1 value (unchanged): " + cells["A1"].Value);
            Console.WriteLine("A2 value (unchanged): " + cells["A2"].Value);
            Console.WriteLine("A3 value (unchanged): " + cells["A3"].Value);

            // Optionally save the workbook to see the calculated result in the file
            workbook.Save("CellCalculateResult.xlsx");
        }
    }
}