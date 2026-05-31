using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some data that will be used by the formula
        cells["A1"].PutValue(5);
        cells["A2"].PutValue(10);
        cells["A3"].PutValue(15);

        // Set a formula in cell B1 that sums the range A1:A3
        Cell formulaCell = cells["B1"];
        formulaCell.Formula = "=SUM(A1:A3)";

        // At this point the formula has not been evaluated yet
        Console.WriteLine("Before calculation, B1 value: " + formulaCell.Value);

        // Evaluate only this single cell using Cell.Calculate
        formulaCell.Calculate(new CalculationOptions());

        // After calculation the cell now holds the computed result
        Console.WriteLine("After calculation, B1 value: " + formulaCell.Value);

        // Verify that other cells remain unchanged
        Console.WriteLine("A1 value: " + cells["A1"].Value);
        Console.WriteLine("A2 value: " + cells["A2"].Value);
        Console.WriteLine("A3 value: " + cells["A3"].Value);

        // Optional: save the workbook to see the result in Excel
        workbook.Save("SingleCellCalculation.xlsx");
    }
}