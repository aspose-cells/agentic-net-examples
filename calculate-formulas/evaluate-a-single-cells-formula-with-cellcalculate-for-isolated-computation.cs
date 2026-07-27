using System;
using Aspose.Cells;

namespace AsposeCellsSingleCellCalculation
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Access a specific cell (A1) and assign a formula to it
            Cell cell = worksheet.Cells["A1"];
            cell.Formula = "=SUM(1, 2, 3)";

            // Perform isolated calculation for this cell using default calculation options
            cell.Calculate(new CalculationOptions());

            // Output the calculated result
            Console.WriteLine($"Calculated value of {cell.Name}: {cell.Value}");

            // Optionally save the workbook to verify the result in Excel
            workbook.Save("SingleCellCalculation.xlsx");
        }
    }
}