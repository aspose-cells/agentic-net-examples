using System;
using Aspose.Cells;

class EvaluateCellFormula
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a formula in cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.Formula = "=SUM(10, 20, 30)";

        // Calculate the formula for this single cell
        cell.Calculate(new CalculationOptions());

        // Display the calculated value
        Console.WriteLine("A1 calculated value: " + cell.Value);

        // Save the workbook (optional)
        workbook.Save("CalculatedResult.xlsx");
    }
}