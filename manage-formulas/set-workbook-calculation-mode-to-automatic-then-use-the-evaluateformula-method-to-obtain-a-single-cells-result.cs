using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue(10);
        cells["B1"].PutValue(20);

        // Set the workbook calculation mode to Automatic
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Evaluate a formula directly (without storing it in a cell)
        object evaluatedResult = worksheet.CalculateFormula("=A1+B1");

        // Display the evaluated result
        Console.WriteLine("Evaluated result of =A1+B1: " + evaluatedResult);

        // Save the workbook (optional)
        workbook.Save("CalculatedResult.xlsx");
    }
}