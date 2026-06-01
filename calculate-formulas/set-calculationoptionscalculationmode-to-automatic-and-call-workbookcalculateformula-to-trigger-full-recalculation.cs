using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet and its cells
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Add sample values and formulas
        cells["A1"].PutValue(5);
        cells["A2"].PutValue(10);
        cells["B1"].Formula = "=A1*2";
        cells["B2"].Formula = "=A2*2";
        cells["C1"].Formula = "=SUM(B1:B2)";

        // Set the calculation mode to Automatic
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Automatic;

        // Trigger a full recalculation of all formulas
        workbook.CalculateFormula();

        // Display the result of the calculated cell
        Console.WriteLine("C1 result: " + cells["C1"].Value);

        // Save the workbook to a file
        workbook.Save("Recalculated.xlsx");
    }
}