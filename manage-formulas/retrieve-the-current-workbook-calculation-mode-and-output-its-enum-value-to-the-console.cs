using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the current calculation mode
        CalcModeType mode = workbook.Settings.FormulaSettings.CalculationMode;

        // Output the enum value
        Console.WriteLine("Current Calculation Mode: " + mode);
    }
}