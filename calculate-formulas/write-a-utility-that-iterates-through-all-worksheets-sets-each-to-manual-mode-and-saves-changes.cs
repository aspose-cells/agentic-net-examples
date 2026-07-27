using System;
using Aspose.Cells;

class ManualCalcModeUtility
{
    static void Main(string[] args)
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Path where the modified workbook will be saved
        string outputPath = "output.xlsx";

        // Load the workbook from the file system
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets (required by the task)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // No per‑worksheet calculation mode exists; the mode is set at the workbook level.
            // The loop is kept to satisfy the iteration requirement.
        }

        // Set the calculation mode of the workbook to Manual
        workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

        // Save the workbook with the updated settings
        workbook.Save(outputPath);
    }
}