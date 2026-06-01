using System;
using Aspose.Cells;

namespace AsposeCellsProgressBarSave
{
    class Program
    {
        static void Main()
        {
            // Load the workbook that already contains the Progress Bar chart.
            // Replace the file name with the actual path to your source workbook.
            Workbook workbook = new Workbook("ProgressBarChart.xlsx");

            // Ensure formulas are not recalculated on save so they are preserved as‑is.
            // This is the default value, but setting it explicitly makes the intention clear.
            workbook.Settings.FormulaSettings.CalculateOnSave = false;

            // Save the workbook to XLSX format.
            // The Save method with (string, SaveFormat) follows the provided rule.
            workbook.Save("ProgressBarChart_Output.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved successfully with formulas preserved.");
        }
    }
}