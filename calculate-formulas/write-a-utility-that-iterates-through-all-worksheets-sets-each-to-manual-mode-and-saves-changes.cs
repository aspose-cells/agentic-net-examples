using System;
using Aspose.Cells;

namespace AsposeCellsUtility
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths (adjust as needed)
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets (demonstration purpose)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // No per‑worksheet calculation mode exists; the mode is workbook‑wide.
                // The loop is kept to satisfy the requirement of iterating worksheets.
                // Any worksheet‑specific logic could be placed here.
            }

            // Set the calculation mode to Manual for the entire workbook
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // Save the modified workbook
            workbook.Save(outputPath);
        }
    }
}