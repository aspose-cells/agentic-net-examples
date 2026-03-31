using System;
using Aspose.Cells;

namespace AsposeCellsOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourcePath = "LargeWorkbook.xlsx";

            // -----------------------------------------------------------------
            // 1. Create LoadOptions and disable formula parsing on open.
            //    This prevents Aspose.Cells from parsing every formula string
            //    while the workbook is being loaded, which significantly speeds
            //    up the loading phase for large files.
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingFormulaOnOpen = false;   // Skip formula parsing during load

            // -----------------------------------------------------------------
            // 2. Load the workbook using the constructor that accepts LoadOptions.
            // -----------------------------------------------------------------
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // -----------------------------------------------------------------
            // 3. (Optional) Enable calculation chain if the workbook will be
            //    recalculated many times after only a small part of the data
            //    changes. This can improve subsequent CalculateFormula calls.
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // -----------------------------------------------------------------
            // 4. Perform any required calculations now.
            //    Since formulas were not parsed on load, we need to calculate them
            //    explicitly before accessing their values.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            // -----------------------------------------------------------------
            // 5. Example: read a calculated cell value.
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Value of cell B2 after calculation: " + sheet.Cells["B2"].Value);

            // -----------------------------------------------------------------
            // 6. Save the workbook if needed.
            // -----------------------------------------------------------------
            string outputPath = "OptimizedResult.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook processed and saved to: " + outputPath);
        }
    }
}