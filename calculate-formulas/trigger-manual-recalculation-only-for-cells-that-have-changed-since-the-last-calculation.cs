// Title: How to trigger manual recalculation for only changed cells with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate only dirty formulas after updating a cell | set workbook to manual calculation mode and recalculate changed cells Aspose.Cells | C# Aspose.Cells selective formula evaluation using CalcModeType.Manual | trigger partial workbook calculation in Aspose.Cells after modifying input data | manual formula recalculation for specific cells in Aspose.Cells .NET
// Tags: Aspose.Cells manual calc mode | dirty formula recalculation C# | CalcModeType.Manual usage | selective formula evaluation Aspose.Cells | partial workbook calculation .NET

using System;
using Aspose.Cells;

namespace ManualRecalculationDemo
{
    // Shows how to set a workbook to manual calculation mode, populate it with data and formulas, perform an initial full calculation, modify a single input cell, and then invoke CalculateFormula to recalculate only the dirty dependent formulas before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the output workbook
            const string outputPath = "ManualRecalcDemo.xlsx";

            // ------------------------------------------------------------
            // 1. Create a new workbook and set calculation mode to Manual
            // ------------------------------------------------------------
            Workbook workbook = new Workbook();
            workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.Manual;

            // ------------------------------------------------------------
            // 2. Populate the worksheet with data and formulas
            // ------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Input data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Formulas that depend on the above data
            cells["B1"].Formula = "=A1*2";
            cells["B2"].Formula = "=A2*2";
            cells["B3"].Formula = "=A3*2";

            // ------------------------------------------------------------
            // 3. Perform the first calculation (all formulas are dirty)
            // ------------------------------------------------------------
            Console.WriteLine("=== Initial calculation (all formulas) ===");
            workbook.CalculateFormula(); // All formulas are evaluated because they are dirty
            PrintValues(cells, new[] { "A1", "A2", "A3", "B1", "B2", "B3" });

            // ------------------------------------------------------------
            // 4. Change only one input cell (A2) – this marks the dependent
            //    formula cell (B2) as dirty, but leaves B1 and B3 clean.
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Changing A2 from 20 to 50 ===");
            cells["A2"].PutValue(50);

            // ------------------------------------------------------------
            // 5. Trigger manual recalculation – only dirty cells (B2) are
            //    recalculated because the workbook is in Manual mode.
            // ------------------------------------------------------------
            Console.WriteLine("\n=== Manual recalculation (only changed cells) ===");
            workbook.CalculateFormula(); // Recalculates only cells whose formulas are dirty
            PrintValues(cells, new[] { "A1", "A2", "A3", "B1", "B2", "B3" });

            // ------------------------------------------------------------
            // 6. Save the workbook to verify the results
            // ------------------------------------------------------------
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }

        /// <param name="cells">The Cells collection.</param>
        /// <param name="addresses">Array of cell addresses to print.</param>
        private static void PrintValues(Cells cells, string[] addresses)
        {
            foreach (string addr in addresses)
            {
                Cell cell = cells[addr];
                string value = cell.IsFormula ? $"{cell.Value} (Formula: {cell.Formula})" : cell.Value?.ToString() ?? "null";
                Console.WriteLine($"{addr}: {value}");
            }
        }
    }
}
