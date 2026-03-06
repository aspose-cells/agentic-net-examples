using System;
using Aspose.Cells;

namespace AsposeCellsFormulaRecalculationDemo
{
    class Program
    {
        static void Main()
        {
            // -----------------------------------------------------------------
            // 1. Load an existing XLSX workbook.
            //    ParsingFormulaOnOpen is set to false so formulas are not
            //    calculated automatically during load – this mimics the need
            //    to calculate them manually later.
            // -----------------------------------------------------------------
            LoadOptions loadOptions = new LoadOptions
            {
                ParsingFormulaOnOpen = false
            };

            // Replace with the actual path of your workbook.
            string inputPath = "InputWorkbook.xlsx";
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // -----------------------------------------------------------------
            // 2. Access the first worksheet and some cells for demonstration.
            // -----------------------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Assume the workbook contains formulas in B1 and C1 that depend on A1.
            // Display the raw values before any calculation (usually 0 or empty).
            Console.WriteLine("=== Before any calculation ===");
            Console.WriteLine($"A1 value : {cells["A1"].Value}");
            Console.WriteLine($"B1 value : {cells["B1"].Value}");
            Console.WriteLine($"C1 value : {cells["C1"].Value}");

            // -----------------------------------------------------------------
            // 3. First calculation pass – evaluate all formulas in the workbook.
            // -----------------------------------------------------------------
            workbook.CalculateFormula();

            Console.WriteLine("\n=== After first calculation ===");
            Console.WriteLine($"A1 value : {cells["A1"].Value}");
            Console.WriteLine($"B1 value : {cells["B1"].Value}");
            Console.WriteLine($"C1 value : {cells["C1"].Value}");

            // -----------------------------------------------------------------
            // 4. Modify a source cell (A1) and recalculate formulas again.
            //    This demonstrates repeated calculation without re‑loading the file.
            // -----------------------------------------------------------------
            cells["A1"].PutValue(100); // Change the source value.

            // Recalculate – the same CalculateFormula method can be called any number of times.
            workbook.CalculateFormula();

            Console.WriteLine("\n=== After modifying A1 and recalculating ===");
            Console.WriteLine($"A1 value : {cells["A1"].Value}");
            Console.WriteLine($"B1 value : {cells["B1"].Value}");
            Console.WriteLine($"C1 value : {cells["C1"].Value}");

            // -----------------------------------------------------------------
            // 5. (Optional) Enable calculation chain for better performance
            //    when many formulas exist and only a small part changes.
            // -----------------------------------------------------------------
            workbook.Settings.FormulaSettings.EnableCalculationChain = true;

            // Another change to demonstrate repeated calculation with the chain enabled.
            cells["A1"].PutValue(250);
            workbook.CalculateFormula();

            Console.WriteLine("\n=== After second modification with calculation chain enabled ===");
            Console.WriteLine($"A1 value : {cells["A1"].Value}");
            Console.WriteLine($"B1 value : {cells["B1"].Value}");
            Console.WriteLine($"C1 value : {cells["C1"].Value}");

            // -----------------------------------------------------------------
            // 6. Save the workbook to a new file.
            // -----------------------------------------------------------------
            string outputPath = "OutputWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
        }
    }
}