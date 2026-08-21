// Title: Calculate all formulas in an Aspose.Cells workbook (C#) and read updated cell values
// Description: Demonstrates how to create a workbook, insert numbers and formulas, invoke Workbook.CalculateFormula() to evaluate every formula, read the resulting .Value of each cell, save the file, reload it, recalculate if needed, and display selected results using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# calculate formulas | Workbook.CalculateFormula | retrieve calculated cell values | save workbook Aspose | load workbook Aspose | recalculate formulas | Excel formula evaluation .NET
// Common Searches: Aspose.Cells calculate all formulas C# | Get formula result after Workbook.CalculateFormula | Recalculate a loaded workbook with Aspose.Cells | Read cell value after calculation Aspose.Cells | Workbook.CalculateFormula vs Workbook.Calculate
// Developer Intent: Evaluate every formula in a workbook programmatically and obtain the computed values.
// Use Cases: Populate cells with raw data and formulas, call Workbook.CalculateFormula(), then read .Value for further processing. | Generate a report by calculating formulas in memory, printing results to the console, and exporting the workbook to XLSX. | Save a calculated workbook, later load it, recalculate to reflect changes, and retrieve specific cell results.
// AI Prompts: Write C# code using Aspose.Cells that creates a workbook, adds numbers and formulas, runs Workbook.CalculateFormula(), and prints each calculated value. | Show how to load an existing .xlsx file with Aspose.Cells, recalculate all formulas, and return the values of cells B1 and C2. | Explain when to use Workbook.CalculateFormula versus Workbook.Calculate in Aspose.Cells and the impact on performance.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaCalculationDemo
{
    // Demonstrates how to create a workbook, insert numbers and formulas, invoke Workbook.CalculateFormula() to evaluate every formula, read the resulting .Value of each cell, save the file, reload it, recalculate if needed, and display selected results using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // -------------------------------------------------
            // 1. Create a new workbook (lifecycle: create)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // Access the first worksheet and its cells
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 2. Populate some data and formulas
            // -------------------------------------------------
            // Simple values
            cells["A1"].PutValue(5);               // A1 = 5
            cells["A2"].PutValue(10);              // A2 = 10

            // Formulas that depend on the above values
            cells["B1"].Formula = "=A1*2";          // B1 = 5 * 2 = 10
            cells["B2"].Formula = "=A2+20";        // B2 = 10 + 20 = 30
            cells["C1"].Formula = "=SUM(A1:A2)";   // C1 = 5 + 10 = 15
            cells["C2"].Formula = "=B1+B2";        // C2 = 10 + 30 = 40

            // -------------------------------------------------
            // 3. Calculate all formulas in the workbook
            //    (using the rule Workbook.CalculateFormula())
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 4. Retrieve and display the updated cell values
            // -------------------------------------------------
            Console.WriteLine("After calculation:");
            Console.WriteLine($"A1 = {cells["A1"].Value}");
            Console.WriteLine($"A2 = {cells["A2"].Value}");
            Console.WriteLine($"B1 (formula '=A1*2') = {cells["B1"].Value}");
            Console.WriteLine($"B2 (formula '=A2+20') = {cells["B2"].Value}");
            Console.WriteLine($"C1 (formula '=SUM(A1:A2)') = {cells["C1"].Value}");
            Console.WriteLine($"C2 (formula '=B1+B2') = {cells["C2"].Value}");

            // -------------------------------------------------
            // 5. Save the workbook to a file (lifecycle: save)
            // -------------------------------------------------
            string outputPath = "FormulaCalculationResult.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");

            // -------------------------------------------------
            // 6. Demonstrate loading an existing workbook,
            //    recalculating, and retrieving values.
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(outputPath); // lifecycle: load
            loadedWorkbook.CalculateFormula(); // recalculate in case data changed
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cells loadedCells = loadedSheet.Cells;

            Console.WriteLine("\nValues from the loaded workbook after recalculation:");
            Console.WriteLine($"B1 = {loadedCells["B1"].Value}");
            Console.WriteLine($"C2 = {loadedCells["C2"].Value}");
        }
    }
}
