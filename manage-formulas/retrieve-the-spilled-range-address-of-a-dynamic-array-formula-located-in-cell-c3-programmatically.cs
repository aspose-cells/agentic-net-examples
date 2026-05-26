using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsDynamicArraySpillDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook (lifecycle: create)
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set a dynamic array formula in cell C3
                // The formula will spill into neighboring cells when calculated
                Cell targetCell = cells["C3"];
                string formula = "=SEQUENCE(3,2)"; // Example dynamic array formula
                targetCell.SetDynamicArrayFormula(formula, new FormulaParseOptions(), calculateValue: true);

                // Calculate all formulas so the dynamic array spills
                wb.CalculateFormula();

                // Retrieve the spilled range of the dynamic array formula
                // GetArrayRange returns a CellArea describing the spill area
                CellArea spillArea = targetCell.GetArrayRange();

                // Convert the CellArea to a readable address string
                int rowCount = spillArea.EndRow - spillArea.StartRow + 1;
                int colCount = spillArea.EndColumn - spillArea.StartColumn + 1;
                AsposeRange spillRange = cells.CreateRange(spillArea.StartRow, spillArea.StartColumn, rowCount, colCount);
                string spillAddress = spillRange.Address;

                // Output the spilled range address
                Console.WriteLine($"Spilled range for dynamic array formula in C3: {spillAddress}");

                // Save the workbook (lifecycle: save)
                string outputPath = "DynamicArraySpillDemo.xlsx";
                wb.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}