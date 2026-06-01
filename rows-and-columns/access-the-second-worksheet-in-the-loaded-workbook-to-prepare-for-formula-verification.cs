using System;
using Aspose.Cells;

namespace AsposeCellsFormulaVerification
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the second worksheet (index is zero‑based)
            Worksheet secondWorksheet = workbook.Worksheets[1];

            // Example: ensure formulas are calculated before verification
            workbook.CalculateFormula();

            // At this point you can perform any verification on the second worksheet,
            // e.g., read a cell value, check a formula, etc.
            Cell sampleCell = secondWorksheet.Cells["A1"];
            Console.WriteLine($"Cell A1 formula: {sampleCell.Formula}");
            Console.WriteLine($"Cell A1 calculated value: {sampleCell.Value}");

            // Save the workbook if any changes were made (optional)
            workbook.Save("output.xlsx");
        }
    }
}