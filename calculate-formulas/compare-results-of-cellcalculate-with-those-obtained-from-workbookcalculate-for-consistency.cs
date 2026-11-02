// Title: Verify that Cell.Calculate and Workbook.CalculateFormula produce identical results in Aspose.Cells for .NET
// Common Searches: Aspose.Cells compare result of Cell.Calculate with Workbook.CalculateFormula in C# | how to ensure formula consistency when using Cell.Calculate versus Workbook.CalculateFormula | C# example checking equality of single‑cell and full‑workbook calculations with Aspose.Cells | using CalculationOptions to validate Excel formula evaluation consistency in Aspose.Cells | detect differences between cell-level and workbook-level calculation in Aspose.Cells
// Tags: cell.calculate result verification | workbook.calculateformula consistency check | aspose.cells formula evaluation comparison | c# calculationoptions for excel formulas | single cell vs full workbook calculation asp.net

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file or creates a new workbook with a simple formula, calculates cell A1 using Cell.Calculate with default CalculationOptions, then recalculates the entire workbook with Workbook.CalculateFormula, compares the two results for equality, prints the values and consistency status, and finally saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        try
        {
            Workbook workbook;

            // Load existing workbook if it exists; otherwise create a new one with a sample formula.
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].Formula = "2+3";
            }

            Worksheet sheet = workbook.Worksheets[0];
            Cell testCell = sheet.Cells["A1"];

            // Calculate the single cell using default calculation options.
            CalculationOptions calcOptions = new CalculationOptions();
            testCell.Calculate(calcOptions);
            object resultFromCellCalc = testCell.Value;

            // Recalculate the entire workbook.
            workbook.CalculateFormula();
            object resultFromWorkbookCalc = testCell.Value;

            // Compare results for consistency.
            bool isConsistent = Equals(resultFromCellCalc, resultFromWorkbookCalc);

            Console.WriteLine($"Result after Cell.Calculate:      {resultFromCellCalc}");
            Console.WriteLine($"Result after Workbook.Calculate: {resultFromWorkbookCalc}");
            Console.WriteLine($"Consistency check: {(isConsistent ? "Consistent" : "Inconsistent")}");

            // Save the workbook.
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
