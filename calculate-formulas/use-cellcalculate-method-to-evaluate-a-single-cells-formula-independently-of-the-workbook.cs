// Title: Evaluate a single cell formula with Cell.Calculate and CalculationOptions in Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# calculate only one cell formula without full workbook recompute | How to use Cell.Calculate with CalculationOptions to evaluate a SUM formula in .NET | Partial formula evaluation example using Aspose.Cells C# | Retrieve result of a specific Excel cell formula with Aspose.Cells API | C# code to compute an Excel cell formula independently using Aspose.Cells
// Tags: Cell.Calculate partial workbook evaluation | Aspose.Cells single cell formula calculation | CalculationOptions evaluate specific Excel cell | C# evaluate SUM formula Aspose.Cells | independent cell formula computation .NET

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaEvaluation
{
    // The example creates a workbook, assigns a SUM formula to cell A1, populates cells B1‑B3 with values, and then uses Cell.Calculate with a CalculationOptions object to evaluate only that cell. The computed result is read from the cell's Value property without triggering a full workbook calculation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set a formula in cell A1 (e.g., sum of B1:B3)
                Cell formulaCell = sheet.Cells["A1"];
                formulaCell.Formula = "=SUM(B1:B3)";

                // Populate the referenced cells with sample data
                sheet.Cells["B1"].PutValue(10);
                sheet.Cells["B2"].PutValue(20);
                sheet.Cells["B3"].PutValue(30);

                // Calculate only the formula cell using CalculationOptions
                CalculationOptions calcOptions = new CalculationOptions();
                formulaCell.Calculate(calcOptions);

                // Retrieve the calculated result
                object result = formulaCell.Value;

                Console.WriteLine($"The result of the formula in A1 is: {result}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
