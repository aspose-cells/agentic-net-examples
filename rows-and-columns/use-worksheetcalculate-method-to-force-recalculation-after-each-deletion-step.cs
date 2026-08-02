// Title: C# – Recalculate formulas after each row deletion using Aspose.Cells Worksheet.Calculate
// Description: Demonstrates how to delete rows from a worksheet and trigger immediate formula recalculation with Worksheet.Calculate (or Workbook.CalculateFormula) in Aspose.Cells for .NET, printing the updated SUM after each removal and saving the final file.
// Keywords: Aspose.Cells C# | Worksheet.Calculate | Workbook.CalculateFormula | recalculate after DeleteRow | dynamic row removal Excel | force formula update .NET | sum formula after row delete | Excel automation Aspose | incremental calculation | performance optimization
// Common Searches: Aspose.Cells recalculate after deleting rows | Worksheet.Calculate vs Workbook.CalculateFormula | C# delete rows and update formulas | force Excel formula refresh after row removal | Aspose.Cells incremental calculation example
// Developer Intent: Keep dependent formulas accurate by forcing a calculation after each row deletion.
// Use Cases: Maintain a running total while pruning data rows. | Ensure financial aggregates stay correct after dynamic row cleanup. | Generate a final workbook that reflects the latest calculations after iterative deletions.
// AI Prompts: Generate C# code that deletes rows from an Aspose.Cells worksheet and calls Worksheet.Calculate after each deletion. | Explain when to prefer Worksheet.Calculate over Workbook.CalculateFormula in Aspose.Cells. | Show how to log formula results after every DeleteRow operation using Aspose.Cells. | Compare performance of batch recalculation versus per‑row recalculation in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to delete rows from a worksheet and trigger immediate formula recalculation with Worksheet.Calculate (or Workbook.CalculateFormula) in Aspose.Cells for .NET, printing the updated SUM after each removal and saving the final file.
class WorksheetCalculateAfterDeletion
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate column A with values 1 to 10
            for (int i = 0; i < 10; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1..A10
            }

            // Add a formula in B1 that sums the values in column A
            cells[0, 1].Formula = "=SUM(A1:A10)";

            // Initial calculation to evaluate the formula
            workbook.CalculateFormula();
            Console.WriteLine("Initial sum (B1): " + cells[0, 1].StringValue);

            // Delete rows from bottom to top, forcing recalculation after each deletion
            for (int rowIndex = 9; rowIndex >= 0; rowIndex--)
            {
                // Delete the current row
                cells.DeleteRow(rowIndex);

                // Recalculate formulas after the deletion
                workbook.CalculateFormula();

                // Output the updated sum after this deletion step
                Console.WriteLine($"After deleting row {rowIndex + 1}, sum (B1): {cells[0, 1].StringValue}");
            }

            // Save the final workbook
            workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
