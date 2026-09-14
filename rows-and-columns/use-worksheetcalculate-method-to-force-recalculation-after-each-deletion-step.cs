// Title: Force worksheet recalculation after each row deletion using Aspose.Cells for .NET
// AI Prompts: Write C# code that deletes a specific row from an Aspose.Cells worksheet and immediately calls Worksheet.CalculateFormula to refresh dependent formulas. | Show how to iterate over multiple row deletions, invoking Workbook.CalculateFormula after each removal to keep SUM and other formulas accurate. | Demonstrate saving the workbook after performing row deletions with intermediate recalculations using Aspose.Cells in a .NET console application.
// Common Searches: How to update SUM formula after deleting rows with Aspose.Cells C# | Aspose.Cells force calculation after DeleteRow operation | C# example of Worksheet.CalculateFormula called after each row removal | Recalculate workbook formulas when rows are removed in Aspose.Cells .NET | Delete rows and keep formulas correct using Aspose.Cells API
// Tags: Worksheet.CalculateFormula after DeleteRow Aspose.Cells | row deletion with intermediate formula recalculation C# | update SUM formula on row removal Aspose.Cells | force workbook calculation after modifying cells .NET | Aspose.Cells DeleteRow example with formula refresh

using System;
using Aspose.Cells;

namespace AsposeCellsRowDeletionDemo
{
    // The sample creates a workbook, fills column A with values, adds a SUM formula in B1, then deletes rows 3, 1, and the original row 5. After each DeleteRow call, Workbook.CalculateFormula is executed to update the formula result, and the final workbook is saved as RowDeletionWithRecalc.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data in column A (rows 1‑5)
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 0].PutValue((i + 1) * 10); // A1=10, A2=20, ..., A5=50
                }

                // Add a formula in B1 that sums the values in column A
                cells[0, 1].Formula = "=SUM(A1:A5)";

                // Initial calculation to evaluate the formula
                workbook.CalculateFormula();
                Console.WriteLine($"Initial B1 value (SUM A1:A5): {cells[0, 1].Value}");

                // Delete the third row (index 2) and recalculate
                sheet.Cells.DeleteRow(2);
                workbook.CalculateFormula();
                Console.WriteLine($"After deleting row 3, B1 value: {cells[0, 1].Value}");

                // Delete the first row (index 0) and recalculate
                sheet.Cells.DeleteRow(0);
                workbook.CalculateFormula();
                Console.WriteLine($"After deleting row 1, B1 value: {cells[0, 1].Value}");

                // Delete the last remaining data row (current index 2) and recalculate
                sheet.Cells.DeleteRow(2);
                workbook.CalculateFormula();
                Console.WriteLine($"After deleting row 5 (original), B1 value: {cells[0, 1].Value}");

                // Save the workbook to verify the final state
                workbook.Save("RowDeletionWithRecalc.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
