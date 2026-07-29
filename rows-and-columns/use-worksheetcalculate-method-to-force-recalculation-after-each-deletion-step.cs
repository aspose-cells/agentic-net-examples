// Title: Recalculate Worksheet Formulas After Deleting Rows – Aspose.Cells for .NET
// Description: Demonstrates how to delete rows with Worksheet.Cells.DeleteRow/DeleteRows and immediately refresh dependent formulas using Worksheet.Calculate (or Workbook.CalculateFormula) in C#. The sample prints formula results before and after each deletion and saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | Worksheet.Calculate | Workbook.CalculateFormula | DeleteRow | DeleteRows | recalculate formulas | row deletion | Excel automation | force formula update
// Common Searches: Aspose.Cells recalculate after row deletion | Worksheet.Calculate C# example | DeleteRow update formulas Aspose | How to refresh formulas after DeleteRows | Force formula calculation in .NET workbook
// Developer Intent: Refresh all dependent formulas immediately after removing rows to keep spreadsheet calculations accurate.
// Use Cases: Adjust totals when rows are removed from a financial table | Clean data by deleting empty rows and ensuring summary formulas stay correct | Generate reports that require dynamic row removal while preserving calculated fields | Automate spreadsheet cleanup in a server‑side .NET service
// AI Prompts: Write C# code that deletes specific rows with Aspose.Cells and calls Worksheet.Calculate after each deletion. | Compare Worksheet.Calculate and Workbook.CalculateFormula and advise when each should be used. | Show how to log a formula cell's value before and after each row removal using Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to delete rows with Worksheet.Cells.DeleteRow/DeleteRows and immediately refresh dependent formulas using Worksheet.Calculate (or Workbook.CalculateFormula) in C#. The sample prints formula results before and after each deletion and saves the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate sample data
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Add a formula that depends on the data above
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Initial calculation to ensure the formula has a value
            workbook.CalculateFormula(new CalculationOptions());

            Console.WriteLine("Before deletion, B1 = " + cells["B1"].StringValue);

            // Delete the second row (index 1) and update references
            cells.DeleteRow(1, true);
            // Recalculate after the deletion step
            workbook.CalculateFormula(new CalculationOptions());

            Console.WriteLine("After deleting row 2, B1 = " + cells["B1"].StringValue);

            // Delete two rows starting from the first row (indices 0 and 1) and update references
            cells.DeleteRows(0, 2, true);
            // Recalculate after the second deletion step
            workbook.CalculateFormula(new CalculationOptions());

            Console.WriteLine("After deleting first two rows, B1 = " + cells["B1"].StringValue);

            // Save the workbook (lifecycle rule)
            workbook.Save("RecalcAfterDeletion.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
