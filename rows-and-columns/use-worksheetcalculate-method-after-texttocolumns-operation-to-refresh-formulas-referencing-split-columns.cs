// Title: Refresh formulas after TextToColumns split using Worksheet.Calculate in Aspose.Cells for .NET
// Description: Demonstrates how to split comma‑separated values with TextToColumns, then invoke Worksheet.Calculate (or Workbook.CalculateFormula) to update formulas that reference the newly created columns, and finally save the workbook.
// Keywords: Aspose.Cells TextToColumns | Worksheet.Calculate | Workbook.CalculateFormula | C# refresh formulas | update dependent formulas after split | Aspose.Cells .NET example
// Common Searches: Aspose.Cells recalculate formulas after TextToColumns | Worksheet.Calculate after column split C# | How to refresh formulas in Aspose.Cells .NET | Update dependent cells after TextToColumns operation | C# Aspose.Cells calculate workbook after data split
// Developer Intent: Run a calculation pass so formulas that depend on columns created by TextToColumns return correct results.
// Use Cases: Import raw CSV strings, split them into separate columns, and automatically adjust summary formulas. | Process user‑entered text data, separate fields with TextToColumns, then recalculate totals or derived values. | Perform data cleansing on a worksheet and ensure all downstream calculations reflect the new layout before exporting.
// AI Prompts: Show C# code that uses TextToColumns and then calls Worksheet.Calculate to update formulas in Aspose.Cells. | Explain when to use Worksheet.Calculate versus Workbook.CalculateFormula after splitting columns. | Generate an Aspose.Cells example that recalculates only the cells affected by a TextToColumns operation.

using System;
using Aspose.Cells;

// Demonstrates how to split comma‑separated values with TextToColumns, then invoke Worksheet.Calculate (or Workbook.CalculateFormula) to update formulas that reference the newly created columns, and finally save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data: comma‑separated values in column A
            sheet.Cells["A1"].PutValue("John,Doe,30");
            sheet.Cells["A2"].PutValue("Jane,Smith,28");

            // Formula that depends on the column that will be created by TextToColumns
            // After splitting, the age will be in column C, so this formula adds 5 to it
            sheet.Cells["D1"].Formula = "=C1+5";

            // Configure TextToColumns to split on comma
            TxtLoadOptions options = new TxtLoadOptions
            {
                Separator = ','
            };

            // Perform the split on the first two rows of column A
            sheet.Cells.TextToColumns(0, 0, 2, options);

            // Refresh all formulas in the workbook after the split operation
            workbook.CalculateFormula();

            // Display the results to verify that the formula was recalculated
            Console.WriteLine("C1 (Age): " + sheet.Cells["C1"].StringValue);
            Console.WriteLine("D1 (Age + 5): " + sheet.Cells["D1"].StringValue);

            // Save the workbook
            workbook.Save("TextToColumns_With_Calculation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
