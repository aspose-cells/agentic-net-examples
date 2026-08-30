// Title: Refresh formulas after splitting a column with TextToColumns in Aspose.Cells for .NET (C#)
// AI Prompts: Split a comma‑delimited column into separate columns using TxtLoadOptions, then call Workbook.CalculateFormula to update any formulas that reference the new columns. | After invoking Worksheet.TextToColumns, trigger a full workbook recalculation with Worksheet.Calculate (or Workbook.CalculateFormula) to refresh dependent cells in C#. | Show how to recalculate a formula that sums the split columns (e.g., =B1+C1) after performing a TextToColumns operation with Aspose.Cells.
// Common Searches: Aspose.Cells C# recalculate formulas after using TextToColumns | How to update dependent cells after splitting a column with TxtLoadOptions in .NET | Workbook.CalculateFormula vs Worksheet.Calculate after column split Aspose.Cells
// Tags: TextToColumns formula recalculation Aspose.Cells | C# split column with TxtLoadOptions | Aspose.Cells Workbook.CalculateFormula example | update dependent cells after column split .NET | Aspose.Cells TextToColumns delimiter comma

using System;
using Aspose.Cells;

// The example creates a workbook, inserts comma‑separated values, uses TxtLoadOptions with TextToColumns to split column A into columns B and C, sets a formula in D1 that adds the split cells, recalculates all formulas via Workbook.CalculateFormula, outputs the result, and saves the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Sample data: combined first name and amount in column A
            cells["A1"].PutValue("John,100");
            cells["A2"].PutValue("Jane,200");

            // Formula that will reference the split columns (B and C) after TextToColumns
            cells["D1"].Formula = "=B1+C1";

            // Configure TextToColumns to split on comma
            TxtLoadOptions options = new TxtLoadOptions
            {
                Separator = ','
            };

            // Perform the split on the first two rows of column A (column index 0, row index 0, 2 rows)
            worksheet.Cells.TextToColumns(0, 0, 2, options);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated result
            Console.WriteLine("D1 value after recalculation: " + cells["D1"].StringValue);

            // Save the workbook (optional)
            workbook.Save("TextToColumnsCalculate.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
