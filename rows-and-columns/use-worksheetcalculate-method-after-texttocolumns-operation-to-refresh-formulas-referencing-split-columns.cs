// Title: Refresh formulas after TextToColumns split with Worksheet.Calculate in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to split a delimited column using TextToColumns, then invoke Worksheet.Calculate (or Workbook.CalculateFormula) to update any formulas that reference the newly created cells, and finally save the workbook.
// Keywords: Aspose.Cells | TextToColumns | Worksheet.Calculate | Workbook.CalculateFormula | recalculate formulas | C# | .NET | split column | calculation options | update dependent cells
// Common Searches: How to recalculate formulas after TextToColumns in Aspose.Cells C# | Worksheet.Calculate after splitting columns Aspose.Cells | Refresh dependent cells after TextToColumns operation .NET | Aspose.Cells recalc formulas after column split
// Developer Intent: Update worksheet formulas so they reflect data created by a TextToColumns operation.
// Use Cases: Split a CSV field into separate columns and automatically adjust sum or average formulas that reference the new columns. | Import a delimited text file, separate its fields with TextToColumns, then recalculate totals, averages, or custom calculations. | Parse address components into distinct columns and refresh distance or cost formulas that depend on those components.
// AI Prompts: Generate C# code that uses Aspose.Cells to split a column with TextToColumns and then calls Worksheet.Calculate to refresh all formulas. | Show how to configure CalculationOptions and invoke Workbook.CalculateFormula after a TextToColumns operation in Aspose.Cells for .NET. | Explain the steps required to ensure formulas referencing newly split columns are updated automatically in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Demonstrates how to split a delimited column using TextToColumns, then invoke Worksheet.Calculate (or Workbook.CalculateFormula) to update any formulas that reference the newly created cells, and finally save the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with comma‑separated values
            sheet.Cells["A1"].PutValue("John,Doe,30");
            sheet.Cells["A2"].PutValue("Jane,Smith,28");

            // Add a formula that will reference the split columns (age will end up in column C)
            sheet.Cells["D1"].Formula = "=C1+C2";

            // Configure TextToColumns to split on commas
            TxtLoadOptions options = new TxtLoadOptions
            {
                Separator = ','
            };

            // Perform the split on the first two rows of column A
            sheet.Cells.TextToColumns(0, 0, 2, options);

            // Recalculate formulas so they reflect the newly split data
            CalculationOptions calcOptions = new CalculationOptions();
            workbook.CalculateFormula(calcOptions);

            // Save the result
            workbook.Save("TextToColumns_Calc.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
