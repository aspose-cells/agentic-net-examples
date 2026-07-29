// Title: Refresh formulas after TextToColumns using Worksheet.Calculate in Aspose.Cells for .NET
// Description: Shows how to split a comma‑separated column with TextToColumns, then recalculate dependent formulas using Worksheet.Calculate (and Workbook.CalculateFormula) in Aspose.Cells for .NET before saving the workbook.
// Keywords: Aspose.Cells | TextToColumns | Worksheet.Calculate | Workbook.CalculateFormula | C# formula refresh | split column | recalculate formulas | Aspose.Cells .NET | Excel automation | update dependent cells
// Common Searches: Aspose.Cells recalculate formulas after TextToColumns | Worksheet.Calculate vs Workbook.CalculateFormula | C# split column and refresh formulas Aspose.Cells | How to update formulas after TextToColumns in .NET | Refresh dependent cells after splitting CSV column Aspose.Cells
// Developer Intent: Update formulas that reference cells created by a TextToColumns operation so they reflect the new data layout.
// Use Cases: Split CSV data into separate columns and ensure existing formulas adjust automatically. | Add new formulas that use the split columns and obtain calculated results immediately. | Perform a full workbook recalculation after column splitting before persisting the file. | Use Worksheet.Calculate for sheet‑level recalculation when only one worksheet is affected.
// AI Prompts: Show C# code that uses Aspose.Cells TextToColumns then calls Worksheet.Calculate to refresh formulas. | Explain when to prefer Worksheet.Calculate over Workbook.CalculateFormula after splitting columns. | Provide a step‑by‑step guide to split a column and update dependent formulas in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to split a comma‑separated column with TextToColumns, then recalculate dependent formulas using Worksheet.Calculate (and Workbook.CalculateFormula) in Aspose.Cells for .NET before saving the workbook.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with comma‑separated values
            sheet.Cells["A1"].PutValue("John,Doe,30");
            sheet.Cells["A2"].PutValue("Jane,Smith,28");

            // Configure TextToColumns to split on commas
            TxtLoadOptions options = new TxtLoadOptions
            {
                Separator = ','
            };

            // Split the data in column A into separate columns (A, B, C)
            sheet.Cells.TextToColumns(0, 0, 2, options);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Example formula that uses the split columns
            sheet.Cells["D1"].Formula = "=B1 & \" \" & C1";

            // Recalculate to obtain the result of the new formula
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("TextToColumns_Calculated.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
