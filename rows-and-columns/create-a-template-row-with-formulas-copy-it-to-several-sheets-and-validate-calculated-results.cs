// Title: Copy a Template Row with a Shared Formula to Multiple Sheets and Verify Results using Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook with a template worksheet, fill columns A and B, set a shared formula in column C (A+B), copy the first row to several new worksheets, recalculate all formulas, read the computed values, and save the file as "TemplateRowCopy.xlsx" using Aspose.Cells for C#.
// Keywords: Aspose.Cells copy row | shared formula Aspose.Cells | C# Aspose.Cells copy rows with formulas | calculate formulas Aspose.Cells | validate formula results .NET | duplicate template worksheet Aspose.Cells | Workbook.CalculateFormula example | copy rows across worksheets C#
// Common Searches: How to copy a row with a shared formula to other sheets in Aspose.Cells | Aspose.Cells copy rows and keep formulas intact | Validate copied formula values after Workbook.CalculateFormula | C# example for replicating a template row across multiple worksheets | Aspose.Cells copy rows between worksheets
// Developer Intent: Copy a template row that contains a shared formula to several worksheets, recalculate the workbook, and confirm that the results are correct.
// Use Cases: Create a master row that sums two columns and reuse it in monthly report sheets. | Automate the propagation of a header row with embedded calculations to new department worksheets. | Run a post‑copy validation to ensure formulas produce expected totals after calling Workbook.CalculateFormula.
// AI Prompts: Show me C# code to set a shared formula in Aspose.Cells and copy the row to multiple worksheets. | Provide an Aspose.Cells example that copies a template row with formulas, recalculates the workbook, and prints the results. | Explain how to use Workbook.CalculateFormula and read cell values to assert correctness after copying rows in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook with a template worksheet, fill columns A and B, set a shared formula in column C (A+B), copy the first row to several new worksheets, recalculate all formulas, read the computed values, and save the file as "TemplateRowCopy.xlsx" using Aspose.Cells for C#.
class Program
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();

        // ---------- Prepare the template worksheet ----------
        Worksheet templateSheet = workbook.Worksheets[0];
        templateSheet.Name = "Template";

        // Fill sample data in columns A and B (rows 1-5)
        for (int i = 0; i < 5; i++)
        {
            // Column A: 1,2,3,4,5
            templateSheet.Cells[i, 0].PutValue(i + 1);
            // Column B: 10,20,30,40,50
            templateSheet.Cells[i, 1].PutValue((i + 1) * 10);
        }

        // Set a shared formula in column C to calculate A+B for the first 5 rows
        // The formula is placed in C1 and propagated to C2:C5
        templateSheet.Cells[0, 2].SetSharedFormula("=A1+B1", 5, 1);

        // ---------- Create additional worksheets and copy the template row ----------
        int additionalSheets = 3; // number of sheets to copy to
        for (int idx = 1; idx <= additionalSheets; idx++)
        {
            // Add a new worksheet
            Worksheet targetSheet = workbook.Worksheets.Add($"Sheet{idx}");

            // Copy the first row (row index 0) from the template to the target sheet
            // Parameters: source cells, source row index, destination row index, number of rows to copy
            targetSheet.Cells.CopyRows(templateSheet.Cells, 0, 0, 1);
        }

        // ---------- Calculate all formulas ----------
        workbook.CalculateFormula();

        // ---------- Validate calculated results ----------
        // Iterate through all sheets (template + copies) and print values from column C
        for (int sheetIdx = 0; sheetIdx < workbook.Worksheets.Count; sheetIdx++)
        {
            Worksheet ws = workbook.Worksheets[sheetIdx];
            Console.WriteLine($"--- {ws.Name} ---");
            for (int row = 0; row < 5; row++)
            {
                // Expected result: (A value) + (B value)
                Console.WriteLine($"C{row + 1} = {ws.Cells[row, 2].Value}");
            }
        }

        // ---------- Save the workbook ----------
        workbook.Save("TemplateRowCopy.xlsx");
    }
}
