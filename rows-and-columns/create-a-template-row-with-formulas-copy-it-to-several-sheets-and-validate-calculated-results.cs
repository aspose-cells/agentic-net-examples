// Title: Copy a template row with formulas to multiple worksheets, recalculate and verify results using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a template row containing a constant and a formula, copies the row to several worksheets, triggers workbook recalculation, and asserts that each sheet's formula result matches the expected value. | Show how to duplicate a row with formulas from a source sheet to target sheets and programmatically validate the calculated values using Aspose.Cells in a .NET application.
// Common Searches: Aspose.Cells copy row with formula to other worksheets C# example | How to recalculate formulas after copying rows in Aspose.Cells .NET | Validate formula results across multiple sheets using Aspose.Cells C# | Duplicate template row with formulas to several sheets in Aspose.Cells | Programmatically check calculated cell values in Aspose.Cells workbook
// Tags: row formula duplication Aspose.Cells C# | workbook formula recalculation Aspose.Cells | cell value verification Aspose.Cells | template worksheet row copy Aspose.Cells | multiple sheet formula propagation Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a workbook, adds a template sheet with a constant in A1 and a formula =A1*2 in B1, copies that row to three additional worksheets, recalculates all formulas, validates that each sheet's B1 equals A1 multiplied by two, and saves the file as TemplateCopyValidation.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (initially contains a default sheet)
            Workbook workbook = new Workbook();

            // Remove the default sheet to avoid name conflicts
            if (workbook.Worksheets.Count > 0)
                workbook.Worksheets.RemoveAt(0);

            // -------------------------------------------------
            // 1. Create a template sheet with a row that contains formulas
            // -------------------------------------------------
            int templateIndex = workbook.Worksheets.Add();
            Worksheet templateSheet = workbook.Worksheets[templateIndex];
            templateSheet.Name = "Template";

            // Set a constant value in column A (e.g., 10)
            Cell a1 = templateSheet.Cells["A1"];
            a1.PutValue(10);

            // Set a formula in column B that references column A (e.g., =A1*2)
            Cell b1 = templateSheet.Cells["B1"];
            b1.Formula = "=A1*2";

            // -------------------------------------------------
            // 2. Create additional sheets and copy the template row to them
            // -------------------------------------------------
            string[] targetSheetNames = { "Sheet1", "Sheet2", "Sheet3" };
            foreach (string sheetName in targetSheetNames)
            {
                // Add a new worksheet and set its name
                int wsIndex = workbook.Worksheets.Add();
                Worksheet ws = workbook.Worksheets[wsIndex];
                ws.Name = sheetName;

                // Copy the values and formulas from the template row (row 0) to the new sheet
                for (int col = 0; col <= templateSheet.Cells.MaxColumn; col++)
                {
                    Cell srcCell = templateSheet.Cells[0, col];
                    Cell destCell = ws.Cells[0, col];

                    if (!srcCell.IsFormula)
                    {
                        destCell.PutValue(srcCell.Value);
                    }
                    else
                    {
                        destCell.Formula = srcCell.Formula;
                    }
                }
            }

            // -------------------------------------------------
            // 3. Recalculate all formulas in the workbook
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 4. Validate the calculated results in each sheet
            // -------------------------------------------------
            foreach (string sheetName in targetSheetNames)
            {
                Worksheet ws = workbook.Worksheets[sheetName];
                double aValue = ws.Cells["A1"].DoubleValue;
                double bValue = ws.Cells["B1"].DoubleValue;
                double expectedB = aValue * 2;

                bool isValid = Math.Abs(bValue - expectedB) < 1e-9;

                Console.WriteLine($"{sheetName}: A1 = {aValue}, B1 = {bValue} (expected {expectedB}) -> Validation {(isValid ? "Passed" : "Failed")}");
            }

            // -------------------------------------------------
            // 5. Save the workbook (optional, demonstrates lifecycle usage)
            // -------------------------------------------------
            workbook.Save("TemplateCopyValidation.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
