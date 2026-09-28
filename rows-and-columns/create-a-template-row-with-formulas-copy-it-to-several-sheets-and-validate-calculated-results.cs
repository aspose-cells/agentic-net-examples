// Title: Copy a template row with a Total formula to multiple worksheets, fill data, calculate and verify results using Aspose.Cells for .NET
// AI Prompts: Generate C# code that creates a workbook, adds a header and a data row where Total = Quantity*Price, copies this row to a set of worksheets, assigns Quantity and Price values, runs workbook.CalculateFormula, and asserts that the Total cell equals the expected product. | Adapt the example to accept a dynamic array of target sheet names and output a pass/fail validation message for each sheet after the formulas are calculated.
// Common Searches: aspocells copy template row with formula to several worksheets c# | how to validate Excel formula results after CalculateFormula in Aspose.Cells .NET | duplicate a row containing a formula across multiple sheets using Aspose.Cells C# | programmatically check calculated Total column equals Quantity times Price with Aspose.Cells
// Tags: copy rows with formulas Aspose.Cells C# | template row duplication across worksheets | calculate and verify Excel formulas programmatically | populate multiple sheets with sample data | formula validation after CalculateFormula

using System;
using Aspose.Cells;

// The program creates a workbook, defines a template sheet with a header and a data row that computes Total = Quantity × Price, copies this row to three additional worksheets, populates Quantity and Price values, triggers calculation of all formulas, validates that each Total matches the expected product, and saves the workbook as Result.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Use the first worksheet as the template sheet
        Worksheet templateSheet = workbook.Worksheets[0];
        templateSheet.Name = "Template";

        // Header row
        templateSheet.Cells["A1"].PutValue("Item");
        templateSheet.Cells["B1"].PutValue("Quantity");
        templateSheet.Cells["C1"].PutValue("Price");
        templateSheet.Cells["D1"].PutValue("Total");

        // Template data row (row 2) with a formula for Total = Quantity * Price
        templateSheet.Cells["A2"].PutValue("Sample");
        templateSheet.Cells["B2"].PutValue(0);
        templateSheet.Cells["C2"].PutValue(0);
        templateSheet.Cells["D2"].Formula = "=B2*C2";

        // Names of sheets that will receive the copied template row
        string[] targetSheetNames = { "Sheet1", "Sheet2", "Sheet3" };

        // Copy the header and template row to each target sheet
        foreach (string sheetName in targetSheetNames)
        {
            Worksheet ws = workbook.Worksheets.Add(sheetName);

            // Copy header (row 1)
            ws.Cells.CopyRows(templateSheet.Cells, 0, 0, 1);

            // Copy template row (row 2)
            ws.Cells.CopyRows(templateSheet.Cells, 1, 1, 1);
        }

        // Populate each sheet with sample data and trigger calculation
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == "Template") continue; // Skip the template sheet

            // Example data: Quantity = 5, Price = 10
            ws.Cells["B2"].PutValue(5);
            ws.Cells["C2"].PutValue(10);
        }

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Validate that the Total column matches Quantity * Price
        foreach (Worksheet ws in workbook.Worksheets)
        {
            if (ws.Name == "Template") continue; // Skip validation for the template sheet

            double quantity = ws.Cells["B2"].DoubleValue;
            double price = ws.Cells["C2"].DoubleValue;
            double total = ws.Cells["D2"].DoubleValue;
            double expected = quantity * price;

            if (Math.Abs(total - expected) > 0.0001)
            {
                Console.WriteLine($"{ws.Name}: Validation FAILED. Expected {expected}, got {total}");
            }
            else
            {
                Console.WriteLine($"{ws.Name}: Validation PASSED. Total = {total}");
            }
        }

        // Save the workbook to a file
        workbook.Save("Result.xlsx");
    }
}
