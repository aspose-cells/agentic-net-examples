// Title: C# – Update a Pivot Table Calculated Field with a Discount‑Rate Cell Using Aspose.Cells
// Description: Loads an existing workbook, finds the first pivot table, writes a discount rate to cell D1, rebuilds the calculated field "DiscountedSales" with a formula that references $D$1, refreshes and recalculates the pivot, then saves the result as output.xlsx.
// Keywords: Aspose.Cells | C# | pivot table | calculated field | update formula | discount rate cell | dynamic discount | RefreshData | CalculateData | Excel automation | replace calculated field
// Common Searches: how to change formula of an existing calculated field in Aspose.Cells pivot table | add discount rate cell to calculated field Aspose.Cells C# | replace calculated field without recreating pivot table Aspose.Cells | refresh pivot table after updating calculated field Aspose.Cells | dynamic discount calculation in Excel pivot using Aspose.Cells
// Developer Intent: Modify the formula of an existing calculated field in a pivot table so that it uses a discount rate stored in a worksheet cell, then refresh the pivot to apply the change.
// Use Cases: Turn a hard‑coded discount into a user‑editable rate by referencing a worksheet cell. | Update a pivot table’s calculated field without rebuilding the entire pivot. | Ensure the pivot reflects the new formula by calling RefreshData and CalculateData. | Validate the presence of the calculated field before overwriting it to avoid duplicates.
// AI Prompts: Write C# code with Aspose.Cells that updates the calculated field "DiscountedSales" in a pivot table to use cell D1 as the discount rate. | Show how to check for an existing calculated field in a pivot table and replace its formula with a new one that references a worksheet cell. | Explain the steps to refresh and recalculate a pivot table after changing a calculated field formula in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an existing workbook, finds the first pivot table, writes a discount rate to cell D1, rebuilds the calculated field "DiscountedSales" with a formula that references $D$1, refreshes and recalculates the pivot, then saves the result as output.xlsx.
class UpdateCalculatedField
{
    static void Main()
    {
        // Load the existing workbook that contains the pivot table
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure the worksheet has at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found.");
            return;
        }

        // Get the first pivot table (adjust index if needed)
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Name of the calculated field that we want to update
        string calculatedFieldName = "DiscountedSales";

        // Locate the existing calculated field (optional, just for verification)
        PivotField existingField = null;
        foreach (PivotField field in pivotTable.DataFields)
        {
            if (field.Name.Equals(calculatedFieldName, StringComparison.OrdinalIgnoreCase) && field.IsCalculatedField)
            {
                existingField = field;
                break;
            }
        }

        // If the field exists, we can retrieve its current formula (for logging)
        if (existingField != null)
        {
            Console.WriteLine($"Current formula for '{calculatedFieldName}': {existingField.GetFormula()}");
        }

        // Define a cell that holds the discount rate parameter (e.g., 10% discount)
        // This cell can be edited by the user without changing the code.
        worksheet.Cells["D1"].PutValue(0.10); // 10% discount

        // Build the new formula that incorporates the discount rate.
        // Assuming the original data field is named "Sales".
        string newFormula = $"=Sales*(1-$D$1)";

        // Add the calculated field again with the same name.
        // Aspose.Cells will replace the existing calculated field with the new definition.
        // The third argument 'true' drags the field to the data area immediately.
        pivotTable.AddCalculatedField(calculatedFieldName, newFormula, true);

        // Refresh and recalculate the pivot table to apply the new formula.
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the updated workbook.
        workbook.Save("output.xlsx");

        Console.WriteLine($"Calculated field '{calculatedFieldName}' updated with new formula: {newFormula}");
    }
}
