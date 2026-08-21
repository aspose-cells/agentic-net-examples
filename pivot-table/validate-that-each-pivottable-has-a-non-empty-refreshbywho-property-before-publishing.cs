// Title: Aspose.Cells for .NET – Verify PivotTable RefreshedByWho is set before saving
// Description: Creates a workbook, adds a pivot table, refreshes it to populate the RefreshedByWho field, then scans every worksheet to ensure the property is not empty. An InvalidOperationException is thrown for any missing value, and the file is saved only after successful validation.
// Keywords: Aspose.Cells pivot validation | RefreshedByWho property .NET | check pivot user name | pivot table metadata verification | Aspose.Cells workbook publishing safety
// Common Searches: how to ensure RefreshedByWho is populated for each PivotTable | Aspose.Cells validate pivot table user info before save | throw error when PivotTable RefreshedByWho is blank | C# check pivot table refreshed by who
// Developer Intent: Confirm that every PivotTable in a workbook has a non‑empty RefreshedByWho value before the workbook is persisted.
// Use Cases: Loop through all worksheets and their PivotTables to enforce RefreshedByWho presence. | Refresh and calculate pivot data first so the property is filled, then perform validation. | Prevent publishing of workbooks that lack user attribution on any pivot table.
// AI Prompts: Generate a helper method that returns a list of PivotTable names with an empty RefreshedByWho in a given Workbook. | Write a unit test that expects InvalidOperationException when a PivotTable's RefreshedByWho is not set. | Suggest an alternative approach to validate RefreshedByWho using Workbook.PivotTables collection without nested loops.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableValidationDemo
{
    // Creates a workbook, adds a pivot table, refreshes it to populate the RefreshedByWho field, then scans every worksheet to ensure the property is not empty. An InvalidOperationException is thrown for any missing value, and the file is saved only after successful validation.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Amount");
            sheet.Cells["A2"].PutValue("Food");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Drink");
            sheet.Cells["B3"].PutValue(80);
            sheet.Cells["A4"].PutValue("Food");
            sheet.Cells["B4"].PutValue(150);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Amount as data field

            // Refresh the pivot table so that RefreshedByWho gets populated
            pivot.RefreshData();
            pivot.CalculateData();

            // Validate that every pivot table in the workbook has a non‑empty RefreshedByWho
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (PivotTable pt in ws.PivotTables)
                {
                    string refreshedBy = pt.RefreshedByWho; // Property provides the last user name
                    if (string.IsNullOrWhiteSpace(refreshedBy))
                    {
                        throw new InvalidOperationException(
                            $"PivotTable '{pt.Name}' in worksheet '{ws.Name}' does not have a valid RefreshedByWho value.");
                    }
                }
            }

            // All validations passed – save the workbook (publishing)
            workbook.Save("ValidatedPivotTable.xlsx");
        }
    }
}
