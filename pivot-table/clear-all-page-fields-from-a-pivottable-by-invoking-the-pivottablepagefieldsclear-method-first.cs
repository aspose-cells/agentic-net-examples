// Title: Aspose.Cells C# – Clear All Page Fields from a PivotTable
// Description: Demonstrates how to create a workbook, add sample data, build a PivotTable with a page (filter) field, remove every page field using PivotTable.PageFields.Clear(), refresh and recalculate the PivotTable, and save the result.
// Keywords: Aspose.Cells | C# | .NET | PivotTable | PageFields.Clear | remove page filters | clear pivot table filters | Aspose.Cells example | programmatic pivot table manipulation | Excel automation
// Common Searches: Aspose.Cells clear page fields C# | how to remove pivot table page filters with Aspose.Cells | PivotTable.PageFields.Clear example | reset pivot table filters Aspose.Cells .NET | C# code to clear pivot table page fields
// Developer Intent: Programmatically delete all page (filter) fields from a PivotTable in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Reset a template workbook before generating a new report. | Remove user‑selected page filters to prevent data leakage. | Dynamically rebuild a PivotTable by clearing existing page fields and adding new ones based on runtime criteria.
// AI Prompts: Show a C# example that clears all page fields from an Aspose.Cells PivotTable and then adds new page fields. | Provide code to check for existing page fields in a PivotTable and safely remove them using Aspose.Cells. | Explain the steps to refresh and recalculate a PivotTable after calling PivotTable.PageFields.Clear in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add sample data, build a PivotTable with a page (filter) field, remove every page field using PivotTable.PageFields.Clear(), refresh and recalculate the PivotTable, and save the result.
class ClearPivotTablePageFields
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");

        sheet.Cells["A2"].PutValue("P1");
        sheet.Cells["B2"].PutValue("North");
        sheet.Cells["C2"].PutValue(120);

        sheet.Cells["A3"].PutValue("P2");
        sheet.Cells["B3"].PutValue("South");
        sheet.Cells["C3"].PutValue(200);

        sheet.Cells["A4"].PutValue("P3");
        sheet.Cells["B4"].PutValue("East");
        sheet.Cells["C4"].PutValue(150);

        sheet.Cells["A5"].PutValue("P1");
        sheet.Cells["B5"].PutValue("West");
        sheet.Cells["C5"].PutValue(180);

        // Add a pivot table that includes a page field (Region)
        int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add fields to the pivot table
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");

        // Clear all page fields from the pivot table
        pivotTable.PageFields.Clear();

        // Refresh and recalculate the pivot table after clearing page fields
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearPageFieldsDemo.xlsx");
    }
}
