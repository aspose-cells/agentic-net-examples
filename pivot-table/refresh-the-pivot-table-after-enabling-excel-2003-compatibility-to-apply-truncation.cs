// Title: Refresh a PivotTable after enabling Excel 2003 compatibility to truncate strings >255 chars – Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add a PivotTable, turn on Excel 2003 compatibility (which truncates text longer than 255 characters), and then refresh and calculate the pivot so the truncation is applied before saving the file.
// Keywords: Aspose.Cells PivotTable refresh | Excel 2003 compatibility | IsExcel2003Compatible | truncate long strings | RefreshData | CalculateData | C# Aspose.Cells example | legacy Excel support | global
// Common Searches: Aspose.Cells refresh pivot after setting IsExcel2003Compatible | how to truncate text in PivotTable for Excel 2003 | C# example for Excel 2003 compatibility mode in Aspose.Cells | pivot table string length limit 255 characters
// Developer Intent: Apply Excel 2003 compatibility to a PivotTable and programmatically refresh it so that any text exceeding 255 characters is truncated.
// Use Cases: Generating workbooks that must open correctly in Excel 2003 or earlier. | Preparing legacy reports where PivotTable fields cannot exceed 255 characters. | Automating data cleanup for long description fields before exporting to older Excel versions.
// AI Prompts: Show C# code that refreshes a PivotTable after enabling IsExcel2003Compatible in Aspose.Cells. | Explain how RefreshData and CalculateData work together to enforce the 255‑character limit in a pivot. | Provide a step‑by‑step guide to truncate long text fields in a PivotTable using Excel 2003 compatibility mode.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add a PivotTable, turn on Excel 2003 compatibility (which truncates text longer than 255 characters), and then refresh and calculate the pivot so the truncation is applied before saving the file.
class RefreshPivot2003Compatibility
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate source data; the description exceeds 255 characters
        worksheet.Cells["A1"].PutValue("Product");
        worksheet.Cells["B1"].PutValue("Description");
        worksheet.Cells["A2"].PutValue("Item1");
        worksheet.Cells["B2"].PutValue(new string('X', 300)); // long string

        // Add a pivot table based on the source data
        int pivotIndex = worksheet.PivotTables.Add("A1:B2", "D4", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Description as data field

        // Enable Excel 2003 compatibility (strings >255 will be truncated)
        pivotTable.IsExcel2003Compatible = true;

        // Refresh the pivot table so the truncation takes effect
        pivotTable.RefreshData();      // Refreshes data from the source
        pivotTable.CalculateData();    // Calculates the refreshed data into the worksheet

        // Save the workbook
        workbook.Save("Pivot2003Compatibility.xlsx");
    }
}
