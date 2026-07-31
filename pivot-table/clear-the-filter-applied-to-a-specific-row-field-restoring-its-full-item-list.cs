// Title: Clear a Pivot Table Row Field Filter with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, add a pivot table, apply a label filter to a row field, then remove that filter using PivotField.ClearFilter, refresh the pivot, and save the file.
// Keywords: Aspose.Cells ClearFilter | PivotTable row field filter .NET | C# remove pivot label filter | Aspose.Cells PivotField.ClearFilter example | reset pivot row items programmatically
// Common Searches: Aspose.Cells how to clear pivot row filter | C# remove label filter from pivot table row field | PivotField.ClearFilter usage Aspose | reset filtered rows in Aspose.Cells pivot | clear specific pivot field filter programmatically
// Developer Intent: Programmatically remove an applied filter from a designated pivot table row field so that all row items become visible again.
// Use Cases: After filtering a pivot table to a single category, clear the filter to show the full category list without rebuilding the pivot. | Toggle row visibility in automated reports: apply a filter for a preview, then clear it before final export. | Ensure pivot totals reflect the complete dataset by clearing filters and recalculating the pivot.
// AI Prompts: Generate C# code using Aspose.Cells that clears a label filter on a pivot table row field and refreshes the pivot. | Explain the effect of PivotField.ClearFilter in Aspose.Cells and the steps required to update the pivot after calling it. | Provide a C# example that clears filters on multiple row fields of a pivot table with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, add a pivot table, apply a label filter to a row field, then remove that filter using PivotField.ClearFilter, refresh the pivot, and save the file.
class ClearPivotRowFieldFilter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].Value = "Category";
        worksheet.Cells["B1"].Value = "Sales";
        worksheet.Cells["A2"].Value = "Fruit";
        worksheet.Cells["B2"].Value = 120;
        worksheet.Cells["A3"].Value = "Vegetable";
        worksheet.Cells["B3"].Value = 80;
        worksheet.Cells["A4"].Value = "Fruit";
        worksheet.Cells["B4"].Value = 150;
        worksheet.Cells["A5"].Value = "Vegetable";
        worksheet.Cells["B5"].Value = 70;

        // Create a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:B5", "D2", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add the "Category" column as a row field
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);

        // Add the "Sales" column as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

        // Apply a filter on the row field to show only "Fruit"
        PivotField rowField = pivotTable.RowFields[0];
        rowField.FilterByLabel(PivotFilterType.CaptionEqual, "Fruit", null);
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Clear the filter on the specific row field, restoring all items
        rowField.ClearFilter();

        // Refresh the pivot table to reflect the cleared filter
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("ClearRowFieldFilter.xlsx");
    }
}
