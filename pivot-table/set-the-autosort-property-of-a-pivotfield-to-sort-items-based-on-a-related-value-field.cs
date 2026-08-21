// Title: Set PivotField AutoSort to order rows by a data field in Aspose.Cells for .NET
// Description: Creates a workbook with Category and Sales columns, adds a pivot table, assigns Category to the row area and Sales to the data area, then enables AutoSort on the row field to sort ascending by the first data field, recalculates the pivot and saves the workbook.
// Keywords: Aspose.Cells AutoSort PivotField | C# pivot table sort by value | IsAutoSort property | AutoSortField Aspose.Cells | pivot row field ascending sort .NET
// Common Searches: Aspose.Cells sort pivot rows by data field | C# set AutoSort on PivotField | How to enable ascending auto‑sort for pivot table rows | Specify data field for pivot field auto sorting Aspose
// Developer Intent: Activate automatic ascending sorting of a pivot row field based on a selected data field.
// Use Cases: Generate Excel reports where product categories are displayed from highest to lowest sales. | Create dynamic pivot tables that re‑order rows whenever source data changes. | Automate workbook creation with consistently sorted pivot rows for dashboards.
// AI Prompts: Show how to configure AutoSort for a PivotField to sort by a specific data field in Aspose.Cells (C#). | Provide a C# example that sorts pivot row items descending using the second data field. | Explain the roles of IsAutoSort, IsAscendSort, and AutoSortField in Aspose.Cells pivot tables.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook with Category and Sales columns, adds a pivot table, assigns Category to the row area and Sales to the data area, then enables AutoSort on the row field to sort ascending by the first data field, recalculates the pivot and saves the workbook.
class SetPivotFieldAutoSort
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].Value = "Category";
        cells["B1"].Value = "Sales";
        cells["A2"].Value = "North";
        cells["B2"].Value = 1000;
        cells["A3"].Value = "South";
        cells["B3"].Value = 1500;
        cells["A4"].Value = "East";
        cells["B4"].Value = 800;
        cells["A5"].Value = "West";
        cells["B5"].Value = 1200;

        // Add a pivot table covering the data range
        int ptIndex = sheet.PivotTables.Add("A1:B5", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];

        // Add a row field (Category) and a data field (Sales)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field to configure auto‑sorting
        PivotField rowField = pivotTable.RowFields["Category"]; // alternatively pivotTable.RowFields[0]

        // Enable auto sort, set ascending order, and sort by the first data field (Sales)
        rowField.IsAutoSort = true;
        rowField.IsAscendSort = true;
        rowField.AutoSortField = 0; // 0 = first data field in the pivot table

        // Refresh the pivot table data and apply the sorting
        pivotTable.CalculateData();

        // Save the workbook with the configured pivot table
        workbook.Save("PivotFieldAutoSortResult.xlsx");
    }
}
