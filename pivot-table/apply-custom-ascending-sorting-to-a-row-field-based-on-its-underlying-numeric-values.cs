// Title: Sort Pivot Table Row Field Ascending by Numeric Data Field Using Aspose.Cells for .NET (C#)
// AI Prompts: Set IsAutoSort = true, IsAscendSort = true, and AutoSortField = 0 on a PivotField to sort rows by the Amount column with Aspose.Cells in C#. | Refresh and calculate the pivot table after configuring custom row sorting based on a data field index using Aspose.Cells. | Apply PivotField.SortBy(SortOrder.Ascending, 0) to achieve the same ascending sort on a pivot row field.
// Common Searches: Aspose.Cells C# how to sort pivot table rows by a numeric data field | custom sort order for pivot row field using data field index in Aspose.Cells | enable auto sort for pivot table row field in .NET workbook | programmatically sort pivot rows by Amount column with Aspose.Cells
// Tags: pivotfield autosort ascending Aspose.Cells C# | custom pivot row sorting by data field index | sort pivot rows by numeric column Aspose.Cells | refresh and calculate pivot table Aspose.Cells | set IsAutoSort property pivotfield

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds Category and Amount data, builds a pivot table, places Category as a row field and Amount as a data field, enables auto‑sorting on the row field, configures it to sort ascending using the first data field (Amount), refreshes and calculates the pivot table, and saves the workbook as SortedPivot.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: a text field and a numeric field
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";
        sheet.Cells["A2"].Value = "C";
        sheet.Cells["B2"].Value = 300;
        sheet.Cells["A3"].Value = "A";
        sheet.Cells["B3"].Value = 100;
        sheet.Cells["A4"].Value = "B";
        sheet.Cells["B4"].Value = 200;

        // Add a pivot table covering the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the row field (Category) and the data field (Amount)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Retrieve the row field object
        PivotField rowField = pivotTable.RowFields[0];

        // Apply custom ascending sorting based on the numeric data field (Amount)
        // - Enable auto‑sorting
        // - Set ascending order
        // - Specify that sorting should be done by the first data field (index 0)
        rowField.IsAutoSort = true;
        rowField.IsAscendSort = true;
        rowField.AutoSortField = 0;               // 0 = first data field (Amount)

        // Alternatively, the same effect can be achieved with the SortBy method:
        // rowField.SortBy(SortOrder.Ascending, 0);

        // Refresh and calculate the pivot table to apply the sorting
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the sorted pivot table
        workbook.Save("SortedPivot.xlsx");
    }
}
