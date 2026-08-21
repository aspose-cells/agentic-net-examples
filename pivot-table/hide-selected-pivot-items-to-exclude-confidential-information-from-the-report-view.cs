// Title: Hide a confidential row item in an Aspose.Cells PivotTable (C#)
// Description: Demonstrates how to create a workbook, build a PivotTable, and hide a selected row field value (e.g., "Confidential") using the PivotField.HideItem method, then refresh and save the file.
// Keywords: Aspose.Cells | C# | PivotTable | HideItem | hide pivot row | exclude confidential data | pivot table filter | Aspose.Cells .NET
// Common Searches: Aspose.Cells hide pivot row item C# | How to hide confidential value in Aspose.Cells pivot table | PivotField.HideItem example Aspose.Cells | Exclude specific items from PivotTable using Aspose.Cells | C# hide pivot table item Aspose.Cells
// Developer Intent: Hide a selected pivot item so confidential information is not displayed in the pivot report.
// Use Cases: Remove the "Confidential" department from a revenue pivot report to protect sensitive data. | Loop through a list of restricted categories and hide each one programmatically. | Automatically hide pivot items whose values exceed a defined threshold before exporting the workbook.
// AI Prompts: Generate C# code that hides multiple pivot items in an Aspose.Cells PivotTable using a loop. | Show how to hide pivot items based on a numeric condition (e.g., revenue > 10000) with Aspose.Cells for .NET. | Provide an example of dynamically retrieving row field names and hiding selected items in a pivot table.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to create a workbook, build a PivotTable, and hide a selected row field value (e.g., "Confidential") using the PivotField.HideItem method, then refresh and save the file.
class HidePivotItemsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (Department and Revenue)
        sheet.Cells["A1"].Value = "Department";
        sheet.Cells["B1"].Value = "Revenue";

        sheet.Cells["A2"].Value = "HR";
        sheet.Cells["B2"].Value = 5000;

        sheet.Cells["A3"].Value = "Finance";
        sheet.Cells["B3"].Value = 12000;

        sheet.Cells["A4"].Value = "R&D";
        sheet.Cells["B4"].Value = 8000;

        // This row contains confidential information that must be hidden
        sheet.Cells["A5"].Value = "Confidential";
        sheet.Cells["B5"].Value = 15000;

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Department field to the row area and Revenue to the data area
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Department");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Revenue");

        // Retrieve the row field (Department) to manipulate its items
        PivotField departmentField = pivotTable.RowFields[0];

        // Hide the confidential item using the string overload of HideItem
        departmentField.HideItem("Confidential", true);

        // Refresh the pivot table to apply changes and recalculate data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the hidden pivot item
        workbook.Save("HiddenPivotItems.xlsx");
    }
}
