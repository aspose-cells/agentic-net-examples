// Title: Hide confidential row items in an Aspose.Cells PivotTable (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to hide specific row field values in a PivotTable and refresh the report. | Show how to programmatically exclude "Finance" and "Legal" items from a PivotTable row field using the PivotField.HideItem method. | Demonstrate saving a workbook after hiding selected pivot items with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# hide specific row values in pivot table | exclude finance and legal departments from pivot report using Aspose.Cells | programmatically hide pivot table items in .NET workbook | how to use PivotField.HideItem method in Aspose.Cells | refresh pivot table after hiding items Aspose.Cells C#
// Tags: Aspose.Cells PivotField.HideItem usage | hide pivot row items C# | confidential data exclusion Aspose.Cells | refresh pivot table after item hide .NET | save workbook with hidden pivot items

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, adds department and expense data, builds a PivotTable, hides the "Finance" and "Legal" entries in the Department row field using PivotField.HideItem, refreshes and recalculates the PivotTable, and saves the file as HiddenPivotItems.xlsx.
class HidePivotItemsDemo
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Department");
        sheet.Cells["B1"].PutValue("Expense");
        sheet.Cells["A2"].PutValue("HR");
        sheet.Cells["B2"].PutValue(5000);
        sheet.Cells["A3"].PutValue("Finance");
        sheet.Cells["B3"].PutValue(12000);
        sheet.Cells["A4"].PutValue("IT");
        sheet.Cells["B4"].PutValue(8000);
        sheet.Cells["A5"].PutValue("Legal");
        sheet.Cells["B5"].PutValue(3000);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add the Department field as a row field and Expense as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Department");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Expense");

        // Retrieve the row field (Department) to manipulate its items
        PivotField departmentField = pivotTable.RowFields[0];

        // Hide confidential items (Finance and Legal) so they do not appear in the report
        departmentField.HideItem("Finance", true);
        departmentField.HideItem("Legal", true);

        // Refresh the pivot table to apply the changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with hidden pivot items
        workbook.Save("HiddenPivotItems.xlsx");
    }
}
