// Title: How to hide the PivotTable field list pane in an Excel file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a workbook, adds a pivot table, and disables the field list pane globally before saving. | Show a C# example that disables the field list pane for a single PivotTable by setting its EnableFieldList property to false using Aspose.Cells.
// Common Searches: Aspose.Cells C# prevent pivot field list from appearing in generated Excel | Disable field list for all pivot tables in Aspose.Cells workbook | Turn off field list on a single PivotTable with Aspose.Cells .NET | How to hide pivot table field list pane programmatically using Aspose.Cells | Aspose.Cells setting to suppress pivot field list in Excel output
// Tags: Aspose.Cells hide pivot field list | Workbook.Settings.HidePivotFieldList C# | PivotTable.EnableFieldList false Aspose.Cells | Excel pivot table field list visibility .NET | Aspose.Cells pivot table configuration C#

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Creates a workbook, adds sample data and a pivot table, disables the field list pane either globally via Workbook.Settings or per pivot table via EnableFieldList, and saves the file as HidePivotFieldListDemo.xlsx.
class HidePivotFieldListDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(150);

        // Add a new worksheet that will contain the pivot table
        int pivotSheetIdx = workbook.Worksheets.Add(SheetType.Worksheet);
        Worksheet pivotSheet = workbook.Worksheets[pivotSheetIdx];
        pivotSheet.Name = "PivotTable";

        // Create the pivot table on the new sheet
        int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Row field: Category
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Data field: Value

        // Hide the field list pane for the workbook (applies to all pivot tables)
        workbook.Settings.HidePivotFieldList = true;

        // Alternatively, hide the field list for this specific pivot table:
        // pivotTable.EnableFieldList = false;

        // Save the workbook
        workbook.Save("HidePivotFieldListDemo.xlsx");
    }
}
