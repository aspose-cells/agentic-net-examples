// Title: Hide the PivotTable field list ribbon in an Aspose.Cells workbook using C#
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, adds a PivotTable, and disables the PivotTable field list ribbon for the entire workbook. | Show how to turn off the field list UI for a specific PivotTable while leaving other workbook settings unchanged in Aspose.Cells C#. | Provide a complete example that sets workbook.Settings.HidePivotFieldList, saves the file, and explains the impact on the Ribbon interface.
// Common Searches: aspnet hide pivot table field list ribbon aspose.cells c# | how to disable pivot table field list in Aspose.Cells workbook | programmatically remove pivot table ribbon UI using Aspose.Cells .NET | Aspose.Cells hide pivot field list for specific pivot table C# example
// Tags: Aspose.Cells hide pivot field list | Workbook.Settings.HidePivotFieldList C# | disable pivot table UI Aspose.Cells | pivot table ribbon suppression .NET | field list visibility control Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotHideRibbonDemo
{
    // The example creates a workbook, adds sample data, builds a PivotTable, sets workbook.Settings.HidePivotFieldList to true to hide the PivotTable field list ribbon for the whole workbook (with an optional line to disable it for a single PivotTable), and saves the result as PivotTable_HideRibbon.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
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

            // Add a new worksheet to host the pivot table
            int pivotSheetIndex = workbook.Worksheets.Add(SheetType.Worksheet);
            Worksheet pivotSheet = workbook.Worksheets[pivotSheetIndex];
            pivotSheet.Name = "PivotTable";

            // Create the pivot table
            int pivotTableIndex = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotTableIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Hide the PivotTable field list (ribbon interface) for the entire workbook
            workbook.Settings.HidePivotFieldList = true;

            // Optionally, also disable the field list for this specific pivot table
            // pivotTable.EnableFieldList = false;

            // Save the workbook
            workbook.Save("PivotTable_HideRibbon.xlsx");
        }
    }
}
