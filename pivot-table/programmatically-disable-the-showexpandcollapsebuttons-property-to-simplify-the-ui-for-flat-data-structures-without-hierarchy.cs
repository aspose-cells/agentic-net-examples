// Title: How to hide expand/collapse field buttons in a PivotChart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, adds a pivot table, links a column chart, and sets chart.PivotOptions.ShowExpandCollapseFieldButtons to false. | Generate an Aspose.Cells example that demonstrates disabling the expand/collapse UI in a PivotChart and saving the workbook. | Provide a step‑by‑step C# snippet to configure a PivotChart’s PivotOptions so the expand/collapse field buttons are not displayed.
// Common Searches: asp.net aspocells hide expand collapse buttons on pivot chart c# | c# Aspose.Cells disable ShowExpandCollapseFieldButtons property for PivotChart | how to remove expand/collapse field buttons from Excel pivot chart using Aspose.Cells | Aspose.Cells PivotChart UI customization hide field buttons | C# example to turn off expand collapse buttons in a pivot chart with Aspose.Cells
// Tags: Aspose.Cells PivotChart ShowExpandCollapseFieldButtons false | C# disable pivot chart expand collapse UI | Aspose.Cells chart.PivotOptions configuration | Excel pivot chart field button removal using Aspose.Cells | PivotChart UI simplification Aspose.Cells | Aspose.Cells hide expand collapse buttons | C# pivot chart customization Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// The example creates a new workbook, populates sample data, adds a pivot table, creates a linked column chart, disables the expand/collapse field buttons via chart.PivotOptions.ShowExpandCollapseFieldButtons = false, refreshes the chart, and saves the file as PivotChart_NoExpandCollapse.xlsx.
class DisableExpandCollapseButtonsDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");

        // Create a chart linked to the pivot table (PivotChart)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 9, 0, 24, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.PivotSource = "Sheet1!PivotTable1";

        // Disable the expand/collapse field buttons in the pivot chart UI
        chart.PivotOptions.ShowExpandCollapseFieldButtons = false;

        // Refresh the chart to apply pivot data changes
        chart.RefreshPivotData();

        // Save the workbook
        workbook.Save("PivotChart_NoExpandCollapse.xlsx");
    }
}
