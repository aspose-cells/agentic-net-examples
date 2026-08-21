// Title: Hide Expand/Collapse Buttons on a Pivot Chart with Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, adds sample data, builds a pivot table, generates a column pivot chart, and disables the chart's expand/collapse field buttons by setting chart.PivotOptions.ShowExpandCollapseFieldButtons to false, then saves the file.
// Keywords: Aspose.Cells C# hide pivot chart expand collapse | ShowExpandCollapseFieldButtons false | disable pivot chart UI elements | Aspose.Cells pivot options | remove expand collapse buttons Excel | C# programmatic pivot chart customization
// Common Searches: Aspose.Cells hide expand collapse buttons pivot chart | Set ShowExpandCollapseFieldButtons to false C# | Disable pivot chart field buttons programmatically | How to simplify pivot chart UI with Aspose.Cells | C# example for removing expand collapse buttons from pivot chart
// Developer Intent: Programmatically turn off the expand/collapse field buttons on a pivot chart to present a flat, non‑hierarchical view.
// Use Cases: Create clean dashboards where hierarchical navigation is unnecessary. | Generate Excel reports for end users that omit extra UI clutter. | Automate bulk export of pivot charts and ensure a consistent, simplified appearance.
// AI Prompts: Write C# code using Aspose.Cells that builds a pivot chart and disables its expand/collapse buttons. | Explain what happens when chart.PivotOptions.ShowExpandCollapseFieldButtons is set to false. | Provide a reusable method that accepts a Worksheet and disables expand/collapse buttons on all its pivot charts.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;

// This example creates a workbook, adds sample data, builds a pivot table, generates a column pivot chart, and disables the chart's expand/collapse field buttons by setting chart.PivotOptions.ShowExpandCollapseFieldButtons to false, then saves the file.
class DisableExpandCollapseDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(150);

        // Add a pivot table
        int ptIndex = sheet.PivotTables.Add("A1:B4", "E3", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[ptIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Value");
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Add a pivot chart based on the pivot table
        int chartIndex = sheet.Charts.Add(ChartType.Column, 9, 0, 24, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.PivotSource = "Data!PivotTable1";

        // Disable expand/collapse field buttons for the pivot chart
        chart.PivotOptions.ShowExpandCollapseFieldButtons = false;

        // Save the workbook
        workbook.Save("PivotChart_NoExpandCollapse.xlsx");
    }
}
