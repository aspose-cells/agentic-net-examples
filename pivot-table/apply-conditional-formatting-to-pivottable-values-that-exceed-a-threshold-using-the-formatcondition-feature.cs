// Title: Conditional Formatting for PivotTable Values in Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, build a pivot table, and use PivotConditionalFormat with a FormatCondition to highlight data field values that are greater than or equal to 100 (yellow background), then refresh, calculate, and save the file.
// Keywords: Aspose.Cells | C# | PivotTable | Conditional Formatting | PivotConditionalFormat | FormatCondition | threshold highlighting | yellow background | .NET
// Common Searches: Aspose.Cells conditional formatting pivot table C# | How to highlight pivot table values above a threshold using Aspose.Cells | PivotConditionalFormat example .NET | Set background color for pivot data field Aspose.Cells | Apply FormatCondition to pivot table in C#
// Developer Intent: Add a conditional format that highlights pivot table data values meeting a specific condition, such as values ≥ 100.
// Use Cases: Emphasize sales figures over a target amount in a financial report. | Flag expense categories that exceed budget limits in a budgeting dashboard. | Color‑code KPI values in a management pivot to quickly show performance tiers.
// AI Prompts: Generate C# code with Aspose.Cells to apply a red font style to pivot table values less than 50. | Show how to create multiple conditional formats for different value ranges in a pivot table using Aspose.Cells. | Explain how to refresh and recalculate a pivot table after adding conditional formatting with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using System.Drawing;

namespace AsposeCellsConditionalFormattingDemo
{
    // Demonstrates how to create a workbook, build a pivot table, and use PivotConditionalFormat with a FormatCondition to highlight data field values that are greater than or equal to 100 (yellow background), then refresh, calculate, and save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "A";
            cells["B2"].Value = 120;
            cells["A3"].Value = "B";
            cells["B3"].Value = 80;
            cells["A4"].Value = "C";
            cells["B4"].Value = 150;
            cells["A5"].Value = "A";
            cells["B5"].Value = 95;
            cells["A6"].Value = "B";
            cells["B6"].Value = 110;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Configure the pivot: rows = Category, data = Amount
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Add a conditional format to the pivot table
            int formatIdx = pivot.ConditionalFormats.Add();
            PivotConditionalFormat pcf = pivot.ConditionalFormats[formatIdx];

            // Apply the format to the data field area
            pcf.AddFieldArea(PivotFieldType.Data, pivot.DataFields[0]);

            // Define the condition: values >= 100 will be highlighted
            int conditionIdx = pcf.FormatConditions.AddCondition(FormatConditionType.CellValue);
            FormatCondition condition = pcf.FormatConditions[conditionIdx];
            condition.Operator = OperatorType.GreaterOrEqual;
            condition.Formula1 = "100";

            // Set the style for the condition (yellow background)
            condition.Style.BackgroundColor = Color.Yellow;

            // Refresh and calculate the pivot table to apply formatting
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotConditionalFormattingDemo.xlsx");
        }
    }
}
