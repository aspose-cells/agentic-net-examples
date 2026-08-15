// Title: Disable Pivot Table Ribbons in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds a simple pivot table, and programmatically hides the PivotTable wizard, field‑list pane, and ribbon field‑list button using EnableWizard, EnableFieldList, and HidePivotFieldList before saving the file.
// Keywords: Aspose.Cells hide pivot ribbon | disable pivot wizard C# | remove pivot field list programmatically | Aspose.Cells pivot UI off | C# hide Excel ribbon elements | Aspose.Cells pivot table settings | Excel UI cleanup Aspose
// Common Searches: how to hide pivot table wizard in Aspose.Cells .NET | disable pivot field list ribbon button C# | programmatically turn off pivot UI Aspose.Cells | Aspose.Cells hide pivot ribbons before saving workbook | C# Aspose.Cells remove pivot table UI elements
// Developer Intent: Hide all pivot‑table related ribbon UI elements before saving the workbook.
// Use Cases: Generate a clean report for end users where pivot editing tools are unnecessary. | Create a template that prevents users from altering the pivot layout via the ribbon. | Distribute an Excel file with a locked pivot UI to avoid accidental changes.
// AI Prompts: Show C# code to hide the PivotTable wizard and field list in Aspose.Cells. | How can I disable the pivot field‑list button on the Excel ribbon using Aspose.Cells? | Provide an example that turns off all pivot UI elements before saving a workbook with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRibbonDisable
{
    // Creates a workbook, adds a simple pivot table, and programmatically hides the PivotTable wizard, field‑list pane, and ribbon field‑list button using EnableWizard, EnableFieldList, and HidePivotFieldList before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1000);
            sheet.Cells["A3"].PutValue("Orange");
            sheet.Cells["B3"].PutValue(1500);
            sheet.Cells["A4"].PutValue("Banana");
            sheet.Cells["B4"].PutValue(2000);

            // Add a pivot table based on the sample data
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Product column
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Sales column

            // Disable UI elements (ribbons) related to the pivot table
            pivotTable.EnableWizard = false;        // Hide the PivotTable Wizard
            pivotTable.EnableFieldList = false;     // Hide the field list pane
            workbook.Settings.HidePivotFieldList = true; // Hide the field list button on the ribbon

            // Save the workbook
            workbook.Save("PivotTable_Ribbons_Disabled.xlsx");
        }
    }
}
