// Title: Aspose.Cells .NET: Refresh PivotTable Data After Modifying Source Range
// Description: Demonstrates how to create a workbook, add source data, build a PivotTable, change a cell value, call RefreshData, recalculate, and save the updated file using Aspose.Cells for C#.
// Keywords: Aspose.Cells RefreshData | C# PivotTable update | Aspose.Cells PivotTable recalc | Refresh pivot after source change | .NET Excel pivot refresh
// Common Searches: Aspose.Cells refresh pivot table C# | RefreshData method example | Update PivotTable after source edit .NET | Programmatic pivot refresh Aspose
// Developer Intent: Update a PivotTable so it reflects changes made to its source range.
// Use Cases: Adjust sales numbers in a worksheet and instantly refresh the summary pivot before exporting. | Automate daily report generation where source data changes and the pivot must stay accurate. | Provide an Excel file that allows end‑users to edit data and have the pivot refreshed programmatically.
// AI Prompts: Generate C# code that modifies multiple source cells and then refreshes and recalculates the Aspose.Cells PivotTable. | Show how to loop through several PivotTables in one sheet and call RefreshData after each source update. | Explain how to detect source range edits and automatically invoke RefreshData for associated pivots using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRefreshDemo
{
    // Demonstrates how to create a workbook, add source data, build a PivotTable, change a cell value, call RefreshData, recalculate, and save the updated file using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate source data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["A4"].PutValue("Apple");
            sheet.Cells["B4"].PutValue(150);

            // Add a pivot table based on the source range A1:B4, placed at E3
            int pivotIndex = sheet.PivotTables.Add("A1:B4", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Initial calculation to populate the pivot table
            pivotTable.CalculateData();

            // Modify the source data (change a sales value)
            sheet.Cells["B2"].PutValue(120); // Apple sales changed from 100 to 120

            // Refresh the pivot table's data from the updated source range
            pivotTable.RefreshData();

            // Recalculate the pivot table after refreshing the data
            pivotTable.CalculateData();

            // Save the workbook with the refreshed pivot table
            workbook.Save("PivotTable_Refreshed.xlsx");
        }
    }
}
