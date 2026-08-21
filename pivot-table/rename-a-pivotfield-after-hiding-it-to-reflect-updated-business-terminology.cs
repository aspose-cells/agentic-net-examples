// Title: Rename a Row PivotField After Hiding Items with Aspose.Cells for .NET
// Description: Creates a workbook, builds a pivot table from a small dataset, hides every row item except "Finance", refreshes and calculates the pivot, then renames the row field from "Category" to "Business Unit" and saves the file.
// Keywords: Aspose.Cells C# pivot table rename field | hide pivot items programmatically | change pivot field caption .NET | PivotField rename after filter | Aspose.Cells PivotTable API
// Common Searches: how to rename a pivot row field after hiding items in Aspose.Cells | Aspose.Cells C# hide pivot items and change field name | update pivot field name after filtering with Aspose.Cells | programmatically rename pivot field in .NET | Aspose.Cells rename pivot field example
// Developer Intent: Update the name of a row PivotField after applying a hide filter so the report reflects new business terminology.
// Use Cases: Filter out unwanted categories in a pivot table, then rename the field for clearer reporting. | Apply custom row‑item visibility, refresh the pivot, and adjust the field caption without rebuilding the table. | Automate terminology changes in generated Excel reports by renaming PivotFields after data manipulation.
// AI Prompts: Generate C# code using Aspose.Cells to hide specific row items in a pivot table and then rename the row field. | Explain how to rename a PivotField after calling RefreshData and CalculateData in Aspose.Cells. | Show the steps to change a pivot table row field caption while preserving hidden items in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotRenameDemo
{
    // Creates a workbook, builds a pivot table from a small dataset, hides every row item except "Finance", refreshes and calculates the pivot, then renames the row field from "Category" to "Business Unit" and saves the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";

            cells["A2"].Value = "Finance";
            cells["B2"].Value = 1200;

            cells["A3"].Value = "Technology";
            cells["B3"].Value = 2500;

            cells["A4"].Value = "Healthcare";
            cells["B4"].Value = 1800;

            cells["A5"].Value = "Finance";
            cells["B5"].Value = 900;

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Add the "Category" field to the row area
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");

            // Add the "Amount" field to the data area
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

            // Hide all categories except "Finance"
            PivotField rowField = pivotTable.RowFields[0];
            for (int i = 0; i < rowField.ItemCount; i++)
            {
                // Use HideItem(string, bool) to hide items that are not "Finance"
                rowField.HideItem(rowField.Items[i], rowField.Items[i] != "Finance");
            }

            // Refresh and calculate the pivot table after hiding items
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Rename the pivot field to reflect new business terminology
            // For example, change "Category" to "Business Unit"
            rowField.Name = "Business Unit";

            // Save the workbook
            workbook.Save("PivotFieldRenamed.xlsx");
        }
    }
}
