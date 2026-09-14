// Title: Disable row grand totals in an Aspose.Cells pivot table using C#
// AI Prompts: Generate C# code that creates a pivot table from a range and sets ShowRowGrandTotals = false with Aspose.Cells. | Write a C# example that adds row and data fields to a pivot table and suppresses the row grand total. | Provide a C# snippet that recalculates the pivot table and saves the workbook after turning off row grand totals.
// Common Searches: Aspose.Cells C# how to hide row grand totals in a pivot table | C# pivot table without row grand total using Aspose.Cells library | disable row grand totals in Aspose.Cells pivot table example code
// Tags: Aspose.Cells ShowRowGrandTotals property | C# pivot table row grand total suppression | Aspose.Cells pivot table configuration | turn off row grand totals Aspose.Cells | Aspose.Cells generate pivot without row totals

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotDemo
{
    // The example creates a new workbook, fills sample data, adds a pivot table on range A1:C6, places 'Product' as a row field and 'Sales' as a data field, disables row grand totals by setting ShowRowGrandTotals to false, recalculates the pivot, and saves the file as PivotTable_NoRowGrandTotals.xlsx.
    class DisableRowGrandTotals
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Product";
            sheet.Cells["B1"].Value = "Region";
            sheet.Cells["C1"].Value = "Sales";

            sheet.Cells["A2"].Value = "Product A";
            sheet.Cells["B2"].Value = "North";
            sheet.Cells["C2"].Value = 1000;

            sheet.Cells["A3"].Value = "Product B";
            sheet.Cells["B3"].Value = "South";
            sheet.Cells["C3"].Value = 2000;

            sheet.Cells["A4"].Value = "Product C";
            sheet.Cells["B4"].Value = "East";
            sheet.Cells["C4"].Value = 3000;

            sheet.Cells["A5"].Value = "Product A";
            sheet.Cells["B5"].Value = "West";
            sheet.Cells["C5"].Value = 1500;

            sheet.Cells["A6"].Value = "Product B";
            sheet.Cells["B6"].Value = "North";
            sheet.Cells["C6"].Value = 2500;

            // Add a pivot table based on the data range
            PivotTableCollection pivotTables = sheet.PivotTables;
            int pivotIndex = pivotTables.Add("A1:C6", "E3", "PivotTable1");
            PivotTable pivotTable = pivotTables[pivotIndex];

            // Configure the pivot table: add row and data fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Disable row grand totals
            pivotTable.ShowRowGrandTotals = false;

            // Recalculate the pivot table to apply changes
            pivotTable.CalculateData();

            // Save the workbook to a file
            workbook.Save("PivotTable_NoRowGrandTotals.xlsx");
        }
    }
}
