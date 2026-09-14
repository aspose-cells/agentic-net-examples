// Title: How to disable column auto‑fit for an Aspose.Cells PivotTable and set fixed column widths in C#
// AI Prompts: Write C# code using Aspose.Cells to create a PivotTable, assign manual column widths, and ensure the widths are preserved when the pivot is refreshed. | Show the steps to prevent a PivotTable from automatically adjusting column sizes in Aspose.Cells, including setting column widths and refreshing the pivot.
// Common Searches: Aspose.Cells keep custom column widths in pivot table after refresh | prevent pivot table column auto‑sizing in C# Aspose.Cells | set manual column width for Aspose.Cells pivot table | how to stop column auto‑fit when updating Aspose.Cells pivot
// Tags: Aspose.Cells PivotTable column width management | C# Aspose.Cells disable column autofit | manual column sizing for Aspose pivot tables | Aspose.Cells refresh pivot without auto‑fit | fixed column width Aspose.Cells pivot example

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotAutoFitDemo
{
    // The example creates a workbook, adds sample data, builds a PivotTable on a separate sheet, manually sets column widths, disables the AutofitColumnWidthOnUpdate property, refreshes the pivot to retain the custom widths, and saves the file as PivotTable_NoAutoFit.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (source data)
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("Electronics");
            dataSheet.Cells["B2"].PutValue("Laptop");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("Electronics");
            dataSheet.Cells["B3"].PutValue("Phone");
            dataSheet.Cells["C3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue("Furniture");
            dataSheet.Cells["B4"].PutValue("Chair");
            dataSheet.Cells["C4"].PutValue(150);

            dataSheet.Cells["A5"].PutValue("Furniture");
            dataSheet.Cells["B5"].PutValue("Table");
            dataSheet.Cells["C5"].PutValue(300);

            // Add a new worksheet to host the pivot table
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Create the pivot table using the source range
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:C5", "A3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Calculate the pivot data
            pivotTable.CalculateData();

            // Manually set column widths to desired values
            // (example: set first three columns to a fixed width)
            pivotSheet.Cells.SetColumnWidth(0, 15); // Column A
            pivotSheet.Cells.SetColumnWidth(1, 20); // Column B
            pivotSheet.Cells.SetColumnWidth(2, 12); // Column C

            // Disable auto‑fit of column widths on pivot table update
            pivotTable.AutofitColumnWidthOnUpdate = false;

            // Refresh pivot tables to apply the setting (no auto‑fit will occur)
            pivotSheet.RefreshPivotTables();

            // Save the workbook
            workbook.Save("PivotTable_NoAutoFit.xlsx");

            Console.WriteLine("Pivot table created with auto‑fit disabled and custom column widths applied.");
        }
    }
}
