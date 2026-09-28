// Title: Programmatically move a page field to the top of a PivotTable filter list using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to reorder page fields in a PivotTable, moving a chosen field to the first position to adjust filter priority. | Show how to configure the page field layout order and wrap count after changing the page field sequence in an Aspose.Cells PivotTable.
// Common Searches: Aspose.Cells C# reorder pivot table page fields to change filter priority | how to move a filter field to the first position in an Aspose.Cells PivotTable | set page field order DownThenOver and wrap count in Aspose.Cells .NET | programmatic pivot table page field positioning with Aspose.Cells | change filter hierarchy of page fields in an Aspose.Cells workbook
// Tags: Aspose.Cells PivotTable page field reordering | C# set PivotTable page field order DownThenOver | Aspose.Cells change filter priority in pivot table | PivotTable page field wrap count .NET | Aspose.Cells move page field position programmatically

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTablePageFieldReorderDemo
{
    // The example creates a workbook, fills it with sample data, adds a PivotTable with Category and Region as page (filter) fields, then moves the Region field to the first position to give it higher filter priority. It also sets the page field layout to DownThenOver, defines a wrap count of two, refreshes and calculates the PivotTable, and saves the result as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Header row
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Region";
            cells["C1"].Value = "Product";
            cells["D1"].Value = "Sales";

            // Data rows
            for (int i = 2; i <= 11; i++)
            {
                cells[$"A{i}"].Value = $"Category{(i % 3) + 1}";
                cells[$"B{i}"].Value = $"Region{(i % 2) + 1}";
                cells[$"C{i}"].Value = $"Product{(i % 4) + 1}";
                cells[$"D{i}"].Value = i * 100;
            }

            // Add a pivot table
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:D11", "F5", "SalesPivot");
            PivotTable pivot = pivots[pivotIndex];

            // Add fields to the pivot table
            // Page fields (filters) – order matters for filter priority
            pivot.AddFieldToArea(PivotFieldType.Page, "Category"); // initially first
            pivot.AddFieldToArea(PivotFieldType.Page, "Region");   // initially second
            // Row and data fields
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Reorder page fields: move "Region" to be the first filter (higher priority)
            // Current positions: Category = 0, Region = 1
            // Move field at position 1 to position 0
            pivot.PageFields.Move(1, 0);

            // Optional: change the layout order of page fields (DownThenOver)
            pivot.PageFieldOrder = PrintOrderType.DownThenOver;
            // Optional: set wrap count if you want multiple columns of page fields
            pivot.PageFieldWrapCount = 2;

            // Refresh and calculate the pivot table to apply changes
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_PageField_Reordered.xlsx");
        }
    }
}
