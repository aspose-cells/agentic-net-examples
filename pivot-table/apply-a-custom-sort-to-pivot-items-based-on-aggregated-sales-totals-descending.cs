// Title: Custom descending sort of pivot row items by aggregated Sales using Aspose.Cells for .NET (C#)
// Description: The example creates a workbook, fills it with product and sales data, builds a pivot table, adds Product as a row field and Sales as a data field, refreshes and calculates the pivot, then sorts the Product rows by total Sales in descending order with PivotField.SortBy, and finally saves the workbook.
// Keywords: Aspose.Cells | C# | .NET | pivot table custom sort | SortBy method | descending order | row field sorting | aggregated sales totals | Excel automation | GitHub example | US developers | European developers | Indian developers
// Common Searches: Aspose.Cells sort pivot rows by total sales | C# custom pivot table sort descending | How to use SortBy with Aspose.Cells pivot | Pivot table row field sorting .NET | Example of descending pivot item order in Aspose.Cells
// Developer Intent: Sort the Product rows of a pivot table in descending order based on the summed Sales values.
// Use Cases: Generate a sales ranking report where products appear from highest to lowest revenue. | Create a reusable workbook that automatically orders rows after data updates. | Prepare a presentation‑ready Excel file with a pre‑sorted pivot for stakeholders.
// AI Prompts: Show how to sort pivot rows by multiple data fields using Aspose.Cells. | Provide a C# example that applies an ascending custom sort to a column field in a pivot table. | Explain how to retrieve the sorted order of pivot items after calling the SortBy method.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotCustomSort
{
    // The example creates a workbook, fills it with product and sales data, builds a pivot table, adds Product as a row field and Sales as a data field, refreshes and calculates the pivot, then sorts the Product rows by total Sales in descending order with PivotField.SortBy, and finally saves the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate sample data
                cells["A1"].Value = "Product";
                cells["B1"].Value = "Sales";

                cells["A2"].Value = "Apple";
                cells["B2"].Value = 1200;

                cells["A3"].Value = "Banana";
                cells["B3"].Value = 800;

                cells["A4"].Value = "Cherry";
                cells["B4"].Value = 1500;

                cells["A5"].Value = "Date";
                cells["B5"].Value = 600;

                // Add a pivot table based on the data range
                int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
                PivotTable pivotTable = sheet.PivotTables[pivotIndex];

                // Add fields to the pivot table
                pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
                pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Refresh and calculate the pivot table before applying custom sort
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Apply custom sort: sort the row field (Product) by the aggregated Sales totals in descending order
                // fieldSortedBy = 0 refers to the first data field (Sales)
                pivotTable.RowFields[0].SortBy(SortOrder.Descending, 0);

                // Save the workbook
                workbook.Save("CustomSortedPivot.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
