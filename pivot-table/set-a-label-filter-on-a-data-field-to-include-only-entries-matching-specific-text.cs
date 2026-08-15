// Title: Aspose.Cells for .NET – Apply a CaptionEqual label filter to a pivot table row field
// Description: Demonstrates creating a workbook, adding a pivot table, and using PivotField.FilterByLabel with PivotFilterType.CaptionEqual to show only rows labeled "Apple", then refreshing and saving the file.
// Keywords: Aspose.Cells | .NET | pivot table label filter | PivotField.FilterByLabel | CaptionEqual | C# example | filter pivot rows | Aspose.Cells API | Excel automation
// Common Searches: Aspose.Cells filter pivot rows by label C# | PivotField.FilterByLabel example .NET | How to apply CaptionEqual filter in Aspose.Cells | Show only specific items in Aspose.Cells pivot table | C# Aspose.Cells label filter pivot table
// Developer Intent: Filter pivot table rows to include only entries with a specific label.
// Use Cases: Generate a sales report that displays only the "Apple" product. | Build an interactive dashboard where users select a product name to filter pivot rows. | Create automated Excel exports that pre‑filter data for predefined categories.
// AI Prompts: Write C# code using Aspose.Cells to apply a CaptionEqual filter for "Banana" on a pivot table row field. | Show how to add multiple label filters (e.g., "Apple" and "Cherry") to a pivot table with Aspose.Cells. | Explain the steps to refresh and recalculate a pivot table after applying a label filter in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsLabelFilterDemo
{
    // Demonstrates creating a workbook, adding a pivot table, and using PivotField.FilterByLabel with PivotFilterType.CaptionEqual to show only rows labeled "Apple", then refreshing and saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Product");
            cells["B1"].PutValue("Sales");
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(120);
            cells["A3"].PutValue("Banana");
            cells["B3"].PutValue(80);
            cells["A4"].PutValue("Apple");
            cells["B4"].PutValue(150);
            cells["A5"].PutValue("Cherry");
            cells["B5"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "ProductPivot");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add the "Product" field to the row area and "Sales" to the data area
            pivot.AddFieldToArea(PivotFieldType.Row, "Product");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Get the row field (Product) and apply a label filter to include only "Apple"
            PivotField productField = pivot.RowFields[0];
            productField.FilterByLabel(PivotFilterType.CaptionEqual, "Apple", null);

            // Refresh and calculate the pivot table to apply the filter
            pivot.RefreshData();
            pivot.CalculateData();

            // Save the workbook
            workbook.Save("LabelFilterPivotDemo.xlsx");
        }
    }
}
