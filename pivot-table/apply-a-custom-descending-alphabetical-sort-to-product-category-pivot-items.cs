// Title: Sort Pivot Table Row Items Descending Alphabetically with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to generate a workbook, insert sample sales data, create a pivot table, and programmatically sort the Category row field in reverse alphabetical order using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# pivot sort | descending row field sort | reverse alphabetical pivot | SortOrder.Descending Aspose | Excel pivot table automation .NET | custom pivot field ordering | programmatic Excel sorting | C# Excel workbook generation
// Common Searches: Aspose.Cells sort pivot rows descending | C# reverse alphabetical pivot field | How to set pivot table row order with Aspose | Programmatic pivot table sorting .NET | Sort pivot table categories Z to A using Aspose.Cells
// Developer Intent: Apply a reverse‑alphabetical sort to the Category row field of a pivot table.
// Use Cases: Produce sales dashboards where categories appear from Z to A for quick scanning. | Generate client‑ready reports that require pivot rows in reverse alphabetical order. | Automate monthly workbook creation with consistent row ordering to match corporate style guides.
// AI Prompts: Show how to modify the code to sort the Category field in ascending order. | Give an example of sorting a pivot data field by numeric totals with Aspose.Cells. | Explain how to turn off auto‑sort after applying a custom sort on a pivot field.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Demonstrates how to generate a workbook, insert sample sales data, create a pivot table, and programmatically sort the Category row field in reverse alphabetical order using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Product");
        worksheet.Cells["C1"].PutValue("Sales");

        worksheet.Cells["A2"].PutValue("Electronics");
        worksheet.Cells["B2"].PutValue("Laptop");
        worksheet.Cells["C2"].PutValue(1200);

        worksheet.Cells["A3"].PutValue("Electronics");
        worksheet.Cells["B3"].PutValue("Phone");
        worksheet.Cells["C3"].PutValue(800);

        worksheet.Cells["A4"].PutValue("Furniture");
        worksheet.Cells["B4"].PutValue("Chair");
        worksheet.Cells["C4"].PutValue(150);

        worksheet.Cells["A5"].PutValue("Furniture");
        worksheet.Cells["B5"].PutValue("Table");
        worksheet.Cells["C5"].PutValue(250);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("A1:C5", "E3", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

        // Add "Category" as a row field and "Sales" as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Retrieve the row field representing the product categories
        PivotField categoryField = pivotTable.RowFields[0];

        // Apply descending alphabetical sort on the category field
        // -1 indicates sorting by the field's own labels (i.e., alphabetical order)
        categoryField.SortBy(SortOrder.Descending, -1);

        // Ensure auto‑sort settings are consistent with the desired order
        categoryField.IsAutoSort = true;
        categoryField.IsAscendSort = false;   // false = descending
        categoryField.AutoSortField = -1;     // sort by the field itself

        // Refresh and calculate the pivot table to apply sorting
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the sorted pivot table
        workbook.Save("PivotCategoryDescAlphabetical.xlsx");
    }
}
