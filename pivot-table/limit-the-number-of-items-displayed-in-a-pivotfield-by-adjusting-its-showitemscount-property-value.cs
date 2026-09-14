// Title: How to display only the top N row items in an Aspose.Cells PivotTable using C# AutoShow settings
// AI Prompts: Generate C# code that creates a workbook, adds a PivotTable, and configures the row field to show only the top 3 categories by enabling IsAutoShow and setting AutoShowCount with Aspose.Cells. | Write a C# snippet that limits a PivotTable row field to a specific number of items using the AutoShow properties of Aspose.Cells PivotField.
// Common Searches: C# Aspose.Cells limit pivot table row field to top 5 items | How to use AutoShowCount in Aspose.Cells PivotField | Show only a certain number of categories in a PivotTable with Aspose.Cells .NET | Aspose.Cells hide low‑value rows in pivot table programmatically | Set row field item count in Aspose.Cells pivot table example
// Tags: Aspose.Cells pivot table AutoShow | C# limit pivot row field items | Aspose.Cells set AutoShowCount | top N items pivot field .NET | limit displayed pivot categories

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a new workbook, populates it with sample category and sales data, adds a PivotTable, and then limits the row field to the top three categories by enabling AutoShow, setting AutoShowCount to 3, and configuring the sort order before refreshing and saving the workbook as LimitedPivotItems.xlsx.
class LimitPivotFieldItems
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data for the pivot table
        cells["A1"].PutValue("Category");
        cells["B1"].PutValue("Sales");
        string[] categories = { "A", "B", "C", "D", "E" };
        int[] sales = { 100, 200, 150, 120, 180 };
        for (int i = 0; i < categories.Length; i++)
        {
            cells[i + 1, 0].PutValue(categories[i]);
            cells[i + 1, 1].PutValue(sales[i]);
        }

        // Add a pivot table to the worksheet
        int pivotIndex = sheet.PivotTables.Add("A1:B6", "D3", "MyPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add row field (Category) and data field (Sales)
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Limit the number of displayed items in the row field to the top 3 categories
        PivotField rowField = pivotTable.RowFields[0];
        rowField.IsAutoShow = true;          // Enable AutoShow
        rowField.AutoShowCount = 3;          // Number of items to display
        rowField.IsAscendShow = false;       // Show top items (descending order)
        rowField.AutoShowField = -1;         // Use the field itself for ranking

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("LimitedPivotItems.xlsx");
    }
}
