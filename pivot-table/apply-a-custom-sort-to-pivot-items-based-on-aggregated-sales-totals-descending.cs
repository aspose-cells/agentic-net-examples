// Title: How to sort pivot table rows by total Sales in descending order using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells to set the RowFields[0] sort order to descending based on the summed Sales data field. | Create a pivot table from a data range and apply a descending sort on the Region row items using the aggregated Sales values. | Refresh and calculate the pivot table after configuring a descending sort on the row field with Aspose.Cells in C#.
// Common Searches: aspnet c# apply descending sort to pivot table row field based on sum of sales using aspose.cells | example of custom sorting pivot rows by aggregated data in Aspose.Cells .NET | how to sort pivot table rows by total sales descending in C# Aspose.Cells
// Tags: aspocells pivot table descending order | c# set pivot row field order by aggregated data | aspocells sort pivot rows by sum of sales | refresh calculate pivot after sort aspocells | pivot table row field descending order using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, fills it with Region and Sales data, adds a pivot table with Region as a row field and Sales as a data field, applies a descending sort to the Region row items based on the summed Sales values, refreshes and calculates the pivot, and saves the workbook as CustomSortedPivot.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data: Region and Sales
        cells["A1"].PutValue("Region");
        cells["B1"].PutValue("Sales");
        cells["A2"].PutValue("North");
        cells["B2"].PutValue(1200);
        cells["A3"].PutValue("South");
        cells["B3"].PutValue(800);
        cells["A4"].PutValue("East");
        cells["B4"].PutValue(1500);
        cells["A5"].PutValue("West");
        cells["B5"].PutValue(600);

        // Add a pivot table covering the data range
        int pivotIndex = sheet.PivotTables.Add("A1:B5", "D3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add Region as a row field and Sales as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Apply custom sort: descending order based on aggregated Sales totals
        // fieldSortedBy = -1 indicates sorting by the data labels of this field (i.e., the sum of Sales)
        pivotTable.RowFields[0].SortBy(SortOrder.Descending, -1);

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("CustomSortedPivot.xlsx");
    }
}
