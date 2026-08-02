using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate sample data
        cells["A1"].Value = "Region";
        cells["B1"].Value = "Category";
        cells["C1"].Value = "Sales";

        string[] regions = { "North", "South", "East", "West" };
        string[] categories = { "Food", "Clothing", "Electronics" };
        int currentRow = 2;

        foreach (string region in regions)
        {
            foreach (string category in categories)
            {
                cells[currentRow, 0].Value = region;
                cells[currentRow, 1].Value = category;
                cells[currentRow, 2].Value = (currentRow - 1) * 100; // sample sales value
                currentRow++;
            }
        }

        // Add a pivot table based on the data range
        PivotTableCollection pivotTables = worksheet.PivotTables;
        int pivotIndex = pivotTables.Add($"A1:C{currentRow - 1}", "E5", "SalesPivot");
        PivotTable pivotTable = pivotTables[pivotIndex];

        // Add page fields: first Region, then Category (default filter priority)
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Page, "Category");

        // Add additional fields for rows and data
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

        // Reorder page fields so that Category becomes the primary filter
        // PageFields collection is zero‑based; move field at index 0 (Region) to index 1
        pivotTable.PageFields.Move(0, 1);

        // Refresh and calculate the pivot table to apply changes
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the reordered page fields
        workbook.Save("ReorderedPageFields.xlsx");
    }
}