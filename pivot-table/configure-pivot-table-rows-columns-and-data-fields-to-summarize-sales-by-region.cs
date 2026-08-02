using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

class PivotTableRegionSummary
{
    static void Main()
    {
        // Create a new workbook and a worksheet for source data
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Add header row
        dataSheet.Cells["A1"].PutValue("Region");
        dataSheet.Cells["B1"].PutValue("Product");
        dataSheet.Cells["C1"].PutValue("Sales");

        // Populate sample data
        string[] regions = { "North", "South", "East", "West" };
        string[] products = { "Bike", "Car", "Truck" };
        int currentRow = 2;
        Random rnd = new Random();

        foreach (string region in regions)
        {
            foreach (string product in products)
            {
                dataSheet.Cells[currentRow, 0].PutValue(region);
                dataSheet.Cells[currentRow, 1].PutValue(product);
                dataSheet.Cells[currentRow, 2].PutValue(rnd.Next(500, 5000));
                currentRow++;
            }
        }

        // Add a worksheet for the pivot table
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");

        // Create the pivot table using the data range
        string sourceRange = $"=Data!A1:C{currentRow - 1}";
        int pivotIndex = pivotSheet.PivotTables.Add(sourceRange, "A3", "SalesByRegion");
        PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

        // Configure rows, columns, and data fields
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");      // Row field
        pivotTable.AddFieldToArea(PivotFieldType.Column, "Product"); // Column field
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");     // Data field

        // Ensure the data field uses Sum aggregation (default, but set explicitly)
        pivotTable.DataFields[0].Function = ConsolidationFunction.Sum;

        // Optional: display the pivot table in tabular form
        pivotTable.ShowInTabularForm();

        // Refresh and calculate the pivot table data
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook
        workbook.Save("SalesByRegionPivot.xlsx");
    }
}